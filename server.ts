import 'dotenv/config';
import express from 'express';
import http from 'http';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { GoogleGenAI, Type, ThinkingLevel } from '@google/genai';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    },
  },
});

const getResponseSchema = (schemaType?: string) => {
  const slideItemSchema = {
    type: Type.OBJECT,
    properties: {
      type: {
        type: Type.STRING,
        description: 'Slide type: open, poll, cloud, quiz, scale, draw, or info',
      },
      q: {
        type: Type.STRING,
        description: 'Question or slide title in Vietnamese',
      },
      b: {
        type: Type.STRING,
        description: 'Body text (only for info slides)',
      },
      opts: {
        type: Type.ARRAY,
        items: { type: Type.STRING },
        description: 'Answer options (3-4 items, only for poll or quiz slides)',
      },
      a: {
        type: Type.INTEGER,
        description: 'Zero-based index of correct option (only for quiz slides)',
      },
      ai: {
        type: Type.BOOLEAN,
        description: 'True if AI live feedback is recommended (for open slides)',
      },
    },
    required: ['type', 'q'],
  };

  if (schemaType === 'autoInsert') {
    return {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          after: {
            type: Type.INTEGER,
            description: 'Zero-based index of the slide to insert AFTER (-1 means at the very beginning)',
          },
          type: {
            type: Type.STRING,
            description: 'One of: cloud, open, poll, quiz, scale, draw',
          },
          q: {
            type: Type.STRING,
            description: 'Interactive question in Vietnamese',
          },
          opts: {
            type: Type.ARRAY,
            items: { type: Type.STRING },
            description: 'Options array (3-4 choices, required for poll and quiz)',
          },
          a: {
            type: Type.INTEGER,
            description: 'Zero-based index of correct answer (required for quiz)',
          },
          ai: {
            type: Type.BOOLEAN,
            description: 'True for open questions needing AI feedback',
          },
        },
        required: ['after', 'type', 'q'],
      },
    };
  }

  if (schemaType === 'lesson') {
    return {
      type: Type.OBJECT,
      properties: {
        title: { type: Type.STRING },
        slides: {
          type: Type.ARRAY,
          items: slideItemSchema,
        },
      },
      required: ['title', 'slides'],
    };
  }

  if (schemaType === 'slidesArray') {
    return {
      type: Type.ARRAY,
      items: slideItemSchema,
    };
  }

  return undefined;
};

async function startServer() {
  const app = express();
  const httpServer = http.createServer(app);
  const PORT = 3000;

  app.use(express.json({ limit: '100mb' }));

  const dataDir = path.resolve(__dirname, '.data');
  const storeFile = path.join(dataDir, 'dashboard-store.json');

  interface OwnerStoreData {
    lessons: Record<string, { id: string; updatedAt: string; serialized: string }>;
    reports: Record<string, any>;
    deletedLessonIds: string[];
    deletedReportIds: string[];
  }

  const readStore = (): Record<string, OwnerStoreData> => {
    try {
      if (fs.existsSync(storeFile)) {
        const raw = fs.readFileSync(storeFile, 'utf-8');
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') return parsed;
      }
    } catch {
      // ignore read errors
    }
    return {};
  };

  const writeStore = (store: Record<string, OwnerStoreData>) => {
    try {
      if (!fs.existsSync(dataDir)) {
        fs.mkdirSync(dataDir, { recursive: true });
      }
      fs.writeFileSync(storeFile, JSON.stringify(store), 'utf-8');
    } catch {
      // ignore write errors
    }
  };

  const getOwnerBucket = (store: Record<string, OwnerStoreData>, ownerId: string): OwnerStoreData => {
    const key = (ownerId || 'sync_quangminh1485').trim() || 'sync_quangminh1485';
    if (!store[key]) {
      store[key] = {
        lessons: {},
        reports: {},
        deletedLessonIds: [],
        deletedReportIds: [],
      };
    }
    return store[key];
  };

  app.get('/api/storage/all', (req, res) => {
    const ownerId = String(req.query.ownerId || 'sync_quangminh1485');
    const store = readStore();
    const bucket = getOwnerBucket(store, ownerId);
    res.json({
      lessons: Object.values(bucket.lessons || {}),
      reports: Object.values(bucket.reports || {}),
      deletedLessonIds: bucket.deletedLessonIds || [],
      deletedReportIds: bucket.deletedReportIds || [],
    });
  });

  app.post('/api/storage/lesson', (req, res) => {
    try {
      const { ownerId, id, updatedAt, serialized } = req.body || {};
      if (!id || typeof serialized !== 'string') {
        res.status(400).json({ error: 'Thiếu dữ liệu bài giảng.' });
        return;
      }
      const store = readStore();
      const bucket = getOwnerBucket(store, String(ownerId || 'sync_quangminh1485'));
      bucket.lessons[id] = {
        id: String(id),
        updatedAt: String(updatedAt || new Date().toISOString()),
        serialized,
      };
      bucket.deletedLessonIds = (bucket.deletedLessonIds || []).filter(x => x !== id);
      writeStore(store);
      res.json({ ok: true, id });
    } catch (e: any) {
      res.status(500).json({ error: e?.message || 'Lỗi lưu bài giảng.' });
    }
  });

  app.delete('/api/storage/lesson/:id', (req, res) => {
    try {
      const id = String(req.params.id || '');
      const ownerId = String(req.query.ownerId || 'sync_quangminh1485');
      const store = readStore();
      const bucket = getOwnerBucket(store, ownerId);
      delete bucket.lessons[id];
      if (!bucket.deletedLessonIds.includes(id)) {
        bucket.deletedLessonIds.push(id);
      }
      writeStore(store);
      res.json({ ok: true });
    } catch (e: any) {
      res.status(500).json({ error: e?.message || 'Lỗi xóa bài giảng.' });
    }
  });

  app.post('/api/storage/report', (req, res) => {
    try {
      const { ownerId, report } = req.body || {};
      if (!report || !report.id) {
        res.status(400).json({ error: 'Thiếu dữ liệu thống kê.' });
        return;
      }
      const store = readStore();
      const bucket = getOwnerBucket(store, String(ownerId || 'sync_quangminh1485'));
      bucket.reports[report.id] = report;
      bucket.deletedReportIds = (bucket.deletedReportIds || []).filter(x => x !== report.id);
      writeStore(store);
      res.json({ ok: true, id: report.id });
    } catch (e: any) {
      res.status(500).json({ error: e?.message || 'Lỗi lưu thống kê.' });
    }
  });

  app.delete('/api/storage/report/:id', (req, res) => {
    try {
      const id = String(req.params.id || '');
      const ownerId = String(req.query.ownerId || 'sync_quangminh1485');
      const store = readStore();
      const bucket = getOwnerBucket(store, ownerId);
      delete bucket.reports[id];
      if (!bucket.deletedReportIds.includes(id)) {
        bucket.deletedReportIds.push(id);
      }
      writeStore(store);
      res.json({ ok: true });
    } catch (e: any) {
      res.status(500).json({ error: e?.message || 'Lỗi xóa thống kê.' });
    }
  });

  app.post('/api/gemini/generate', async (req, res) => {
    try {
      if (!process.env.GEMINI_API_KEY) {
        res.status(500).json({ error: 'Chưa cấu hình GEMINI_API_KEY trên máy chủ.' });
        return;
      }

      const { prompt, json, schemaType, imgBase64, imgMimeType } = req.body || {};
      if (!prompt || typeof prompt !== 'string') {
        res.status(400).json({ error: 'Thiếu nội dung yêu cầu (prompt).' });
        return;
      }

      const parts: any[] = [{ text: prompt }];
      if (imgBase64 && imgMimeType) {
        parts.push({
          inlineData: {
            mimeType: imgMimeType,
            data: imgBase64,
          },
        });
      }

      const config: any = {
        thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
      };

      if (json) {
        config.responseMimeType = 'application/json';
        const schema = getResponseSchema(schemaType);
        if (schema) {
          config.responseSchema = schema;
        }
      }

      const runGenerate = async (modelName: string) => {
        const modelConfig: any = {
          ...config,
          thinkingConfig: {
            thinkingLevel: modelName === 'gemini-3.1-flash-lite' ? ThinkingLevel.MINIMAL : ThinkingLevel.LOW,
          },
        };
        return ai.models.generateContent({
          model: modelName,
          contents: [{ role: 'user', parts }],
          config: modelConfig,
        });
      };

      const extractErrorInfo = (err: any): { status: number; message: string } => {
        let status = typeof err?.status === 'number' ? err.status : 500;
        let rawMsg = String(err?.message || 'Lỗi kết nối Gemini AI.');
        try {
          const parsed = JSON.parse(rawMsg);
          if (parsed?.error) {
            if (typeof parsed.error.code === 'number') status = parsed.error.code;
            if (parsed.error.message) rawMsg = parsed.error.message;
          }
        } catch {
          // keep rawMsg
        }
        return { status, message: rawMsg };
      };

      const modelsToTry = ['gemini-4-hydrogen', 'gemini-3.1-flash-lite', 'gemini-2.5-flash', 'gemini-2.0-flash'];
      let response: any = null;
      let lastErr: any = null;

      for (let attempt = 0; attempt < modelsToTry.length; attempt++) {
        try {
          response = await runGenerate(modelsToTry[attempt]);
          lastErr = null;
          break;
        } catch (err: any) {
          lastErr = err;
          const { status } = extractErrorInfo(err);
          const isRetryable = status === 404 || status === 400 || status === 503 || status === 429 || status === 500 || status === 502 || status === 504;
          if (!isRetryable || attempt === modelsToTry.length - 1) {
            break;
          }
          await new Promise(r => setTimeout(r, 800 * (attempt + 1)));
        }
      }

      if (lastErr || !response) {
        const { status, message } = extractErrorInfo(lastErr);
        res.status(status).json({ error: message });
        return;
      }

      const text = (response.text ?? '').trim();
      if (!text) {
        res.status(502).json({ error: 'AI không trả về nội dung.' });
        return;
      }

      if (json) {
        const cleaned = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
        try {
          const parsed = JSON.parse(cleaned);
          res.json({ text: cleaned, data: parsed });
          return;
        } catch {
          // Try extracting first JSON array or object
          const match = cleaned.match(/(\[[\s\S]*\]|\{[\s\S]*\})/);
          if (match) {
            const parsed = JSON.parse(match[1]);
            res.json({ text: match[1], data: parsed });
            return;
          }
          res.status(502).json({ error: 'Dữ liệu JSON từ AI không hợp lệ.' });
          return;
        }
      }

      res.json({ text });
    } catch (err: any) {
      const status = err?.status && typeof err.status === 'number' ? err.status : 500;
      const message = err?.message || 'Lỗi kết nối Gemini AI.';
      res.status(status).json({ error: message });
    }
  });

  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        ws: { server: httpServer },
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      const indexFile = path.join(distDir, 'index.html');
      if (fs.existsSync(indexFile)) {
        res.sendFile(indexFile);
      } else {
        res.status(404).send('Not found');
      }
    });
  }

  httpServer.listen(PORT, '0.0.0.0', () => {
    console.log(`CuriVN server running on http://localhost:${PORT}`);
  });
}

startServer();
