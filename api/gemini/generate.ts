import { GoogleGenAI, Type } from '@google/genai';

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
            enum: ['cloud', 'open', 'poll', 'quiz', 'scale', 'draw'],
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

export default async function handler(req: any, res: any) {
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method Not Allowed' });
    return;
  }

  try {
    const apiKey =
      process.env.GEMINI_API_KEY ||
      process.env.VITE_GEMINI_API_KEY ||
      process.env.GOOGLE_API_KEY;

    if (!apiKey) {
      res.status(500).json({
        error:
          'Chưa cấu hình GEMINI_API_KEY trên máy chủ (Vercel Environment Variables -> thêm GEMINI_API_KEY).',
      });
      return;
    }

    const ai = new GoogleGenAI({ apiKey });

    const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body || {};
    const { prompt, json, schemaType, imgBase64, imgMimeType } = body;

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

    const config: any = {};
    if (json) {
      config.responseMimeType = 'application/json';
      const schema = getResponseSchema(schemaType);
      if (schema) {
        config.responseSchema = schema;
      }
    }

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

    const modelsToTry = [
      'gemini-2.5-flash',
      'gemini-4-hydrogen',
      'gemini-3.1-flash-lite',
      'gemini-2.0-flash',
    ];
    let response: any = null;
    let lastErr: any = null;

    for (let attempt = 0; attempt < modelsToTry.length; attempt++) {
      try {
        response = await ai.models.generateContent({
          model: modelsToTry[attempt],
          contents: [{ role: 'user', parts }],
          config,
        });
        lastErr = null;
        break;
      } catch (err: any) {
        lastErr = err;
        const { status } = extractErrorInfo(err);
        const isRetryable =
          status === 404 ||
          status === 400 ||
          status === 503 ||
          status === 429 ||
          status === 500 ||
          status === 502 ||
          status === 504;
        if (!isRetryable || attempt === modelsToTry.length - 1) {
          break;
        }
        await new Promise(r => setTimeout(r, 500 * (attempt + 1)));
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
      const cleaned = text
        .replace(/^```(?:json)?\s*/i, '')
        .replace(/\s*```$/i, '')
        .trim();
      try {
        const parsed = JSON.parse(cleaned);
        res.status(200).json({ text: cleaned, data: parsed });
        return;
      } catch {
        const match = cleaned.match(/(\[[\s\S]*\]|\{[\s\S]*\})/);
        if (match) {
          const parsed = JSON.parse(match[1]);
          res.status(200).json({ text: match[1], data: parsed });
          return;
        }
        res.status(502).json({ error: 'Dữ liệu JSON từ AI không hợp lệ.' });
        return;
      }
    }

    res.status(200).json({ text });
  } catch (err: any) {
    const status = err?.status && typeof err.status === 'number' ? err.status : 500;
    const message = err?.message || 'Lỗi kết nối Gemini AI.';
    res.status(status).json({ error: message });
  }
}
