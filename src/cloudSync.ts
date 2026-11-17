import {
  auth,
  db,
  googleProvider,
  signInWithPopup,
  fbSignOut,
  onAuthStateChanged,
  doc,
  getDoc,
  getDocs,
  setDoc,
  updateDoc,
  deleteDoc,
  collection,
  query,
  where,
  serverTimestamp,
  handleFirestoreError,
  OperationType,
  type User,
} from './firebase';

const CHUNK_MAX_CHARS = 600000; // Strictly <= 750000 in firebase-blueprint.json & firestore.rules
const REPORT_MAX_CHARS = 800000; // Strictly <= 850000 in firebase-blueprint.json & firestore.rules
const SYNC_KEY_STORAGE = 'curivn_cloud_sync_key';
const DEFAULT_TEACHER_SYNC_KEY = 'quangminh1485';

function sanitizeId(rawId: string): string {
  const clean = (rawId || '')
    .replace(/[^a-zA-Z0-9_-]/g, '_')
    .slice(0, 120);
  return clean || 'doc_' + Math.random().toString(36).slice(2, 10);
}

export function normalizeSyncKey(rawInput: string): string {
  const trimmed = (rawInput || '').trim().toLowerCase();
  if (!trimmed) return '';
  const cleaned = trimmed
    .replace(/@gmail\.com$/i, '')
    .replace(/[^a-z0-9_-]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 80);
  if (cleaned.length < 3) return '';
  return cleaned;
}

export function getSavedSyncKey(): string {
  try {
    const stored = localStorage.getItem(SYNC_KEY_STORAGE);
    if (stored === null) {
      localStorage.setItem(SYNC_KEY_STORAGE, DEFAULT_TEACHER_SYNC_KEY);
      return DEFAULT_TEACHER_SYNC_KEY;
    }
    return stored || DEFAULT_TEACHER_SYNC_KEY;
  } catch {
    return DEFAULT_TEACHER_SYNC_KEY;
  }
}

export function setSavedSyncKey(rawKey: string): string {
  const norm = normalizeSyncKey(rawKey) || DEFAULT_TEACHER_SYNC_KEY;
  try {
    localStorage.setItem(SYNC_KEY_STORAGE, norm);
  } catch {
    // ignore storage errors
  }
  return norm;
}

export function getActiveCloudOwnerId(): string | null {
  const savedKey = getSavedSyncKey();
  if (savedKey && savedKey.length >= 3) {
    return `sync_${savedKey}`;
  }
  return `sync_${DEFAULT_TEACHER_SYNC_KEY}`;
}

export function getActiveCloudLabel(): string | null {
  return getSavedSyncKey() || DEFAULT_TEACHER_SYNC_KEY;
}

async function blobToBase64(blob: Blob): Promise<string> {
  const ab = await blob.arrayBuffer();
  const bytes = new Uint8Array(ab);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function base64ToBlob(b64: string, mimeType: string): Blob {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes.buffer as ArrayBuffer], { type: mimeType || 'application/octet-stream' });
}

export async function serializeLessonWithMedia(lesson: any): Promise<string> {
  const serializedM: Record<string, { _b64: string; type: string }> = {};
  if (lesson.M) {
    for (const [k, val] of Object.entries(lesson.M)) {
      if (val instanceof Blob) {
        try {
          const b64 = await blobToBase64(val);
          serializedM[k] = { _b64: b64, type: val.type || 'application/octet-stream' };
        } catch {
          // skip unreadable blob
        }
      } else if (val && (val as any)._ab) {
        try {
          const blob = new Blob([(val as any)._ab], { type: (val as any).type || 'application/octet-stream' });
          const b64 = await blobToBase64(blob);
          serializedM[k] = { _b64: b64, type: blob.type };
        } catch {
          // skip
        }
      } else if (val && (val as any)._b64) {
        serializedM[k] = val as any;
      }
    }
  }
  return JSON.stringify({
    ...lesson,
    M: serializedM,
  });
}

export function deserializeLessonWithMedia(jsonStr: string): any {
  const parsed = JSON.parse(jsonStr);
  const nextM: Record<string, Blob> = {};
  if (parsed?.M && typeof parsed.M === 'object') {
    for (const [k, val] of Object.entries(parsed.M)) {
      if (val && typeof (val as any)._b64 === 'string') {
        try {
          nextM[k] = base64ToBlob((val as any)._b64, (val as any).type || 'image/png');
        } catch {
          // ignore
        }
      }
    }
  }
  return {
    ...parsed,
    M: nextM,
  };
}

export async function cloudSaveLesson(lesson: any): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return;

  const rawId = sanitizeId(lesson.id || Math.random().toString(36).slice(2, 10));
  const docId = sanitizeId(`${ownerId}_${rawId}`);
  const fullJson = await serializeLessonWithMedia({ ...lesson, id: rawId });
  const chunks: string[] = [];
  for (let i = 0; i < fullJson.length; i += CHUNK_MAX_CHARS) {
    chunks.push(fullJson.slice(i, i + CHUNK_MAX_CHARS));
  }
  if (chunks.length === 0) {
    chunks.push('{}');
  }
  const safeChunks = chunks.slice(0, 200);
  const updatedAtMs = lesson.updatedAt ? Date.parse(lesson.updatedAt) || Date.now() : Date.now();

  const lessonPath = `lessons/${docId}`;
  const lessonRef = doc(db, 'lessons', docId);

  let existingSnap;
  try {
    existingSnap = await getDoc(lessonRef);
  } catch {
    existingSnap = null;
  }

  const metaFields = {
    title: String(lesson.title || 'Chưa đặt tên').slice(0, 500),
    className: String(lesson.className || '').slice(0, 200),
    customCode: String(lesson.customCode || '').slice(0, 32),
    slideCount: Array.isArray(lesson.slides) ? Math.min(lesson.slides.length, 2000) : 0,
    chunkCount: safeChunks.length,
    updatedAtMs: Math.max(0, Math.floor(updatedAtMs)),
  };

  try {
    if (existingSnap && existingSnap.exists()) {
      await updateDoc(lessonRef, {
        ...metaFields,
        updatedAt: serverTimestamp(),
      });
    } else {
      await setDoc(lessonRef, {
        id: docId,
        ownerId,
        ...metaFields,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      });
    }
  } catch (error) {
    handleFirestoreError(error, existingSnap?.exists() ? OperationType.UPDATE : OperationType.CREATE, lessonPath);
  }

  for (let idx = 0; idx < safeChunks.length; idx++) {
    const chunkId = `chunk_${idx}`;
    const chunkPath = `lessons/${docId}/chunks/${chunkId}`;
    const chunkRef = doc(db, 'lessons', docId, 'chunks', chunkId);
    let existingChunk;
    try {
      existingChunk = await getDoc(chunkRef);
    } catch {
      existingChunk = null;
    }

    try {
      if (existingChunk && existingChunk.exists()) {
        await updateDoc(chunkRef, {
          idx,
          data: safeChunks[idx],
          updatedAt: serverTimestamp(),
        });
      } else {
        await setDoc(chunkRef, {
          ownerId,
          lessonId: docId,
          idx,
          data: safeChunks[idx],
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
        });
      }
    } catch (error) {
      handleFirestoreError(error, existingChunk?.exists() ? OperationType.UPDATE : OperationType.CREATE, chunkPath);
    }
  }
}

export async function cloudDeleteLesson(rawLessonId: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return;
  const rawId = sanitizeId(rawLessonId);
  const docIdsToTry = [sanitizeId(`${ownerId}_${rawId}`), rawId];

  for (const docId of docIdsToTry) {
    const chunksCol = collection(db, 'lessons', docId, 'chunks');
    try {
      const q = query(chunksCol, where('ownerId', '==', ownerId));
      const snap = await getDocs(q);
      for (const d of snap.docs) {
        await deleteDoc(d.ref);
      }
      await deleteDoc(doc(db, 'lessons', docId));
    } catch {
      // ignore if doc does not exist
    }
  }
}

export async function cloudSaveReport(report: any): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId || !report?.id) return;

  const rawId = sanitizeId(report.id);
  const docId = sanitizeId(`${ownerId}_${rawId}`);
  let payloadJson = JSON.stringify(report);
  if (payloadJson.length > REPORT_MAX_CHARS) {
    const trimmed = {
      ...report,
      slides: (report.slides || []).map((sl: any) => ({
        ...sl,
        responses: (sl.responses || []).map((res: any) => ({
          ...res,
          value:
            typeof res.value === 'string' && res.value.length > 4000
              ? res.value.startsWith('data:image/')
                ? '[Bản vẽ đã lưu]'
                : res.value.slice(0, 4000)
              : res.value,
        })),
      })),
    };
    payloadJson = JSON.stringify(trimmed);
  }
  if (payloadJson.length > REPORT_MAX_CHARS) {
    payloadJson = payloadJson.slice(0, REPORT_MAX_CHARS);
  }

  const updatedAtMs = report.updatedAt ? Date.parse(report.updatedAt) || Date.now() : Date.now();
  const reportPath = `reports/${docId}`;
  const reportRef = doc(db, 'reports', docId);

  let existingSnap;
  try {
    existingSnap = await getDoc(reportRef);
  } catch {
    existingSnap = null;
  }

  const fields = {
    lessonTitle: String(report.lessonTitle || 'Chưa đặt tên').slice(0, 500),
    className: String(report.className || 'Chưa phân lớp').slice(0, 200),
    roomCode: String(report.roomCode || '').slice(0, 32),
    updatedAtMs: Math.max(0, Math.floor(updatedAtMs)),
    payloadJson,
  };

  try {
    if (existingSnap && existingSnap.exists()) {
      await updateDoc(reportRef, {
        ...fields,
        updatedAt: serverTimestamp(),
      });
    } else {
      await setDoc(reportRef, {
        id: docId,
        ownerId,
        ...fields,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      });
    }
  } catch (error) {
    handleFirestoreError(error, existingSnap?.exists() ? OperationType.UPDATE : OperationType.CREATE, reportPath);
  }
}

export async function cloudDeleteReport(rawReportId: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return;
  const rawId = sanitizeId(rawReportId);
  const docIdsToTry = [sanitizeId(`${ownerId}_${rawId}`), rawId];
  for (const docId of docIdsToTry) {
    try {
      await deleteDoc(doc(db, 'reports', docId));
    } catch {
      // ignore
    }
  }
}

export async function cloudSyncAll(
  localLessons: any[],
  localReports: any[],
  saveLocalLesson: (l: any) => Promise<void>,
  saveLocalReport: (r: any) => Promise<void>
): Promise<{ lessons: any[]; reports: any[] }> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) {
    return { lessons: localLessons, reports: localReports };
  }

  const stripPrefix = (docId: string) => {
    const prefix = `${ownerId}_`;
    return docId.startsWith(prefix) ? docId.slice(prefix.length) : docId;
  };

  // 1. Fetch remote lessons metadata
  const lessonsMap = new Map<string, any>();
  for (const l of localLessons) {
    if (l?.id) {
      const cleanLocalId = sanitizeId(l.id);
      lessonsMap.set(cleanLocalId, { ...l, id: cleanLocalId });
    }
  }

  let remoteLessonDocs: any[] = [];
  try {
    const qLessons = query(collection(db, 'lessons'), where('ownerId', '==', ownerId));
    const snap = await getDocs(qLessons);
    remoteLessonDocs = snap.docs.map(d => d.data());
  } catch (error) {
    // Fail gracefully without throwing errors to keep operations lightning fast
    remoteLessonDocs = [];
  }

  const remoteLessonIds = new Set<string>();
  for (const meta of remoteLessonDocs) {
    const docId = meta.id;
    const lid = stripPrefix(docId);
    remoteLessonIds.add(lid);
    const local = lessonsMap.get(lid);
    const localMs = local?.updatedAt ? Date.parse(local.updatedAt) || 0 : 0;
    const remoteMs = typeof meta.updatedAtMs === 'number' ? meta.updatedAtMs : 0;

    if (!local || remoteMs > localMs + 1000) {
      try {
        const qChunks = query(
          collection(db, 'lessons', docId, 'chunks'),
          where('ownerId', '==', ownerId)
        );
        const chunkSnap = await getDocs(qChunks);
        const sortedChunks = chunkSnap.docs
          .map(d => d.data())
          .sort((a, b) => (a.idx || 0) - (b.idx || 0));
        const expectedCount = typeof meta.chunkCount === 'number' ? meta.chunkCount : sortedChunks.length;
        const joined = sortedChunks
          .slice(0, expectedCount)
          .map(c => c.data || '')
          .join('');
        if (joined) {
          const hydrated = deserializeLessonWithMedia(joined);
          hydrated.id = lid;
          lessonsMap.set(lid, hydrated);
          await saveLocalLesson(hydrated);
        }
      } catch {
        // keep local if chunk assembly fails
      }
    } else if (local && localMs > remoteMs + 1000) {
      try {
        await cloudSaveLesson(local);
      } catch {}
    }
  }

  for (const [lid, local] of lessonsMap.entries()) {
    if (!remoteLessonIds.has(lid)) {
      try {
        await cloudSaveLesson(local);
      } catch {}
    }
  }

  // 2. Fetch remote session reports
  const reportsMap = new Map<string, any>();
  for (const r of localReports) {
    if (r?.id) {
      const cleanId = sanitizeId(r.id);
      reportsMap.set(cleanId, { ...r, id: cleanId });
    }
  }

  let remoteReportDocs: any[] = [];
  try {
    const qReports = query(collection(db, 'reports'), where('ownerId', '==', ownerId));
    const snap = await getDocs(qReports);
    remoteReportDocs = snap.docs.map(d => d.data());
  } catch (error) {
    remoteReportDocs = [];
  }

  const remoteReportIds = new Set<string>();
  for (const rdoc of remoteReportDocs) {
    const docId = rdoc.id;
    const rid = stripPrefix(docId);
    remoteReportIds.add(rid);
    const local = reportsMap.get(rid);
    const localMs = local?.updatedAt ? Date.parse(local.updatedAt) || 0 : 0;
    const remoteMs = typeof rdoc.updatedAtMs === 'number' ? rdoc.updatedAtMs : 0;

    if (!local || remoteMs > localMs + 1000) {
      try {
        const parsed = JSON.parse(rdoc.payloadJson);
        if (parsed && parsed.id) {
          parsed.id = rid;
          reportsMap.set(rid, parsed);
          await saveLocalReport(parsed);
        }
      } catch {
        // skip invalid JSON
      }
    } else if (local && localMs > remoteMs + 1000) {
      try {
        await cloudSaveReport(local);
      } catch {}
    }
  }

  for (const [rid, local] of reportsMap.entries()) {
    if (!remoteReportIds.has(rid)) {
      try {
        await cloudSaveReport(local);
      } catch {}
    }
  }

  const mergedLessons = Array.from(lessonsMap.values()).sort((a, b) =>
    (b.updatedAt || '').localeCompare(a.updatedAt || '')
  );
  const mergedReports = Array.from(reportsMap.values()).sort((a, b) =>
    (b.updatedAt || b.id).localeCompare(a.updatedAt || a.id)
  );

  return { lessons: mergedLessons, reports: mergedReports };
}

export async function exportBackupJson(lessons: any[], reports: any[]): Promise<void> {
  const serializedLessons: string[] = [];
  for (const l of lessons) {
    serializedLessons.push(await serializeLessonWithMedia(l));
  }
  const payload = JSON.stringify({
    version: 2,
    exportedAt: new Date().toISOString(),
    lessons: serializedLessons,
    reports,
  });
  const blob = new Blob([payload], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `lop-hoc-truc-tuyen-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

export async function importBackupJson(file: File): Promise<{ lessons: any[]; reports: any[] }> {
  const text = await file.text();
  const parsed = JSON.parse(text);
  const lessons: any[] = [];
  if (Array.isArray(parsed?.lessons)) {
    for (const item of parsed.lessons) {
      if (typeof item === 'string') {
        lessons.push(deserializeLessonWithMedia(item));
      } else if (item && typeof item === 'object') {
        lessons.push(item);
      }
    }
  }
  const reports: any[] = Array.isArray(parsed?.reports) ? parsed.reports : [];
  return { lessons, reports };
}

export async function cloudFetchAllLessons(): Promise<any[]> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return [];
  const stripPrefix = (docId: string) => {
    const prefix = `${ownerId}_`;
    return docId.startsWith(prefix) ? docId.slice(prefix.length) : docId;
  };
  try {
    const qLessons = query(collection(db, 'lessons'), where('ownerId', '==', ownerId));
    const snap = await getDocs(qLessons);
    const remoteLessons: any[] = [];
    for (const dSnap of snap.docs) {
      const meta = dSnap.data();
      const docId = meta.id;
      const lid = stripPrefix(docId);
      try {
        const qChunks = query(collection(db, 'lessons', docId, 'chunks'), where('ownerId', '==', ownerId));
        const chunkSnap = await getDocs(qChunks);
        const sortedChunks = chunkSnap.docs
          .map(d => d.data())
          .sort((a, b) => (a.idx || 0) - (b.idx || 0));
        const expectedCount = typeof meta.chunkCount === 'number' ? meta.chunkCount : sortedChunks.length;
        const joined = sortedChunks.slice(0, expectedCount).map(c => c.data || '').join('');
        if (joined) {
          const hydrated = deserializeLessonWithMedia(joined);
          hydrated.id = lid;
          remoteLessons.push(hydrated);
        }
      } catch {}
    }
    return remoteLessons.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
  } catch {
    return [];
  }
}

export async function cloudFetchAllReports(): Promise<any[]> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return [];
  const stripPrefix = (docId: string) => {
    const prefix = `${ownerId}_`;
    return docId.startsWith(prefix) ? docId.slice(prefix.length) : docId;
  };
  try {
    const qReports = query(collection(db, 'reports'), where('ownerId', '==', ownerId));
    const snap = await getDocs(qReports);
    const reports: any[] = [];
    for (const rdoc of snap.docs) {
      const data = rdoc.data();
      const rid = stripPrefix(data.id);
      try {
        const parsed = JSON.parse(data.payloadJson);
        if (parsed && parsed.id) {
          parsed.id = rid;
          reports.push(parsed);
        }
      } catch {}
    }
    return reports.sort((a, b) => (b.updatedAt || b.id).localeCompare(a.updatedAt || a.id));
  } catch {
    return [];
  }
}

export {
  auth,
  googleProvider,
  signInWithPopup,
  fbSignOut,
  onAuthStateChanged,
  type User,
};
