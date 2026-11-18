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

let firestoreWritePaused = false;

async function runFirestoreWrite<T>(fn: () => Promise<T>, timeoutMs = 2500): Promise<boolean> {
  if (firestoreWritePaused) return false;
  let timer: any = null;
  try {
    const timeoutPromise = new Promise<'TIMEOUT'>(resolve => {
      timer = setTimeout(() => resolve('TIMEOUT'), timeoutMs);
    });
    const res = await Promise.race([fn(), timeoutPromise]);
    if (res === 'TIMEOUT') {
      firestoreWritePaused = true;
      return false;
    }
    return true;
  } catch {
    firestoreWritePaused = true;
    return false;
  } finally {
    if (timer) clearTimeout(timer);
  }
}

async function serverFetchAll(): Promise<{
  lessons: any[];
  reports: any[];
  deletedLessonIds: Set<string>;
  deletedReportIds: Set<string>;
}> {
  const ownerId = getActiveCloudOwnerId() || `sync_${DEFAULT_TEACHER_SYNC_KEY}`;
  try {
    const resp = await fetch(`/api/storage/all?ownerId=${encodeURIComponent(ownerId)}`);
    if (!resp.ok) {
      return { lessons: [], reports: [], deletedLessonIds: new Set(), deletedReportIds: new Set() };
    }
    const data = await resp.json();
    const lessons: any[] = [];
    if (Array.isArray(data?.lessons)) {
      for (const item of data.lessons) {
        if (item && typeof item.serialized === 'string') {
          try {
            const hydrated = deserializeLessonWithMedia(item.serialized);
            hydrated.id = sanitizeId(item.id || hydrated.id);
            lessons.push(hydrated);
          } catch {
            // skip invalid serialized lesson
          }
        }
      }
    }
    const reports: any[] = Array.isArray(data?.reports) ? data.reports : [];
    const deletedLessonIds = new Set<string>(
      Array.isArray(data?.deletedLessonIds) ? data.deletedLessonIds.map((x: any) => sanitizeId(String(x))) : []
    );
    const deletedReportIds = new Set<string>(
      Array.isArray(data?.deletedReportIds) ? data.deletedReportIds.map((x: any) => sanitizeId(String(x))) : []
    );
    return { lessons, reports, deletedLessonIds, deletedReportIds };
  } catch {
    return { lessons: [], reports: [], deletedLessonIds: new Set(), deletedReportIds: new Set() };
  }
}

async function serverSaveLesson(lesson: any, rawId: string, fullJson: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId() || `sync_${DEFAULT_TEACHER_SYNC_KEY}`;
  try {
    await fetch('/api/storage/lesson', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ownerId,
        id: rawId,
        updatedAt: lesson.updatedAt || new Date().toISOString(),
        serialized: fullJson,
      }),
    });
  } catch {
    // ignore network errors
  }
}

async function serverDeleteLesson(rawId: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId() || `sync_${DEFAULT_TEACHER_SYNC_KEY}`;
  try {
    await fetch(`/api/storage/lesson/${encodeURIComponent(rawId)}?ownerId=${encodeURIComponent(ownerId)}`, {
      method: 'DELETE',
    });
  } catch {
    // ignore
  }
}

async function serverSaveReport(report: any, rawId: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId() || `sync_${DEFAULT_TEACHER_SYNC_KEY}`;
  try {
    await fetch('/api/storage/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ownerId,
        report: { ...report, id: rawId },
      }),
    });
  } catch {
    // ignore
  }
}

async function serverDeleteReport(rawId: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId() || `sync_${DEFAULT_TEACHER_SYNC_KEY}`;
  try {
    await fetch(`/api/storage/report/${encodeURIComponent(rawId)}?ownerId=${encodeURIComponent(ownerId)}`, {
      method: 'DELETE',
    });
  } catch {
    // ignore
  }
}

export async function cloudSaveLesson(lesson: any): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return;

  const rawId = sanitizeId(lesson.id || Math.random().toString(36).slice(2, 10));
  const fullJson = await serializeLessonWithMedia({ ...lesson, id: rawId });

  // 1. Immediately persist to backend system storage (fast, reliable, zero browser storage)
  await serverSaveLesson(lesson, rawId, fullJson);

  // 2. Non-blocking Firestore sync if quota is available
  if (firestoreWritePaused) return;

  const docId = sanitizeId(`${ownerId}_${rawId}`);
  const chunks: string[] = [];
  for (let i = 0; i < fullJson.length; i += CHUNK_MAX_CHARS) {
    chunks.push(fullJson.slice(i, i + CHUNK_MAX_CHARS));
  }
  if (chunks.length === 0) {
    chunks.push('{}');
  }
  const safeChunks = chunks.slice(0, 200);
  const updatedAtMs = lesson.updatedAt ? Date.parse(lesson.updatedAt) || Date.now() : Date.now();
  const lessonRef = doc(db, 'lessons', docId);

  let existingSnap: any = null;
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

  const metaOk = await runFirestoreWrite(() =>
    existingSnap && existingSnap.exists()
      ? updateDoc(lessonRef, {
          ...metaFields,
          updatedAt: serverTimestamp(),
        })
      : setDoc(lessonRef, {
          id: docId,
          ownerId,
          ...metaFields,
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
        })
  );
  if (!metaOk) return;

  for (let idx = 0; idx < safeChunks.length; idx++) {
    if (firestoreWritePaused) return;
    const chunkId = `chunk_${idx}`;
    const chunkRef = doc(db, 'lessons', docId, 'chunks', chunkId);
    let existingChunk: any = null;
    try {
      existingChunk = await getDoc(chunkRef);
    } catch {
      existingChunk = null;
    }

    const chunkOk = await runFirestoreWrite(() =>
      existingChunk && existingChunk.exists()
        ? updateDoc(chunkRef, {
            idx,
            data: safeChunks[idx],
            updatedAt: serverTimestamp(),
          })
        : setDoc(chunkRef, {
            ownerId,
            lessonId: docId,
            idx,
            data: safeChunks[idx],
            createdAt: serverTimestamp(),
            updatedAt: serverTimestamp(),
          })
    );
    if (!chunkOk) return;
  }
}

export async function cloudDeleteLesson(rawLessonId: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return;
  const rawId = sanitizeId(rawLessonId);

  await serverDeleteLesson(rawId);

  if (firestoreWritePaused) return;
  const docIdsToTry = [sanitizeId(`${ownerId}_${rawId}`), rawId];

  for (const docId of docIdsToTry) {
    if (firestoreWritePaused) break;
    const chunksCol = collection(db, 'lessons', docId, 'chunks');
    try {
      const q = query(chunksCol, where('ownerId', '==', ownerId));
      const snap = await getDocs(q);
      for (const d of snap.docs) {
        const ok = await runFirestoreWrite(() => deleteDoc(d.ref));
        if (!ok) break;
      }
      await runFirestoreWrite(() => deleteDoc(doc(db, 'lessons', docId)));
    } catch {
      // ignore if doc does not exist
    }
  }
}

export async function cloudSaveReport(report: any): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId || !report?.id) return;

  const rawId = sanitizeId(report.id);
  await serverSaveReport(report, rawId);

  if (firestoreWritePaused) return;

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
  const reportRef = doc(db, 'reports', docId);

  let existingSnap: any = null;
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

  await runFirestoreWrite(() =>
    existingSnap && existingSnap.exists()
      ? updateDoc(reportRef, {
          ...fields,
          updatedAt: serverTimestamp(),
        })
      : setDoc(reportRef, {
          id: docId,
          ownerId,
          ...fields,
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
        })
  );
}

export async function cloudDeleteReport(rawReportId: string): Promise<void> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return;
  const rawId = sanitizeId(rawReportId);

  await serverDeleteReport(rawId);

  if (firestoreWritePaused) return;
  const docIdsToTry = [sanitizeId(`${ownerId}_${rawId}`), rawId];
  for (const docId of docIdsToTry) {
    if (firestoreWritePaused) break;
    await runFirestoreWrite(() => deleteDoc(doc(db, 'reports', docId)));
  }
}

export async function cloudSyncAll(
  localLessons: any[],
  localReports: any[],
  _saveLocalLesson?: (l: any) => Promise<void>,
  _saveLocalReport?: (r: any) => Promise<void>
): Promise<{ lessons: any[]; reports: any[] }> {
  const [fetchedLessons, fetchedReports] = await Promise.all([
    cloudFetchAllLessons(),
    cloudFetchAllReports(),
  ]);

  const lessonsMap = new Map<string, any>();
  for (const l of localLessons || []) {
    if (l?.id) {
      const cid = sanitizeId(l.id);
      lessonsMap.set(cid, { ...l, id: cid });
    }
  }
  for (const rl of fetchedLessons || []) {
    if (rl?.id) {
      const cid = sanitizeId(rl.id);
      const existing = lessonsMap.get(cid);
      const existingMs = existing?.updatedAt ? Date.parse(existing.updatedAt) || 0 : 0;
      const remoteMs = rl?.updatedAt ? Date.parse(rl.updatedAt) || 0 : 0;
      if (!existing || remoteMs >= existingMs) {
        lessonsMap.set(cid, { ...rl, id: cid });
      }
    }
  }

  const reportsMap = new Map<string, any>();
  for (const r of localReports || []) {
    if (r?.id) {
      const cid = sanitizeId(r.id);
      reportsMap.set(cid, { ...r, id: cid });
    }
  }
  for (const rr of fetchedReports || []) {
    if (rr?.id) {
      const cid = sanitizeId(rr.id);
      const existing = reportsMap.get(cid);
      const existingMs = existing?.updatedAt ? Date.parse(existing.updatedAt) || 0 : 0;
      const remoteMs = rr?.updatedAt ? Date.parse(rr.updatedAt) || 0 : 0;
      if (!existing || remoteMs >= existingMs) {
        reportsMap.set(cid, { ...rr, id: cid });
      }
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

  const serverData = await serverFetchAll();
  const byId = new Map<string, any>();
  for (const sl of serverData.lessons) {
    if (sl?.id && !serverData.deletedLessonIds.has(sanitizeId(sl.id))) {
      const cid = sanitizeId(sl.id);
      byId.set(cid, { ...sl, id: cid });
    }
  }

  try {
    const qLessons = query(collection(db, 'lessons'), where('ownerId', '==', ownerId));
    const snap = await getDocs(qLessons);
    for (const dSnap of snap.docs) {
      const meta = dSnap.data();
      const docId = meta.id;
      const lid = sanitizeId(stripPrefix(docId));
      if (serverData.deletedLessonIds.has(lid)) continue;

      const existing = byId.get(lid);
      const existingMs = existing?.updatedAt ? Date.parse(existing.updatedAt) || 0 : 0;
      const remoteMs = typeof meta.updatedAtMs === 'number' ? meta.updatedAtMs : 0;
      if (existing && existingMs >= remoteMs) continue;

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
          byId.set(lid, hydrated);
          serverSaveLesson(hydrated, lid, joined).catch(() => {});
        }
      } catch {}
    }
  } catch {
    // ignore Firestore read error
  }

  return Array.from(byId.values()).sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}

export async function cloudFetchAllReports(): Promise<any[]> {
  const ownerId = getActiveCloudOwnerId();
  if (!ownerId) return [];
  const stripPrefix = (docId: string) => {
    const prefix = `${ownerId}_`;
    return docId.startsWith(prefix) ? docId.slice(prefix.length) : docId;
  };

  const serverData = await serverFetchAll();
  const byId = new Map<string, any>();
  for (const sr of serverData.reports) {
    if (sr?.id && !serverData.deletedReportIds.has(sanitizeId(sr.id))) {
      const cid = sanitizeId(sr.id);
      byId.set(cid, { ...sr, id: cid });
    }
  }

  try {
    const qReports = query(collection(db, 'reports'), where('ownerId', '==', ownerId));
    const snap = await getDocs(qReports);
    for (const rdoc of snap.docs) {
      const data = rdoc.data();
      const rid = sanitizeId(stripPrefix(data.id));
      if (serverData.deletedReportIds.has(rid)) continue;

      const existing = byId.get(rid);
      const existingMs = existing?.updatedAt ? Date.parse(existing.updatedAt) || 0 : 0;
      const remoteMs = typeof data.updatedAtMs === 'number' ? data.updatedAtMs : 0;
      if (existing && existingMs >= remoteMs) continue;

      try {
        const parsed = JSON.parse(data.payloadJson);
        if (parsed && parsed.id) {
          parsed.id = rid;
          byId.set(rid, parsed);
          serverSaveReport(parsed, rid).catch(() => {});
        }
      } catch {}
    }
  } catch {
    // ignore Firestore read error
  }

  return Array.from(byId.values()).sort((a, b) => (b.updatedAt || b.id).localeCompare(a.updatedAt || a.id));
}

export {
  auth,
  googleProvider,
  signInWithPopup,
  fbSignOut,
  onAuthStateChanged,
  type User,
};
