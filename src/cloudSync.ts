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

function sanitizeId(rawId: string): string {
  const clean = (rawId || '')
    .replace(/[^a-zA-Z0-9_-]/g, '_')
    .slice(0, 120);
  return clean || 'doc_' + Math.random().toString(36).slice(2, 10);
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
  const user = auth.currentUser;
  if (!user || !user.emailVerified) return;

  const lessonId = sanitizeId(lesson.id || Math.random().toString(36).slice(2, 10));
  const fullJson = await serializeLessonWithMedia({ ...lesson, id: lessonId });
  const chunks: string[] = [];
  for (let i = 0; i < fullJson.length; i += CHUNK_MAX_CHARS) {
    chunks.push(fullJson.slice(i, i + CHUNK_MAX_CHARS));
  }
  if (chunks.length === 0) {
    chunks.push('{}');
  }
  const safeChunks = chunks.slice(0, 200);
  const updatedAtMs = lesson.updatedAt ? Date.parse(lesson.updatedAt) || Date.now() : Date.now();

  const lessonPath = `lessons/${lessonId}`;
  const lessonRef = doc(db, 'lessons', lessonId);

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
        id: lessonId,
        ownerId: user.uid,
        ...metaFields,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      });
    }
  } catch (error) {
    handleFirestoreError(error, existingSnap?.exists() ? OperationType.UPDATE : OperationType.CREATE, lessonPath);
  }

  // Write chunks
  for (let idx = 0; idx < safeChunks.length; idx++) {
    const chunkId = `chunk_${idx}`;
    const chunkPath = `lessons/${lessonId}/chunks/${chunkId}`;
    const chunkRef = doc(db, 'lessons', lessonId, 'chunks', chunkId);
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
          ownerId: user.uid,
          lessonId,
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
  const user = auth.currentUser;
  if (!user || !user.emailVerified) return;
  const lessonId = sanitizeId(rawLessonId);
  const chunksCol = collection(db, 'lessons', lessonId, 'chunks');
  try {
    const q = query(chunksCol, where('ownerId', '==', user.uid));
    const snap = await getDocs(q);
    for (const d of snap.docs) {
      await deleteDoc(d.ref);
    }
  } catch {
    // ignore chunk cleanup errors
  }
  const lessonPath = `lessons/${lessonId}`;
  try {
    await deleteDoc(doc(db, 'lessons', lessonId));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, lessonPath);
  }
}

export async function cloudSaveReport(report: any): Promise<void> {
  const user = auth.currentUser;
  if (!user || !user.emailVerified || !report?.id) return;

  const reportId = sanitizeId(report.id);
  let payloadJson = JSON.stringify(report);
  if (payloadJson.length > REPORT_MAX_CHARS) {
    // Trim oversize data-URL drawings so report fits within Firestore 1MB limit
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
  const reportPath = `reports/${reportId}`;
  const reportRef = doc(db, 'reports', reportId);

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
        id: reportId,
        ownerId: user.uid,
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
  const user = auth.currentUser;
  if (!user || !user.emailVerified) return;
  const reportId = sanitizeId(rawReportId);
  const reportPath = `reports/${reportId}`;
  try {
    await deleteDoc(doc(db, 'reports', reportId));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, reportPath);
  }
}

export async function cloudSyncAll(
  localLessons: any[],
  localReports: any[],
  saveLocalLesson: (l: any) => Promise<void>,
  saveLocalReport: (r: any) => Promise<void>
): Promise<{ lessons: any[]; reports: any[] }> {
  const user = auth.currentUser;
  if (!user || !user.emailVerified) {
    return { lessons: localLessons, reports: localReports };
  }

  // 1. Fetch remote lessons metadata
  const lessonsMap = new Map<string, any>();
  for (const l of localLessons) {
    if (l?.id) lessonsMap.set(sanitizeId(l.id), { ...l, id: sanitizeId(l.id) });
  }

  let remoteLessonDocs: any[] = [];
  try {
    const qLessons = query(collection(db, 'lessons'), where('ownerId', '==', user.uid));
    const snap = await getDocs(qLessons);
    remoteLessonDocs = snap.docs.map(d => d.data());
  } catch (error) {
    handleFirestoreError(error, OperationType.LIST, 'lessons');
  }

  const remoteLessonIds = new Set<string>();
  for (const meta of remoteLessonDocs) {
    const lid = meta.id;
    remoteLessonIds.add(lid);
    const local = lessonsMap.get(lid);
    const localMs = local?.updatedAt ? Date.parse(local.updatedAt) || 0 : 0;
    const remoteMs = typeof meta.updatedAtMs === 'number' ? meta.updatedAtMs : 0;

    if (!local || remoteMs > localMs + 1000) {
      // Download chunks from cloud
      try {
        const qChunks = query(
          collection(db, 'lessons', lid, 'chunks'),
          where('ownerId', '==', user.uid)
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
          lessonsMap.set(lid, hydrated);
          await saveLocalLesson(hydrated);
        }
      } catch {
        // keep local if chunk assembly fails
      }
    } else if (local && localMs > remoteMs + 1000) {
      await cloudSaveLesson(local);
    }
  }

  // Upload local lessons that do not exist in cloud yet
  for (const [lid, local] of lessonsMap.entries()) {
    if (!remoteLessonIds.has(lid)) {
      await cloudSaveLesson(local);
    }
  }

  // 2. Fetch remote session reports
  const reportsMap = new Map<string, any>();
  for (const r of localReports) {
    if (r?.id) reportsMap.set(sanitizeId(r.id), { ...r, id: sanitizeId(r.id) });
  }

  let remoteReportDocs: any[] = [];
  try {
    const qReports = query(collection(db, 'reports'), where('ownerId', '==', user.uid));
    const snap = await getDocs(qReports);
    remoteReportDocs = snap.docs.map(d => d.data());
  } catch (error) {
    handleFirestoreError(error, OperationType.LIST, 'reports');
  }

  const remoteReportIds = new Set<string>();
  for (const rdoc of remoteReportDocs) {
    const rid = rdoc.id;
    remoteReportIds.add(rid);
    const local = reportsMap.get(rid);
    const localMs = local?.updatedAt ? Date.parse(local.updatedAt) || 0 : 0;
    const remoteMs = typeof rdoc.updatedAtMs === 'number' ? rdoc.updatedAtMs : 0;

    if (!local || remoteMs > localMs + 1000) {
      try {
        const parsed = JSON.parse(rdoc.payloadJson);
        if (parsed && parsed.id) {
          reportsMap.set(rid, parsed);
          await saveLocalReport(parsed);
        }
      } catch {
        // skip invalid JSON
      }
    } else if (local && localMs > remoteMs + 1000) {
      await cloudSaveReport(local);
    }
  }

  for (const [rid, local] of reportsMap.entries()) {
    if (!remoteReportIds.has(rid)) {
      await cloudSaveReport(local);
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

export {
  auth,
  googleProvider,
  signInWithPopup,
  fbSignOut,
  onAuthStateChanged,
  type User,
};
