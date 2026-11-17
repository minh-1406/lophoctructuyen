/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import { 
  Sparkles, BookOpen, Plus, Trash2, ArrowRight, ArrowLeft, 
  Check, X, Trophy, Play, Upload, FileText, QrCode, RefreshCw, BarChart2, 
  HelpCircle, MessageSquare, Cloud, Sliders, Edit3, Image as ImageIcon, Info,
  Share2, Award, ChevronRight, CheckCircle2, AlertCircle, Save, FolderOpen,
  Clock, Pause, RotateCcw, Maximize2, Minimize2, Eye
} from 'lucide-react';
import {
  PptxSlideData,
  PptxSlideView,
  parsePptxFile,
  getPptxMaxStep,
  repairAndVerifyMediaBlob,
} from './pptxEngine';
import {
  auth,
  googleProvider,
  signInWithPopup,
  fbSignOut,
  onAuthStateChanged,
  cloudSaveLesson,
  cloudDeleteLesson,
  cloudSaveReport,
  cloudDeleteReport,
  cloudFetchAllLessons,
  cloudFetchAllReports,
  cloudSyncAll,
  getSavedSyncKey,
  setSavedSyncKey,
  getActiveCloudOwnerId,
  getActiveCloudLabel,
  exportBackupJson,
  importBackupJson,
  type User,
} from './cloudSync';

// Declare external globals from CDN scripts
declare const Peer: any;
declare const qrcode: any;
declare const mammoth: any;
declare const pdfjsLib: any;
declare const JSZip: any;

interface Slide {
  type: 'open' | 'poll' | 'cloud' | 'quiz' | 'scale' | 'media' | 'draw' | 'info';
  q: string;
  b?: string; // body for info or media
  opts?: string[];
  a?: number; // correct answer for quiz
  imgs?: string[]; // media image ids
  vid?: string; // media video id
  tx?: string; // text for pdf/doc
  auto?: boolean;
  ai?: boolean;
  aiGenerated?: boolean;
  duration?: number; // countdown seconds (0/undefined = unlimited)
  orig?: boolean;
  _fb?: number;
  pptx?: PptxSlideData;
}

interface Lesson {
  id?: string;
  title: string;
  className?: string;
  customCode?: string; // Teacher-editable alphanumeric class code (e.g., 11A1, TOAN10, AB123)
  classCodes?: string[]; // Multiple class codes saved/taught for this lesson
  slides: Slide[];
  M?: Record<string, Blob>;
  updatedAt?: string;
}

interface SavedActivityResponse {
  sid: string;
  name: string;
  className?: string;
  value: any;
  grade?: number;
  submittedAt?: string;
}

interface StudentSessionStat {
  sid: string;
  name: string;
  className?: string;
  answeredCount: number;
  participationPct: number;
  score: number;
}

interface SlideSessionStat {
  idx: number;
  type: Slide['type'];
  q: string;
  b?: string;
  opts?: string[];
  correctIdx?: number;
  isInteractive?: boolean;
  submittedCount: number;
  participationPct: number;
  isAi?: boolean;
  responses?: SavedActivityResponse[];
}

interface ClassSessionReport {
  id: string;
  lessonId?: string;
  className: string;
  studentClassNames?: string[];
  lessonTitle: string;
  roomCode: string;
  startedAt?: string;
  updatedAt?: string;
  endedAt: string;
  status?: 'live' | 'ended';
  totalStudents: number;
  totalInteractiveSlides: number;
  overallParticipationPct: number;
  students: StudentSessionStat[];
  slides: SlideSessionStat[];
}

const SLIDE_TYPES: Record<string, string> = {
  open: 'Câu hỏi mở',
  poll: 'Bình chọn',
  cloud: 'Đám mây từ',
  quiz: 'Trắc nghiệm',
  scale: 'Thang điểm',
  media: 'Slide gốc',
  draw: 'Vẽ',
  info: 'Nội dung'
};

const TEMPLATES: Lesson[] = [
  {
    title: 'Địa lý 11 – Khởi động: Đông Nam Á',
    slides: [
      { type: 'cloud', q: 'Nhắc đến Đông Nam Á, em nghĩ ngay đến từ nào?' },
      { type: 'poll', q: 'Quốc gia nào KHÔNG thuộc ASEAN?', opts: ['Thái Lan', 'Việt Nam', 'Bangladesh', 'Philippines'] },
      { type: 'quiz', q: 'Sông Mê Công chảy qua bao nhiêu quốc gia?', opts: ['4', '5', '6', '7'], a: 2 },
      { type: 'draw', q: 'Vẽ sơ đồ vị trí Việt Nam trong khu vực Đông Nam Á' },
      { type: 'open', q: 'Theo em, vì sao Đông Nam Á có vị trí chiến lược?' }
    ]
  },
  {
    title: 'Biến đổi khí hậu – thảo luận',
    slides: [
      { type: 'info', q: 'Mục tiêu bài học', b: 'Hiểu nguyên nhân, biểu hiện và giải pháp ứng phó biến đổi khí hậu.' },
      { type: 'scale', q: 'Em lo ngại mức nào về biến đổi khí hậu ở địa phương?' },
      { type: 'cloud', q: 'Một từ mô tả hậu quả của biến đổi khí hậu' },
      { type: 'open', q: 'Em có thể làm gì trong tuần này để giảm tác động?' }
    ]
  },
  {
    title: 'Bài khởi động đầu giờ (mọi môn)',
    slides: [
      { type: 'cloud', q: 'Hôm nay em cảm thấy thế nào? (một từ)' },
      { type: 'scale', q: 'Em tự tin với bài học hôm trước ở mức nào?' },
      { type: 'open', q: 'Điều em còn thắc mắc từ bài trước là gì?' }
    ]
  }
];

export default function App() {
  const [view, setView] = useState<'home' | 'dash' | 'build' | 'host' | 'join' | 'stats'>('home');

  // Lessons / Dash
  const [mine, setMine] = useState<Lesson[]>([]);
  const [lesson, setLesson] = useState<Lesson>({ title: '', className: '', customCode: '', classCodes: [], slides: [{ type: 'open', q: 'Em đã biết gì về chủ đề hôm nay?' }], M: {} });
  const [classCodeDraft, setClassCodeDraft] = useState('');
  const [classCodeMsg, setClassCodeMsg] = useState('');
  const [slideIdx, setSlideIdx] = useState(0);
  const [topicInput, setTopicInput] = useState('');
  const [genMsg, setGenMsg] = useState('');
  const [isAutoInserting, setIsAutoInserting] = useState(false);
  const [autoInsertMsg, setAutoInsertMsg] = useState('');
  const [defaultAiTimerSec, setDefaultAiTimerSec] = useState(45);
  const [fileMsg, setFileMsg] = useState('.pdf, .docx, .txt, .md, .csv hoặc ảnh trang sách');
  const [docTextSrc, setDocTextSrc] = useState('');
  const [saveMsg, setSaveMsg] = useState('');
  const [importingImage, setImportingImage] = useState<File | null>(null);
  const [aiQuestionCount, setAiQuestionCount] = useState<number>(4);

  // Host State
  const [roomCode, setRoomCode] = useState('');
  const [hostCodeDraft, setHostCodeDraft] = useState('');
  const [roomLink, setRoomLink] = useState('');
  const [qrDataUrl, setQrDataUrl] = useState('');
  const [hostStatus, setHostStatus] = useState('Đang khởi tạo phòng...');
  const [teacherStartedLive, setTeacherStartedLive] = useState(false);
  const [showQrModal, setShowQrModal] = useState(false);
  const [connectedStudents, setConnectedStudents] = useState<Record<string, string>>({});
  const [allSessionStudents, setAllSessionStudents] = useState<Record<string, string>>({});
  const [allSessionStudentClasses, setAllSessionStudentClasses] = useState<Record<string, string>>({});
  const [responses, setResponses] = useState<Record<number, Record<string, { n: string; cls?: string; v: any; g?: number; ts?: string }>>>({});
  const [peerPeer, setPeerPeer] = useState<any>(null);
  const [connections, setConnections] = useState<any[]>([]);
  const [mediaCache, setMediaCache] = useState<Record<string, string>>({});
  const [aiFeedbackText, setAiFeedbackText] = useState('');
  const [showScoreboard, setShowScoreboard] = useState(false);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [hostAnimStep, setHostAnimStep] = useState(0);
  const [builderPreviewIdx, setBuilderPreviewIdx] = useState<number | null>(null);
  const [builderAnimStep, setBuilderAnimStep] = useState(0);
  const liveStageRef = useRef<HTMLDivElement | null>(null);
  const hasNativeFullscreenRef = useRef(false);

  // Host Countdown Timer State
  const [hostTimerLeft, setHostTimerLeft] = useState<number | null>(null);
  const [hostTimerTotal, setHostTimerTotal] = useState(0);
  const [hostTimerRunning, setHostTimerRunning] = useState(false);

  // Session Statistics History (stored in IndexedDB 'reports' with localStorage migration)
  const [sessionReports, setSessionReports] = useState<ClassSessionReport[]>(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('curivn_session_reports') || '[]');
      return Array.isArray(saved) ? saved : [];
    } catch {
      return [];
    }
  });
  const [activeReportId, setActiveReportId] = useState<string | null>(null);
  const [statsClassFilter, setStatsClassFilter] = useState<string>('ALL');
  const [expandedActivitySlides, setExpandedActivitySlides] = useState<Record<string, boolean>>({});
  const [cloudUser, setCloudUser] = useState<User | null>(null);
  const [cloudSyncKey, setCloudSyncKey] = useState<string>(() => getSavedSyncKey());
  const [cloudKeyInput, setCloudKeyInput] = useState<string>(() => getSavedSyncKey());
  const [showCloudModal, setShowCloudModal] = useState(false);
  const [cloudSyncing, setCloudSyncing] = useState(false);
  const [cloudSyncMsg, setCloudSyncMsg] = useState('');
  const [revealedSlides, setRevealedSlides] = useState<Record<number, boolean>>({});
  const [votingActive, setVotingActive] = useState<Record<number, boolean>>({});
  const [studentVotes, setStudentVotes] = useState<Record<number, Record<string, string>>>({});
  const [studentVotingActive, setStudentVotingActive] = useState(false);
  const [studentSubmissions, setStudentSubmissions] = useState<any[]>([]);
  const [studentMyVote, setStudentMyVote] = useState<string | null>(null);

  const getSlideVoteTally = (sIdx: number) => {
    const tally: Record<string, number> = {};
    const slideMap = studentVotes?.[sIdx] || {};
    for (const targetSid of Object.values(slideMap)) {
      if (targetSid) tally[targetSid] = (tally[targetSid] || 0) + 1;
    }
    return tally;
  };

  const startVotingPhase = () => {
    const curIdx = slideIdxRef.current;
    const respMap = responsesRef.current?.[curIdx] || {};
    const r = Object.entries(respMap).map(([sid, val]) => ({
      sid,
      n: val?.n || sid,
      cls: val?.cls || '',
      v: val?.v || '',
    }));
    setVotingActive(prev => ({ ...(prev || {}), [curIdx]: true }));
    connectionsRef.current.forEach(c => {
      if (c && c.open) c.send({ t: 'voting_start', slideIdx: curIdx, submissions: r });
    });
  };

  const forceRevealResults = () => {
    setRevealedSlides(prev => ({ ...(prev || {}), [slideIdxRef.current]: true }));
    setHostTimerRunning(false);
    hostTimerRunningRef.current = false;
    setHostTimerLeft(0);
    hostTimerLeftRef.current = 0;
  };



  // Student State
  const [joinCodeInput, setJoinCodeInput] = useState('');
  const [studentName, setStudentName] = useState(() => {
    try {
      const p = JSON.parse(localStorage.getItem('curivn_student_profile') || '{}');
      return typeof p.name === 'string' ? p.name : '';
    } catch {
      return '';
    }
  });
  const [studentClassName, setStudentClassName] = useState(() => {
    try {
      const p = JSON.parse(localStorage.getItem('curivn_student_profile') || '{}');
      return typeof p.className === 'string' ? p.className : '';
    } catch {
      return '';
    }
  });
  const [joinMsg, setJoinMsg] = useState('');
  const [studentRoomJoined, setStudentRoomJoined] = useState(false);
  const [studentProfileConfirmed, setStudentProfileConfirmed] = useState(false);
  const [studentProfileError, setStudentProfileError] = useState('');
  const [studentConnected, setStudentConnected] = useState(false);
  const [studentCurState, setStudentCurState] = useState<any>(null);
  const [studentMyAns, setStudentMyAns] = useState<Record<number, any>>({});
  const [studentMyScore, setStudentMyScore] = useState<number | null>(null);
  const [studentPeer, setStudentPeer] = useState<any>(null);
  const [studentConn, setStudentConn] = useState<any>(null);
  const [studentTimerLeft, setStudentTimerLeft] = useState<number | null>(null);
  const [studentTimerTotal, setStudentTimerTotal] = useState(0);
  const [studentTimerRunning, setStudentTimerRunning] = useState(false);
  const [studentAnimStep, setStudentAnimStep] = useState(0);
  const studentSid = useRef(Math.random().toString(36).slice(2, 8)).current;

  // Canvas drawing ref for student
  const studentCanvasRef = useRef<HTMLCanvasElement | null>(null);

  // Refs to prevent stale closures in PeerJS callbacks & timers
  const lessonRef = useRef(lesson);
  const slideIdxRef = useRef(slideIdx);
  const hostAnimStepRef = useRef(hostAnimStep);
  const responsesRef = useRef(responses);
  const connectionsRef = useRef(connections);
  const connectedStudentsRef = useRef(connectedStudents);
  const allSessionStudentsRef = useRef(allSessionStudents);
  const allSessionStudentClassesRef = useRef(allSessionStudentClasses);
  const roomCodeRef = useRef(roomCode);
  const currentSessionIdRef = useRef<string | null>(null);
  const currentSessionStartedAtRef = useRef<string>('');
  const lastSavedLessonSigRef = useRef<string>('');
  const teacherStartedLiveRef = useRef(teacherStartedLive);
  const hostTimerLeftRef = useRef(hostTimerLeft);
  const hostTimerTotalRef = useRef(hostTimerTotal);
  const hostTimerRunningRef = useRef(hostTimerRunning);
  const studentNameRef = useRef(studentName);
  const studentClassNameRef = useRef(studentClassName);
  const studentProfileConfirmedRef = useRef(studentProfileConfirmed);
  const studentPeerRef = useRef(studentPeer);
  const studentConnRef = useRef(studentConn);
  const studentTimerLeftRef = useRef(studentTimerLeft);
  const askedMediaRef = useRef<Record<string, number>>({});
  const lastAutoJoinedRef = useRef('');

  useEffect(() => { lessonRef.current = lesson; }, [lesson]);
  useEffect(() => { slideIdxRef.current = slideIdx; }, [slideIdx]);
  useEffect(() => { hostAnimStepRef.current = hostAnimStep; }, [hostAnimStep]);
  useEffect(() => { responsesRef.current = responses; }, [responses]);
  useEffect(() => { connectionsRef.current = connections; }, [connections]);
  useEffect(() => { connectedStudentsRef.current = connectedStudents; }, [connectedStudents]);
  useEffect(() => { allSessionStudentsRef.current = allSessionStudents; }, [allSessionStudents]);
  useEffect(() => { allSessionStudentClassesRef.current = allSessionStudentClasses; }, [allSessionStudentClasses]);
  useEffect(() => { roomCodeRef.current = roomCode; }, [roomCode]);
  useEffect(() => { teacherStartedLiveRef.current = teacherStartedLive; }, [teacherStartedLive]);
  useEffect(() => { hostTimerLeftRef.current = hostTimerLeft; }, [hostTimerLeft]);
  useEffect(() => { hostTimerTotalRef.current = hostTimerTotal; }, [hostTimerTotal]);
  useEffect(() => { hostTimerRunningRef.current = hostTimerRunning; }, [hostTimerRunning]);
  useEffect(() => { studentNameRef.current = studentName; }, [studentName]);
  useEffect(() => { studentClassNameRef.current = studentClassName; }, [studentClassName]);
  useEffect(() => { studentProfileConfirmedRef.current = studentProfileConfirmed; }, [studentProfileConfirmed]);
  useEffect(() => { studentPeerRef.current = studentPeer; }, [studentPeer]);
  useEffect(() => { studentConnRef.current = studentConn; }, [studentConn]);
  useEffect(() => { studentTimerLeftRef.current = studentTimerLeft; }, [studentTimerLeft]);

  const formatTimer = (sec: number) => {
    const s = Math.max(0, Math.floor(sec));
    const m = Math.floor(s / 60);
    const r = s % 60;
    return `${String(m).padStart(2, '0')}:${String(r).padStart(2, '0')}`;
  };

  const applyAiSlideDefaults = (slides: Slide[], timerSec = defaultAiTimerSec): Slide[] =>
    slides.map(s => {
      const isInteractive = s.type !== 'info' && s.type !== 'media';
      return {
        ...s,
        aiGenerated: isInteractive ? true : s.aiGenerated,
        duration: isInteractive ? (s.duration && s.duration > 0 ? s.duration : (timerSec > 0 ? timerSec : undefined)) : s.duration,
      };
    });

  const localMediaUrlsRef = useRef<Record<string, string>>({});
  const repairedMediaIdsRef = useRef<Record<string, boolean>>({});

  // Pre-cache and auto-repair host media blob URLs (fixes TIFF/SVG/EMF blobs even in already-saved IndexedDB lessons)
  useEffect(() => {
    if (!lesson.M) return;
    let cancelled = false;
    const immediateUrls: Record<string, string> = {};
    let immediateChanged = false;

    for (const [id, rawEntry] of Object.entries(lesson.M)) {
      if (!mediaCache[id] && !localMediaUrlsRef.current[id]) {
        const blob =
          rawEntry instanceof Blob
            ? rawEntry
            : (rawEntry as any)?._ab
            ? new Blob([(rawEntry as any)._ab], { type: (rawEntry as any).type || 'image/png' })
            : null;
        if (blob && blob.size > 0) {
          const u = URL.createObjectURL(blob);
          localMediaUrlsRef.current[id] = u;
          immediateUrls[id] = u;
          immediateChanged = true;
        }
      }
    }
    if (immediateChanged) {
      setMediaCache(prev => ({ ...prev, ...immediateUrls }));
    }

    // Async inspection & repair for any TIFF/EMF/SVG or IndexedDB-restored image blobs
    (async () => {
      const repairedUrls: Record<string, string> = {};
      let anyRepaired = false;
      for (const [id, rawEntry] of Object.entries(lesson.M || {})) {
        if (cancelled || repairedMediaIdsRef.current[id]) continue;
        repairedMediaIdsRef.current[id] = true;
        const blob =
          rawEntry instanceof Blob
            ? rawEntry
            : (rawEntry as any)?._ab
            ? new Blob([(rawEntry as any)._ab], { type: (rawEntry as any).type || 'image/png' })
            : null;
        if (!blob || blob.size === 0) continue;
        if (blob.type.startsWith('video/') || blob.type.startsWith('audio/')) continue;

        const fixed = await repairAndVerifyMediaBlob(blob);
        if (!cancelled && fixed) {
          if (lessonRef.current.M) {
            lessonRef.current.M[id] = fixed;
          }
          const newUrl = URL.createObjectURL(fixed);
          localMediaUrlsRef.current[id] = newUrl;
          repairedUrls[id] = newUrl;
          anyRepaired = true;
        }
      }
      if (!cancelled && anyRepaired) {
        setMediaCache(prev => ({ ...prev, ...repairedUrls }));
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [Object.keys(lesson.M || {}).sort().join(',')]);

  const normalizeClassCode = (raw: string) =>
    (raw || '')
      .toUpperCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/Đ/g, 'D')
      .replace(/[^A-Z0-9]/g, '')
      .slice(0, 15);

  const isValidClassCode = (code: string) => /^[A-Z0-9]{2,15}$/.test(code);

  const generateRandomClassCode = () => Math.random().toString(36).slice(2, 7).toUpperCase();

  // Check QR / URL hash on mount (supports any 2-15 char alphanumeric class code)
  useEffect(() => {
    const checkUrlRoom = () => {
      const hashCode = normalizeClassCode(window.location.hash.slice(1).trim());
      const searchCode = normalizeClassCode(new URLSearchParams(window.location.search).get('code')?.trim() || '');
      const code = isValidClassCode(hashCode) ? hashCode : (isValidClassCode(searchCode) ? searchCode : '');
      if (code) {
        setJoinCodeInput(code);
        setView('join');
        if (lastAutoJoinedRef.current !== code) {
          lastAutoJoinedRef.current = code;
          joinClassroom(code);
        }
      }
    };
    checkUrlRoom();
    window.addEventListener('hashchange', checkUrlRoom);
    return () => window.removeEventListener('hashchange', checkUrlRoom);
  }, []);

  // IndexedDB helpers (v2: 'l' for lessons, 'reports' for full session activities & stats)
  const openIdb = () => new Promise<IDBDatabase>((ok, no) => {
    const r = indexedDB.open('curivn', 2);
    r.onupgradeneeded = () => {
      const db = r.result;
      if (!db.objectStoreNames.contains('l')) db.createObjectStore('l', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('reports')) db.createObjectStore('reports', { keyPath: 'id' });
    };
    r.onsuccess = () => ok(r.result);
    r.onerror = no;
  });

  const hydrateLessonMedia = (l: Lesson): Lesson => {
    if (!l?.M) return l;
    const nextM: Record<string, Blob> = {};
    for (const [k, val] of Object.entries(l.M)) {
      if (val instanceof Blob) {
        nextM[k] = val;
      } else if (val && (val as any)._ab) {
        nextM[k] = new Blob([(val as any)._ab], { type: (val as any).type || 'image/png' });
      }
    }
    return { ...l, M: nextM };
  };

  const isBogusPlaceholderLesson = (l: Lesson): boolean => {
    if (!l) return true;
    if ((l.title || '').trim().length > 0) return false;
    const slides = l.slides || [];
    if (slides.length > 1) return false;
    if (slides.length === 0) return true;
    const s = slides[0];
    if ((s.b || '').trim().length > 0) return false;
    if ((s.imgs && s.imgs.length > 0) || s.vid || s.pptx) return false;
    const q = (s.q || '').trim();
    return q === '' || q === 'Em đã biết gì về chủ đề hôm nay?';
  };

  const dbAll = async () => {
    try {
      const list = await cloudFetchAllLessons();
      return (list || []).map((item: any) => hydrateLessonMedia(item));
    } catch {
      return [];
    }
  };

  const dbPut = async (l: Lesson) => {
    try {
      await cloudSaveLesson(l);
    } catch {}
  };

  const dbDel = async (id: string) => {
    try {
      await cloudDeleteLesson(id);
    } catch {}
  };

  const reportsAll = async () => {
    try {
      const list = await cloudFetchAllReports();
      return list || [];
    } catch {
      return [];
    }
  };

  const reportsPut = async (report: ClassSessionReport, skipCloud = false) => {
    try {
      await cloudSaveReport(report);
    } catch {}
  };

  const reportsDel = async (id: string) => {
    try {
      await cloudDeleteReport(id);
    } catch {}
  };

  const saveReportsSummaryFallback = (reports: ClassSessionReport[]) => {
    try {
      // Save lightweight summary fallback to localStorage without giant drawing data URLs
      const lite = reports.slice(0, 20).map(r => ({
        ...r,
        slides: r.slides.map(sl => ({
          ...sl,
          responses: sl.responses?.map(res => ({
            ...res,
            value: typeof res.value === 'string' && res.value.startsWith('data:image/') ? '[Bản vẽ]' : res.value,
          })),
        })),
      }));
      localStorage.setItem('curivn_session_reports', JSON.stringify(lite));
    } catch {}
  };

  // Load saved lessons & session reports on mount (migrating localStorage reports if needed)
  useEffect(() => {
    let active = true;
    (async () => {
      const [savedLessons, idbReports] = await Promise.all([dbAll(), reportsAll()]);
      if (!active) return;
      setMine(savedLessons);

      const mergedMap = new Map<string, ClassSessionReport>();
      idbReports.forEach(r => mergedMap.set(r.id, r));
      try {
        const lsReports = JSON.parse(localStorage.getItem('curivn_session_reports') || '[]');
        if (Array.isArray(lsReports)) {
          for (const r of lsReports) {
            if (r?.id && !mergedMap.has(r.id)) {
              mergedMap.set(r.id, r);
              reportsPut(r);
            }
          }
        }
      } catch {}
      const merged = Array.from(mergedMap.values()).sort((a, b) =>
        (b.updatedAt || b.id).localeCompare(a.updatedAt || a.id)
      );
      if (active) setSessionReports(merged);
    })();
    return () => { active = false; };
  }, []);

  const syncWithCloudNow = async (showNotice = true) => {
    if (!getActiveCloudOwnerId()) return;
    setCloudSyncing(true);
    if (showNotice) setCloudSyncMsg('Đang đồng bộ dữ liệu đám mây...');
    try {
      const [localLessons, localReps] = await Promise.all([dbAll(), reportsAll()]);
      const { lessons: mergedLessons, reports: mergedReps } = await cloudSyncAll(
        localLessons,
        localReps,
        async l => { await dbPut(l); },
        async r => { await reportsPut(r, true); }
      );
      setMine(mergedLessons);
      setSessionReports(mergedReps);
      saveReportsSummaryFallback(mergedReps);
      if (showNotice) {
        setCloudSyncMsg(`Đã đồng bộ đám mây (${mergedLessons.length} bài giảng, ${mergedReps.length} thống kê).`);
      }
    } catch (err: any) {
      if (showNotice) {
        setCloudSyncMsg('Đồng bộ đám mây tạm gián đoạn, dữ liệu trên máy vẫn an toàn.');
      }
    } finally {
      setCloudSyncing(false);
    }
  };

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, u => {
      setCloudUser(u);
      if (u && u.email && !getSavedSyncKey()) {
        const norm = setSavedSyncKey(u.email);
        setCloudSyncKey(norm);
        setCloudKeyInput(norm);
      }
      if (u || getSavedSyncKey()) {
        syncWithCloudNow(true);
      }
    });
    if (getSavedSyncKey()) {
      syncWithCloudNow(true);
    }
    return () => unsub();
  }, []);

  const handleSaveCloudKeyAndSync = async () => {
    const norm = setSavedSyncKey(cloudKeyInput);
    if (!norm) {
      setCloudSyncMsg('Vui lòng nhập Mã đồng bộ hoặc Email (tối thiểu 3 ký tự chữ/số, ví dụ: quangminh1485).');
      return;
    }
    setCloudSyncKey(norm);
    setCloudKeyInput(norm);
    await syncWithCloudNow(true);
    setShowCloudModal(false);
  };

  const handleCloudLogin = async () => {
    try {
      setCloudSyncing(true);
      const cred = await signInWithPopup(auth, googleProvider);
      if (cred?.user?.email) {
        const norm = setSavedSyncKey(cred.user.email);
        setCloudSyncKey(norm);
        setCloudKeyInput(norm);
      }
      await syncWithCloudNow(true);
      setShowCloudModal(false);
    } catch (err: any) {
      const msg = String(err?.message || err?.code || '');
      if (msg.includes('unauthorized-domain') || msg.includes('popup')) {
        setCloudSyncMsg(
          'Tên miền hiện tại (Vercel/TV) chặn cửa sổ Google Popup. Hãy nhập trực tiếp Email hoặc Mã đồng bộ của bạn ở ô bên trên để đồng bộ ngay!'
        );
      } else {
        setCloudSyncMsg('Hãy nhập Mã đồng bộ / Email ở ô bên trên để đồng bộ trực tiếp.');
      }
    } finally {
      setCloudSyncing(false);
    }
  };

  const handleCloudLogout = async () => {
    try {
      setSavedSyncKey('');
      setCloudSyncKey('');
      setCloudKeyInput('');
      await fbSignOut(auth);
      setCloudUser(null);
      setCloudSyncMsg('Đã ngắt kết nối đồng bộ đám mây trên thiết bị này.');
    } catch {}
  };

  const handleExportBackupFile = async () => {
    try {
      setCloudSyncing(true);
      const [localLessons, localReps] = await Promise.all([dbAll(), reportsAll()]);
      await exportBackupJson(localLessons, localReps);
      setCloudSyncMsg('Đã tải file sao lưu toàn bộ bài giảng & thống kê về máy.');
    } catch (e: any) {
      setCloudSyncMsg('Lỗi xuất file sao lưu: ' + (e?.message || ''));
    } finally {
      setCloudSyncing(false);
    }
  };

  const handleImportBackupFile = async (file: File) => {
    try {
      setCloudSyncing(true);
      const { lessons: impLessons, reports: impReps } = await importBackupJson(file);
      for (const l of impLessons) {
        if (l?.id) await dbPut(l);
      }
      for (const r of impReps) {
        if (r?.id) await reportsPut(r, false);
      }
      const [allLessons, allReps] = await Promise.all([dbAll(), reportsAll()]);
      setMine(allLessons);
      setSessionReports(allReps);
      saveReportsSummaryFallback(allReps);
      if (getActiveCloudOwnerId()) {
        await syncWithCloudNow(false);
      }
      setCloudSyncMsg(`Đã khôi phục thành công ${impLessons.length} bài giảng và ${impReps.length} thống kê!`);
      setShowCloudModal(false);
    } catch (e: any) {
      setCloudSyncMsg('File sao lưu không hợp lệ: ' + (e?.message || ''));
    } finally {
      setCloudSyncing(false);
    }
  };

  // Auto-save lesson whenever it has real content and changes
  const getLessonSignature = (l: Lesson) =>
    JSON.stringify({
      title: l.title,
      className: l.className || '',
      customCode: l.customCode || '',
      classCodes: l.classCodes || [],
      slides: l.slides,
      mediaKeys: Object.keys(l.M || {}),
    });

  const persistLessonNow = async (targetLesson: Lesson, showToast = false) => {
    if (!showToast && isBogusPlaceholderLesson(targetLesson)) {
      return targetLesson;
    }

    const hasMeaningfulContent =
      targetLesson.title.trim().length > 0 ||
      targetLesson.slides.length > 1 ||
      targetLesson.slides.some(
        s =>
          ((s.q || '').trim().length > 0 && (s.q || '').trim() !== 'Em đã biết gì về chủ đề hôm nay?') ||
          (s.b || '').trim().length > 0 ||
          (s.imgs && s.imgs.length > 0) ||
          !!s.vid ||
          !!s.pptx
      );
    if (!hasMeaningfulContent && !showToast) return targetLesson;

    const assignedId = targetLesson.id || Math.random().toString(36).slice(2, 10);
    const lToSave: Lesson = {
      ...targetLesson,
      id: assignedId,
      updatedAt: new Date().toISOString(),
    };
    const sig = getLessonSignature(lToSave);
    lastSavedLessonSigRef.current = sig;
    await dbPut(lToSave);
    if (showToast && getActiveCloudOwnerId()) {
      try {
        await cloudSaveLesson(lToSave);
      } catch {}
    } else if (getActiveCloudOwnerId()) {
      // Background cloud sync debounced/deferred silently
      setTimeout(() => {
        cloudSaveLesson(lToSave).catch(() => {});
      }, 3000);
    }
    setMine(prev => {
      const idx = prev.findIndex(x => x.id === assignedId);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = lToSave;
        return next;
      }
      return [lToSave, ...prev];
    });
    if (!targetLesson.id) {
      setLesson(prev => (prev.id ? prev : { ...prev, id: assignedId, updatedAt: lToSave.updatedAt }));
      lessonRef.current = { ...lessonRef.current, id: assignedId, updatedAt: lToSave.updatedAt };
    }
    if (showToast) {
      setSaveMsg(
        getActiveCloudOwnerId()
          ? 'Đã lưu bài giảng và đồng bộ lên Đám mây (Firebase).'
          : 'Đã lưu bài giảng trên trình duyệt.'
      );
    }
    return lToSave;
  };

  useEffect(() => {
    if (view !== 'build' && view !== 'host') return;
    if (isBogusPlaceholderLesson(lesson)) return;
    const sig = getLessonSignature(lesson);
    if (sig === lastSavedLessonSigRef.current) return;
    const t = window.setTimeout(() => {
      persistLessonNow(lesson, false).catch(() => {});
    }, 600);
    return () => window.clearTimeout(t);
  }, [lesson, view]);

  const loadDashboard = async () => {
    const [allLessons, allReps] = await Promise.all([dbAll(), reportsAll()]);
    setMine(allLessons);
    if (allReps.length) {
      const sorted = [...allReps].sort((a, b) => (b.updatedAt || b.id).localeCompare(a.updatedAt || a.id));
      setSessionReports(sorted);
    }
    setView('dash');
    if (getActiveCloudOwnerId()) {
      syncWithCloudNow(false);
    }
  };

  const saveCurrentLesson = async () => {
    try {
      const saved = await persistLessonNow(lesson, true);
      setLesson(saved);
    } catch (e: any) {
      setSaveMsg('Không lưu được: ' + e.message);
    }
  };

  const deleteSavedLesson = async (index: number) => {
    const l = mine[index];
    if (l && l.id) {
      await dbDel(l.id);
      if (getActiveCloudOwnerId()) {
        cloudDeleteLesson(l.id).catch(() => {});
      }
      loadDashboard();
    }
  };

  const useTemplate = (index: number) => {
    const copy: Lesson = JSON.parse(JSON.stringify(TEMPLATES[index]));
    copy.id = Math.random().toString(36).slice(2, 10);
    lastSavedLessonSigRef.current = '';
    setLesson(copy);
    setSlideIdx(0);
    setView('build');
  };

  const openSavedLesson = (index: number, preferredCode?: string, startImmediately = false) => {
    const l = mine[index];
    if (!l) return;
    const cleanPreferred = preferredCode ? normalizeClassCode(preferredCode) : '';
    const mergedCodes = Array.from(
      new Set([
        ...(l.classCodes || []),
        ...(l.customCode ? [l.customCode] : []),
        ...(cleanPreferred && isValidClassCode(cleanPreferred) ? [cleanPreferred] : []),
      ].map(normalizeClassCode).filter(isValidClassCode))
    );
    const activeCode = cleanPreferred && isValidClassCode(cleanPreferred)
      ? cleanPreferred
      : (l.customCode && isValidClassCode(normalizeClassCode(l.customCode)) ? normalizeClassCode(l.customCode) : (mergedCodes[0] || ''));
    const nextClassName = cleanPreferred && isValidClassCode(cleanPreferred)
      ? `Lớp ${cleanPreferred}`
      : (l.className || (activeCode ? `Lớp ${activeCode}` : ''));

    const loadedLesson: Lesson = {
      ...l,
      customCode: activeCode,
      classCodes: mergedCodes,
      className: nextClassName,
      slides: l.slides.map(s => ({ ...s, opts: s.opts ? [...s.opts] : undefined })),
    };
    lastSavedLessonSigRef.current = getLessonSignature(loadedLesson);
    setLesson(loadedLesson);
    lessonRef.current = loadedLesson;
    setClassCodeDraft(activeCode);
    setClassCodeMsg('');
    setSlideIdx(0);
    if (startImmediately) {
      startClassroom(activeCode || undefined, loadedLesson);
    } else {
      setView('build');
    }
  };

  const newLesson = () => {
    const fresh: Lesson = { id: Math.random().toString(36).slice(2, 10), title: '', className: '', customCode: '', classCodes: [], slides: [{ type: 'open', q: '' }], M: {} };
    lastSavedLessonSigRef.current = getLessonSignature(fresh);
    setLesson(fresh);
    lessonRef.current = fresh;
    setClassCodeDraft('');
    setClassCodeMsg('');
    setSlideIdx(0);
    setSaveMsg('');
  };

  const addOrSelectClassCode = (rawInput?: string, syncClassName = true) => {
    const code = normalizeClassCode(rawInput !== undefined ? rawInput : classCodeDraft);
    if (!isValidClassCode(code)) {
      setClassCodeMsg('Mã lớp cần từ 2–15 ký tự chữ và số (VD: 11A1, 10A2, TOAN11, AB123).');
      return '';
    }
    setClassCodeDraft(code);
    setClassCodeMsg(`Đã chọn và lưu mã lớp ${code} cho bài giảng.`);
    setLesson(prev => {
      const existing = (prev.classCodes || []).map(normalizeClassCode).filter(isValidClassCode);
      const nextCodes = Array.from(new Set([code, ...existing]));
      const nextClassName = syncClassName && (!prev.className?.trim() || prev.className.trim().startsWith('Lớp '))
        ? `Lớp ${code}`
        : prev.className;
      const updated: Lesson = {
        ...prev,
        customCode: code,
        classCodes: nextCodes,
        className: nextClassName,
      };
      lessonRef.current = updated;
      return updated;
    });
    return code;
  };

  const removeClassCodeFromLesson = (codeToRemove: string) => {
    const target = normalizeClassCode(codeToRemove);
    setLesson(prev => {
      const nextCodes = (prev.classCodes || []).map(normalizeClassCode).filter(c => c !== target && isValidClassCode(c));
      const nextActive = prev.customCode === target ? (nextCodes[0] || '') : (prev.customCode || '');
      const updated: Lesson = {
        ...prev,
        customCode: nextActive,
        classCodes: nextCodes,
      };
      lessonRef.current = updated;
      setClassCodeDraft(nextActive);
      return updated;
    });
    setClassCodeMsg(`Đã xóa mã lớp ${target} khỏi danh sách mã của bài giảng.`);
  };

  // Server-side Gemini AI Call helper
  const callAi = async (
    prompt: string,
    opts: { json?: boolean; schemaType?: 'autoInsert' | 'lesson' | 'slidesArray'; img?: File | null | any } = {}
  ) => {
    let imgBase64: string | undefined;
    let imgMimeType: string | undefined;
    if (opts.img) {
      imgBase64 = await new Promise<string>(r => {
        const reader = new FileReader();
        reader.onload = () => r((reader.result as string).split(',')[1]);
        reader.readAsDataURL(opts.img);
      });
      imgMimeType = opts.img.type || 'image/jpeg';
    }

    const res = await fetch('/api/gemini/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        prompt,
        json: !!opts.json,
        schemaType: opts.schemaType,
        imgBase64,
        imgMimeType,
      }),
    });

    const payload = await res.json().catch(() => ({}));
    if (!res.ok) {
      throw new Error(payload?.error || `mã lỗi ${res.status}`);
    }

    if (opts.json) {
      if (payload?.data !== undefined) return payload.data;
      const cleaned = String(payload?.text || '').replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
      return JSON.parse(cleaned);
    }
    return payload?.text || '';
  };

  const handleAiError = (e: any, setMsg: (s: string) => void) => {
    setMsg('AI gặp lỗi: ' + (e?.message || 'Không xác định'));
  };

  // Generate Lesson by Topic
  const generateLessonByTopic = async () => {
    const t = topicInput.trim();
    if (!t) return;
    setGenMsg('AI đang soạn bài giảng...');
    try {
      const j = await callAi(`Soạn bài giảng tương tác tiếng Việt cho học sinh THPT về chủ đề: "${t}". Trả về JSON: {"title":string,"slides":[{"type":"open|poll|cloud|quiz|scale|info","q":string,"b":string (chỉ info),"opts":[string] (chỉ poll/quiz, 3-4 mục),"a":số thứ tự đúng bắt đầu từ 0 (chỉ quiz)}]}. Gồm 6 slide, đa dạng, bắt đầu bằng cloud hoặc open.`, { json: true, schemaType: 'lesson' });
      if (j && j.slides && j.slides.length) {
        setLesson(prev => ({ ...prev, title: j.title || t, slides: applyAiSlideDefaults(j.slides), M: {} }));
        setGenMsg('Đã soạn xong – các câu hỏi tương tác đã được gắn sẵn thời gian đếm ngược.');
      } else {
        throw new Error('Kết quả không hợp lệ');
      }
    } catch (e: any) {
      handleAiError(e, setGenMsg);
    }
  };

  const loadExternalScript = (url: string) => new Promise<void>((ok, no) => {
    if (document.querySelector(`script[src="${url}"]`)) {
      ok();
      return;
    }
    const script = document.createElement('script');
    script.src = url;
    script.onload = () => ok();
    script.onerror = no;
    document.head.appendChild(script);
  });

  // Handle Document / File processing (also supports .pptx & video by routing to full preservation import)
  const handleFileUpload = async (file: File) => {
    if (!file) return;
    const name = file.name.toLowerCase();
    if (name.endsWith('.pptx') || file.type.startsWith('video/') || /\.(mp4|webm|mov|m4v|ogv)$/i.test(name)) {
      await handlePresentationImport(file);
      setFileMsg(`Đã giữ nguyên toàn bộ nội dung, hình ảnh, video và hiệu ứng từ tệp ${file.name}.`);
      return;
    }
    setImportingImage(null);
    setFileMsg('Đang đọc tệp...');
    try {
      let t = '';
      if (file.type.startsWith('image/')) {
        setImportingImage(file);
        setFileMsg('Đã chọn ảnh: ' + file.name + ' – AI sẽ đọc nội dung ảnh (hoặc tải lên ở mục Nhập bài giảng để giữ nguyên ảnh).');
        return;
      }
      if (/\.(txt|md|csv)$/.test(name)) {
        t = await file.text();
      } else if (name.endsWith('.docx')) {
        await loadExternalScript('https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.4.2/mammoth.browser.min.js');
        t = (await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() })).value;
      } else if (name.endsWith('.pdf')) {
        await loadExternalScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        const d = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
        for (let i = 1; i <= Math.min(d.numPages, 30); i++) {
          const pg = await d.getPage(i);
          const ctn = await pg.getTextContent();
          t += ctn.items.map((z: any) => z.str).join(' ') + '\n';
        }
      } else {
        throw new Error('Định dạng không hỗ trợ');
      }
      if (!t.trim()) {
        setFileMsg('Không đọc được chữ (có thể là PDF scan – hãy dùng mục Nhập bài giảng ở trên để giữ nguyên hình từng trang).');
        return;
      }
      setDocTextSrc(t.trim());
      setFileMsg('Đã đọc: ' + file.name + ' (' + t.length + ' ký tự)');
    } catch (e: any) {
      setFileMsg('Không đọc được tệp này. Hãy thử .txt hoặc dán nội dung vào ô bên dưới.');
    }
  };

  const generateFromDoc = async () => {
    const text = docTextSrc.trim();
    if (!text && !importingImage) {
      setGenMsg('Hãy tải tệp lên hoặc dán nội dung trước.');
      return;
    }
    setGenMsg('AI đang soạn bài giảng từ tài liệu...');
    try {
      const j = await callAi(`Dựa trên tài liệu${importingImage ? ' (ảnh đính kèm)' : ''}${text ? ' sau' : ''}, soạn bài giảng tương tác tiếng Việt cho học sinh THPT. Trả về JSON: {"title":string,"slides":[{"type":"open|poll|cloud|quiz|scale|info","q":string,"b":string (chỉ info),"opts":[string] (chỉ poll/quiz),"a":số thứ tự đúng từ 0 (chỉ quiz)}]}. Gồm 7 slide đa dạng, bám sát nội dung tài liệu.\n\n${text.slice(0, 8000)}`, { json: true, schemaType: 'lesson', img: importingImage });
      if (j && j.slides && j.slides.length) {
        setLesson(prev => ({ ...prev, title: j.title || '', slides: applyAiSlideDefaults(j.slides), M: {} }));
        setGenMsg('Đã soạn xong – các câu hỏi tương tác đã được gắn sẵn thời gian đếm ngược.');
      } else {
        throw new Error('Kết quả không hợp lệ');
      }
    } catch (e: any) {
      handleAiError(e, setGenMsg);
    }
  };

  const generateFromTextType = async (type: string) => {
    const text = docTextSrc.trim();
    if (!text && !importingImage) return;
    setGenMsg('AI đang tạo...');
    try {
      const j = await callAi(`Từ văn bản sau, tạo 3 slide tiếng Việt kiểu "${type}". Trả về JSON mảng [{"type":"${type}","q":string${type === 'open' ? '' : ',"opts":[3-4 mục]'}${type === 'quiz' ? ',"a":số thứ tự đúng từ 0' : ''}}].\n\nVăn bản:\n${text.slice(0, 6000)}`, { json: true, schemaType: 'slidesArray', img: importingImage });
      if (Array.isArray(j)) {
        const enriched = applyAiSlideDefaults(j);
        setLesson(prev => ({ ...prev, slides: [...prev.slides, ...enriched] }));
        setGenMsg('Đã thêm slide có thời gian đếm ngược.');
      } else {
        throw new Error('Kết quả không hợp lệ');
      }
    } catch (e: any) {
      handleAiError(e, setGenMsg);
    }
  };

  // Import PowerPoint (.pptx), PDF (.pdf), Word (.docx), Video or Image files retaining 100% original images, videos, layouts & effects
  const handlePresentationImport = async (fileOrFiles: File | FileList) => {
    const files: File[] = fileOrFiles instanceof FileList ? Array.from(fileOrFiles) : [fileOrFiles];
    if (!files.length) return;
    const firstFile = files[0];
    const name = firstFile.name.toLowerCase();
    const M: Record<string, Blob> = { ...(lesson.M || {}) };
    setGenMsg('Đang đọc tệp và giữ nguyên toàn bộ ảnh, video, hiệu ứng...');
    const uid = () => Math.random().toString(36).slice(2, 10);
    const freshUrls: Record<string, string> = {};
    const putBlob = (b: Blob) => {
      const id = uid();
      M[id] = b;
      const objUrl = URL.createObjectURL(b);
      localMediaUrlsRef.current[id] = objUrl;
      repairedMediaIdsRef.current[id] = true;
      freshUrls[id] = objUrl;
      return id;
    };

    // Preserve PNG transparency, animated GIFs/WebPs, SVGs, and high-resolution slide graphics
    const shrinkImage = async (b: ArrayBuffer, ext: string) => {
      const mimeMap: Record<string, string> = {
        png: 'image/png',
        jpg: 'image/jpeg',
        jpeg: 'image/jpeg',
        jfif: 'image/jpeg',
        gif: 'image/gif',
        webp: 'image/webp',
        svg: 'image/svg+xml',
        bmp: 'image/bmp',
      };
      const mime = mimeMap[ext] || 'image/png';
      const blob = new Blob([b], { type: mime });
      // Never re-encode GIF, SVG, WebP, or any image under 1.8MB so transparency and original quality stay 100% intact
      if (ext === 'gif' || ext === 'svg' || ext === 'webp' || blob.size < 1800000) return blob;
      try {
        const bm = await createImageBitmap(blob);
        const scale = Math.min(1, 1920 / Math.max(bm.width, 1));
        if (scale >= 1) return blob;
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(bm.width * scale);
        canvas.height = Math.round(bm.height * scale);
        canvas.getContext('2d')?.drawImage(bm, 0, 0, canvas.width, canvas.height);
        const outMime = ext === 'png' ? 'image/png' : 'image/jpeg';
        return await new Promise<Blob>(r => canvas.toBlob(res => r(res || blob), outMime, outMime === 'image/png' ? undefined : 0.9));
      } catch (e) {
        return blob;
      }
    };

    try {
      const out: Slide[] = [];

      // If multiple image/video files were selected, or a single image/video file
      if (files.every(f => f.type.startsWith('image/') || f.type.startsWith('video/') || /\.(png|jpe?g|gif|webp|svg|bmp|mp4|webm|mov|m4v|ogv)$/i.test(f.name))) {
        for (const [idx, f] of files.entries()) {
          const isVid = f.type.startsWith('video/') || /\.(mp4|webm|mov|m4v|ogv)$/i.test(f.name);
          const mediaId = putBlob(f);
          const cleanTitle = f.name.replace(/\.\w+$/, '') || `Slide ${idx + 1}`;
          out.push({
            type: 'media',
            q: cleanTitle,
            b: '',
            imgs: isVid ? [] : [mediaId],
            vid: isVid ? mediaId : undefined,
            tx: cleanTitle,
            orig: true,
          });
        }
      } else if (name.endsWith('.pptx')) {
        await loadExternalScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
        const zip = await JSZip.loadAsync(await firstFile.arrayBuffer());
        const parsedSlides = await parsePptxFile(
          zip,
          putBlob,
          shrinkImage,
          (idx, total) => setGenMsg(`Đang xử lý slide ${idx + 1}/${total} (giữ nguyên toàn bộ ảnh, video, bảng biểu & hiệu ứng)...`)
        );
        for (const item of parsedSlides) {
          out.push({
            type: 'media',
            q: item.title,
            b: item.body,
            imgs: item.imgs,
            vid: item.vid,
            tx: item.tx,
            orig: true,
            pptx: item.pptx,
          });
        }
      } else if (name.endsWith('.pdf')) {
        await loadExternalScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        const pdfDoc = await pdfjsLib.getDocument({ data: await firstFile.arrayBuffer() }).promise;
        for (let i = 1; i <= Math.min(pdfDoc.numPages, 60); i++) {
          setGenMsg(`Đang xử lý trang ${i}/${pdfDoc.numPages} (giữ nguyên độ nét & bố cục gốc)...`);
          const pg = await pdfDoc.getPage(i);
          const vp = pg.getViewport({ scale: 2.0 });
          const canvas = document.createElement('canvas');
          canvas.width = vp.width;
          canvas.height = vp.height;
          await pg.render({ canvasContext: canvas.getContext('2d')!, viewport: vp }).promise;
          const tx = (await pg.getTextContent()).items.map((z: any) => z.str).join(' ');
          const canvasBlob = await new Promise<Blob>(r => canvas.toBlob(res => r(res || new Blob()), 'image/jpeg', 0.9));
          const pageImgId = putBlob(canvasBlob);
          const wPx = 960;
          const hPx = Math.max(360, Math.round(wPx * (vp.height / Math.max(1, vp.width))));
          out.push({
            type: 'media',
            q: tx.trim().slice(0, 70) || `Trang ${i}`,
            b: '',
            tx,
            imgs: [pageImgId],
            orig: true,
            pptx: {
              widthPx: wPx,
              heightPx: hPx,
              background: '#ffffff',
              bgImgId: pageImgId,
              elements: [],
              animations: [],
            },
          });
        }
      } else if (name.endsWith('.docx')) {
        await loadExternalScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
        const zip = await JSZip.loadAsync(await firstFile.arrayBuffer());
        const docXmlStr = (await zip.file('word/document.xml')?.async('string')) || '';
        const relsStr = (await zip.file('word/_rels/document.xml.rels')?.async('string')) || '';
        const docDom = new DOMParser().parseFromString(docXmlStr, 'text/xml');
        const relsDom = new DOMParser().parseFromString(relsStr, 'text/xml');

        const relMap: Record<string, string> = {};
        const relNodes = relsDom.getElementsByTagName('*');
        for (let i = 0; i < relNodes.length; i++) {
          if (relNodes[i].localName === 'Relationship') {
            const rId = relNodes[i].getAttribute('Id') || '';
            const target = relNodes[i].getAttribute('Target') || '';
            if (rId && target) {
              relMap[rId] = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^(\.\/)+/, '')}`;
            }
          }
        }

        const blocks: Array<{ text: string; imgs: string[] }> = [];
        let curText: string[] = [];
        let curImgs: string[] = [];

        const pNodes = docDom.getElementsByTagName('*');
        for (let i = 0; i < pNodes.length; i++) {
          const el = pNodes[i];
          if (el.localName !== 'p') continue;
          const tList: string[] = [];
          const pImgs: string[] = [];
          const sub = el.getElementsByTagName('*');
          for (let j = 0; j < sub.length; j++) {
            const sEl = sub[j];
            if (sEl.localName === 't' && sEl.textContent) {
              tList.push(sEl.textContent);
            } else if (sEl.localName === 'blip') {
              const rId = sEl.getAttribute('r:embed') || sEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || '';
              const zPath = rId ? relMap[rId] : '';
              if (zPath && zip.file(zPath)) {
                const ext = zPath.split('.').pop()?.toLowerCase() || 'png';
                const ab = await zip.file(zPath).async('arraybuffer');
                const imgBlob = await shrinkImage(ab, ext);
                pImgs.push(putBlob(imgBlob));
              }
            }
          }
          const line = tList.join('').trim();
          if (line) curText.push(line);
          if (pImgs.length) curImgs.push(...pImgs);

          if (curText.join('\n').length > 360 || (curImgs.length > 0 && curText.length >= 2)) {
            blocks.push({ text: curText.join('\n'), imgs: curImgs });
            curText = [];
            curImgs = [];
          }
        }
        if (curText.length > 0 || curImgs.length > 0) {
          blocks.push({ text: curText.join('\n'), imgs: curImgs });
        }

        blocks.forEach((b, idx) => {
          const lines = b.text.split('\n').filter(Boolean);
          out.push({
            type: 'media',
            q: lines[0]?.slice(0, 80) || `Phần ${idx + 1}`,
            b: lines.slice(1).join('\n'),
            imgs: b.imgs,
            tx: b.text,
            orig: true,
          });
        });
      } else {
        throw new Error('Unsupported');
      }

      if (!out.length) throw new Error('Empty');
      if (Object.keys(freshUrls).length > 0) {
        setMediaCache(prev => ({ ...prev, ...freshUrls }));
      }
      setLesson(prev => ({
        ...prev,
        title: prev.title || firstFile.name.replace(/\.\w+$/, ''),
        slides: out,
        M,
      }));
      setGenMsg(`Đã nhập ${out.length} slide (giữ nguyên toàn bộ hình ảnh, video, bố cục & hiệu ứng gốc).`);
    } catch (e: any) {
      setGenMsg('Không nhập được tệp này (hỗ trợ .pptx, .pdf, .docx, ảnh và video).');
    }
  };

  // AI Auto-Insert Interactive Slides (works for imported .pptx/.pdf and any existing lesson slides)
  const autoInsertInteractiveSlides = async (targetCountOverride?: number) => {
    const base = lesson.slides.filter(x => !x.auto);
    const hasContent = base.some(
      x => (x.q || '').trim() || (x.b || '').trim() || (x.tx || '').trim() || (x.imgs && x.imgs.length > 0) || !!x.vid
    );
    if (!base.length || !hasContent) {
      const msg = 'Hãy nhập bài giảng (.pptx/.pdf) hoặc thêm nội dung slide trước khi dùng AI chèn câu hỏi.';
      setAutoInsertMsg(msg);
      setGenMsg(msg);
      return;
    }

    const countToInsert = Math.max(1, Math.min(20, Number(targetCountOverride || aiQuestionCount || 4)));

    setIsAutoInserting(true);
    setAutoInsertMsg(`AI đang đọc nội dung từng slide và chọn chỗ chèn ${countToInsert} câu hỏi tương tác...`);
    setGenMsg(`AI đang đọc nội dung từng slide và chọn chỗ chèn ${countToInsert} câu hỏi tương tác...`);

    try {
      const list = base
        .map((x, i) => {
          const detail = (x.b || x.tx || (x.opts ? x.opts.join(' / ') : '') || '').slice(0, 300).replace(/\n/g, ' ');
          return `${i}: [${x.type}] ${x.q || `Slide ${i + 1}`} | ${detail}`;
        })
        .slice(0, 45)
        .join('\n');

      const raw = await callAi(
        `Bạn là trợ lý thiết kế bài giảng sư phạm tương tác thông minh và sáng tạo tiếng Việt.
Tên bài giảng: "${lesson.title || 'Bài giảng trên lớp'}".
Dưới đây là danh sách các slide hiện có (số thứ tự từ 0 đến ${base.length - 1}):
${list}

Các slide hiện có phải được GIỮ NGUYÊN. Hãy chọn vị trí phù hợp để chèn thêm đúng ${countToInsert} slide câu hỏi tương tác bám sát nội dung bài giảng.
Quy tắc cực kỳ quan trọng:
- "after": số nguyên từ -1 đến ${base.length - 1} (-1 là chèn vào đầu bài, i là chèn ngay sau slide thứ i).
- "type": phải sử dụng đa dạng các loại câu hỏi phong phú sau đây:
  1. "cloud" (Đám mây từ / từ khóa nhanh)
  2. "open" (Câu hỏi mở tự luận, thảo luận sâu, yêu cầu đặt "ai": true)
  3. "poll" (Bình chọn / Khảo sát ý kiến học sinh, bắt buộc có "opts" gồm 3–4 lựa chọn)
  4. "quiz" (Trắc nghiệm kiểm tra kiến thức đúng/sai, bắt buộc có "opts" gồm 3–4 lựa chọn và "a" là chỉ số đáp án đúng từ 0)
  5. "scale" (Thang điểm đánh giá từ 1 đến 5 mức độ hiểu bài / tự tin)
  6. "draw" (Vẽ sơ đồ tư duy hoặc minh họa trực quan nhanh)
- "q": câu hỏi ngắn gọn, rõ ràng, sâu sắc bằng tiếng Việt cho học sinh trả lời.
Hãy đảm bảo phân bổ đều các loại câu hỏi (cloud, open, poll, quiz, scale, draw) xuyên suốt bài giảng.`,
        { json: true, schemaType: 'autoInsert' }
      );

      const items = Array.isArray(raw) ? raw : Array.isArray(raw?.slides) ? raw.slides : Array.isArray(raw?.items) ? raw.items : null;
      if (!items) throw new Error('Dữ liệu AI trả về không đúng định dạng danh sách.');

      const validTypes: Slide['type'][] = ['cloud', 'open', 'poll', 'quiz', 'scale', 'draw'];
      const ins: Record<number, Slide[]> = {};
      let insertedCount = 0;

      items.forEach((x: any) => {
        if (!x || !x.q || !validTypes.includes(x.type)) return;
        const slideType = x.type as Slide['type'];
        const normalized: Slide = {
          type: slideType,
          q: String(x.q).trim(),
          auto: true,
          aiGenerated: true,
          duration: defaultAiTimerSec > 0 ? defaultAiTimerSec : undefined,
        };
        if (slideType === 'poll' || slideType === 'quiz') {
          const cleanOpts = Array.isArray(x.opts)
            ? x.opts.map((o: any) => String(o || '').trim()).filter(Boolean)
            : [];
          const resolvedOpts = cleanOpts.length >= 2 ? cleanOpts : ['Lựa chọn A', 'Lựa chọn B', 'Lựa chọn C', 'Lựa chọn D'];
          normalized.opts = resolvedOpts;
          if (slideType === 'quiz') {
            const ansIdx = Number.isFinite(Number(x.a)) ? Math.trunc(Number(x.a)) : 0;
            normalized.a = Math.max(0, Math.min(resolvedOpts.length - 1, ansIdx));
          }
        }
        if (slideType === 'open' && x.ai) {
          normalized.ai = true;
        }

        let afterIdx = Number(x.after);
        if (Number.isNaN(afterIdx)) afterIdx = base.length - 1;
        afterIdx = Math.max(-1, Math.min(base.length - 1, Math.trunc(afterIdx)));

        (ins[afterIdx] = ins[afterIdx] || []).push(normalized);
        insertedCount++;
      });

      if (!insertedCount) {
        throw new Error('AI không tạo được câu hỏi phù hợp.');
      }

      const r2: Slide[] = [...(ins[-1] || [])];
      base.forEach((x, i) => {
        r2.push(x);
        (ins[i] || []).forEach(y => r2.push(y));
      });

      setLesson(prev => ({ ...prev, slides: r2 }));
      const okMsg = `AI đã tự động chèn thành công ${insertedCount} câu hỏi tương tác đa dạng vào bài giảng.`;
      setAutoInsertMsg(okMsg);
      setGenMsg(okMsg);
    } catch (e: any) {
      handleAiError(e, setAutoInsertMsg);
      handleAiError(e, setGenMsg);
    } finally {
      setIsAutoInserting(false);
    }
  };

  const addSlideType = (t: Slide['type']) => {
    setLesson(prev => ({
      ...prev,
      slides: [
        ...prev.slides,
        {
          type: t,
          q: '',
          opts: t === 'poll' || t === 'quiz' ? ['Lựa chọn A', 'Lựa chọn B', 'Lựa chọn C'] : undefined,
          a: t === 'quiz' ? 0 : undefined
        }
      ]
    }));
  };

  const initTimerForSlide = (idx: number, autoStart: boolean) => {
    const s = lessonRef.current.slides[idx];
    const dur = s?.duration && s.duration > 0 ? Math.round(s.duration) : 0;
    if (dur > 0) {
      setHostTimerTotal(dur);
      hostTimerTotalRef.current = dur;
      setHostTimerLeft(dur);
      hostTimerLeftRef.current = dur;
      setHostTimerRunning(autoStart);
      hostTimerRunningRef.current = autoStart;
      return { remainingSec: dur, totalSec: dur, running: autoStart };
    } else {
      setHostTimerTotal(0);
      hostTimerTotalRef.current = 0;
      setHostTimerLeft(null);
      hostTimerLeftRef.current = null;
      setHostTimerRunning(false);
      hostTimerRunningRef.current = false;
      return { remainingSec: null, totalSec: 0, running: false };
    }
  };

  const broadcastTimer = (idx: number, remainingSec: number | null, totalSec: number, running: boolean) => {
    const msg = { t: 'timer', idx, remainingSec, totalSec, running };
    connectionsRef.current.forEach(c => {
      if (c.open) c.send(msg);
    });
  };

  const enterTeacherLiveScreen = () => {
    setTeacherStartedLive(true);
    teacherStartedLiveRef.current = true;
    setView('host');
    if (hostTimerLeftRef.current !== null && hostTimerLeftRef.current > 0 && !hostTimerRunningRef.current) {
      setHostTimerRunning(true);
      hostTimerRunningRef.current = true;
      broadcastTimer(slideIdxRef.current, hostTimerLeftRef.current, hostTimerTotalRef.current, true);
    }
    broadcastState(slideIdxRef.current);
  };

  const updateLiveSlideTimer = (nextDurationSec: number, autoStart = true) => {
    const curIdx = slideIdxRef.current;
    const cleanDur = Math.max(0, Math.round(nextDurationSec));
    setLesson(prev => {
      const nextSlides = prev.slides.map((sl, i) => (i === curIdx ? { ...sl, duration: cleanDur > 0 ? cleanDur : undefined } : sl));
      const updated = { ...prev, slides: nextSlides };
      lessonRef.current = updated;
      return updated;
    });
    if (cleanDur > 0) {
      setHostTimerTotal(cleanDur);
      hostTimerTotalRef.current = cleanDur;
      setHostTimerLeft(cleanDur);
      hostTimerLeftRef.current = cleanDur;
      setHostTimerRunning(autoStart);
      hostTimerRunningRef.current = autoStart;
      broadcastTimer(curIdx, cleanDur, cleanDur, autoStart);
    } else {
      setHostTimerTotal(0);
      hostTimerTotalRef.current = 0;
      setHostTimerLeft(null);
      hostTimerLeftRef.current = null;
      setHostTimerRunning(false);
      hostTimerRunningRef.current = false;
      broadcastTimer(curIdx, null, 0, false);
    }
  };

  const toggleLiveTimerPause = () => {
    if (hostTimerLeftRef.current === null || hostTimerLeftRef.current <= 0) return;
    const nextRun = !hostTimerRunningRef.current;
    setHostTimerRunning(nextRun);
    hostTimerRunningRef.current = nextRun;
    broadcastTimer(slideIdxRef.current, hostTimerLeftRef.current, hostTimerTotalRef.current, nextRun);
  };

  const addLiveTimerSeconds = (extraSec: number) => {
    const curLeft = hostTimerLeftRef.current ?? 0;
    const nextLeft = Math.max(0, curLeft + extraSec);
    const nextTotal = Math.max(hostTimerTotalRef.current, nextLeft);
    setHostTimerTotal(nextTotal);
    hostTimerTotalRef.current = nextTotal;
    setHostTimerLeft(nextLeft);
    hostTimerLeftRef.current = nextLeft;
    const shouldRun = nextLeft > 0;
    setHostTimerRunning(shouldRun);
    hostTimerRunningRef.current = shouldRun;
    broadcastTimer(slideIdxRef.current, nextLeft, nextTotal, shouldRun);
  };

  // Host timer tick effect
  useEffect(() => {
    if (view !== 'host' || !teacherStartedLive || !hostTimerRunning || hostTimerLeft === null) return;
    if (hostTimerLeft <= 0) {
      setHostTimerRunning(false);
      hostTimerRunningRef.current = false;
      setRevealedSlides(prev => ({ ...prev, [slideIdxRef.current]: true }));
      broadcastTimer(slideIdxRef.current, 0, hostTimerTotalRef.current, false);
      return;
    }
    const id = window.setInterval(() => {
      const cur = hostTimerLeftRef.current;
      if (cur === null || !hostTimerRunningRef.current) return;
      if (cur <= 1) {
        setHostTimerLeft(0);
        hostTimerLeftRef.current = 0;
        setHostTimerRunning(false);
        hostTimerRunningRef.current = false;
        setRevealedSlides(prev => ({ ...prev, [slideIdxRef.current]: true }));
        broadcastTimer(slideIdxRef.current, 0, hostTimerTotalRef.current, false);
      } else {
        const next = cur - 1;
        setHostTimerLeft(next);
        hostTimerLeftRef.current = next;
      }
    }, 1000);
    return () => window.clearInterval(id);
  }, [view, teacherStartedLive, hostTimerRunning, hostTimerLeft !== null]);

  // Student timer tick effect
  useEffect(() => {
    if (!studentConnected || !studentTimerRunning || studentTimerLeft === null) return;
    if (studentTimerLeft <= 0) {
      setStudentTimerRunning(false);
      return;
    }
    const id = window.setInterval(() => {
      setStudentTimerLeft(prev => {
        if (prev === null) return null;
        if (prev <= 1) {
          setStudentTimerRunning(false);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
    return () => window.clearInterval(id);
  }, [studentConnected, studentTimerRunning, studentTimerLeft !== null]);

  // Bind PeerJS Host for a given alphanumeric class code
  const bindHostPeer = (code: string) => {
    try {
      if (peerPeer) peerPeer.destroy();
    } catch (e) {}
    setConnections([]);
    connectionsRef.current = [];

    const peer = new Peer('curivn-' + code.toLowerCase());
    setPeerPeer(peer);

    peer.on('open', () => {
      setHostStatus(`Phòng mã ${code} đã sẵn sàng – đang chờ học sinh quét QR hoặc nhập mã...`);
    });
    peer.on('error', (e: any) => {
      if (e.type === 'unavailable-id') {
        peer.destroy();
        const fallbackCode = normalizeClassCode(`${code}${Math.floor(10 + Math.random() * 89)}`).slice(0, 15);
        setRoomCode(fallbackCode);
        setHostCodeDraft(fallbackCode);
        roomCodeRef.current = fallbackCode;
        const baseUrl = window.location.href.split('#')[0].split('?')[0];
        setRoomLink(`${baseUrl}#${fallbackCode}`);
        setHostStatus(`Mã ${code} đang bận, hệ thống tự chuyển sang ${fallbackCode}.`);
        bindHostPeer(fallbackCode);
      } else {
        setHostStatus('Lỗi kết nối: ' + e.type);
      }
    });
    peer.on('connection', (c: any) => {
      setConnections(prev => {
        const next = [...prev, c];
        connectionsRef.current = next;
        return next;
      });
      c.on('data', (d: any) => handleHostData(c, d));
      c.on('close', () => {
        setConnections(prev => {
          const next = prev.filter(x => x !== c);
          connectionsRef.current = next;
          return next;
        });
        if (c._sid) {
          setConnectedStudents(prev => {
            const next = { ...prev };
            delete next[c._sid];
            connectedStudentsRef.current = next;
            return next;
          });
        }
      });
    });
  };

  // Allow teacher to edit/switch the alphanumeric room code right in the Host Lobby
  const applyHostRoomCode = async (rawNewCode?: string) => {
    const cleanCode = normalizeClassCode(rawNewCode !== undefined ? rawNewCode : hostCodeDraft);
    if (!isValidClassCode(cleanCode)) {
      setHostStatus('Mã lớp cần từ 2–15 ký tự chữ hoặc số (VD: 11A1, 10A2, TOAN11).');
      return;
    }
    const prevCode = roomCodeRef.current;
    setRoomCode(cleanCode);
    setHostCodeDraft(cleanCode);
    roomCodeRef.current = cleanCode;
    setClassCodeDraft(cleanCode);

    const baseUrl = window.location.href.split('#')[0].split('?')[0];
    setRoomLink(`${baseUrl}#${cleanCode}`);

    const cur = lessonRef.current;
    const nextCodes = Array.from(
      new Set([cleanCode, ...(cur.classCodes || []).map(normalizeClassCode).filter(isValidClassCode)])
    );
    const shouldSyncClassName =
      !cur.className?.trim() ||
      cur.className.trim() === `Lớp ${prevCode}` ||
      cur.className.trim() === prevCode;
    const updatedLesson: Lesson = {
      ...cur,
      customCode: cleanCode,
      classCodes: nextCodes,
      className: shouldSyncClassName ? `Lớp ${cleanCode}` : cur.className,
    };
    setLesson(updatedLesson);
    lessonRef.current = updatedLesson;
    await persistLessonNow(updatedLesson, false);

    setConnectedStudents({});
    connectedStudentsRef.current = {};
    setHostStatus(`Đang chuyển sang mã lớp ${cleanCode}...`);
    bindHostPeer(cleanCode);
    syncLiveSessionReport('live');
  };

  // Teacher Host & PeerJS setup
  const startClassroom = async (overrideCode?: string | any, overrideLesson?: Lesson) => {
    const sourceLesson = overrideLesson || lessonRef.current || lesson;
    const validSlides = sourceLesson.slides.filter(s => (s.q || '').trim() || s.type === 'media');
    if (!validSlides.length) {
      alert('Hãy thêm ít nhất một slide có nội dung.');
      return;
    }

    const explicitCode = typeof overrideCode === 'string' ? normalizeClassCode(overrideCode) : '';
    const draftCode = normalizeClassCode(classCodeDraft);
    const customLessonCode = normalizeClassCode(sourceLesson.customCode || '');
    const firstSavedCode = (sourceLesson.classCodes || []).map(normalizeClassCode).find(isValidClassCode) || '';
    const inferredFromClassName = normalizeClassCode((sourceLesson.className || '').replace(/^lớp\s+/i, ''));

    const code =
      (isValidClassCode(explicitCode) && explicitCode) ||
      (isValidClassCode(draftCode) && draftCode) ||
      (isValidClassCode(customLessonCode) && customLessonCode) ||
      (isValidClassCode(firstSavedCode) && firstSavedCode) ||
      (isValidClassCode(inferredFromClassName) && inferredFromClassName) ||
      generateRandomClassCode();

    const mergedCodes = Array.from(
      new Set([
        code,
        ...(sourceLesson.classCodes || []).map(normalizeClassCode).filter(isValidClassCode),
      ])
    );
    const resolvedClassName = (sourceLesson.className || '').trim() || `Lớp ${code}`;

    const activeLesson = await persistLessonNow(
      {
        ...sourceLesson,
        slides: validSlides,
        customCode: code,
        classCodes: mergedCodes,
        className: resolvedClassName,
      },
      false
    );
    setLesson(activeLesson);
    lessonRef.current = activeLesson;
    setClassCodeDraft(code);

    setRoomCode(code);
    setHostCodeDraft(code);
    roomCodeRef.current = code;
    const sessionId = `${Date.now()}-${code}`;
    currentSessionIdRef.current = sessionId;
    currentSessionStartedAtRef.current = new Date().toLocaleString('vi-VN');

    setSlideIdx(0);
    slideIdxRef.current = 0;
    setHostAnimStep(0);
    hostAnimStepRef.current = 0;
    setResponses({});
    responsesRef.current = {};
    setAiFeedbackText('');
    setShowScoreboard(false);
    setTeacherStartedLive(false);
    teacherStartedLiveRef.current = false;
    setShowQrModal(false);
    setConnectedStudents({});
    connectedStudentsRef.current = {};
    setAllSessionStudents({});
    allSessionStudentsRef.current = {};
    setAllSessionStudentClasses({});
    allSessionStudentClassesRef.current = {};
    initTimerForSlide(0, false);

    // Immediately create and save a live session report so all lesson activities are tracked from start
    syncLiveSessionReport('live');

    const baseUrl = window.location.href.split('#')[0].split('?')[0];
    const fullLink = `${baseUrl}#${code}`;
    setRoomLink(fullLink);
    setView('host');

    bindHostPeer(code);
  };

  const broadcastState = (newIdx = slideIdxRef.current, timerOverride?: { remainingSec: number | null; totalSec: number; running: boolean }, animOverride?: number) => {
    const curLesson = lessonRef.current;
    const timer = timerOverride ?? {
      remainingSec: hostTimerLeftRef.current,
      totalSec: hostTimerTotalRef.current,
      running: hostTimerRunningRef.current,
    };
    const animStep = animOverride ?? hostAnimStepRef.current;
    const msg = {
      t: 'state',
      started: teacherStartedLiveRef.current,
      L: { title: curLesson.title, className: curLesson.className, slides: curLesson.slides },
      idx: newIdx,
      animStep,
      timer,
    };
    connectionsRef.current.forEach(c => {
      if (c.open) c.send(msg);
    });
  };

  const setLiveAnimStep = (nextStep: number) => {
    const cleanStep = Math.max(0, nextStep);
    setHostAnimStep(cleanStep);
    hostAnimStepRef.current = cleanStep;
    const msg = { t: 'anim', idx: slideIdxRef.current, step: cleanStep };
    connectionsRef.current.forEach(c => {
      if (c.open) c.send(msg);
    });
  };

  const calculateScores = (curLesson = lessonRef.current, curResponses = responsesRef.current) => {
    const t: Record<string, number> = {};
    const nm: Record<string, string> = {};
    const slides = curLesson?.slides || [];
    slides.forEach((s, i) => {
      const respMap = curResponses?.[i] || {};
      for (const [k, a] of Object.entries(respMap)) {
        if (!a) continue;
        nm[k] = a.n || k;
        let p = 0;
        if (s?.type === 'quiz') p = (a.v === s.a) ? 1 : 0;
        else if (a.g !== undefined && a.g !== null) p = Number(a.g) || 0;
        t[k] = (t[k] || 0) + p;
      }
    });
    return { t, nm };
  };

  const handleHostData = (c: any, d: any) => {
    if (!d) return;
    const curLesson = lessonRef.current;
    const curIdx = slideIdxRef.current;
    if (d.t === 'hi') {
      c._sid = d.s;
      const cleanName = String(d.n || 'Học sinh').slice(0, 40);
      const cleanCls = String(d.cls || '').trim().slice(0, 30);
      setConnectedStudents(prev => {
        const next = { ...prev, [d.s]: cleanName };
        connectedStudentsRef.current = next;
        setHostStatus(`Đã có ${Object.keys(next).length} học sinh truy cập lớp`);
        return next;
      });
      setAllSessionStudents(prev => {
        const next = { ...prev, [d.s]: cleanName };
        allSessionStudentsRef.current = next;
        return next;
      });
      if (cleanCls) {
        setAllSessionStudentClasses(prev => {
          const next = { ...prev, [d.s]: cleanCls };
          allSessionStudentClassesRef.current = next;
          return next;
        });
      }
      c.send({
        t: 'state',
        started: teacherStartedLiveRef.current,
        L: { title: curLesson.title, className: curLesson.className, slides: curLesson.slides },
        idx: curIdx,
        animStep: hostAnimStepRef.current,
        timer: {
          remainingSec: hostTimerLeftRef.current,
          totalSec: hostTimerTotalRef.current,
          running: hostTimerRunningRef.current,
        },
      });
      c.send({ t: 'scores', sc: calculateScores(curLesson, responsesRef.current).t });
      syncLiveSessionReport('live');
    } else if (d.t === 'ans') {
      const targetIdx = typeof d.i === 'number' ? d.i : curIdx;
      const cleanName = String(d.n || 'Học sinh').slice(0, 40);
      const cleanCls = String(d.cls || allSessionStudentClassesRef.current[d.s] || '').trim().slice(0, 30);
      setConnectedStudents(prev => {
        const next = { ...prev, [d.s]: cleanName };
        connectedStudentsRef.current = next;
        return next;
      });
      setAllSessionStudents(prev => {
        const next = { ...prev, [d.s]: cleanName };
        allSessionStudentsRef.current = next;
        return next;
      });
      if (cleanCls) {
        setAllSessionStudentClasses(prev => {
          const next = { ...prev, [d.s]: cleanCls };
          allSessionStudentClassesRef.current = next;
          return next;
        });
      }
      const nextResponses = {
        ...responsesRef.current,
        [targetIdx]: {
          ...(responsesRef.current[targetIdx] || {}),
          [d.s]: {
            n: cleanName,
            cls: cleanCls || undefined,
            v: d.v,
            ts: new Date().toLocaleTimeString('vi-VN'),
          },
        },
      };
      responsesRef.current = nextResponses;
      setResponses(nextResponses);

      const scores = calculateScores(curLesson, nextResponses).t;
      connectionsRef.current.forEach(conn => {
        if (conn.open) conn.send({ t: 'scores', sc: scores });
      });
      syncLiveSessionReport('live');
    } else if (d.t === 'name' || d.t === 'profile') {
      const cleanName = String(d.n || '').trim().slice(0, 40);
      const cleanCls = typeof d.cls === 'string' ? d.cls.trim().slice(0, 30) : undefined;
      if (cleanName) {
        setConnectedStudents(prev => {
          const next = { ...prev, [d.s]: cleanName };
          connectedStudentsRef.current = next;
          return next;
        });
        setAllSessionStudents(prev => {
          const next = { ...prev, [d.s]: cleanName };
          allSessionStudentsRef.current = next;
          return next;
        });
      }
      if (cleanCls !== undefined) {
        setAllSessionStudentClasses(prev => {
          const next = { ...prev };
          if (cleanCls) next[d.s] = cleanCls;
          else delete next[d.s];
          allSessionStudentClassesRef.current = next;
          return next;
        });
      }
      const nextResponses: typeof responsesRef.current = {};
      for (const [sKey, ansMap] of Object.entries(responsesRef.current)) {
        const idxKey = Number(sKey);
        if (ansMap[d.s]) {
          nextResponses[idxKey] = {
            ...ansMap,
            [d.s]: {
              ...ansMap[d.s],
              n: cleanName || ansMap[d.s].n,
              cls: cleanCls !== undefined ? (cleanCls || undefined) : ansMap[d.s].cls,
            },
          };
        } else {
          nextResponses[idxKey] = ansMap;
        }
      }
      responsesRef.current = nextResponses;
      setResponses(nextResponses);
      syncLiveSessionReport('live');
    } else if (d.t === 'vote') {
      const sIdx = typeof d.slideIdx === 'number' ? d.slideIdx : curIdx;
      const tSid = d.targetSid;
      const voterId = d.s || c._sid;
      if (tSid && voterId) {
        setStudentVotes(prev => {
          const slideMap = prev[sIdx] || {};
          return { ...prev, [sIdx]: { ...slideMap, [voterId]: tSid } };
        });
      }
    } else if (d.t === 'need') {
      const b = curLesson.M && curLesson.M[d.id];
      if (b) {
        b.arrayBuffer().then(buf => c.send({ t: 'media', id: d.id, mt: b.type, data: buf }));
      }
    }
  };

  const moveSlide = (delta: number) => {
    const nextIdx = Math.max(0, Math.min(lessonRef.current.slides.length - 1, slideIdxRef.current + delta));
    setSlideIdx(nextIdx);
    slideIdxRef.current = nextIdx;
    setHostAnimStep(0);
    hostAnimStepRef.current = 0;
    setShowScoreboard(false);
    setAiFeedbackText('');
    setVotingActive({});
    setStudentVotingActive(false);
    setStudentSubmissions([]);
    setStudentMyVote(null);
    const timerInfo = initTimerForSlide(nextIdx, teacherStartedLiveRef.current);
    const shouldReveal = timerInfo.remainingSec === null || timerInfo.remainingSec <= 0;
    setRevealedSlides(prev => ({ ...prev, [nextIdx]: shouldReveal }));
    const scores = calculateScores().t;
    connectionsRef.current.forEach(c => {
      if (c.open) c.send({ t: 'scores', sc: scores });
    });
    broadcastState(nextIdx, timerInfo, 0);
  };

  const gradeAnswer = (jIdx: number, grade: number) => {
    const curIdx = slideIdxRef.current;
    const prevList = responsesRef.current[curIdx] || {};
    const keys = Object.keys(prevList);
    const k = keys[jIdx];
    if (!k) return;
    const nextResponses = {
      ...responsesRef.current,
      [curIdx]: {
        ...prevList,
        [k]: { ...prevList[k], g: grade }
      }
    };
    setResponses(nextResponses);
    responsesRef.current = nextResponses;
    const scores = calculateScores(lessonRef.current, nextResponses).t;
    connectionsRef.current.forEach(c => {
      if (c.open) c.send({ t: 'scores', sc: scores });
    });
    syncLiveSessionReport('live');
  };

  const deleteAnswer = (jIdx: number) => {
    const curIdx = slideIdxRef.current;
    const prevList = { ...(responsesRef.current[curIdx] || {}) };
    const keys = Object.keys(prevList);
    const k = keys[jIdx];
    if (k) {
      delete prevList[k];
      const nextResponses = { ...responsesRef.current, [curIdx]: prevList };
      setResponses(nextResponses);
      responsesRef.current = nextResponses;
      syncLiveSessionReport('live');
    }
  };

  const requestAiFeedback = async () => {
    const r = Object.values(responses[slideIdx] || {});
    if (!r.length) return;
    setAiFeedbackText('AI đang phân tích...');
    try {
      const text = await callAi(`Câu hỏi: ${lesson.slides[slideIdx].q}\nCác câu trả lời của học sinh:\n${r.map(a => a.n + ': ' + String(a.v).slice(0, 300)).join('\n')}\nHãy viết bằng tiếng Việt, ngắn gọn: (1) điểm chung, (2) hiểu lầm phổ biến, (3) gợi ý một câu hỏi tiếp theo cho giáo viên.`);
      setAiFeedbackText(text);
    } catch (e: any) {
      setAiFeedbackText('Lỗi AI: ' + e.message);
    }
  };

  const buildSessionReport = (status: 'live' | 'ended' = 'ended'): ClassSessionReport => {
    const curLesson = lessonRef.current;
    const curResponses = responsesRef.current;
    const curCode = roomCodeRef.current;
    const { t: scoreMap, nm: scoreNames } = calculateScores(curLesson, curResponses);

    // Collect every unique student & className across session joins & responses
    const studentMap: Record<string, string> = {
      ...allSessionStudentsRef.current,
      ...connectedStudentsRef.current,
      ...scoreNames,
    };
    const studentClsMap: Record<string, string> = {
      ...allSessionStudentClassesRef.current,
    };
    for (const ansBySlide of Object.values(curResponses)) {
      for (const [sid, ans] of Object.entries(ansBySlide || {})) {
        if (!studentMap[sid]) studentMap[sid] = ans.n;
        if (ans.cls && !studentClsMap[sid]) studentClsMap[sid] = ans.cls;
      }
    }

    const interactiveIndices = curLesson.slides
      .map((s, idx) => ({ s, idx }))
      .filter(({ s }) => s.type !== 'info' && s.type !== 'media');

    const totalInteractiveSlides = interactiveIndices.length;
    const studentEntries = Object.entries(studentMap);
    const totalStudents = studentEntries.length;

    let totalInteractiveSubmissions = 0;
    const slides: SlideSessionStat[] = curLesson.slides.map((s, idx) => {
      const isInteractive = s.type !== 'info' && s.type !== 'media';
      const slideRespEntries = Object.entries(curResponses[idx] || {});
      const submittedCount = slideRespEntries.length;
      if (isInteractive) totalInteractiveSubmissions += submittedCount;
      const participationPct = isInteractive && totalStudents > 0
        ? Math.min(100, Math.round((submittedCount / totalStudents) * 100))
        : 0;
      return {
        idx,
        type: s.type,
        q: s.q || `Slide ${idx + 1}`,
        b: s.b,
        opts: s.opts,
        correctIdx: s.type === 'quiz' ? s.a : undefined,
        isInteractive,
        submittedCount,
        participationPct,
        isAi: !!(s.auto || s.aiGenerated),
        responses: slideRespEntries.map(([sid, ans]) => ({
          sid,
          name: studentMap[sid] || ans.n,
          className: studentClsMap[sid] || ans.cls,
          value: ans.v,
          grade: ans.g,
          submittedAt: ans.ts,
        })),
      };
    });

    const students: StudentSessionStat[] = studentEntries
      .map(([sid, name]) => {
        let answeredCount = 0;
        interactiveIndices.forEach(({ idx }) => {
          if (curResponses[idx] && curResponses[idx][sid] !== undefined) {
            answeredCount++;
          }
        });
        const participationPct = totalInteractiveSlides > 0 ? Math.min(100, Math.round((answeredCount / totalInteractiveSlides) * 100)) : 0;
        return {
          sid,
          name,
          className: studentClsMap[sid] || undefined,
          answeredCount,
          participationPct,
          score: scoreMap[sid] || 0,
        };
      })
      .sort((a, b) => b.participationPct - a.participationPct || b.score - a.score);

    const studentClassNames = Array.from(
      new Set(Object.values(studentClsMap).map(c => c.trim()).filter(Boolean))
    );
    const resolvedClassName =
      (curLesson.className || '').trim() ||
      (studentClassNames.length === 1 ? studentClassNames[0] : '') ||
      `Lớp ${curCode}`;

    const overallParticipationPct =
      totalStudents > 0 && totalInteractiveSlides > 0
        ? Math.min(100, Math.round((totalInteractiveSubmissions / (totalStudents * totalInteractiveSlides)) * 100))
        : 0;

    const nowIso = new Date().toISOString();
    const nowLocal = new Date().toLocaleString('vi-VN');
    const sessionId = currentSessionIdRef.current || `${Date.now()}-${curCode || Math.random().toString(36).slice(2, 6)}`;
    currentSessionIdRef.current = sessionId;

    return {
      id: sessionId,
      lessonId: curLesson.id,
      className: resolvedClassName,
      studentClassNames,
      lessonTitle: curLesson.title || 'Bài giảng trực tiếp',
      roomCode: curCode,
      startedAt: currentSessionStartedAtRef.current || nowLocal,
      updatedAt: nowIso,
      endedAt: status === 'ended' ? nowLocal : 'Đang diễn ra',
      status,
      totalStudents,
      totalInteractiveSlides,
      overallParticipationPct,
      students,
      slides,
    };
  };

  const syncLiveSessionReport = (status: 'live' | 'ended' = 'live') => {
    const report = buildSessionReport(status);
    setSessionReports(prev => {
      const idx = prev.findIndex(r => r.id === report.id);
      const next = idx >= 0
        ? prev.map((r, i) => (i === idx ? report : r))
        : [report, ...prev];
      saveReportsSummaryFallback(next);
      return next;
    });
    reportsPut(report);
    return report;
  };

  const endClassroom = (cancelOnly = false) => {
    const hadSessionActivity =
      teacherStartedLiveRef.current ||
      Object.keys(allSessionStudentsRef.current).length > 0 ||
      Object.keys(responsesRef.current).length > 0;

    let finalReport: ClassSessionReport | null = null;
    if (!cancelOnly || hadSessionActivity) {
      finalReport = syncLiveSessionReport('ended');
      setActiveReportId(finalReport.id);
      setStatsClassFilter(finalReport.className);
    } else if (currentSessionIdRef.current) {
      // Cancelled before any student joined: remove empty draft session
      const draftId = currentSessionIdRef.current;
      setSessionReports(prev => {
        const next = prev.filter(r => r.id !== draftId);
        saveReportsSummaryFallback(next);
        return next;
      });
      reportsDel(draftId);
    }

    currentSessionIdRef.current = null;
    try {
      if (peerPeer) peerPeer.destroy();
    } catch (e) {}
    setPeerPeer(null);
    setConnections([]);
    connectionsRef.current = [];
    setTeacherStartedLive(false);
    teacherStartedLiveRef.current = false;
    setShowQrModal(false);
    setConnectedStudents({});
    connectedStudentsRef.current = {};
    setHostTimerRunning(false);
    hostTimerRunningRef.current = false;
    setHostTimerLeft(null);
    hostTimerLeftRef.current = null;
    setIsFullscreen(false);
    hasNativeFullscreenRef.current = false;
    try {
      if (document.fullscreenElement) {
        document.exitFullscreen().catch(() => {});
      }
    } catch {}
    setView(finalReport ? 'stats' : 'build');
  };

  // Student joining: Step 1 (QR/Code) -> Step 2 (Enter Name & Class) -> Step 3 (Wait for teacher to start)
  const joinClassroom = (overrideCode?: string) => {
    const c = normalizeClassCode(overrideCode ?? joinCodeInput);
    if (!isValidClassCode(c)) {
      setJoinMsg('Hãy nhập mã lớp gồm chữ hoặc số (từ 2 đến 15 ký tự, VD: 11A1, 10A2, AB123).');
      return;
    }
    setJoinCodeInput(c);
    lastAutoJoinedRef.current = c;
    setStudentRoomJoined(true);
    setStudentProfileConfirmed(false);
    studentProfileConfirmedRef.current = false;
    setStudentProfileError('');
    setJoinMsg(`Đang kết nối lớp ${c}...`);
    setStudentConnected(false);
    setStudentCurState(null);
    setStudentMyAns({});
    setStudentMyScore(null);
    askedMediaRef.current = {};

    try {
      if (studentPeerRef.current) studentPeerRef.current.destroy();
    } catch (e) {}

    const p = new Peer();
    setStudentPeer(p);
    studentPeerRef.current = p;

    p.on('error', (err: any) => {
      setJoinMsg(err.type === 'peer-unavailable' ? `Không tìm thấy lớp "${c}". Hãy kiểm tra lại mã lớp.` : 'Lỗi kết nối: ' + err.type);
    });
    p.on('open', () => hookStudentConnection(p, c));
  };

  const confirmStudentProfile = () => {
    const trimmedName = studentNameRef.current.trim();
    const trimmedCls = studentClassNameRef.current.trim();
    if (!trimmedName || !trimmedCls) {
      setStudentProfileError('Vui lòng nhập đầy đủ Họ và tên cùng Tên lớp của em trước khi vào lớp.');
      return;
    }
    setStudentProfileError('');
    saveStudentProfileLocal(trimmedName, trimmedCls);
    setStudentProfileConfirmed(true);
    studentProfileConfirmedRef.current = true;
    if (studentConnRef.current && studentConnRef.current.open) {
      studentConnRef.current.send({
        t: 'hi',
        s: studentSid,
        n: trimmedName,
        cls: trimmedCls,
      });
    }
  };

  const leaveStudentClassroom = () => {
    try {
      if (studentPeerRef.current) studentPeerRef.current.destroy();
    } catch (e) {}
    setStudentPeer(null);
    studentPeerRef.current = null;
    setStudentConn(null);
    studentConnRef.current = null;
    setStudentRoomJoined(false);
    setStudentProfileConfirmed(false);
    studentProfileConfirmedRef.current = false;
    setStudentProfileError('');
    setStudentConnected(false);
    setStudentCurState(null);
    setStudentMyAns({});
    setStudentMyScore(null);
    setJoinCodeInput('');
    setJoinMsg('');
    lastAutoJoinedRef.current = '';
    if (window.location.hash) {
      window.history.replaceState(null, '', window.location.pathname + window.location.search);
    }
  };

  const hookStudentConnection = (p: any, code: string) => {
    const conn = p.connect('curivn-' + code.toLowerCase(), { reliable: true });
    setStudentConn(conn);
    studentConnRef.current = conn;

    conn.on('open', () => {
      const displayName = studentNameRef.current.trim() || `Học sinh ${studentSid.slice(0, 3).toUpperCase()}`;
      const displayCls = studentClassNameRef.current.trim();
      conn.send({ t: 'hi', s: studentSid, n: displayName, cls: displayCls });
      setJoinMsg('');
      setStudentConnected(true);
    });
    conn.on('data', (d: any) => {
      if (!d) return;
      if (d.t === 'state') {
        setStudentCurState(d);
        setStudentConnected(true);
        setJoinMsg('');
        setStudentAnimStep(typeof d.animStep === 'number' ? d.animStep : 0);
        if (d.timer && d.timer.remainingSec !== null && d.timer.remainingSec !== undefined) {
          setStudentTimerTotal(Number(d.timer.totalSec) || Number(d.timer.remainingSec) || 0);
          setStudentTimerLeft(Number(d.timer.remainingSec));
          setStudentTimerRunning(!!d.timer.running);
        } else {
          setStudentTimerTotal(0);
          setStudentTimerLeft(null);
          setStudentTimerRunning(false);
        }
      } else if (d.t === 'anim') {
        if (typeof d.step === 'number') setStudentAnimStep(d.step);
      } else if (d.t === 'timer') {
        if (d.remainingSec !== null && d.remainingSec !== undefined) {
          setStudentTimerTotal(Number(d.totalSec) || Number(d.remainingSec) || 0);
          setStudentTimerLeft(Number(d.remainingSec));
          setStudentTimerRunning(!!d.running);
        } else {
          setStudentTimerTotal(0);
          setStudentTimerLeft(null);
          setStudentTimerRunning(false);
        }
      } else if (d.t === 'scores') {
        setStudentMyScore((d.sc && d.sc[studentSid]) || 0);
      } else if (d.t === 'voting_start') {
        setStudentVotingActive(true);
        setStudentSubmissions(Array.isArray(d.submissions) ? d.submissions : []);
        setStudentMyVote(null);
      } else if (d.t === 'media') {
        setMediaCache(prev => ({ ...prev, [d.id]: URL.createObjectURL(new Blob([d.data], { type: d.mt })) }));
      }
    });

    conn.on('close', () => {
      setJoinMsg('Mất kết nối – đang thử kết nối lại...');
      askedMediaRef.current = {};
      setTimeout(() => {
        if (p && !p.destroyed && studentPeerRef.current === p) hookStudentConnection(p, code);
      }, 3000);
    });
  };

  const sendStudentVote = (targetSid: string) => {
    if (!studentConnRef.current || !studentConnRef.current.open) return;
    const curIdx = studentCurState?.idx ?? slideIdxRef.current;
    studentConnRef.current.send({
      t: 'vote',
      slideIdx: curIdx,
      targetSid,
      s: studentSid,
    });
    setStudentMyVote(targetSid);
  };

  const saveStudentProfileLocal = (name: string, className: string) => {
    try {
      localStorage.setItem('curivn_student_profile', JSON.stringify({ name, className }));
    } catch {}
  };

  const handleStudentNameChange = (newName: string) => {
    setStudentName(newName);
    studentNameRef.current = newName;
    if (studentProfileError) setStudentProfileError('');
    saveStudentProfileLocal(newName, studentClassNameRef.current);
    const trimmed = newName.trim();
    if (studentProfileConfirmedRef.current && studentConnRef.current && studentConnRef.current.open) {
      studentConnRef.current.send({
        t: 'profile',
        s: studentSid,
        n: trimmed || `Học sinh ${studentSid.slice(0, 3).toUpperCase()}`,
        cls: studentClassNameRef.current.trim(),
      });
    }
  };

  const handleStudentClassChange = (newClassName: string) => {
    setStudentClassName(newClassName);
    studentClassNameRef.current = newClassName;
    if (studentProfileError) setStudentProfileError('');
    saveStudentProfileLocal(studentNameRef.current, newClassName);
    const trimmedName = studentNameRef.current.trim() || `Học sinh ${studentSid.slice(0, 3).toUpperCase()}`;
    if (studentProfileConfirmedRef.current && studentConnRef.current && studentConnRef.current.open) {
      studentConnRef.current.send({
        t: 'profile',
        s: studentSid,
        n: trimmedName,
        cls: newClassName.trim(),
      });
    }
  };

  const submitStudentAnswer = (slideIndex: number, val: any) => {
    if (studentTimerLeftRef.current === 0) return;
    if (val === '' || studentMyAns[slideIndex] !== undefined || !studentConnRef.current || !studentConnRef.current.open) return;
    const displayName = studentNameRef.current.trim() || `Học sinh ${studentSid.slice(0, 3).toUpperCase()}`;
    const displayCls = studentClassNameRef.current.trim();
    setStudentMyAns(prev => ({ ...prev, [slideIndex]: val }));
    studentConnRef.current.send({ t: 'ans', i: slideIndex, s: studentSid, n: displayName, cls: displayCls, v: val });
  };

  // Media URL resolver (requests all images, videos, and audio from host over PeerJS)
  useEffect(() => {
    if (!studentCurState || !studentConnRef.current?.open) return;
    const s: Slide | undefined = studentCurState.L?.slides?.[studentCurState.idx];
    if (!s) return;
    const pptxMediaIds = s.pptx
      ? [
          ...(s.pptx.bgImgId ? [s.pptx.bgImgId] : []),
          ...(s.pptx.audioIds || []),
          ...s.pptx.elements.flatMap(el => [el.imgId, el.vidId, el.audioId].filter((x): x is string => Boolean(x))),
        ]
      : [];
    const ids = Array.from(new Set([...(s.imgs || []), ...(s.vid ? [s.vid] : []), ...pptxMediaIds]));
    ids.forEach(id => {
      if (!mediaCache[id] && !askedMediaRef.current[id]) {
        askedMediaRef.current[id] = 1;
        studentConnRef.current.send({ t: 'need', id });
      }
    });
  }, [studentCurState?.idx, studentConnected]);

  const resolveMediaUrl = (id: string) => {
    if (!id) return '';
    if (mediaCache[id]) return mediaCache[id];
    if (localMediaUrlsRef.current[id]) return localMediaUrlsRef.current[id];
    const raw = lesson.M?.[id] || lessonRef.current?.M?.[id];
    if (raw) {
      const blob =
        raw instanceof Blob
          ? raw
          : (raw as any)?._ab
          ? new Blob([(raw as any)._ab], { type: (raw as any).type || 'image/png' })
          : null;
      if (blob && blob.size > 0) {
        const u = URL.createObjectURL(blob);
        localMediaUrlsRef.current[id] = u;
        return u;
      }
    }
    return '';
  };

  // Render Media HTML (preserves PowerPoint layouts, transitions & animations, and auto-fits any TV screen)
  const renderMediaSlide = (s: Slide, isHostView = false, keyPrefix: string | number = slideIdx) => {
    if (s.pptx) {
      return (
        <PptxSlideView
          pptx={s.pptx}
          slideKey={`${keyPrefix}-${s.q || ''}`}
          resolveMediaUrl={resolveMediaUrl}
          animStep={isHostView ? hostAnimStep : studentAnimStep}
          onSetAnimStep={isHostView ? setLiveAnimStep : undefined}
          showControls={isHostView}
          isFullscreen={isHostView && isFullscreen}
          fitViewport={true}
        />
      );
    }

    const hasText = Boolean(s.b && s.b.trim());
    const imgCount = s.imgs?.length || 0;
    const hasMedia = imgCount > 0 || Boolean(s.vid);
    const maxMediaH = isHostView && isFullscreen
      ? (hasText ? 'max-h-[62vh]' : 'max-h-[78vh]')
      : (hasText ? 'max-h-[52vh]' : 'max-h-[70vh]');
    const textSizeClass =
      (s.b?.length || 0) > 450
        ? 'text-sm sm:text-base'
        : (s.b?.length || 0) > 220
        ? 'text-base sm:text-lg'
        : 'text-lg sm:text-xl';

    return (
      <div className={hasText && hasMedia ? 'grid grid-cols-1 lg:grid-cols-12 gap-4 items-center' : 'space-y-3'}>
        {hasText && (
          <div className={hasMedia ? 'lg:col-span-5' : 'w-full'}>
            <p className={`${textSizeClass} whitespace-pre-line leading-relaxed break-words`}>{s.b}</p>
          </div>
        )}
        {hasMedia && (
          <div className={`${hasText ? 'lg:col-span-7' : 'w-full'} space-y-3`}>
            {imgCount > 0 && (
              <div className={imgCount > 1 ? 'grid grid-cols-1 sm:grid-cols-2 gap-3 items-center' : 'flex justify-center'}>
                {s.imgs?.map((id, idx) => {
                  const u = resolveMediaUrl(id);
                  return u ? (
                    <img
                      key={idx}
                      src={u}
                      alt=""
                      onError={e => {
                        (e.currentTarget as HTMLImageElement).style.display = 'none';
                      }}
                      className={`w-auto max-w-full ${imgCount > 1 ? 'max-h-[44vh]' : maxMediaH} rounded-xl mx-auto shadow-md object-contain bg-white`}
                    />
                  ) : (
                    <p key={idx} className="text-gray-500 italic text-center">Đang tải hình ảnh...</p>
                  );
                })}
              </div>
            )}
            {s.vid && (
              resolveMediaUrl(s.vid) ? (
                <video
                  src={resolveMediaUrl(s.vid)}
                  controls
                  playsInline
                  className={`w-auto max-w-full ${maxMediaH} rounded-xl mx-auto shadow-md object-contain`}
                />
              ) : (
                <p className="text-gray-500 italic text-center">Đang tải video...</p>
              )
            )}
          </div>
        )}
      </div>
    );
  };

  // QR Code generator helper for host (usable in lobby & modal)
  useEffect(() => {
    if (!roomLink) {
      setQrDataUrl('');
      return;
    }
    try {
      if (typeof qrcode !== 'undefined') {
        const q = qrcode(0, 'M');
        q.addData(roomLink);
        q.make();
        setQrDataUrl(q.createDataURL(6, 3));
      } else {
        setQrDataUrl('');
      }
    } catch (e) {
      setQrDataUrl('');
    }
  }, [roomLink]);

  // Fullscreen toggle & native fullscreenchange sync for teacher presentation
  const toggleFullscreen = async () => {
    if (!isFullscreen) {
      setIsFullscreen(true);
      try {
        const targetEl = liveStageRef.current || document.documentElement;
        if (targetEl?.requestFullscreen && !document.fullscreenElement) {
          await targetEl.requestFullscreen();
          hasNativeFullscreenRef.current = true;
        }
      } catch {
        hasNativeFullscreenRef.current = false;
      }
    } else {
      setIsFullscreen(false);
      hasNativeFullscreenRef.current = false;
      try {
        if (document.fullscreenElement && document.exitFullscreen) {
          await document.exitFullscreen();
        }
      } catch {}
    }
  };

  useEffect(() => {
    const onFullscreenChange = () => {
      const active = !!document.fullscreenElement;
      if (active) {
        hasNativeFullscreenRef.current = true;
        setIsFullscreen(true);
      } else if (hasNativeFullscreenRef.current) {
        hasNativeFullscreenRef.current = false;
        setIsFullscreen(false);
      }
    };
    document.addEventListener('fullscreenchange', onFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', onFullscreenChange);
  }, []);

  // Keyboard navigation & fullscreen shortcuts during live teacher lesson mode
  useEffect(() => {
    if (view !== 'host' || !teacherStartedLive) return;
    const onKeyDown = (e: KeyboardEvent) => {
      const tag = (e.target as HTMLElement)?.tagName;
      if (tag === 'INPUT' || tag === 'TEXTAREA') return;
      const curSlide = lessonRef.current.slides[slideIdxRef.current];
      const maxAnim = getPptxMaxStep(curSlide?.pptx);

      if (e.key === 'ArrowRight' || e.key === ' ') {
        if (e.key === ' ') e.preventDefault();
        if (maxAnim > 0 && hostAnimStepRef.current < maxAnim) {
          setLiveAnimStep(hostAnimStepRef.current + 1);
        } else {
          moveSlide(1);
        }
      } else if (e.key === 'ArrowLeft') {
        if (maxAnim > 0 && hostAnimStepRef.current > 0) {
          setLiveAnimStep(hostAnimStepRef.current - 1);
        } else {
          moveSlide(-1);
        }
      } else if ((e.key === 'f' || e.key === 'F') && !e.ctrlKey && !e.metaKey) {
        e.preventDefault();
        toggleFullscreen();
      } else if (e.key === 'Escape' && isFullscreen && !document.fullscreenElement) {
        setIsFullscreen(false);
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [view, teacherStartedLive, isFullscreen]);

  const onlineStudentCount = Math.max(connections.filter(c => c.open).length, Object.keys(connectedStudents).length);
  const isTeacherLiveMode = view === 'host' && teacherStartedLive;
  const isStudentLiveMode = view === 'join' && studentRoomJoined && studentProfileConfirmed && studentConnected && !!studentCurState?.started;
  const isFullLessonMode = isTeacherLiveMode || isStudentLiveMode;

  return (
    <div className="min-h-screen bg-stone-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 flex flex-col font-sans selection:bg-purple-500 selection:text-white">
      {/* Top Navbar (hidden when teaching or attending live lesson so only the lesson screen is shown) */}
      {!isFullLessonMode && (
        <nav className="border-b border-stone-200 dark:border-slate-800 bg-white/80 dark:bg-slate-900/80 backdrop-blur sticky top-0 z-50 px-4 py-3">
          <div className="max-w-5xl mx-auto flex justify-between items-center flex-wrap gap-4">
            <div className="text-xl sm:text-2xl font-black tracking-tight text-purple-600 dark:text-purple-400 cursor-pointer flex items-center gap-2" onClick={() => setView('home')}>
              Lớp học <span className="text-slate-900 dark:text-white">trực tuyến</span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {peerPeer && roomCode && (
                <button
                  onClick={() => setView('host')}
                  className="px-3.5 py-2 text-xs font-bold rounded-xl bg-emerald-50 dark:bg-emerald-950/50 border border-emerald-200 dark:border-emerald-800 text-emerald-700 dark:text-emerald-300 hover:bg-emerald-100 transition flex items-center gap-1.5"
                >
                  <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                  Phòng {roomCode}: {onlineStudentCount} HS truy cập
                </button>
              )}
              <button onClick={() => setView('join')} className="px-4 py-2 text-sm font-semibold rounded-xl border border-stone-300 dark:border-slate-700 hover:bg-stone-100 dark:hover:bg-slate-800 transition">
                Tham gia lớp
              </button>
              <button onClick={() => setView('stats')} className="px-4 py-2 text-sm font-semibold rounded-xl border border-stone-300 dark:border-slate-700 hover:bg-stone-100 dark:hover:bg-slate-800 transition flex items-center gap-1.5">
                <BarChart2 className="w-4 h-4 text-purple-600" /> Thống kê lớp
                {sessionReports.length > 0 && (
                  <span className="px-1.5 py-0.5 text-[10px] font-bold bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 rounded-full">
                    {sessionReports.length}
                  </span>
                )}
              </button>
              <button onClick={loadDashboard} className="px-4 py-2 text-sm font-semibold rounded-xl bg-purple-50 dark:bg-purple-950/50 text-purple-600 dark:text-purple-300 hover:bg-purple-100 transition flex items-center gap-1.5">
                <BookOpen className="w-4 h-4" /> Bảng điều khiển
              </button>
              <button
                onClick={() => setShowCloudModal(true)}
                title="Cài đặt đồng bộ đám mây"
                className="px-3.5 py-2 text-sm font-semibold rounded-xl bg-purple-50 dark:bg-purple-950/50 border border-purple-200 dark:border-purple-800 text-purple-700 dark:text-purple-300 hover:bg-purple-100 transition flex items-center gap-1.5"
              >
                <Cloud className="w-4 h-4 text-purple-600" />
                <span>Đồng bộ đám mây</span>
              </button>
              <button onClick={() => { newLesson(); setView('build'); }} className="px-4 py-2 text-sm font-semibold rounded-xl bg-purple-600 text-white hover:bg-purple-700 shadow-sm transition flex items-center gap-1.5">
                <Plus className="w-4 h-4" /> Tạo bài giảng
              </button>
            </div>
          </div>
        </nav>
      )}

      {/* Main Content Area */}
      <main className={isFullLessonMode ? 'w-full flex-1 flex flex-col' : 'max-w-5xl w-full mx-auto px-4 py-6 flex-1'}>
        {/* HOME VIEW */}
        {view === 'home' && (
          <div className="space-y-12 py-8">
            <div className="text-center space-y-4 max-w-2xl mx-auto">
              <h1 className="text-4xl md:text-6xl font-extrabold tracking-tight leading-tight">
                Bài giảng tương tác mà học sinh yêu thích
              </h1>
              <p className="text-lg text-stone-600 dark:text-slate-400">
                Giáo viên dẫn dắt, cả lớp cùng đọc, viết và thảo luận trên điện thoại. Học sinh quét mã QR là vào lớp, không cần tài khoản rườm rà.
              </p>
              <div className="flex justify-center gap-3 pt-2 flex-wrap">
                <button onClick={loadDashboard} className="px-6 py-3 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl shadow-lg shadow-purple-500/20 transition flex items-center gap-2">
                  <Play className="w-5 h-5 fill-current" /> Bắt đầu ngay
                </button>
                <button onClick={() => setView('join')} className="px-6 py-3 bg-stone-100 dark:bg-slate-800 text-stone-700 dark:text-stone-300 hover:bg-stone-200 font-bold rounded-xl transition">
                  Tôi là học sinh
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 pt-6">
              <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-2">
                <div className="w-10 h-10 rounded-xl bg-purple-100 dark:bg-purple-950 flex items-center justify-center text-purple-600 font-bold mb-4">1</div>
                <h3 className="text-lg font-bold">Nhập hoặc tạo bài giảng</h3>
                <p className="text-sm text-stone-600 dark:text-slate-400">Nhập file .pptx/.pdf (giữ nguyên chữ, hình, video) hoặc để AI soạn từ chủ đề.</p>
              </div>
              <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-2">
                <div className="w-10 h-10 rounded-xl bg-purple-100 dark:bg-purple-950 flex items-center justify-center text-purple-600 font-bold mb-4">2</div>
                <h3 className="text-lg font-bold">AI chèn câu hỏi</h3>
                <p className="text-sm text-stone-600 dark:text-slate-400">AI đọc nội dung từng slide và tự động chèn đám mây từ, trắc nghiệm, câu hỏi mở, thang điểm.</p>
              </div>
              <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-2">
                <div className="w-10 h-10 rounded-xl bg-purple-100 dark:bg-purple-950 flex items-center justify-center text-purple-600 font-bold mb-4">3</div>
                <h3 className="text-lg font-bold">Học sinh quét QR</h3>
                <p className="text-sm text-stone-600 dark:text-slate-400">Slide hiện trên điện thoại, các em trả lời và kết quả hiển thị trực quan lên tivi.</p>
              </div>
            </div>
          </div>
        )}

        {/* DASHBOARD VIEW */}
        {view === 'dash' && (
          <div className="space-y-6 py-4">
            <div className="flex justify-between items-center flex-wrap gap-4">
              <div>
                <h2 className="text-2xl font-bold">Bảng điều khiển</h2>
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={() => { newLesson(); setView('build'); }} className="px-4 py-2 bg-stone-200 dark:bg-slate-800 font-semibold rounded-xl hover:bg-stone-300 transition flex items-center gap-2">
                  <Plus className="w-4 h-4" /> Bài giảng mới
                </button>
                <button onClick={() => { newLesson(); setView('build'); }} className="px-4 py-2 bg-purple-600 text-white font-semibold rounded-xl hover:bg-purple-700 shadow-sm transition flex items-center gap-2">
                  <Sparkles className="w-4 h-4" /> Tạo bằng AI
                </button>
              </div>
            </div>



            <div className="space-y-4">
              <h3 className="text-lg font-bold text-stone-700 dark:text-slate-300">Bài giảng của tôi</h3>
              {mine.length === 0 ? (
                <p className="text-stone-500 dark:text-slate-500 italic">Chưa có bài giảng nào được lưu trên trình duyệt.</p>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                  {mine.map((l, i) => {
                    const lessonReps = sessionReports.filter(
                      r => (l.id && r.lessonId === l.id) || (l.title && r.lessonTitle === l.title)
                    );
                    const avgPct = lessonReps.length
                      ? Math.round(lessonReps.reduce((acc, r) => acc + r.overallParticipationPct, 0) / lessonReps.length)
                      : null;
                    const savedClassCodes = Array.from(
                      new Set([
                        ...(l.classCodes || []),
                        ...(l.customCode ? [l.customCode] : []),
                        ...lessonReps.map(r => r.roomCode),
                      ].map(normalizeClassCode).filter(isValidClassCode))
                    );
                    return (
                      <div key={l.id || i} className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-5 rounded-2xl flex flex-col justify-between shadow-sm">
                        <div className="space-y-2">
                          <div className="flex items-center justify-between gap-2 flex-wrap">
                            <h4 className="font-bold text-lg">{l.title || 'Chưa đặt tên'}</h4>
                            {l.className && (
                              <span className="px-2 py-0.5 bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 text-[11px] font-bold rounded-full">
                                {l.className}
                              </span>
                            )}
                          </div>
                          <p className="text-sm text-stone-500 dark:text-slate-400">
                            {l.slides.length} slide · {l.slides.filter(s => s.type !== 'info' && s.type !== 'media').length} câu hỏi tương tác
                          </p>
                          {/* Multi-class codes stored on this lesson */}
                          <div className="pt-1 space-y-1">
                            <div className="text-[11px] font-semibold text-stone-500 dark:text-slate-400">
                              Mã lớp đã lưu ({savedClassCodes.length} lớp – bấm mã để dạy ngay):
                            </div>
                            <div className="flex items-center gap-1.5 flex-wrap">
                              {savedClassCodes.map(code => (
                                <button
                                  key={code}
                                  type="button"
                                  onClick={() => openSavedLesson(i, code, true)}
                                  title={`Dạy bài giảng này với mã lớp ${code}`}
                                  className="px-2.5 py-0.5 rounded-lg bg-emerald-50 dark:bg-emerald-950/60 border border-emerald-200 dark:border-emerald-800 text-emerald-700 dark:text-emerald-300 font-mono text-xs font-extrabold hover:bg-emerald-100 transition inline-flex items-center gap-1"
                                >
                                  <Play className="w-2.5 h-2.5 fill-current" /> {code}
                                </button>
                              ))}
                              <button
                                type="button"
                                onClick={() => openSavedLesson(i)}
                                className="px-2 py-0.5 rounded-lg bg-stone-100 dark:bg-slate-800 text-stone-600 dark:text-slate-300 text-[11px] font-bold hover:bg-purple-100 hover:text-purple-700 transition"
                              >
                                + Thêm / chỉnh mã lớp
                              </button>
                            </div>
                          </div>
                          <div className="text-xs font-semibold text-purple-600 dark:text-purple-400">
                            {lessonReps.length > 0
                              ? `${lessonReps.length} buổi đã lưu (${Array.from(new Set(lessonReps.map(r => r.className))).join(', ')}) · Tương tác TB: ${avgPct}%`
                              : 'Đã lưu tự động · Chưa có buổi dạy'}
                          </div>
                        </div>
                        <div className="flex gap-2 mt-5">
                          <button onClick={() => openSavedLesson(i)} className="flex-1 py-2 bg-purple-600 text-white text-sm font-semibold rounded-xl hover:bg-purple-700 transition">
                            Mở & Chọn lớp
                          </button>
                          <button
                            onClick={() => {
                              if (lessonReps.length > 0) {
                                setStatsClassFilter('ALL');
                                setActiveReportId(lessonReps[0].id);
                              }
                              setView('stats');
                            }}
                            className="px-3 py-2 bg-purple-50 dark:bg-purple-950/50 text-purple-600 dark:text-purple-300 text-xs font-bold rounded-xl hover:bg-purple-100 transition flex items-center gap-1"
                            title="Xem thống kê bài giảng"
                          >
                            <BarChart2 className="w-3.5 h-3.5" /> Thống kê
                          </button>
                          <button onClick={() => deleteSavedLesson(i)} className="p-2 border border-stone-300 dark:border-slate-700 rounded-xl text-red-500 hover:bg-red-50 dark:hover:bg-red-950/30 transition">
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div className="space-y-4 pt-4">
              <h3 className="text-lg font-bold text-stone-700 dark:text-slate-300">Thư viện mẫu</h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                {TEMPLATES.map((tmpl, i) => (
                  <div key={i} className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-5 rounded-2xl flex flex-col justify-between shadow-sm">
                    <div>
                      <h4 className="font-bold text-base mb-1">{tmpl.title}</h4>
                      <p className="text-xs text-stone-500 dark:text-slate-400">
                        {tmpl.slides.length} slide · {[...new Set(tmpl.slides.map(s => SLIDE_TYPES[s.type]))].join(', ')}
                      </p>
                    </div>
                    <button onClick={() => useTemplate(i)} className="mt-6 w-full py-2 bg-purple-50 dark:bg-purple-950/50 text-purple-600 dark:text-purple-300 text-sm font-semibold rounded-xl hover:bg-purple-100 transition">
                      Dùng mẫu này
                    </button>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* BUILDER VIEW */}
        {view === 'build' && (
          <div className="space-y-6 py-4">
            <div className="flex justify-between items-center flex-wrap gap-4">
              <h2 className="text-2xl font-bold">Tạo bài giảng</h2>
              <div className="flex gap-3">
                <button onClick={startClassroom} className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl shadow transition flex items-center gap-2">
                  <Play className="w-4 h-4 fill-current" /> Bắt đầu lớp học
                </button>
                <button onClick={saveCurrentLesson} className="px-5 py-2.5 bg-stone-200 dark:bg-slate-800 hover:bg-stone-300 font-semibold rounded-xl transition flex items-center gap-2">
                  <Save className="w-4 h-4" /> Lưu bài giảng
                </button>
              </div>
            </div>
            {saveMsg && <p className="text-sm font-medium text-emerald-600">{saveMsg}</p>}

            {/* AI Topic Generator */}
            <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-3">
              <h3 className="font-bold text-base flex items-center gap-2 text-purple-600 dark:text-purple-400">
                <Sparkles className="w-5 h-5" /> Soạn bài giảng tự động bằng AI
              </h3>
              <div className="flex gap-2 flex-wrap">
                <input 
                  type="text" 
                  value={topicInput} 
                  onChange={e => setTopicInput(e.target.value)} 
                  placeholder="Chủ đề (vd: Địa lý 11 – Đông Nam Á)" 
                  className="flex-1 min-w-[240px]" 
                />
                <button onClick={generateLessonByTopic} className="px-5 py-2.5 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl transition">
                  Soạn bằng AI
                </button>
              </div>
              {genMsg && <p className="text-sm font-medium text-purple-600 dark:text-purple-400">{genMsg}</p>}
            </div>

            {/* Import .pptx / .pdf / .docx / video / image */}
            <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-3">
              <h3 className="font-bold text-base flex items-center gap-2">
                <Upload className="w-5 h-5 text-purple-600" /> Nhập bài giảng có sẵn (.pptx, .pdf, .docx, video, hình ảnh)
              </h3>
              <p className="text-sm text-stone-500 dark:text-slate-400">
                Giữ nguyên 100% bố cục, chữ, hình ảnh, video, âm thanh, bảng biểu và hiệu ứng gốc của tệp khi tải lên. Sau đó bạn có thể bấm nút để AI tự chèn thêm câu hỏi tương tác.
              </p>
              <div className="flex gap-3 flex-wrap items-center pt-2">
                <input 
                  type="file" 
                  multiple
                  accept=".pptx,.pdf,.docx,video/*,image/*" 
                  onChange={e => e.target.files?.length && handlePresentationImport(e.target.files)} 
                  className="file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-semibold file:bg-purple-50 file:text-purple-700 hover:file:bg-purple-100 cursor-pointer"
                />
                <div className="flex items-center gap-2 bg-purple-50 dark:bg-purple-950/50 border border-purple-200 dark:border-purple-800 px-3 py-1.5 rounded-xl">
                  <span className="text-xs font-bold text-purple-800 dark:text-purple-300">Số câu AI chèn:</span>
                  <input
                    type="number"
                    min="1"
                    max="15"
                    value={aiQuestionCount}
                    onChange={e => setAiQuestionCount(Math.max(1, Math.min(15, Number(e.target.value) || 4)))}
                    className="w-16 !py-1 !px-2 text-xs font-bold text-center bg-white dark:bg-slate-900 border border-purple-300 dark:border-purple-700 rounded-lg"
                  />
                </div>
                <button
                  onClick={() => autoInsertInteractiveSlides(aiQuestionCount)}
                  disabled={isAutoInserting}
                  className="px-4 py-2.5 bg-purple-600 hover:bg-purple-700 disabled:opacity-60 text-white font-semibold rounded-xl text-sm transition flex items-center gap-1.5 shadow-sm"
                >
                  <Sparkles className={`w-4 h-4 ${isAutoInserting ? 'animate-spin' : ''}`} />
                  {isAutoInserting ? 'AI đang chèn...' : `AI chèn ${aiQuestionCount} câu hỏi`}
                </button>
              </div>
              {autoInsertMsg && (
                <p className={`text-sm font-medium ${autoInsertMsg.startsWith('AI gặp lỗi') || autoInsertMsg.startsWith('Hãy nhập') ? 'text-amber-600 dark:text-amber-400' : 'text-emerald-600 dark:text-emerald-400'}`}>
                  {autoInsertMsg}
                </p>
              )}
            </div>

            {/* Tool AI from document */}
            <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-3">
              <h3 className="font-bold text-base">Công cụ AI từ tài liệu</h3>
              <div className="flex items-center gap-3 flex-wrap">
                <input 
                  type="file" 
                  accept=".txt,.md,.csv,.docx,.pdf,image/*" 
                  onChange={e => e.target.files?.[0] && handleFileUpload(e.target.files[0])} 
                  className="text-sm text-stone-500"
                />
                <span className="text-xs text-stone-500">{fileMsg}</span>
              </div>
              <textarea 
                rows={3} 
                value={docTextSrc} 
                onChange={e => setDocTextSrc(e.target.value)} 
                placeholder="Hoặc dán nội dung vào đây..." 
              />
              <div className="flex gap-2 flex-wrap pt-1">
                <button onClick={generateFromDoc} className="px-4 py-2 bg-purple-600 text-white text-sm font-semibold rounded-xl hover:bg-purple-700 transition">
                  ✨ Tạo cả bài giảng
                </button>
                <button onClick={() => generateFromTextType('quiz')} className="px-4 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 text-sm font-semibold rounded-xl hover:bg-purple-100 transition">
                  Tạo trắc nghiệm
                </button>
                <button onClick={() => generateFromTextType('open')} className="px-4 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 text-sm font-semibold rounded-xl hover:bg-purple-100 transition">
                  Tạo câu hỏi thảo luận
                </button>
                <button onClick={() => generateFromTextType('poll')} className="px-4 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 text-sm font-semibold rounded-xl hover:bg-purple-100 transition">
                  Tạo bình chọn
                </button>
              </div>
            </div>

            {/* Slides Editor List */}
            <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold uppercase tracking-wider text-stone-500 mb-1">Tên bài giảng</label>
                  <input 
                    type="text" 
                    value={lesson.title} 
                    onChange={e => setLesson(prev => ({ ...prev, title: e.target.value }))} 
                    placeholder="Nhập tên bài giảng..." 
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold uppercase tracking-wider text-stone-500 mb-1">Tên lớp hiển thị (dùng để thống kê sau buổi học)</label>
                  <input 
                    type="text" 
                    value={lesson.className || ''} 
                    onChange={e => setLesson(prev => ({ ...prev, className: e.target.value }))} 
                    placeholder="Ví dụ: Lớp 11A1, Khối 10..." 
                  />
                </div>
              </div>

              {/* Editable Alphanumeric Class Codes & Multi-Class Manager */}
              <div className="p-4 rounded-2xl bg-emerald-50/70 dark:bg-emerald-950/30 border border-emerald-200 dark:border-emerald-900/60 space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <div className="text-sm font-bold text-emerald-900 dark:text-emerald-200 flex items-center gap-2">
                      <QrCode className="w-4 h-4 text-emerald-600" />
                      <span>Mã lớp bằng chữ & số (tùy chỉnh và lưu nhiều mã lớp cho cùng 1 bài giảng)</span>
                    </div>
                    <p className="text-xs text-stone-600 dark:text-slate-400">
                      Giáo viên có thể tự đặt mã lớp gồm chữ và số (VD: <b>11A1</b>, <b>10A2</b>, <b>TOAN11</b>, <b>AB123</b>). Một bài giảng lưu được nhiều mã lớp để dạy nhiều lớp và thống kê riêng từng lớp.
                    </p>
                  </div>
                  {lesson.customCode && (
                    <span className="px-3 py-1 rounded-xl bg-emerald-600 text-white font-mono text-xs font-black tracking-wider">
                      Mã đang chọn: {lesson.customCode}
                    </span>
                  )}
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <input
                    type="text"
                    value={classCodeDraft}
                    onChange={e => {
                      const val = normalizeClassCode(e.target.value);
                      setClassCodeDraft(val);
                      if (classCodeMsg) setClassCodeMsg('');
                    }}
                    onKeyDown={e => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        addOrSelectClassCode(classCodeDraft, true);
                      }
                    }}
                    maxLength={15}
                    placeholder="Nhập mã lớp chữ & số (VD: 11A1, 12B2, SINH10...)"
                    className="!w-64 sm:!w-72 font-mono font-bold uppercase tracking-wider text-sm"
                  />
                  <button
                    type="button"
                    onClick={() => addOrSelectClassCode(classCodeDraft, true)}
                    className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5"
                  >
                    <Plus className="w-3.5 h-3.5" /> Lưu / Chọn mã lớp
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const rnd = generateRandomClassCode();
                      addOrSelectClassCode(rnd, false);
                    }}
                    className="px-3 py-2 bg-white dark:bg-slate-900 border border-stone-300 dark:border-slate-700 text-stone-700 dark:text-slate-200 text-xs font-semibold rounded-xl hover:bg-stone-100 transition flex items-center gap-1"
                  >
                    <RefreshCw className="w-3.5 h-3.5" /> Tạo mã ngẫu nhiên
                  </button>
                  {classCodeDraft && isValidClassCode(classCodeDraft) && (
                    <button
                      type="button"
                      onClick={() => startClassroom(classCodeDraft)}
                      className="px-3.5 py-2 bg-purple-600 hover:bg-purple-700 text-white text-xs font-bold rounded-xl transition flex items-center gap-1.5"
                    >
                      <Play className="w-3.5 h-3.5 fill-current" /> Dạy ngay lớp {classCodeDraft}
                    </button>
                  )}
                </div>

                {/* Saved Multi-Class Codes List */}
                {(lesson.classCodes && lesson.classCodes.length > 0) ? (
                  <div className="flex items-center gap-2 flex-wrap pt-1">
                    <span className="text-xs font-semibold text-stone-500">
                      Các mã lớp đã lưu của bài giảng ({lesson.classCodes.length}):
                    </span>
                    {lesson.classCodes.map(code => {
                      const isCurrent = lesson.customCode === code;
                      return (
                        <div
                          key={code}
                          className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-xs font-mono font-bold border transition ${
                            isCurrent
                              ? 'bg-emerald-600 text-white border-emerald-600 shadow-xs'
                              : 'bg-white dark:bg-slate-900 text-stone-700 dark:text-slate-200 border-stone-200 dark:border-slate-700'
                          }`}
                        >
                          <button
                            type="button"
                            onClick={() => {
                              setClassCodeDraft(code);
                              setLesson(prev => {
                                const updated: Lesson = {
                                  ...prev,
                                  customCode: code,
                                  className: !prev.className?.trim() || prev.className.startsWith('Lớp ') ? `Lớp ${code}` : prev.className,
                                };
                                lessonRef.current = updated;
                                return updated;
                              });
                              setClassCodeMsg(`Đã chuyển sang mã lớp ${code}.`);
                            }}
                            title={`Chọn mã lớp ${code}`}
                            className="hover:underline"
                          >
                            {code}
                          </button>
                          <button
                            type="button"
                            onClick={() => startClassroom(code)}
                            title={`Mở phòng dạy ngay với mã ${code}`}
                            className={`px-1.5 py-0.5 rounded text-[10px] font-sans font-bold ${
                              isCurrent ? 'bg-white/20 text-white hover:bg-white/30' : 'bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 hover:bg-purple-200'
                            }`}
                          >
                            Dạy
                          </button>
                          <button
                            type="button"
                            onClick={() => removeClassCodeFromLesson(code)}
                            title={`Xóa mã ${code}`}
                            className={`p-0.5 rounded hover:bg-black/10 ${isCurrent ? 'text-white/80' : 'text-stone-400 hover:text-red-500'}`}
                          >
                            <X className="w-3 h-3" />
                          </button>
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  <p className="text-xs text-stone-500 italic">
                    Chưa có mã lớp cố định. Bạn có thể nhập mã lớp (VD: 11A1, 11A2) để lưu nhiều lớp cho bài giảng này, hoặc để hệ thống tự tạo khi bắt đầu.
                  </p>
                )}

                {classCodeMsg && (
                  <p className="text-xs font-semibold text-emerald-700 dark:text-emerald-300">{classCodeMsg}</p>
                )}
              </div>

              {/* Global AI Question Timer Setting (Giao diện tổng) */}
              <div className="p-4 rounded-2xl bg-purple-50/70 dark:bg-purple-950/30 border border-purple-100 dark:border-purple-900/50 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2 text-sm font-semibold text-purple-800 dark:text-purple-200">
                  <Clock className="w-4 h-4 text-purple-600" />
                  <span>Thời gian đếm ngược câu hỏi tương tác AI (giây):</span>
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  {[0, 30, 45, 60, 90].map(sec => (
                    <button
                      key={sec}
                      type="button"
                      onClick={() => setDefaultAiTimerSec(sec)}
                      className={`px-2.5 py-1 rounded-lg text-xs font-bold transition ${defaultAiTimerSec === sec ? 'bg-purple-600 text-white' : 'bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-700 text-stone-600 dark:text-slate-300 hover:border-purple-400'}`}
                    >
                      {sec === 0 ? 'Không giới hạn' : `${sec}s`}
                    </button>
                  ))}
                  <input
                    type="number"
                    min="0"
                    max="600"
                    value={defaultAiTimerSec}
                    onChange={e => setDefaultAiTimerSec(Math.max(0, Math.min(600, parseInt(e.target.value || '0', 10) || 0)))}
                    className="!w-20 !py-1 text-center text-xs font-bold"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      setLesson(prev => ({
                        ...prev,
                        slides: prev.slides.map(sl =>
                          sl.type !== 'info' && sl.type !== 'media' && (sl.auto || sl.aiGenerated)
                            ? { ...sl, duration: defaultAiTimerSec > 0 ? defaultAiTimerSec : undefined }
                            : sl
                        ),
                      }));
                      setAutoInsertMsg(
                        defaultAiTimerSec > 0
                          ? `Đã đặt thời gian ${defaultAiTimerSec} giây cho tất cả câu hỏi tương tác AI.`
                          : 'Đã tắt giới hạn thời gian cho tất cả câu hỏi tương tác AI.'
                      );
                    }}
                    className="px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white text-xs font-bold rounded-lg transition"
                  >
                    Áp dụng cho câu hỏi AI
                  </button>
                </div>
              </div>

              <div className="space-y-4 pt-2">
                {lesson.slides.map((s, idx) => (
                  <div key={idx} className="border border-stone-200 dark:border-slate-800 rounded-2xl p-4 bg-stone-50 dark:bg-slate-950/50 space-y-3">
                    <div className="flex justify-between items-center flex-wrap gap-2">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="px-2.5 py-1 bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 rounded-full text-xs font-bold">
                          {SLIDE_TYPES[s.type]}
                        </span>
                        {(s.auto || s.aiGenerated) && <span className="px-2.5 py-1 bg-amber-100 dark:bg-amber-950 text-amber-700 rounded-full text-xs font-bold">AI tạo</span>}
                        {s.ai && <span className="px-2.5 py-1 bg-blue-100 dark:bg-blue-950 text-blue-700 rounded-full text-xs font-bold">+ AI phản hồi</span>}
                        {s.duration && s.duration > 0 && (
                          <span className="px-2.5 py-1 bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300 rounded-full text-xs font-bold inline-flex items-center gap-1">
                            <Clock className="w-3 h-3" /> {s.duration}s
                          </span>
                        )}
                        <span className="text-xs text-stone-500 font-medium">Slide {idx + 1}</span>
                      </div>
                      <button 
                        onClick={() => setLesson(prev => ({ ...prev, slides: prev.slides.filter((_, i) => i !== idx) }))} 
                        className="text-red-500 hover:text-red-700 p-1 rounded-lg hover:bg-red-50 transition"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>

                    <input 
                      type="text" 
                      value={s.q} 
                      onChange={e => {
                        const val = e.target.value;
                        setLesson(prev => {
                          const nextSlides = [...prev.slides];
                          nextSlides[idx].q = val;
                          return { ...prev, slides: nextSlides };
                        });
                      }} 
                      placeholder="Nội dung câu hỏi / tiêu đề slide" 
                    />

                    {s.type === 'media' && (
                      <div className="space-y-2">
                        <div className="flex items-center justify-between flex-wrap gap-2">
                          <p className="text-xs text-stone-500">
                            {(s.imgs || []).length} ảnh{s.vid ? ', 1 video' : ''}{s.b ? ' · có chữ' : ''} – giữ nguyên khi trình chiếu
                            {s.pptx?.transition && ` · Chuyển slide: ${s.pptx.transition.label || s.pptx.transition.type}`}
                            {getPptxMaxStep(s.pptx) > 0 && ` · ${getPptxMaxStep(s.pptx)} bước hiệu ứng`}
                          </p>
                          {s.pptx && (
                            <button
                              type="button"
                              onClick={() => {
                                setBuilderPreviewIdx(builderPreviewIdx === idx ? null : idx);
                                setBuilderAnimStep(0);
                              }}
                              className="px-2.5 py-1 rounded-lg bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 text-[11px] font-bold hover:bg-purple-200 transition"
                            >
                              {builderPreviewIdx === idx ? 'Ẩn xem thử PPT' : 'Xem thử hiệu ứng PPT'}
                            </button>
                          )}
                        </div>
                        {s.pptx && builderPreviewIdx === idx && (
                          <div className="pt-1">
                            <PptxSlideView
                              pptx={s.pptx}
                              slideKey={`preview-${idx}`}
                              resolveMediaUrl={resolveMediaUrl}
                              animStep={builderAnimStep}
                              onSetAnimStep={setBuilderAnimStep}
                              showControls={true}
                              maxHeightClass="max-h-[380px]"
                            />
                          </div>
                        )}
                      </div>
                    )}

                    {s.type === 'info' && (
                      <textarea 
                        rows={3} 
                        value={s.b || ''} 
                        onChange={e => {
                          const val = e.target.value;
                          setLesson(prev => {
                            const nextSlides = [...prev.slides];
                            nextSlides[idx].b = val;
                            return { ...prev, slides: nextSlides };
                          });
                        }} 
                        placeholder="Nội dung trình bày chi tiết" 
                      />
                    )}

                    {s.opts && (
                      <textarea 
                        rows={3} 
                        value={s.opts.join('\n')} 
                        onChange={e => {
                          const opts = e.target.value.split('\n');
                          setLesson(prev => {
                            const nextSlides = [...prev.slides];
                            nextSlides[idx].opts = opts;
                            return { ...prev, slides: nextSlides };
                          });
                        }} 
                        placeholder="Mỗi dòng một lựa chọn" 
                      />
                    )}

                    {s.type === 'quiz' && (
                      <div className="flex items-center gap-3">
                        <span className="text-xs font-semibold text-stone-500">Đáp án đúng (số thứ tự 1, 2, 3...):</span>
                        <input 
                          type="number" 
                          min="1" 
                          style={{ width: '80px' }} 
                          value={(s.a ?? 0) + 1} 
                          onChange={e => {
                            const a = Math.max(0, parseInt(e.target.value || '1') - 1);
                            setLesson(prev => {
                              const nextSlides = [...prev.slides];
                              nextSlides[idx].a = a;
                              return { ...prev, slides: nextSlides };
                            });
                          }} 
                        />
                      </div>
                    )}

                    {s.type !== 'info' && s.type !== 'media' && (
                      <div className="flex items-center gap-2 flex-wrap pt-1">
                        <span className="text-xs font-semibold text-stone-500 flex items-center gap-1">
                          <Clock className="w-3.5 h-3.5 text-purple-600" /> Đếm ngược (giây):
                        </span>
                        {[0, 30, 45, 60, 90].map(sec => (
                          <button
                            key={sec}
                            type="button"
                            onClick={() => {
                              setLesson(prev => {
                                const nextSlides = prev.slides.map((sl, i) => (i === idx ? { ...sl, duration: sec > 0 ? sec : undefined } : sl));
                                return { ...prev, slides: nextSlides };
                              });
                            }}
                            className={`px-2 py-0.5 rounded-md text-xs font-semibold transition ${(s.duration || 0) === sec ? 'bg-purple-600 text-white' : 'bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-700 text-stone-600 dark:text-slate-300'}`}
                          >
                            {sec === 0 ? 'Không' : `${sec}s`}
                          </button>
                        ))}
                        <input
                          type="number"
                          min="0"
                          max="600"
                          value={s.duration || 0}
                          onChange={e => {
                            const dur = Math.max(0, Math.min(600, parseInt(e.target.value || '0', 10) || 0));
                            setLesson(prev => {
                              const nextSlides = prev.slides.map((sl, i) => (i === idx ? { ...sl, duration: dur > 0 ? dur : undefined } : sl));
                              return { ...prev, slides: nextSlides };
                            });
                          }}
                          className="!w-20 !py-1 text-center text-xs font-bold"
                        />
                      </div>
                    )}
                  </div>
                ))}
              </div>

              {/* Add Slide Toolbar */}
              <div className="flex flex-wrap gap-2 pt-4 border-t border-stone-200 dark:border-slate-800 items-center">
                <button
                  onClick={() => autoInsertInteractiveSlides(aiQuestionCount)}
                  disabled={isAutoInserting}
                  className="px-3.5 py-2 bg-purple-600 hover:bg-purple-700 disabled:opacity-60 text-white font-semibold rounded-xl text-xs transition flex items-center gap-1.5"
                >
                  <Sparkles className={`w-3.5 h-3.5 ${isAutoInserting ? 'animate-spin' : ''}`} />
                  {isAutoInserting ? 'AI đang chèn...' : 'AI tự chèn câu hỏi tương tác'}
                </button>
                <button onClick={() => addSlideType('open')} className="px-3 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 font-semibold rounded-xl text-xs hover:bg-purple-100 transition">
                  + Câu hỏi mở
                </button>
                <button onClick={() => addSlideType('poll')} className="px-3 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 font-semibold rounded-xl text-xs hover:bg-purple-100 transition">
                  + Bình chọn
                </button>
                <button onClick={() => addSlideType('cloud')} className="px-3 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 font-semibold rounded-xl text-xs hover:bg-purple-100 transition">
                  + Đám mây từ
                </button>
                <button onClick={() => addSlideType('quiz')} className="px-3 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 font-semibold rounded-xl text-xs hover:bg-purple-100 transition">
                  + Trắc nghiệm
                </button>
                <button onClick={() => addSlideType('scale')} className="px-3 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 font-semibold rounded-xl text-xs hover:bg-purple-100 transition">
                  + Thang điểm 1–5
                </button>
                <button onClick={() => addSlideType('draw')} className="px-3 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 font-semibold rounded-xl text-xs hover:bg-purple-100 transition">
                  + Vẽ
                </button>
                <button onClick={() => addSlideType('info')} className="px-3 py-2 bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 font-semibold rounded-xl text-xs hover:bg-purple-100 transition">
                  + Nội dung
                </button>
              </div>
              {autoInsertMsg && (
                <p className={`text-sm font-medium ${autoInsertMsg.startsWith('AI gặp lỗi') || autoInsertMsg.startsWith('Hãy nhập') ? 'text-amber-600 dark:text-amber-400' : 'text-emerald-600 dark:text-emerald-400'}`}>
                  {autoInsertMsg}
                </p>
              )}
            </div>
          </div>
        )}

        {/* HOST VIEW (Lobby stays visible until teacher clicks Vào bài giảng) */}
        {view === 'host' && !teacherStartedLive && (
          <div className="space-y-6 py-4 max-w-3xl mx-auto">
            <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-8 rounded-3xl shadow-sm space-y-6">
              <div className="flex flex-wrap gap-8 items-center justify-between">
                <div className="w-[210px] h-[210px] flex items-center justify-center bg-white rounded-2xl border border-stone-200 p-2 shadow-inner mx-auto sm:mx-0">
                  {qrDataUrl ? (
                    <img alt="QR vào lớp" src={qrDataUrl} className="w-full h-full object-contain rounded-lg" />
                  ) : (
                    <span className="text-xs text-stone-400 text-center px-2">Đang tạo mã QR...</span>
                  )}
                </div>
                <div className="flex-1 min-w-[240px] space-y-3">
                  <div className="flex items-center gap-2 flex-wrap">
                    <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 text-xs font-bold">
                      Phòng học đang mở
                    </div>
                    <div className={`inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-extrabold ${onlineStudentCount > 0 ? 'bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300' : 'bg-stone-100 dark:bg-slate-800 text-stone-600 dark:text-slate-300'}`}>
                      <span>{onlineStudentCount} học sinh đang truy cập</span>
                    </div>
                  </div>
                  <p className="text-sm text-stone-600 dark:text-slate-400">
                    Học sinh quét mã QR hoặc nhập mã bên dưới để vào hàng chờ. Khi sẵn sàng, bấm <b>Vào bài giảng</b> để bắt đầu trình chiếu.
                  </p>
                  <div className="text-4xl sm:text-5xl font-black tracking-widest text-purple-600 dark:text-purple-400 font-mono">
                    {roomCode}
                  </div>
                  {/* Editable Alphanumeric Room Code & Saved Class Codes Switcher in Host Lobby */}
                  <div className="p-3.5 rounded-2xl bg-stone-50 dark:bg-slate-950 border border-stone-200 dark:border-slate-800 space-y-2.5">
                    <label className="block text-xs font-bold text-stone-600 dark:text-slate-300">
                      Điều chỉnh mã lớp (chữ & số) hoặc chọn lớp dạy:
                    </label>
                    <div className="flex items-center gap-2 flex-wrap">
                      <input
                        type="text"
                        value={hostCodeDraft}
                        onChange={e => setHostCodeDraft(normalizeClassCode(e.target.value))}
                        onKeyDown={e => e.key === 'Enter' && applyHostRoomCode(hostCodeDraft)}
                        maxLength={15}
                        placeholder="VD: 11A1, 10A2, TOAN11"
                        className="!w-44 font-mono font-bold uppercase tracking-wider !py-1.5 text-sm"
                      />
                      <button
                        type="button"
                        onClick={() => applyHostRoomCode(hostCodeDraft)}
                        className="px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white text-xs font-bold rounded-xl transition"
                      >
                        Đổi & Lưu mã lớp
                      </button>
                      <button
                        type="button"
                        onClick={() => applyHostRoomCode(generateRandomClassCode())}
                        className="px-2.5 py-1.5 bg-white dark:bg-slate-900 border border-stone-300 dark:border-slate-700 text-xs font-semibold rounded-xl hover:bg-stone-100 transition"
                      >
                        Mã ngẫu nhiên
                      </button>
                    </div>
                    {lesson.classCodes && lesson.classCodes.length > 0 && (
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <span className="text-[11px] font-semibold text-stone-500">Các mã lớp của bài giảng:</span>
                        {lesson.classCodes.map(code => (
                          <button
                            key={code}
                            type="button"
                            onClick={() => applyHostRoomCode(code)}
                            className={`px-2.5 py-0.5 rounded-lg font-mono text-xs font-bold transition ${
                              roomCode === code
                                ? 'bg-purple-600 text-white'
                                : 'bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-700 text-stone-700 dark:text-slate-300 hover:border-purple-400'
                            }`}
                          >
                            {code}
                          </button>
                        ))}
                      </div>
                    )}
                    <div>
                      <label className="block text-xs font-semibold text-stone-500 mb-1">Tên lớp hiển thị (dùng cho thống kê cuối buổi)</label>
                      <input
                        type="text"
                        value={lesson.className || ''}
                        onChange={e => {
                          const val = e.target.value;
                          setLesson(prev => {
                            const updated = { ...prev, className: val };
                            lessonRef.current = updated;
                            return updated;
                          });
                          syncLiveSessionReport('live');
                        }}
                        placeholder={`Ví dụ: Lớp 11A1 (mặc định: Lớp ${roomCode})`}
                        className="!py-1.5 text-sm"
                      />
                    </div>
                  </div>
                  <div className="text-xs text-stone-500 break-all">{roomLink}</div>
                  <div className="text-xs font-semibold text-emerald-600">{hostStatus}</div>
                  <div className="flex flex-wrap gap-2 pt-2">
                    <button onClick={enterTeacherLiveScreen} className="px-5 py-2.5 bg-purple-600 hover:bg-purple-700 text-white text-sm font-bold rounded-xl shadow transition flex items-center gap-2">
                      <Play className="w-4 h-4 fill-current" /> Vào bài giảng ({onlineStudentCount} HS)
                    </button>
                    <button onClick={() => endClassroom(true)} className="px-4 py-2.5 border border-stone-300 dark:border-slate-700 text-stone-600 dark:text-stone-300 text-sm font-semibold rounded-xl hover:bg-stone-100 dark:hover:bg-slate-800 transition">
                      Hủy lớp
                    </button>
                  </div>
                </div>
              </div>

              {/* Real-time Accessed Students Counter & List */}
              <div className="border-t border-stone-200 dark:border-slate-800 pt-5 space-y-3">
                <div className="flex items-center justify-between flex-wrap gap-2">
                  <div className="flex items-center gap-2">
                    <span className={`w-2.5 h-2.5 rounded-full ${onlineStudentCount > 0 ? 'bg-emerald-500 animate-pulse' : 'bg-stone-300 dark:bg-slate-700'}`} />
                    <span className="text-sm font-bold">
                      Số lượng học sinh truy cập: <span className="text-purple-600 dark:text-purple-400 text-base font-black">{onlineStudentCount}</span>
                    </span>
                  </div>
                  {Object.keys(allSessionStudents).length > onlineStudentCount && (
                    <span className="text-xs text-stone-500">
                      Tổng lượt đã kết nối: {Object.keys(allSessionStudents).length}
                    </span>
                  )}
                </div>
                {Object.keys(connectedStudents).length > 0 ? (
                  <div className="flex items-center gap-2 flex-wrap">
                    {Object.entries(connectedStudents).map(([sid, sname]) => {
                      const scls = allSessionStudentClasses[sid];
                      return (
                        <span key={sid} className="px-3 py-1 rounded-full bg-purple-50 dark:bg-purple-950/50 border border-purple-200 dark:border-purple-800 text-xs font-semibold text-purple-800 dark:text-purple-200 inline-flex items-center gap-1.5">
                          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                          <span>{sname}</span>
                          {scls && <span className="px-1.5 py-0.2 bg-white dark:bg-slate-900 rounded-full text-[10px] font-bold text-purple-600">{scls}</span>}
                        </span>
                      );
                    })}
                  </div>
                ) : (
                  <p className="text-xs text-stone-400 italic">Chưa có học sinh nào truy cập. Số lượng và danh sách học sinh sẽ hiện ngay tại đây khi các em quét mã QR hoặc nhập mã.</p>
                )}
              </div>
            </div>
          </div>
        )}

        {/* TEACHER FULL-SCREEN LIVE LESSON & INTERACTION MODE */}
        {isTeacherLiveMode && (
          <div
            ref={liveStageRef}
            className={
              isFullscreen
                ? 'fixed inset-0 z-50 overflow-y-auto flex flex-col min-h-screen bg-stone-50 dark:bg-slate-950'
                : 'flex-1 flex flex-col min-h-screen bg-stone-50 dark:bg-slate-950'
            }
          >
            {/* Slim Presentation Top Bar (no outer website chrome) */}
            <header className="bg-white/90 dark:bg-slate-900/90 backdrop-blur border-b border-stone-200 dark:border-slate-800 px-3 sm:px-6 py-2 sticky top-0 z-40">
              <div className="max-w-[96vw] 2xl:max-w-[92vw] mx-auto flex items-center justify-between flex-wrap gap-2">
                <div className="flex items-center gap-3 min-w-0">
                  <span className="px-3 py-1 bg-purple-600 text-white font-mono font-black tracking-widest rounded-xl text-sm">
                    {roomCode}
                  </span>
                  <div className="min-w-0">
                    <div className="flex items-center gap-2 flex-wrap">
                      <h1 className="font-bold text-base sm:text-lg truncate">
                        {lesson.title || 'Bài giảng trực tiếp'}
                      </h1>
                      <input
                        type="text"
                        value={lesson.className || ''}
                        onChange={e => {
                          const val = e.target.value;
                          setLesson(prev => {
                            const updated = { ...prev, className: val };
                            lessonRef.current = updated;
                            return updated;
                          });
                          syncLiveSessionReport('live');
                        }}
                        placeholder={`Lớp ${roomCode}`}
                        className="!w-28 sm:!w-36 !py-0.5 !px-2 text-xs font-bold rounded-lg bg-purple-50 dark:bg-purple-950/60 border border-purple-200 dark:border-purple-800 text-purple-700 dark:text-purple-300"
                        title="Tên lớp (chỉnh sửa trực tiếp)"
                      />
                    </div>
                    <div className="text-xs text-stone-500 dark:text-slate-400 flex items-center gap-2">
                      <span className="inline-flex items-center gap-1 text-emerald-600 dark:text-emerald-400 font-semibold">
                        <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                        {Math.max(connections.filter(c => c.open).length, Object.keys(connectedStudents).length)} học sinh đang học
                      </span>
                      <span>·</span>
                      <span>Slide {slideIdx + 1} / {lesson.slides.length}</span>
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  <button
                    onClick={toggleFullscreen}
                    className={`px-3 py-2 text-xs font-bold rounded-xl border transition flex items-center gap-1.5 ${
                      isFullscreen
                        ? 'bg-purple-600 text-white border-purple-600 hover:bg-purple-700'
                        : 'border-stone-300 dark:border-slate-700 hover:bg-stone-100 dark:hover:bg-slate-800'
                    }`}
                    title={isFullscreen ? 'Thoát toàn màn hình (Esc / F)' : 'Chế độ toàn màn hình (F)'}
                  >
                    {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4 text-purple-600" />}
                    <span>{isFullscreen ? 'Thu nhỏ' : 'Toàn màn hình'}</span>
                  </button>
                  <button onClick={() => setShowQrModal(true)} className="px-3 py-2 text-xs font-bold rounded-xl border border-stone-300 dark:border-slate-700 hover:bg-stone-100 dark:hover:bg-slate-800 transition flex items-center gap-1.5">
                    <QrCode className="w-4 h-4 text-purple-600" /> Mã QR
                  </button>
                  <button onClick={requestAiFeedback} className="px-3 py-2 text-xs font-bold rounded-xl bg-purple-50 dark:bg-purple-950 text-purple-600 dark:text-purple-300 hover:bg-purple-100 transition flex items-center gap-1.5">
                    <Sparkles className="w-4 h-4" /> Phản hồi AI
                  </button>
                  <button onClick={() => setShowScoreboard(v => !v)} className="px-3 py-2 text-xs font-bold rounded-xl bg-amber-50 dark:bg-amber-950/40 text-amber-600 hover:bg-amber-100 transition flex items-center gap-1.5">
                    <Trophy className="w-4 h-4" /> Bảng điểm
                  </button>
                  <button onClick={() => endClassroom(false)} className="px-3.5 py-2 text-xs font-bold rounded-xl bg-red-50 dark:bg-red-950/40 text-red-600 dark:text-red-400 hover:bg-red-100 transition">
                    Kết thúc & Thống kê
                  </button>
                </div>
              </div>
            </header>

            {/* Main Lesson & Direct Student Interaction Stage (auto-fits TV aspect ratio) */}
            <div className="flex-1 max-w-[96vw] 2xl:max-w-[92vw] w-full mx-auto px-3 sm:px-6 py-3 space-y-3 flex flex-col justify-center">
              {/* Online Student Pills (hidden in immersive fullscreen to keep stage uncluttered) */}
              {!isFullscreen && Object.keys(connectedStudents).length > 0 && (
                <div className="flex items-center gap-1.5 flex-wrap">
                  <span className="text-xs font-semibold text-stone-500">Học sinh trong lớp:</span>
                  {Object.entries(connectedStudents).map(([sid, sname]) => {
                    const scls = allSessionStudentClasses[sid];
                    return (
                      <span key={sid} className="px-2.5 py-0.5 rounded-full bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 text-xs font-semibold text-stone-700 dark:text-slate-300 shadow-xs inline-flex items-center gap-1">
                        <span>{sname}</span>
                        {scls && <span className="px-1.5 py-0.2 bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 rounded-full text-[10px] font-bold">{scls}</span>}
                      </span>
                    );
                  })}
                </div>
              )}

              {/* Active Slide + Live Responses */}
              <div className={`bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 ${isFullscreen ? 'p-4 sm:p-6' : 'p-4 sm:p-6'} rounded-3xl shadow-sm space-y-3`}>
                {(() => {
                  const curSlide = lesson.slides[slideIdx];
                  const submittedCount = Object.values(responses[slideIdx] || {}).length;
                  const totalStudentsCount = Math.max(
                    connections.filter(c => c.open).length,
                    Object.keys(connectedStudents).length,
                    Object.keys(allSessionStudents).length,
                    submittedCount
                  );
                  const submittedPct = totalStudentsCount > 0 ? Math.min(100, Math.round((submittedCount / totalStudentsCount) * 100)) : 0;
                  const isInteractive = curSlide && curSlide.type !== 'info' && curSlide.type !== 'media';
                  const timerPct = hostTimerTotal > 0 && hostTimerLeft !== null ? Math.max(0, Math.min(100, Math.round((hostTimerLeft / hostTimerTotal) * 100))) : 0;

                  return (
                    <>
                      <div className="flex justify-between items-center flex-wrap gap-3">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="px-3.5 py-1 bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 rounded-full text-xs font-bold">
                            {SLIDE_TYPES[curSlide?.type]}
                          </span>
                          {(curSlide?.auto || curSlide?.aiGenerated) && (
                            <span className="px-2.5 py-1 bg-amber-100 dark:bg-amber-950 text-amber-700 rounded-full text-xs font-bold">
                              Câu hỏi AI
                            </span>
                          )}
                        </div>

                        {/* Submission Counter & Live Timer Controls */}
                        <div className="flex items-center gap-3 flex-wrap">
                          {isInteractive && (
                            <div className="flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-emerald-50 dark:bg-emerald-950/40 border border-emerald-200 dark:border-emerald-900 text-emerald-700 dark:text-emerald-300">
                              <span className="text-sm font-extrabold">
                                Đã nộp: {submittedCount} / {totalStudentsCount} học sinh ({submittedPct}%)
                              </span>
                            </div>
                          )}

                          {isInteractive && (
                            <div className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-stone-100 dark:bg-slate-800 border border-stone-200 dark:border-slate-700 flex-wrap">
                              <Clock className={`w-4 h-4 ${hostTimerLeft === 0 ? 'text-red-500 animate-pulse' : hostTimerLeft !== null && hostTimerLeft <= 10 ? 'text-amber-500' : 'text-purple-600'}`} />
                              {hostTimerLeft !== null ? (
                                <>
                                  <span className={`font-mono font-black text-base ${hostTimerLeft === 0 ? 'text-red-600 dark:text-red-400' : hostTimerLeft <= 10 ? 'text-amber-600' : 'text-slate-900 dark:text-white'}`}>
                                    {hostTimerLeft === 0 ? 'Hết giờ!' : formatTimer(hostTimerLeft)}
                                  </span>
                                  <button
                                    type="button"
                                    onClick={toggleLiveTimerPause}
                                    disabled={hostTimerLeft === 0}
                                    className="px-2 py-0.5 text-xs font-bold rounded-md bg-white dark:bg-slate-900 border border-stone-300 dark:border-slate-600 disabled:opacity-40"
                                    title={hostTimerRunning ? 'Tạm dừng' : 'Tiếp tục'}
                                  >
                                    {hostTimerRunning ? 'Tạm dừng' : 'Chạy'}
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => addLiveTimerSeconds(15)}
                                    className="px-2 py-0.5 text-xs font-bold rounded-md bg-purple-600 text-white hover:bg-purple-700"
                                  >
                                    +15s
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => updateLiveSlideTimer(curSlide?.duration || hostTimerTotal || 45, true)}
                                    className="p-1 text-xs rounded-md bg-white dark:bg-slate-900 border border-stone-300 dark:border-slate-600"
                                    title="Đặt lại thời gian"
                                  >
                                    <RotateCcw className="w-3.5 h-3.5" />
                                  </button>
                                </>
                              ) : (
                                <span className="text-xs font-semibold text-stone-500">Không giới hạn</span>
                              )}
                              <input
                                type="number"
                                min="0"
                                max="600"
                                value={curSlide?.duration || 0}
                                onChange={e => updateLiveSlideTimer(parseInt(e.target.value || '0', 10) || 0, true)}
                                className="!w-16 !py-0.5 !px-1.5 text-center text-xs font-bold ml-1"
                                title="Chỉnh thời gian đếm ngược (giây, 0 = tắt)"
                              />
                              <span className="text-[11px] text-stone-500">giây</span>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Countdown Progress Bar */}
                      {isInteractive && hostTimerLeft !== null && hostTimerTotal > 0 && (
                        <div className="w-full h-2 bg-stone-100 dark:bg-slate-800 rounded-full overflow-hidden">
                          <div
                            className={`h-full transition-all duration-500 ${hostTimerLeft === 0 ? 'bg-red-500' : hostTimerLeft <= 10 ? 'bg-amber-500' : 'bg-purple-600'}`}
                            style={{ width: `${timerPct}%` }}
                          />
                        </div>
                      )}
                    </>
                  );
                })()}

                {!lesson.slides[slideIdx]?.pptx && (
                  <h2 className="text-[clamp(1.25rem,2.2vw,2.25rem)] font-extrabold tracking-tight leading-snug break-words">
                    {lesson.slides[slideIdx]?.q}
                  </h2>
                )}

                <div className="pt-1">
                  {(() => {
                    const s = lesson.slides[slideIdx];
                    const r = Object.values(responses[slideIdx] || {});
                    if (!s) return null;

                    if (s.type === 'media') return renderMediaSlide(s, true, `host-${slideIdx}`);
                    if (s.type === 'info') {
                      const infoLen = (s.b || '').length;
                      const infoSize =
                        infoLen > 500
                          ? 'text-base sm:text-lg'
                          : infoLen > 250
                          ? 'text-lg sm:text-xl'
                          : 'text-xl sm:text-2xl';
                      return <p className={`${infoSize} whitespace-pre-line leading-relaxed break-words text-stone-700 dark:text-slate-300`}>{s.b || ''}</p>;
                    }

                    const isInteractive = !['info', 'media'].includes(s.type);
                    const isRevealed = revealedSlides[slideIdx] || hostTimerLeft === 0 || hostTimerLeft === null;

                    if (isInteractive && !isRevealed) {
                      return (
                        <div className="py-12 px-6 bg-purple-50/80 dark:bg-purple-950/30 border border-purple-200 dark:border-purple-900 rounded-3xl text-center space-y-4 my-4 shadow-sm">
                          <Clock className="w-12 h-12 text-purple-600 animate-pulse mx-auto" />
                          <h3 className="text-xl font-bold">Đang trong thời gian học sinh trả lời...</h3>
                          <p className="text-sm text-stone-600 dark:text-slate-400">
                            Đã có <b>{r.length}</b> học sinh nộp bài. Toàn bộ câu trả lời sẽ tự động xuất hiện khi hết giờ hoặc khi giáo viên bấm hiện kết quả.
                          </p>
                          <button
                            onClick={forceRevealResults}
                            className="px-6 py-3 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl shadow-md transition flex items-center gap-2 mx-auto"
                          >
                            <Eye className="w-5 h-5" /> Hiện kết quả ngay (Hết giờ / Kết thúc)
                          </button>
                        </div>
                      );
                    }

                    if (s.type === 'draw' || s.type === 'open') {
                      const tally = getSlideVoteTally(slideIdx);
                      const isVotingOpen = !!votingActive[slideIdx];
                      const sortedSubmissions = r.map((a, j) => {
                        const sid = Object.keys(responses[slideIdx] || {})[j] || String(j);
                        const votesCount = tally[sid] || 0;
                        return { ...a, sid, votesCount };
                      }).sort((x, y) => y.votesCount - x.votesCount);

                      return r.length ? (
                        <div className="space-y-6">
                          <div className="flex justify-between items-center flex-wrap gap-3 p-4 bg-purple-50 dark:bg-purple-950/40 rounded-2xl border border-purple-200 dark:border-purple-800">
                            <div>
                              <h4 className="font-bold text-base">🗳️ Vòng bình chọn trong lớp</h4>
                              <p className="text-xs text-stone-500">
                                {isVotingOpen ? 'Học sinh đang bình chọn bài vẽ/câu hay nhất trên điện thoại...' : 'Hết thời gian! Bấm bắt đầu để học sinh bình chọn bài đẹp/hay nhất.'}
                              </p>
                            </div>
                            {!isVotingOpen ? (
                              <button
                                onClick={startVotingPhase}
                                className="px-5 py-2.5 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl shadow transition flex items-center gap-2"
                              >
                                <Award className="w-4 h-4" /> Bắt đầu bình chọn
                              </button>
                            ) : (
                              <span className="px-4 py-2 bg-emerald-100 dark:bg-emerald-950 text-emerald-700 font-bold rounded-xl text-xs flex items-center gap-1.5">
                                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" /> Đang mở bình chọn
                              </span>
                            )}
                          </div>

                          {/* Top 3 podium highlight if votes exist */}
                          {sortedSubmissions.length > 0 && sortedSubmissions[0].votesCount > 0 && (
                            <div className="space-y-2">
                              <h5 className="font-extrabold text-sm text-stone-600 dark:text-slate-400 flex items-center gap-1.5">
                                <Trophy className="w-4 h-4 text-amber-500 fill-current" /> Top 3 bức tranh / câu trả lời được yêu thích nhất:
                              </h5>
                              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                {sortedSubmissions.slice(0, 3).map((item, rank) => {
                                  if (item.votesCount === 0) return null;
                                  const medals = ['🥇 Top 1', '🥈 Top 2', '🥉 Top 3'];
                                  const borderColors = ['border-amber-400 bg-amber-50/40 ring-2 ring-amber-400 shadow-lg', 'border-slate-300 bg-slate-50/40 shadow', 'border-amber-700/40 bg-amber-900/10 shadow'];
                                  return (
                                    <div key={rank} className={`p-4 rounded-2xl border-2 ${borderColors[rank]} space-y-2 text-center relative`}>
                                      <div className="font-black text-xs px-2 py-0.5 bg-amber-500 text-white rounded-full inline-block shadow">{medals[rank]} ({item.votesCount} lượt)</div>
                                      <div className="font-bold text-xs truncate">{item.n}</div>
                                      {s.type === 'draw' && String(item.v).startsWith('data:image/jpeg') ? (
                                        <img src={item.v} alt="Top drawing" className="w-full h-28 object-contain bg-white rounded-lg border border-stone-200 mx-auto" />
                                      ) : (
                                        <div className="text-sm italic font-medium">"{item.v}"</div>
                                      )}
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}

                          {/* Full submission list */}
                          <div className={s.type === 'draw' ? "grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4" : "grid grid-cols-1 md:grid-cols-2 gap-3"}>
                            {sortedSubmissions.map((item, j) => {
                              const rank = item.votesCount > 0 ? sortedSubmissions.findIndex(x => x.votesCount === item.votesCount) : -1;
                              const isTop3 = rank >= 0 && rank < 3 && item.votesCount > 0;
                              return s.type === 'draw' ? (
                                <div key={j} className={`bg-stone-50 dark:bg-slate-950 border ${isTop3 ? 'border-amber-400 ring-2 ring-amber-400 shadow-lg' : 'border-stone-200 dark:border-slate-800'} p-3 rounded-2xl space-y-2 relative`}>
                                  {isTop3 && (
                                    <div className="absolute -top-3 -right-2 px-2.5 py-0.5 bg-amber-500 text-white text-[11px] font-black rounded-full shadow">
                                      {rank === 0 ? '🥇 Top 1' : rank === 1 ? '🥈 Top 2' : '🥉 Top 3'} ({item.votesCount})
                                    </div>
                                  )}
                                  <img src={item.v} alt="Drawing" className="w-full h-40 object-contain bg-white rounded-xl border border-stone-200" />
                                  <div className="text-sm font-bold truncate">
                                    {item.n} {item.cls && <span className="text-xs font-semibold text-purple-600">({item.cls})</span>}
                                  </div>
                                  <div className="flex items-center justify-between text-xs font-bold text-purple-600 pt-1">
                                    <span>{item.votesCount} bình chọn</span>
                                    {item.g !== undefined ? <span className="text-emerald-600">{item.g}đ</span> : (
                                      <button onClick={() => gradeAnswer(j, 1)} className="px-2 py-0.5 bg-purple-600 text-white rounded text-[11px]">Chấm 1đ</button>
                                    )}
                                  </div>
                                </div>
                              ) : (
                                <div key={j} className={`p-4 bg-purple-50/60 dark:bg-purple-950/20 border ${isTop3 ? 'border-amber-400 ring-2 ring-amber-400 shadow-lg' : 'border-purple-100 dark:border-purple-900/40'} rounded-2xl space-y-2 relative`}>
                                  {isTop3 && (
                                    <div className="absolute -top-3 right-3 px-2.5 py-0.5 bg-amber-500 text-white text-[11px] font-black rounded-full shadow">
                                      {rank === 0 ? '🥇 Top 1' : rank === 1 ? '🥈 Top 2' : '🥉 Top 3'} ({item.votesCount})
                                    </div>
                                  )}
                                  <div className="text-lg">
                                    <b>{item.n}{item.cls ? ` (${item.cls})` : ''}:</b> {item.v}
                                  </div>
                                  <div className="flex items-center justify-between text-xs font-bold text-purple-600 pt-1">
                                    <span>{item.votesCount} bình chọn</span>
                                    {item.g !== undefined ? <span className="text-emerald-600">{item.g} điểm</span> : (
                                      <button onClick={() => gradeAnswer(j, 1)} className="px-2.5 py-0.5 bg-purple-600 text-white rounded text-[11px]">Chấm 1đ</button>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      ) : (
                        <p className="text-stone-400 italic py-6 text-center">Đang chờ học sinh gửi bài...</p>
                      );
                    }

                    if (s.type === 'cloud') {
                      const counts: Record<string, number> = {};
                      r.forEach(a => {
                        const w = String(a.v).trim().toLowerCase();
                        if (w) counts[w] = (counts[w] || 0) + 1;
                      });
                      const entries = Object.entries(counts);
                      return entries.length ? (
                        <div className="flex flex-wrap gap-3 py-8 justify-center items-center">
                          {entries.map(([w, n], idx) => (
                            <span key={idx} style={{ fontSize: `${Math.min(54, 18 + n * 10)}px` }} className="font-extrabold text-purple-600 dark:text-purple-400 px-4 py-1.5 bg-purple-50 dark:bg-purple-950/50 rounded-2xl shadow-sm">
                              {w}
                            </span>
                          ))}
                        </div>
                      ) : (
                        <p className="text-stone-400 italic py-6 text-center">Đang chờ học sinh gửi từ khóa...</p>
                      );
                    }

                    // Poll, Quiz, Scale
                    const O = s.type === 'scale' ? ['1', '2', '3', '4', '5'] : s.opts || [];
                    const off = s.type === 'scale' ? 1 : 0;
                    const avg = (off && r.length) ? (r.reduce((acc, a) => acc + Number(a.v), 0) / r.length).toFixed(1) : null;

                    return (
                      <div className="space-y-3">
                        {avg && <p className="text-base font-semibold text-stone-600 dark:text-stone-400">Trung bình: <span className="text-purple-600 text-2xl font-bold">{avg}</span> / 5</p>}
                        {O.map((o: string, k: number) => {
                          const n = r.filter(a => Number(a.v) === (k + off)).length;
                          const p = r.length ? (n / r.length) * 100 : 0;
                          const isCorrect = s.type === 'quiz' && k === s.a;
                          return (
                            <div key={k} className={`relative overflow-hidden rounded-2xl border ${isCorrect ? 'border-emerald-500 bg-emerald-50/20' : 'border-stone-200 dark:border-slate-800'} p-4`}>
                              <div className={`absolute left-0 top-0 bottom-0 transition-all duration-300 ${isCorrect ? 'bg-emerald-500/20' : 'bg-purple-500/20'}`} style={{ width: `${p}%` }} />
                              <div className="relative flex justify-between items-center text-lg font-semibold">
                                <span>{o} {isCorrect && '✓ (Đáp án đúng)'}</span>
                                <span className="font-bold">{n} ({p.toFixed(0)}%)</span>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    );
                  })()}
                </div>
              </div>

              {/* AI Feedback Box */}
              {aiFeedbackText && (
                <div className="bg-purple-50 dark:bg-purple-950/40 border border-purple-200 dark:border-purple-900 p-6 rounded-2xl space-y-2">
                  <div className="flex justify-between items-center">
                    <h4 className="font-bold text-purple-700 dark:text-purple-300 flex items-center gap-2">
                      <Sparkles className="w-5 h-5" /> Đánh giá & Phản hồi từ AI
                    </h4>
                    <button onClick={() => setAiFeedbackText('')} className="text-xs text-stone-500 hover:text-stone-800">Đóng</button>
                  </div>
                  <div className="text-sm leading-relaxed whitespace-pre-line text-stone-800 dark:text-slate-200">{aiFeedbackText}</div>
                </div>
              )}

              {/* Scoreboard Panel */}
              {showScoreboard && (
                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-lg space-y-4">
                  <div className="flex justify-between items-center">
                    <h3 className="text-xl font-bold flex items-center gap-2">
                      <Trophy className="w-6 h-6 text-amber-500" /> Bảng điểm lớp học
                    </h3>
                    <button onClick={() => setShowScoreboard(false)} className="px-3 py-1 bg-stone-100 dark:bg-slate-800 text-xs font-semibold rounded-lg">Đóng</button>
                  </div>
                  {(() => {
                    const { t, nm } = calculateScores();
                    const sorted = Object.entries(t).sort((x, y) => y[1] - x[1]);
                    return sorted.length ? (
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                        {sorted.map(([id, score], i) => (
                          <div key={id} className="flex justify-between items-center p-3 bg-stone-50 dark:bg-slate-950 rounded-xl">
                            <span className="font-bold">
                              {i + 1}. {nm[id]}
                              {allSessionStudentClasses[id] && (
                                <span className="ml-1.5 text-xs font-semibold text-purple-600">({allSessionStudentClasses[id]})</span>
                              )}
                            </span>
                            <span className="font-extrabold text-purple-600">{score} điểm</span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <p className="text-stone-500 italic">Chưa có điểm tổng kết.</p>
                    );
                  })()}
                </div>
              )}
            </div>

            {/* Bottom Presentation Navigation Bar */}
            <footer className="bg-white/90 dark:bg-slate-900/90 backdrop-blur border-t border-stone-200 dark:border-slate-800 px-4 sm:px-8 py-3 sticky bottom-0 z-40">
              <div className="max-w-6xl mx-auto flex justify-between items-center gap-4">
                <button onClick={() => moveSlide(-1)} disabled={slideIdx === 0} className="px-5 py-2.5 bg-stone-100 dark:bg-slate-800 font-bold rounded-xl disabled:opacity-40 hover:bg-stone-200 transition flex items-center gap-2">
                  <ArrowLeft className="w-4 h-4" /> Slide trước
                </button>
                <div className="flex items-center gap-1.5 overflow-x-auto max-w-full py-1">
                  {lesson.slides.map((_, i) => (
                    <button
                      key={i}
                      onClick={() => moveSlide(i - slideIdx)}
                      className={`w-8 h-8 rounded-lg text-xs font-bold transition ${i === slideIdx ? 'bg-purple-600 text-white shadow' : 'bg-stone-100 dark:bg-slate-800 text-stone-600 dark:text-slate-400 hover:bg-purple-100'}`}
                    >
                      {i + 1}
                    </button>
                  ))}
                </div>
                <button onClick={() => moveSlide(1)} disabled={slideIdx === lesson.slides.length - 1} className="px-5 py-2.5 bg-purple-600 text-white font-bold rounded-xl hover:bg-purple-700 disabled:opacity-40 transition flex items-center gap-2">
                  Slide tiếp <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            </footer>

            {/* QR Modal for latecomers */}
            {showQrModal && (
              <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4" onClick={() => setShowQrModal(false)}>
                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-3xl max-w-sm w-full text-center space-y-4 shadow-2xl" onClick={e => e.stopPropagation()}>
                  <div className="flex justify-between items-center">
                    <h3 className="font-bold text-lg">Tham gia lớp học</h3>
                    <button onClick={() => setShowQrModal(false)} className="p-1.5 rounded-lg hover:bg-stone-100 dark:hover:bg-slate-800">
                      <X className="w-5 h-5" />
                    </button>
                  </div>
                  {qrDataUrl && (
                    <img alt="QR vào lớp" src={qrDataUrl} className="w-48 h-48 mx-auto rounded-xl border border-stone-200 bg-white p-1" />
                  )}
                  <div className="text-4xl font-black tracking-widest text-purple-600 font-mono">{roomCode}</div>
                  <p className="text-xs text-stone-500 break-all">{roomLink}</p>
                </div>
              </div>
            )}

            {/* Cloud Sync & Privacy Settings Modal */}
            {showCloudModal && (
              <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4" onClick={() => setShowCloudModal(false)}>
                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-3xl max-w-lg w-full space-y-5 shadow-2xl" onClick={e => e.stopPropagation()}>
                  <div className="flex justify-between items-center">
                    <h3 className="font-bold text-xl flex items-center gap-2">
                      <Cloud className="w-5 h-5 text-purple-600" /> Đồng bộ đám mây & Quyền riêng tư
                    </h3>
                    <button onClick={() => setShowCloudModal(false)} className="p-1.5 rounded-lg hover:bg-stone-100 dark:hover:bg-slate-800">
                      <X className="w-5 h-5" />
                    </button>
                  </div>

                  <div className="space-y-3 text-sm text-stone-600 dark:text-slate-300">
                    <p>
                      Đăng nhập tài khoản Google để lưu trữ và đồng bộ tự động bài giảng & báo cáo thống kê lên đám mây an toàn, truy cập mượt mà trên mọi thiết bị và Vercel.
                    </p>

                    {cloudSyncMsg && (
                      <p className="text-xs font-semibold text-purple-600 dark:text-purple-400">{cloudSyncMsg}</p>
                    )}

                    <div className="space-y-2 pt-2 border-t border-stone-200 dark:border-slate-800">
                      <label className="text-xs font-bold text-stone-500 dark:text-slate-400">Mã đồng bộ / Email của bạn:</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={cloudKeyInput}
                          onChange={e => setCloudKeyInput(e.target.value)}
                          placeholder="vd: quangminh1485"
                          className="flex-1 font-mono text-sm uppercase px-3 py-2 rounded-xl border border-stone-300 dark:border-slate-700 bg-white dark:bg-slate-800"
                        />
                        <button
                          onClick={handleSaveCloudKeyAndSync}
                          disabled={cloudSyncing}
                          className="px-4 py-2 bg-purple-600 text-white font-bold rounded-xl text-xs hover:bg-purple-700 transition"
                        >
                          Lưu & Đồng bộ
                        </button>
                      </div>
                      <p className="text-[11px] text-stone-400">Mặc định: quangminh1485. Dùng chung mã này trên mọi thiết bị để đồng bộ tự động.</p>
                    </div>

                    <div className="pt-3 border-t border-stone-200 dark:border-slate-800 space-y-2">
                      <label className="text-xs font-bold text-stone-500 dark:text-slate-400">Hoặc đăng nhập bằng Google (tùy chọn):</label>
                      <button
                        onClick={handleCloudLogin}
                        disabled={cloudSyncing}
                        className="w-full py-2.5 bg-white dark:bg-slate-800 border border-stone-300 dark:border-slate-700 hover:bg-stone-50 text-stone-800 dark:text-slate-200 font-bold rounded-xl transition flex items-center justify-center gap-2 shadow-xs"
                      >
                        <svg className="w-4 h-4" viewBox="0 0 24 24">
                          <path fill="#4285F4" d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.52-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.66-5.17 3.66-9.17z"/>
                          <path fill="#34A853" d="M12 24c3.24 0 5.95-1.08 7.93-2.91l-3.88-3.05c-1.08.72-2.45 1.16-4.05 1.16-3.13 0-5.78-2.11-6.73-4.96H1.18v3.15C3.15 21.32 7.23 24 12 24z"/>
                          <path fill="#FBBC05" d="M5.27 14.24c-.25-.72-.38-1.49-.38-2.24s.13-1.52.38-2.24V6.61H1.18C.43 8.13 0 9.83 0 12s.43 3.87 1.18 5.39l4.09-3.15z"/>
                          <path fill="#EA4335" d="M12 4.75c1.77 0 3.35.61 4.6 1.8l3.42-3.42C17.95 1.19 15.24 0 12 0 7.23 0 3.15 2.68 1.18 6.61l4.09 3.15c.95-2.85 3.6-4.96 6.73-4.96z"/>
                        </svg>
                        {cloudUser ? `Đang kết nối (${cloudUser.email})` : 'Đăng nhập tài khoản Google'}
                      </button>
                    </div>

                    <div className="pt-2 border-t border-stone-200 dark:border-slate-800 flex flex-wrap gap-2 justify-between items-center">
                      <div className="flex gap-2 flex-wrap">
                        <button
                          onClick={handleExportBackupFile}
                          disabled={cloudSyncing}
                          className="px-3.5 py-2 rounded-xl bg-stone-100 dark:bg-slate-800 text-xs font-semibold hover:bg-stone-200 transition flex items-center gap-1.5"
                        >
                          <Save className="w-3.5 h-3.5" /> Xuất file sao lưu (.json)
                        </button>
                        <label className="px-3.5 py-2 rounded-xl bg-stone-100 dark:bg-slate-800 text-xs font-semibold hover:bg-stone-200 transition flex items-center gap-1.5 cursor-pointer">
                          <FolderOpen className="w-3.5 h-3.5" /> Khôi phục sao lưu
                          <input type="file" accept=".json" onChange={e => e.target.files?.[0] && handleImportBackupFile(e.target.files[0])} className="hidden" />
                        </label>
                      </div>
                      <button
                        onClick={() => syncWithCloudNow(true)}
                        disabled={cloudSyncing}
                        className="px-4 py-2 rounded-xl bg-emerald-600 text-white text-xs font-bold hover:bg-emerald-700 transition flex items-center gap-1.5"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 ${cloudSyncing ? 'animate-spin' : ''}`} /> Đồng bộ ngay
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* JOIN / STUDENT LESSON VIEW */}
        {view === 'join' && (
          <div className="space-y-6 max-w-2xl mx-auto py-4">
            {!studentRoomJoined ? (
              /* STEP 1: Enter 5-character class code or scan QR */
              <div className="space-y-5">
                <h2 className="text-2xl font-bold">Tham gia lớp học</h2>
                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-4">
                  <p className="text-sm text-stone-500 dark:text-slate-400">
                    Quét mã QR của giáo viên hoặc nhập mã lớp (chữ và số, VD: <b>11A1</b>, <b>10A2</b>, <b>AB123</b>) để tiếp tục nhập tên và lớp của em.
                  </p>
                  <div className="flex gap-3 flex-wrap">
                    <input 
                      type="text" 
                      value={joinCodeInput} 
                      onChange={e => {
                        const val = normalizeClassCode(e.target.value);
                        setJoinCodeInput(val);
                      }}
                      onKeyDown={e => e.key === 'Enter' && joinClassroom()}
                      placeholder="Mã lớp (vd: 11A1, AB123)" 
                      maxLength={15} 
                      className="uppercase font-mono font-bold tracking-widest flex-1 min-w-[180px] text-lg" 
                    />
                    <button onClick={() => joinClassroom()} className="px-6 py-2.5 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl transition">
                      Tiếp tục
                    </button>
                  </div>
                  {joinMsg && (
                    <div className="flex items-center gap-2 text-sm font-medium text-purple-600 dark:text-purple-400">
                      <span>{joinMsg}</span>
                    </div>
                  )}
                </div>
              </div>
            ) : !studentProfileConfirmed ? (
              /* STEP 2: After scanning QR / entering code -> Enter Student Name & Class */
              <div className="space-y-5">
                <div className="flex items-center justify-between flex-wrap gap-2">
                  <h2 className="text-2xl font-bold">Thông tin học sinh</h2>
                  <button
                    onClick={leaveStudentClassroom}
                    className="px-3 py-1.5 text-xs font-semibold border border-stone-300 dark:border-slate-700 rounded-xl hover:bg-stone-100 dark:hover:bg-slate-800 transition"
                  >
                    Đổi mã lớp ({joinCodeInput})
                  </button>
                </div>

                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 sm:p-8 rounded-3xl shadow-sm space-y-5">
                  <div className="flex items-center justify-between flex-wrap gap-2 pb-3 border-b border-stone-200 dark:border-slate-800">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="px-3 py-1 rounded-full bg-purple-600 text-white font-mono text-xs font-black tracking-widest">
                        MÃ LỚP: {joinCodeInput}
                      </span>
                      <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold ${studentConnected ? 'bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300' : 'bg-amber-100 dark:bg-amber-950/60 text-amber-700 dark:text-amber-300'}`}>
                        <span className={`w-2 h-2 rounded-full ${studentConnected ? 'bg-emerald-500 animate-pulse' : 'bg-amber-500 animate-pulse'}`} />
                        {studentConnected ? 'Đã kết nối lớp học' : (joinMsg || 'Đang kết nối lớp học...')}
                      </span>
                    </div>
                    {studentCurState?.L?.title && (
                      <span className="text-xs font-bold text-stone-500 dark:text-slate-400">
                        Bài giảng: {studentCurState.L.title}
                      </span>
                    )}
                  </div>

                  <p className="text-sm text-stone-600 dark:text-slate-400">
                    Hãy nhập <b>Họ và tên</b> cùng <b>Tên lớp</b> của em để điểm danh trước khi vào bài giảng.
                  </p>

                  <div className="space-y-4">
                    <div>
                      <label className="block text-xs font-bold text-stone-600 dark:text-slate-300 mb-1.5">
                        Họ và tên của em <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="text"
                        value={studentName}
                        onChange={e => handleStudentNameChange(e.target.value)}
                        onKeyDown={e => e.key === 'Enter' && confirmStudentProfile()}
                        placeholder="Ví dụ: Nguyễn Văn An"
                        autoFocus
                        className="w-full"
                      />
                    </div>

                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <label className="block text-xs font-bold text-stone-600 dark:text-slate-300">
                          Lớp của em <span className="text-red-500">*</span>
                        </label>
                        {studentCurState?.L?.className && studentClassName !== studentCurState.L.className && (
                          <button
                            type="button"
                            onClick={() => handleStudentClassChange(studentCurState.L.className)}
                            className="text-xs font-bold text-purple-600 hover:underline"
                          >
                            Chọn nhanh: {studentCurState.L.className}
                          </button>
                        )}
                      </div>
                      <input
                        type="text"
                        value={studentClassName}
                        onChange={e => handleStudentClassChange(e.target.value)}
                        onKeyDown={e => e.key === 'Enter' && confirmStudentProfile()}
                        placeholder={studentCurState?.L?.className ? `Ví dụ: ${studentCurState.L.className}` : 'Ví dụ: 11A1'}
                        className="w-full"
                      />
                    </div>
                  </div>

                  {studentProfileError && (
                    <div className="p-3 rounded-xl bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-900 text-xs font-semibold text-red-600 dark:text-red-300">
                      {studentProfileError}
                    </div>
                  )}

                  {joinMsg && !studentConnected && (
                    <div className="text-xs font-medium text-amber-600 dark:text-amber-400">
                      {joinMsg}
                    </div>
                  )}

                  <button
                    onClick={confirmStudentProfile}
                    className="w-full py-3 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl shadow transition"
                  >
                    Xác nhận & Vào phòng chờ
                  </button>
                </div>
              </div>
            ) : (!studentConnected || !studentCurState || !studentCurState.started) ? (
              /* STEP 3: Waiting Room until Teacher clicks "Vào bài giảng" */
              <div className="space-y-5">
                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-8 rounded-3xl shadow-sm text-center space-y-5">
                  <div className="w-16 h-16 rounded-2xl bg-purple-100 dark:bg-purple-950/80 text-purple-600 dark:text-purple-300 flex items-center justify-center mx-auto">
                    <Clock className="w-8 h-8 animate-pulse" />
                  </div>

                  <div className="space-y-2">
                    <div className="inline-flex items-center gap-2 px-3.5 py-1 rounded-full bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300 text-xs font-extrabold">
                      <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                      {studentConnected ? 'Đã báo danh vào lớp thành công' : 'Đang kết nối với máy giáo viên...'}
                    </div>
                    <h2 className="text-2xl font-extrabold">
                      {studentCurState?.L?.title || `Phòng học ${joinCodeInput}`}
                    </h2>
                    <p className="text-sm text-stone-600 dark:text-slate-400 max-w-md mx-auto">
                      Đang đợi giáo viên bấm <b>Vào bài giảng</b>. Màn hình bài giảng và câu hỏi tương tác sẽ tự động mở ngay khi giáo viên bắt đầu!
                    </p>
                  </div>

                  <div className="p-4 rounded-2xl bg-stone-50 dark:bg-slate-950 border border-stone-200 dark:border-slate-800 max-w-md mx-auto flex items-center justify-between flex-wrap gap-3 text-left">
                    <div>
                      <div className="text-xs text-stone-500">Học sinh tham gia</div>
                      <div className="font-bold text-sm">
                        {studentName} <span className="text-purple-600 dark:text-purple-400">· Lớp {studentClassName}</span>
                      </div>
                      <div className="text-[11px] text-stone-400 font-mono">Mã phòng: {joinCodeInput}</div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setStudentProfileConfirmed(false)}
                        className="px-3 py-1.5 text-xs font-bold rounded-xl bg-white dark:bg-slate-900 border border-stone-300 dark:border-slate-700 hover:bg-stone-100 transition"
                      >
                        Sửa tên / lớp
                      </button>
                      <button
                        onClick={leaveStudentClassroom}
                        className="px-3 py-1.5 text-xs font-semibold rounded-xl border border-stone-300 dark:border-slate-700 text-stone-500 hover:bg-stone-100 transition"
                      >
                        Đổi lớp
                      </button>
                    </div>
                  </div>

                  {joinMsg && (
                    <p className="text-xs font-medium text-amber-600 dark:text-amber-400">{joinMsg}</p>
                  )}
                </div>
              </div>
            ) : (
              /* STEP 4: Teacher has started the live lesson */
              <div className="space-y-4">
                {/* Compact Lesson Top Bar for Student */}
                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 px-5 py-3.5 rounded-2xl shadow-sm flex flex-wrap items-center justify-between gap-3">
                  <div className="min-w-0">
                    <div className="text-xs font-bold text-purple-600 dark:text-purple-400 uppercase tracking-wider">
                      {studentCurState.L.className ? `${studentCurState.L.className} · ` : ''}Mã {joinCodeInput} · Slide {studentCurState.idx + 1}/{studentCurState.L.slides.length}
                    </div>
                    <div className="font-bold text-base truncate">
                      {studentCurState.L.title || 'Bài giảng trực tiếp'}
                    </div>
                  </div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <input
                      type="text"
                      value={studentName}
                      onChange={e => handleStudentNameChange(e.target.value)}
                      placeholder="Tên của em..."
                      className="text-xs px-3 py-1.5 w-32 sm:w-40 rounded-lg border border-stone-300 dark:border-slate-700 bg-stone-50 dark:bg-slate-950"
                    />
                    <input
                      type="text"
                      value={studentClassName}
                      onChange={e => handleStudentClassChange(e.target.value)}
                      placeholder="Lớp (vd: 11A1)"
                      className="text-xs px-3 py-1.5 w-28 sm:w-32 rounded-lg border border-stone-300 dark:border-slate-700 bg-stone-50 dark:bg-slate-950"
                    />
                    {studentMyScore !== null && (
                      <span className="px-3 py-1.5 bg-purple-50 dark:bg-purple-950/60 text-purple-700 dark:text-purple-300 rounded-lg text-xs font-bold">
                        {studentMyScore} điểm
                      </span>
                    )}
                    <button
                      onClick={leaveStudentClassroom}
                      className="px-3 py-1.5 text-xs font-semibold border border-stone-300 dark:border-slate-700 rounded-lg hover:bg-stone-100 dark:hover:bg-slate-800 transition"
                    >
                      Đổi lớp
                    </button>
                  </div>
                </div>

                {joinMsg && <p className="text-xs font-medium text-amber-600 dark:text-amber-400 px-1">{joinMsg}</p>}

                {/* Direct Lesson Slide Card */}
                <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-2xl shadow-sm space-y-4">
                  {(() => {
                    const s = studentCurState.L.slides[studentCurState.idx];
                    const k = studentCurState.idx;
                    const mySubmittedVal = studentMyAns[k];

                    if (!s) return <p className="text-stone-500">Đang chờ giáo viên trình chiếu slide...</p>;

                    return (
                      <div className="space-y-4">
                        <div className="flex items-center justify-between flex-wrap gap-2">
                          <span className="px-3 py-1 bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 rounded-full text-xs font-bold">
                            {SLIDE_TYPES[s.type]}
                          </span>
                          {s.type !== 'info' && s.type !== 'media' && studentTimerLeft !== null && (
                            <span className={`px-3 py-1 rounded-full text-xs font-mono font-black inline-flex items-center gap-1.5 ${studentTimerLeft === 0 ? 'bg-red-100 text-red-700 dark:bg-red-950/60 dark:text-red-300' : studentTimerLeft <= 10 ? 'bg-amber-100 text-amber-700' : 'bg-purple-50 dark:bg-purple-950/60 text-purple-700 dark:text-purple-300'}`}>
                              <Clock className="w-3.5 h-3.5" />
                              {studentTimerLeft === 0 ? 'Hết giờ!' : formatTimer(studentTimerLeft)}
                            </span>
                          )}
                        </div>
                        {s.type !== 'info' && s.type !== 'media' && studentTimerLeft !== null && studentTimerTotal > 0 && (
                          <div className="w-full h-1.5 bg-stone-100 dark:bg-slate-800 rounded-full overflow-hidden">
                            <div
                              className={`h-full transition-all duration-500 ${studentTimerLeft === 0 ? 'bg-red-500' : studentTimerLeft <= 10 ? 'bg-amber-500' : 'bg-purple-600'}`}
                              style={{ width: `${Math.max(0, Math.min(100, Math.round((studentTimerLeft / studentTimerTotal) * 100)))}%` }}
                            />
                          </div>
                        )}
                        {!s.pptx && <h3 className="text-xl sm:text-2xl font-bold">{s.q}</h3>}

                        {studentVotingActive && (s.type === 'draw' || s.type === 'open') ? (
                          <div className="space-y-4 py-2">
                            <div className="p-4 bg-amber-50 dark:bg-amber-950/40 border border-amber-300 dark:border-amber-700 rounded-2xl text-amber-800 dark:text-amber-200 font-bold text-center space-y-1">
                              <div className="text-base">🗳️ Vòng bình chọn bài xuất sắc!</div>
                              <div className="text-xs font-normal">Hãy bấm chọn bức tranh hoặc câu trả lời đẹp/hay nhất của bạn bè:</div>
                            </div>
                            <div className="space-y-3">
                              {studentSubmissions.map((item: any, idx: number) => {
                                const isMyVote = studentMyVote === item.sid;
                                return (
                                  <div key={idx} className={`p-4 rounded-2xl border transition ${isMyVote ? 'border-purple-600 bg-purple-50 dark:bg-purple-950/60 ring-2 ring-purple-500' : 'border-stone-200 dark:border-slate-800 bg-white dark:bg-slate-900'} space-y-2`}>
                                    <div className="text-sm font-bold text-stone-700 dark:text-slate-300">{item.n}</div>
                                    {s.type === 'draw' && String(item.v).startsWith('data:image/jpeg') ? (
                                      <img src={item.v} alt="Peer submission" className="w-full h-36 object-contain bg-white rounded-xl border border-stone-200 mx-auto" />
                                    ) : (
                                      <div className="text-base font-medium">"{item.v}"</div>
                                    )}
                                    <button
                                      onClick={() => sendStudentVote(item.sid)}
                                      className={`w-full py-2.5 rounded-xl font-bold text-xs transition flex items-center justify-center gap-1.5 ${isMyVote ? 'bg-purple-600 text-white shadow' : 'bg-stone-100 dark:bg-slate-800 text-stone-700 dark:text-stone-300 hover:bg-purple-100'}`}
                                    >
                                      {isMyVote ? <Check className="w-4 h-4" /> : <Award className="w-4 h-4" />}
                                      {isMyVote ? 'Đã bình chọn bài này' : 'Bình chọn bài này'}
                                    </button>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        ) : mySubmittedVal !== undefined ? (
                          <div className="p-4 bg-emerald-50 dark:bg-emerald-950/30 border border-emerald-200 rounded-xl text-emerald-700 dark:text-emerald-300 font-medium">
                            Em đã gửi câu trả lời. Chờ giáo viên chuyển slide...
                          </div>
                        ) : (s.type !== 'info' && s.type !== 'media' && studentTimerLeft === 0) ? (
                          <div className="p-4 bg-red-50 dark:bg-red-950/30 border border-red-200 dark:border-red-900 rounded-xl text-red-700 dark:text-red-300 font-semibold">
                            Đã hết thời gian trả lời cho câu hỏi này. Chờ giáo viên chuyển slide hoặc cộng thêm giờ...
                          </div>
                        ) : (
                          <>
                            {s.type === 'media' && renderMediaSlide(s, false, `stu-${k}`)}
                            {s.type === 'info' && <p className="text-lg whitespace-pre-line">{s.b || ''}</p>}
                            {s.type === 'scale' && (
                              <div className="space-y-3">
                                <div className="flex gap-2 justify-center">
                                  {[1, 2, 3, 4, 5].map((n: number) => (
                                    <button key={n} onClick={() => submitStudentAnswer(k, n)} className="w-12 h-12 bg-purple-50 dark:bg-purple-950 hover:bg-purple-600 hover:text-white font-bold rounded-xl text-lg transition">
                                      {n}
                                    </button>
                                  ))}
                                </div>
                                <p className="text-xs text-stone-500 text-center">1 = hoàn toàn không, 5 = hoàn toàn có</p>
                              </div>
                            )}
                            {s.type === 'draw' && (
                              <div className="space-y-3 text-center">
                                <canvas 
                                  ref={(el: HTMLCanvasElement | null) => {
                                    if (el && !el.dataset.init) {
                                      el.dataset.init = 'true';
                                      const ctx = el.getContext('2d');
                                      if (ctx) {
                                        ctx.fillStyle = '#ffffff';
                                        ctx.fillRect(0, 0, el.width, el.height);
                                        ctx.lineWidth = 3;
                                        ctx.lineCap = 'round';
                                        ctx.strokeStyle = '#1e1b2e';
                                        let drawing = false;
                                        const getPos = (e: PointerEvent) => {
                                          const r = el.getBoundingClientRect();
                                          return [(e.clientX - r.left) * el.width / r.width, (e.clientY - r.top) * el.height / r.height];
                                        };
                                        el.onpointerdown = e => { drawing = true; const [x, y] = getPos(e); ctx.beginPath(); ctx.moveTo(x, y); };
                                        el.onpointermove = e => { if (!drawing) return; const [x, y] = getPos(e); ctx.lineTo(x, y); ctx.stroke(); };
                                        el.onpointerup = el.onpointerleave = () => { drawing = false; };
                                      }
                                    }
                                    studentCanvasRef.current = el;
                                  }}
                                  width={320} 
                                  height={220} 
                                  className="bg-white border border-stone-200 rounded-xl mx-auto touch-none shadow-inner"
                                />
                                <div className="flex justify-center gap-2">
                                  <button onClick={() => {
                                    const c = studentCanvasRef.current;
                                    if (c) {
                                      const ctx = c.getContext('2d');
                                      if (ctx) {
                                        ctx.fillStyle = '#ffffff';
                                        ctx.fillRect(0, 0, c.width, c.height);
                                      }
                                    }
                                  }} className="px-4 py-2 bg-stone-200 dark:bg-slate-800 text-xs font-bold rounded-xl">
                                    Xóa vẽ lại
                                  </button>
                                  <button onClick={() => {
                                    const c = studentCanvasRef.current;
                                    if (c) submitStudentAnswer(k, c.toDataURL('image/jpeg', 0.5));
                                  }} className="px-4 py-2 bg-purple-600 text-white text-xs font-bold rounded-xl">
                                    Gửi bản vẽ
                                  </button>
                                </div>
                              </div>
                            )}
                            {s.opts && (
                              <div className="space-y-2">
                                {s.opts.map((o: string, j: number) => (
                                  <button key={j} onClick={() => submitStudentAnswer(k, j)} className="w-full text-left p-4 bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 rounded-xl hover:border-purple-500 font-semibold transition">
                                    {o}
                                  </button>
                                ))}
                              </div>
                            )}
                            {!s.opts && s.type !== 'scale' && s.type !== 'draw' && s.type !== 'media' && s.type !== 'info' && (
                              <div className="space-y-3">
                                {s.type === 'open' ? (
                                  <textarea id={`sa-${k}`} rows={3} placeholder="Câu trả lời của em..." />
                                ) : (
                                  <input id={`sa-${k}`} type="text" placeholder="Một từ hoặc cụm từ ngắn..." />
                                )}
                                <button onClick={() => {
                                  const val = (document.getElementById(`sa-${k}`) as HTMLInputElement | HTMLTextAreaElement)?.value || '';
                                  submitStudentAnswer(k, val);
                                }} className="w-full py-3 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl transition">
                                  Gửi câu trả lời
                                </button>
                              </div>
                            )}
                          </>
                        )}
                      </div>
                    );
                  })()}
                </div>
              </div>
            )}
          </div>
        )}

        {/* POST-LESSON CLASS INTERACTION STATISTICS VIEW */}
        {view === 'stats' && (
          <div className="space-y-6 py-4">
            <div className="flex justify-between items-center flex-wrap gap-3">
              <div>
                <h2 className="text-2xl font-bold flex items-center gap-2">
                  <BarChart2 className="w-6 h-6 text-purple-600" /> Thống kê & Hoạt động bài giảng
                </h2>
                <p className="text-sm text-stone-500 dark:text-slate-400">
                  Toàn bộ bài giảng, câu trả lời của học sinh và mức độ tương tác (%) theo tên lớp đều được lưu tự động.
                </p>
              </div>
              <div className="flex gap-2">
                <button onClick={() => setView('build')} className="px-4 py-2 text-sm font-semibold rounded-xl border border-stone-300 dark:border-slate-700 hover:bg-stone-100 dark:hover:bg-slate-800 transition">
                  Quay lại soạn bài
                </button>
                <button onClick={loadDashboard} className="px-4 py-2 text-sm font-semibold rounded-xl bg-purple-600 text-white hover:bg-purple-700 transition">
                  Bảng điều khiển
                </button>
              </div>
            </div>

            {sessionReports.length === 0 ? (
              <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-10 rounded-3xl text-center space-y-4">
                <p className="text-stone-500">Chưa có buổi dạy nào được ghi nhận. Bạn có {mine.length} bài giảng đã lưu sẵn sàng giảng dạy.</p>
                {mine.length > 0 && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 max-w-2xl mx-auto text-left">
                    {mine.map((l, idx) => (
                      <div key={l.id || idx} className="p-4 rounded-2xl border border-stone-200 dark:border-slate-800 bg-stone-50 dark:bg-slate-950 flex items-center justify-between gap-3">
                        <div className="min-w-0">
                          <div className="font-bold text-sm truncate">{l.title || 'Chưa đặt tên'}</div>
                          <div className="text-xs text-stone-500">
                            {l.className ? `${l.className} · ` : ''}{l.slides.length} slide · {l.slides.filter(s => s.type !== 'info' && s.type !== 'media').length} câu hỏi tương tác
                          </div>
                        </div>
                        <button onClick={() => openSavedLesson(idx)} className="px-3 py-1.5 bg-purple-600 text-white text-xs font-bold rounded-lg shrink-0">
                          Mở dạy
                        </button>
                      </div>
                    ))}
                  </div>
                )}
                <button onClick={() => setView('build')} className="px-5 py-2.5 bg-purple-600 text-white text-sm font-bold rounded-xl">
                  Mở bài giảng để bắt đầu dạy
                </button>
              </div>
            ) : (
              (() => {
                const allClassLabels = Array.from(
                  new Set(
                    sessionReports.flatMap(r => [
                      r.className,
                      r.roomCode ? `Mã ${r.roomCode}` : '',
                      ...(r.studentClassNames || []),
                      ...r.students.map(st => st.className || '').filter(Boolean),
                    ]).filter(Boolean)
                  )
                );
                const classNames = ['ALL', ...allClassLabels];
                const matchesClass = (r: ClassSessionReport, cls: string) =>
                  cls === 'ALL' ||
                  r.className === cls ||
                  (r.roomCode && `Mã ${r.roomCode}` === cls) ||
                  r.roomCode === cls ||
                  r.studentClassNames?.includes(cls) ||
                  r.students.some(st => st.className === cls);

                const filteredReports = sessionReports.filter(r => matchesClass(r, statsClassFilter));
                const selectedReport = filteredReports.find(r => r.id === activeReportId) || filteredReports[0];

                const classSummaries = allClassLabels.map(cls => {
                  const list = sessionReports.filter(r => matchesClass(r, cls));
                  const avgPct = Math.round(list.reduce((acc, r) => acc + r.overallParticipationPct, 0) / list.length);
                  const roomCodes = Array.from(new Set(list.map(r => r.roomCode).filter(Boolean)));
                  return { className: cls, sessions: list.length, avgPct, roomCodes };
                });

                const levelBadge = (pct: number) => {
                  if (pct >= 80) return { label: 'Rất tích cực', cls: 'bg-emerald-100 text-emerald-700 dark:bg-emerald-950/60 dark:text-emerald-300', bar: 'bg-emerald-500' };
                  if (pct >= 50) return { label: 'Tương tác khá', cls: 'bg-purple-100 text-purple-700 dark:bg-purple-950/60 dark:text-purple-300', bar: 'bg-purple-600' };
                  return { label: 'Cần khuyến khích', cls: 'bg-amber-100 text-amber-700 dark:bg-amber-950/60 dark:text-amber-300', bar: 'bg-amber-500' };
                };

                const formatSavedValue = (sl: SlideSessionStat, val: any) => {
                  if (val === undefined || val === null) return '';
                  if (typeof val === 'string' && val.startsWith('data:image/')) return val;
                  if (sl.type === 'quiz' || sl.type === 'poll') {
                    const optIdx = Number(val);
                    const optText = sl.opts?.[optIdx] ?? `Lựa chọn ${optIdx + 1}`;
                    if (sl.type === 'quiz' && sl.correctIdx !== undefined) {
                      return `${optText} ${optIdx === sl.correctIdx ? '✓ (Đúng)' : '✗ (Sai)'}`;
                    }
                    return optText;
                  }
                  if (sl.type === 'scale') return `${val} / 5`;
                  return String(val);
                };

                return (
                  <div className="space-y-6">
                    {/* Class Filter & Aggregate Percentage Cards */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                      {classSummaries.map(cs => {
                        const lv = levelBadge(cs.avgPct);
                        const isSel = statsClassFilter === cs.className;
                        return (
                          <button
                            key={cs.className}
                            onClick={() => {
                              setStatsClassFilter(cs.className);
                              const first = sessionReports.find(r => matchesClass(r, cs.className));
                              if (first) setActiveReportId(first.id);
                            }}
                            className={`text-left p-4 rounded-2xl border transition bg-white dark:bg-slate-900 ${isSel ? 'border-purple-600 ring-2 ring-purple-500/20' : 'border-stone-200 dark:border-slate-800 hover:border-purple-300'}`}
                          >
                            <div className="flex justify-between items-center mb-2">
                              <span className="font-extrabold text-base">{cs.className}</span>
                              <span className={`px-2 py-0.5 rounded-full text-[11px] font-bold ${lv.cls}`}>
                                {cs.avgPct}%
                              </span>
                            </div>
                            <div className="w-full h-2 bg-stone-100 dark:bg-slate-800 rounded-full overflow-hidden mb-1.5">
                              <div className={`h-full ${lv.bar}`} style={{ width: `${cs.avgPct}%` }} />
                            </div>
                            <div className="text-xs text-stone-500">
                              {cs.sessions} buổi học · {lv.label}
                              {cs.roomCodes.length > 0 && ` · Mã: ${cs.roomCodes.join(', ')}`}
                            </div>
                          </button>
                        );
                      })}
                    </div>

                    {/* Class Tabs */}
                    <div className="flex items-center gap-2 flex-wrap">
                      {classNames.map(c => (
                        <button
                          key={c}
                          onClick={() => {
                            setStatsClassFilter(c);
                            const first = c === 'ALL' ? sessionReports[0] : sessionReports.find(r => matchesClass(r, c));
                            if (first) setActiveReportId(first.id);
                          }}
                          className={`px-3 py-1.5 rounded-xl text-xs font-bold transition ${statsClassFilter === c ? 'bg-purple-600 text-white' : 'bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 text-stone-600 dark:text-slate-300'}`}
                        >
                          {c === 'ALL' ? 'Tất cả các lớp' : c}
                        </button>
                      ))}
                    </div>

                    {/* Session Selector Pills */}
                    {filteredReports.length > 0 && (
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="text-xs font-semibold text-stone-500">Các buổi học đã lưu ({filteredReports.length}):</span>
                        {filteredReports.map(r => (
                          <button
                            key={r.id}
                            onClick={() => setActiveReportId(r.id)}
                            className={`px-3 py-1 rounded-lg text-xs font-semibold transition ${selectedReport?.id === r.id ? 'bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 border border-purple-300' : 'bg-stone-100 dark:bg-slate-800 text-stone-600 dark:text-slate-400'}`}
                          >
                            {r.className} · {r.lessonTitle} ({r.status === 'live' ? 'Đang diễn ra' : r.endedAt})
                          </button>
                        ))}
                      </div>
                    )}

                    {/* Detailed Report */}
                    {selectedReport && (
                      <div className="space-y-6">
                        {(() => {
                          const lv = levelBadge(selectedReport.overallParticipationPct);
                          const reportStudents = statsClassFilter === 'ALL' || selectedReport.className === statsClassFilter
                            ? selectedReport.students
                            : selectedReport.students.filter(st => st.className === statsClassFilter);

                          return (
                            <div className="bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 p-6 rounded-3xl shadow-sm space-y-5">
                              <div className="flex justify-between items-start flex-wrap gap-3">
                                <div>
                                  <div className="flex items-center gap-2 flex-wrap mb-1">
                                    <span className="px-3 py-1 rounded-full bg-purple-600 text-white text-xs font-extrabold">
                                      {selectedReport.className}
                                    </span>
                                    {selectedReport.studentClassNames?.filter(c => c !== selectedReport.className).map(sc => (
                                      <span key={sc} className="px-2.5 py-0.5 rounded-full bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 text-xs font-bold">
                                        HS lớp {sc}
                                      </span>
                                    ))}
                                    <span className={`px-3 py-1 rounded-full text-xs font-bold ${lv.cls}`}>
                                      {lv.label}
                                    </span>
                                    {selectedReport.status === 'live' && (
                                      <span className="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold">
                                        Đang diễn ra
                                      </span>
                                    )}
                                    <span className="text-xs text-stone-400 font-mono">Mã: {selectedReport.roomCode}</span>
                                  </div>
                                  <h3 className="text-xl font-extrabold">{selectedReport.lessonTitle}</h3>
                                  <p className="text-xs text-stone-500">
                                    {selectedReport.startedAt ? `Bắt đầu: ${selectedReport.startedAt} · ` : ''}
                                    {selectedReport.status === 'live' ? 'Đang tự động lưu hoạt động' : `Kết thúc: ${selectedReport.endedAt}`}
                                  </p>
                                </div>
                                <button
                                  onClick={() => {
                                    const next = sessionReports.filter(r => r.id !== selectedReport.id);
                                    setSessionReports(next);
                                    saveReportsSummaryFallback(next);
                                    reportsDel(selectedReport.id);
                                    setActiveReportId(next[0]?.id || null);
                                  }}
                                  className="px-3 py-1.5 text-xs font-semibold text-red-500 border border-red-200 dark:border-red-900 rounded-xl hover:bg-red-50 dark:hover:bg-red-950/30 transition flex items-center gap-1"
                                >
                                  <Trash2 className="w-3.5 h-3.5" /> Xóa báo cáo
                                </button>
                              </div>

                              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                                <div className="p-4 rounded-2xl bg-stone-50 dark:bg-slate-950 border border-stone-200 dark:border-slate-800">
                                  <div className="text-xs font-semibold text-stone-500">Sĩ số tham gia</div>
                                  <div className="text-2xl font-black mt-1">{selectedReport.totalStudents} học sinh</div>
                                </div>
                                <div className="p-4 rounded-2xl bg-stone-50 dark:bg-slate-950 border border-stone-200 dark:border-slate-800">
                                  <div className="text-xs font-semibold text-stone-500">Hoạt động tương tác</div>
                                  <div className="text-2xl font-black mt-1">{selectedReport.totalInteractiveSlides} câu hỏi</div>
                                </div>
                                <div className="p-4 rounded-2xl bg-purple-50 dark:bg-purple-950/30 border border-purple-200 dark:border-purple-900">
                                  <div className="text-xs font-semibold text-purple-700 dark:text-purple-300">Mức độ tương tác chung</div>
                                  <div className="text-2xl font-black text-purple-600 dark:text-purple-400 mt-1">
                                    {selectedReport.overallParticipationPct}%
                                  </div>
                                </div>
                              </div>

                              {/* Per-Student Interaction Breakdown */}
                              <div className="space-y-3">
                                <h4 className="font-bold text-base">
                                  Chi tiết mức độ tương tác của học sinh ({statsClassFilter === 'ALL' ? selectedReport.className : statsClassFilter})
                                </h4>
                                {reportStudents.length === 0 ? (
                                  <p className="text-sm text-stone-500 italic">Chưa có học sinh nào tham gia trong buổi học này.</p>
                                ) : (
                                  <div className="space-y-2">
                                    {reportStudents.map((st, i) => {
                                      const stLv = levelBadge(st.participationPct);
                                      return (
                                        <div key={st.sid} className="p-3.5 rounded-2xl bg-stone-50 dark:bg-slate-950 border border-stone-200 dark:border-slate-800 space-y-2">
                                          <div className="flex justify-between items-center flex-wrap gap-2">
                                            <div className="font-bold text-sm flex items-center gap-2 flex-wrap">
                                              <span>{i + 1}. {st.name}</span>
                                              {st.className && (
                                                <span className="px-2 py-0.5 rounded-full bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 text-[11px] font-bold">
                                                  Lớp {st.className}
                                                </span>
                                              )}
                                            </div>
                                            <div className="flex items-center gap-3 text-xs font-bold">
                                              <span className="text-stone-500">
                                                Đã nộp: {st.answeredCount}/{selectedReport.totalInteractiveSlides}
                                              </span>
                                              <span className="text-purple-600">{st.score} điểm</span>
                                              <span className={`px-2.5 py-0.5 rounded-full ${stLv.cls}`}>
                                                {st.participationPct}%
                                              </span>
                                            </div>
                                          </div>
                                          <div className="w-full h-2 bg-stone-200 dark:bg-slate-800 rounded-full overflow-hidden">
                                            <div className={`h-full transition-all ${stLv.bar}`} style={{ width: `${st.participationPct}%` }} />
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>
                                )}
                              </div>

                              {/* All Saved Lesson Activities & Responses Breakdown */}
                              {selectedReport.slides.length > 0 && (
                                <div className="space-y-3 pt-2">
                                  <h4 className="font-bold text-base">Toàn bộ hoạt động bài giảng & câu trả lời đã lưu ({selectedReport.slides.length} slide)</h4>
                                  <div className="space-y-2">
                                    {selectedReport.slides.map((sl, i) => {
                                      const isInteractive = sl.isInteractive !== undefined ? sl.isInteractive : (sl.type !== 'info' && sl.type !== 'media');
                                      const savedResps = sl.responses || [];
                                      const slideKey = `${selectedReport.id}-${i}`;
                                      const isExpanded = expandedActivitySlides[slideKey] ?? savedResps.length > 0;
                                      return (
                                        <div key={i} className="p-3.5 rounded-2xl bg-stone-50 dark:bg-slate-950 border border-stone-200 dark:border-slate-800 space-y-2">
                                          <div className="flex justify-between items-center flex-wrap gap-2">
                                            <div className="flex items-center gap-2 min-w-0 flex-wrap">
                                              <span className="text-xs font-bold text-stone-400">#{sl.idx + 1}</span>
                                              <span className="px-2 py-0.5 bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 rounded-md text-[11px] font-bold">
                                                {SLIDE_TYPES[sl.type]}
                                              </span>
                                              {sl.isAi && (
                                                <span className="px-2 py-0.5 bg-amber-100 dark:bg-amber-950 text-amber-700 rounded-md text-[11px] font-bold">
                                                  AI
                                                </span>
                                              )}
                                              <span className="font-semibold text-sm">{sl.q}</span>
                                            </div>
                                            <div className="flex items-center gap-2">
                                              {isInteractive ? (
                                                <span className="text-xs font-extrabold text-purple-600">
                                                  {sl.submittedCount}/{selectedReport.totalStudents} HS ({sl.participationPct}%)
                                                </span>
                                              ) : (
                                                <span className="text-xs font-semibold text-stone-400">Slide nội dung</span>
                                              )}
                                              {savedResps.length > 0 && (
                                                <button
                                                  type="button"
                                                  onClick={() => setExpandedActivitySlides(prev => ({ ...prev, [slideKey]: !isExpanded }))}
                                                  className="px-2.5 py-1 rounded-lg bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-700 text-[11px] font-bold text-purple-600 hover:bg-purple-50"
                                                >
                                                  {isExpanded ? 'Ẩn bài nộp' : `Xem ${savedResps.length} bài nộp`}
                                                </button>
                                              )}
                                            </div>
                                          </div>
                                          {isInteractive && (
                                            <div className="w-full h-2 bg-stone-200 dark:bg-slate-800 rounded-full overflow-hidden">
                                              <div className="h-full bg-purple-600" style={{ width: `${sl.participationPct}%` }} />
                                            </div>
                                          )}
                                          {sl.b && !isInteractive && (
                                            <p className="text-xs text-stone-500 whitespace-pre-line line-clamp-3">{sl.b}</p>
                                          )}
                                          {isExpanded && savedResps.length > 0 && (
                                            <div className="pt-2 border-t border-stone-200 dark:border-slate-800 grid grid-cols-1 sm:grid-cols-2 gap-2">
                                              {savedResps.map((res, rIdx) => {
                                                const formatted = formatSavedValue(sl, res.value);
                                                const isImg = typeof formatted === 'string' && formatted.startsWith('data:image/');
                                                return (
                                                  <div key={rIdx} className="p-2.5 rounded-xl bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-800 text-xs space-y-1">
                                                    <div className="flex items-center justify-between gap-2">
                                                      <span className="font-bold">
                                                        {res.name}
                                                        {res.className && <span className="ml-1 text-purple-600">({res.className})</span>}
                                                      </span>
                                                      <div className="flex items-center gap-1.5">
                                                        {res.grade !== undefined && (
                                                          <span className="px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-bold">{res.grade}đ</span>
                                                        )}
                                                        {res.submittedAt && <span className="text-[10px] text-stone-400">{res.submittedAt}</span>}
                                                      </div>
                                                    </div>
                                                    {isImg ? (
                                                      <img src={formatted} alt="Bài vẽ" className="w-full max-h-32 object-contain rounded-lg border border-stone-200 bg-white" />
                                                    ) : (
                                                      <div className="text-stone-700 dark:text-slate-300 break-words">{formatted}</div>
                                                    )}
                                                  </div>
                                                );
                                              })}
                                            </div>
                                          )}
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                              )}
                            </div>
                          );
                        })()}
                      </div>
                    )}
                  </div>
                );
              })()
            )}
          </div>
        )}
      </main>

      {/* Footer (hidden during live lesson mode) */}
      {!isFullLessonMode && (
        <footer className="py-8 text-center text-stone-500 dark:text-slate-500 text-xs border-t border-stone-200 dark:border-slate-800 mt-auto">
          Lớp học trực tuyến · Toàn bộ bài giảng, hoạt động tương tác và thống kê lớp được lưu tự động
        </footer>
      )}
    </div>
  );
}
