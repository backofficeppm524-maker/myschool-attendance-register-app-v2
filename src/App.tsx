import React, { useState, useMemo, useRef, useEffect } from 'react';
import { format, getDaysInMonth, startOfMonth } from 'date-fns';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import html2canvas from 'html2canvas';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { Printer, Upload, Download, Plus, Trash2, Pencil, CalendarDays, FileText, LogOut, Lock, Users, Percent, Activity, Settings, Loader2, AlertTriangle, Search, CheckCircle2, XCircle, ShieldCheck, BookOpen, BarChart3, Home, Info, Phone, GraduationCap, Sun, TrendingUp, TrendingDown, User, Copy, MessageCircle } from 'lucide-react';
import { 
  auth, db, signInWithPopup, signOut, onAuthStateChanged,
  createUserWithEmailAndPassword, signInWithEmailAndPassword,
  doc, collection, onSnapshot, setDoc, updateDoc, deleteDoc, getDoc, writeBatch, arrayUnion
} from './firebase';

type AttendanceValue = 'P' | 'A' | 'WH' | 'GH' | 'OH' | 'OT' | 'RM' | '';
type Session = 'fn' | 'an';
type AttendanceDay = { fn: AttendanceValue; an: AttendanceValue };
type AttendanceMap = Record<string, Record<number, AttendanceDay>>;

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  // In a real app, you might show a toast or a modal here
}

function isPermissionDeniedError(error: unknown) {
  const message = error instanceof Error ? error.message : String(error);
  return message.toLowerCase().includes('missing or insufficient permissions') ||
    message.toLowerCase().includes('permission-denied');
}

interface Student {
  id: string;
  sno: string;
  admNo: string;
  name: string;
  gender?: 'Boy' | 'Girl';
  comments?: string;
  createdMonth?: string;
  deletedMonth?: string;
  deletedMonths?: string[];
}

const HOLIDAY_TYPES = ['WH', 'GH', 'OH', 'OT'];

const MONTH_WORKING_DAYS_OVERRIDES: Record<string, number> = {
  '2026-06': 18,
  '2026-07': 23,
  '2026-08': 20,
  '2026-09': 18,
  '2026-10': 18,
};

const getMonthWorkingDaysOverride = (date: Date) => MONTH_WORKING_DAYS_OVERRIDES[format(date, 'yyyy-MM')];

const getFullStudentData = (student: Student) => ({
  sno: String(student.sno || ''),
  admNo: String(student.admNo || ''),
  name: String(student.name || ''),
  gender: student.gender || 'Boy',
  ...(student.createdMonth ? { createdMonth: student.createdMonth } : {})
});

const getLegacyStudentData = (student: Student) => ({
  sno: String(student.sno || ''),
  admNo: String(student.admNo || ''),
  name: String(student.name || '')
});

const getLocalStudentsKey = (userId: string) => `attendance-register-local-students:${userId}`;

const loadLocalStudents = (userId: string): Student[] => {
  try {
    const raw = localStorage.getItem(getLocalStudentsKey(userId));
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const saveLocalStudents = (userId: string, list: Student[]) => {
  localStorage.setItem(getLocalStudentsKey(userId), JSON.stringify(list));
};

const isStudentActiveInMonth = (student: Student, monthKey: string, attendanceMap?: AttendanceMap) => {
  const isCreated = !student.createdMonth || student.createdMonth <= monthKey;
  const deletedOnlyInMonth = student.deletedMonths?.includes(monthKey) ?? false;
  const deletedInLegacyMonth = student.deletedMonth === monthKey;
  const removedInAttendance = attendanceMap?.[`${monthKey}_${student.id}`]?.[0]?.fn === 'RM';
  return isCreated && !deletedOnlyInMonth && !deletedInLegacyMonth && !removedInAttendance;
};

interface AttendanceSelectorProps {
  x: number;
  y: number;
  onSelect: (value: AttendanceValue) => void;
  onClose: () => void;
}

const AttendanceSelector = ({ x, y, onSelect, onClose }: AttendanceSelectorProps) => {
  const options = [
    { label: 'Present (P)', value: 'P' as AttendanceValue, icon: '✅', bg: 'hover:bg-green-50' },
    { label: 'Absent (A)', value: 'A' as AttendanceValue, icon: '❌', bg: 'hover:bg-red-50' },
    { label: 'Clear', value: '' as AttendanceValue, icon: '🗑️', bg: 'hover:bg-gray-100' },
  ];

  return (
    <div className="fixed inset-0 z-[100]" onClick={onClose}>
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 5 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 5 }}
        className="absolute bg-white shadow-[0_4px_20px_rgba(0,0,0,0.08)] rounded-xl p-1.5 min-w-[130px] pointer-events-auto border border-gray-100"
        style={{ left: x - 65, top: y + 8 }}
      >
        <div className="flex flex-col gap-0.5">
          {options.map((opt) => (
            <button
              key={opt.label}
              onClick={(e) => {
                e.stopPropagation();
                onSelect(opt.value);
              }}
              className={cn(
                "flex items-center gap-2 px-2.5 py-1 text-[12px] font-bold text-gray-800 rounded-lg transition-all active:scale-95 text-left",
                opt.bg
              )}
            >
              <span className="text-base flex items-center justify-center w-5 h-5">{opt.icon}</span>
              {opt.label}
            </button>
          ))}
        </div>
      </motion.div>
    </div>
  );
};

interface BulkMarkingPopupProps {
  x: number;
  y: number;
  day: number;
  session: Session;
  onSelect: (value: AttendanceValue) => void;
  onClose: () => void;
}

const BulkMarkingPopup = ({ x, y, day, session, onSelect, onClose }: BulkMarkingPopupProps) => {
  const options = [
    { label: `All ${session.toUpperCase()} = P`, value: 'P' as AttendanceValue, icon: '✅', bg: 'hover:bg-green-50' },
    { label: `All ${session.toUpperCase()} = A`, value: 'A' as AttendanceValue, icon: '❌', bg: 'hover:bg-red-50' },
    { label: `Clear ${session.toUpperCase()}`, value: '' as AttendanceValue, icon: '🗑️', bg: 'hover:bg-gray-100' },
  ];

  return (
    <div className="fixed inset-0 z-[100]" onClick={onClose}>
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 5 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 5 }}
        className="absolute bg-white shadow-[0_4px_20px_rgba(0,0,0,0.08)] rounded-xl p-1.5 min-w-[155px] pointer-events-auto border border-gray-100"
        style={{ left: x - 77, top: y + 8 }}
      >
        <div className="px-2.5 py-1 text-[8px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50 mb-1">Bulk Mark</div>
        <div className="flex flex-col gap-0.5">
          {options.map((opt) => (
            <button
              key={opt.label}
              onClick={(e) => {
                e.stopPropagation();
                onSelect(opt.value);
              }}
              className={cn(
                "flex items-center gap-2 px-2.5 py-1 text-[12px] font-bold text-gray-800 rounded-lg transition-all active:scale-95 text-left",
                opt.bg
              )}
            >
              <span className="text-base flex items-center justify-center w-5 h-5">{opt.icon}</span>
              {opt.label}
            </button>
          ))}
        </div>
      </motion.div>
    </div>
  );
};

export default function App() {
  const [user, setUser] = useState<any>(null);
  const [isAuthLoading, setIsAuthLoading] = useState(true);
  const [authMode, setAuthMode] = useState<'login' | 'signup'>('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [adminUsername, setAdminUsername] = useState('admin');
  const [adminPassword, setAdminPassword] = useState('admin123');
  const [showSettings, setShowSettings] = useState(false);
  const [newUsername, setNewUsername] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [showDeleteConfirm, setShowDeleteConfirm] = useState<string | null>(null);
  const [isDeletingStudent, setIsDeletingStudent] = useState(false);
  const [lastAddedStudentId, setLastAddedStudentId] = useState<string | null>(null);
  const tableEndRef = useRef<HTMLDivElement>(null);

  // Add Student Modal state
  const [showAddStudentModal, setShowAddStudentModal] = useState(false);
  const [newStudentName, setNewStudentName] = useState('');
  const [newStudentAdmNo, setNewStudentAdmNo] = useState('');
  const [newStudentGender, setNewStudentGender] = useState<'Boy' | 'Girl'>('Boy');
  const [isAddingStudent, setIsAddingStudent] = useState(false);
  const [isImporting, setIsImporting] = useState(false);

  // Edit Student Modal state
  const [showEditStudentModal, setShowEditStudentModal] = useState<string | null>(null);
  const [editStudentName, setEditStudentName] = useState('');
  const [editStudentAdmNo, setEditStudentAdmNo] = useState('');
  const [editStudentGender, setEditStudentGender] = useState<'Boy' | 'Girl'>('Boy');

  // Scroll to bottom when student is added
  useEffect(() => {
    if (lastAddedStudentId) {
      tableEndRef.current?.scrollIntoView({ behavior: 'smooth' });
      const timer = setTimeout(() => setLastAddedStudentId(null), 2000);
      return () => clearTimeout(timer);
    }
  }, [lastAddedStudentId]);
  const [isExporting, setIsExporting] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [viewMode, setViewMode] = useState<'register' | 'daily' | 'monthly' | 'annual' | 'genderwise' | 'monthwise' | 'about'>('register');
  const [reportDay, setReportDay] = useState<number>(1);
  
  const [attendanceHistory, setAttendanceHistory] = useState<AttendanceMap[]>([]);
  const [showHistoryModal, setShowHistoryModal] = useState<string | null>(null);
  const [showHolidayManager, setShowHolidayManager] = useState(false);
  
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [holidayDate, setHolidayDate] = useState<string>(format(new Date(), 'yyyy-MM-dd'));
  const [holidayType, setHolidayType] = useState<AttendanceValue>('GH');
  const [printSize, setPrintSize] = useState<'a4' | 'legal' | 'letter'>('a4');
  const [printOrientation, setPrintOrientation] = useState<'p' | 'l'>('l');
  const [printMargin, setPrintMargin] = useState<number>(14);

  // New state for circular menu and bulk marking
  const [activeCell, setActiveCell] = useState<{ studentId: string, day: number, session: Session, x: number, y: number } | null>(null);
  const [activeBulkColumn, setActiveBulkColumn] = useState<{ day: number, session: Session, x: number, y: number } | null>(null);

  const [monthDate, setMonthDate] = useState<Date>(startOfMonth(new Date()));
  const [institutionName, setInstitutionName] = useState('GOVT HR SEC SCHOOL PANNAIPURAM');
  const [standard, setStandard] = useState('');
  const [section, setSection] = useState('');
  const [remoteStudents, setRemoteStudents] = useState<Student[]>([]);
  const [localStudents, setLocalStudents] = useState<Student[]>([]);
  const [attendance, setAttendance] = useState<AttendanceMap>({});

  const fileInputRef = useRef<HTMLInputElement>(null);
  const configLoadedRef = useRef(false);

  const students = useMemo(() => {
    const merged = new Map<string, Student>();
    remoteStudents.forEach(student => merged.set(student.id, student));
    localStudents.forEach(student => merged.set(student.id, student));
    return Array.from(merged.values()).sort((a, b) => {
      const snoA = parseInt(a.sno) || 0;
      const snoB = parseInt(b.sno) || 0;
      return snoA - snoB;
    });
  }, [remoteStudents, localStudents]);

  const persistLocalStudents = (updater: (current: Student[]) => Student[]) => {
    if (!user) return;
    setLocalStudents(current => {
      const next = updater(current);
      saveLocalStudents(user.uid, next);
      return next;
    });
  };

  // Auth Listener
  useEffect(() => {
    // Safety timeout - ensure loading state doesn't get stuck
    const timeout = setTimeout(() => {
      setIsAuthLoading(false);
    }, 10000);

    const unsubscribe = onAuthStateChanged(auth,
      (currentUser) => {
        clearTimeout(timeout);
        setUser(currentUser);
        setIsAuthLoading(false);
      },
      (error) => {
        clearTimeout(timeout);
        console.error("Auth state change error:", error);
        setIsAuthLoading(false);
      }
    );
    return () => {
      unsubscribe();
      clearTimeout(timeout);
    };
  }, []);

  // Firestore Listeners
  useEffect(() => {
    if (!user) {
      setRemoteStudents([]);
      setLocalStudents([]);
      return;
    }

    const userId = user.uid;
    configLoadedRef.current = false;
    setLocalStudents(loadLocalStudents(userId));

    // Config Listener
    const configUnsubscribe = onSnapshot(doc(db, `users/${userId}/config/settings`), (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        setInstitutionName(data.institutionName || 'GOVT HR SEC SCHOOL PANNAIPURAM');
        setStandard(data.standard || '');
        setSection(data.section || '');
        setAdminUsername(data.adminUsername || 'admin');
        setAdminPassword(data.adminPassword || 'admin123');
      }
      // Mark config as loaded — even for a brand-new user with no saved
      // config yet — so the auto-save effect below knows it's safe to write.
      configLoadedRef.current = true;
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, `users/${userId}/config/settings`);
      // Ensure config is marked as loaded even on error to prevent blocking
      configLoadedRef.current = true;
    });

    // Students Listener
    const studentsUnsubscribe = onSnapshot(collection(db, `users/${userId}/students`), (snapshot) => {
      const studentsList: Student[] = [];
      snapshot.forEach((doc) => {
        studentsList.push({ id: doc.id, ...doc.data() } as Student);
      });
      // Sort students by sno if possible
      studentsList.sort((a, b) => {
        const snoA = parseInt(a.sno) || 0;
        const snoB = parseInt(b.sno) || 0;
        return snoA - snoB;
      });
      setRemoteStudents(studentsList);
    }, (error) => handleFirestoreError(error, OperationType.LIST, `users/${userId}/students`));

    // Attendance Listener
    const attendanceUnsubscribe = onSnapshot(collection(db, `users/${userId}/attendance`), (snapshot) => {
      const attendanceMap: AttendanceMap = {};
      snapshot.forEach((doc) => {
        attendanceMap[doc.id] = doc.data().days || {};
      });
      setAttendance(attendanceMap);
    }, (error) => handleFirestoreError(error, OperationType.LIST, `users/${userId}/attendance`));

    return () => {
      configUnsubscribe();
      studentsUnsubscribe();
      attendanceUnsubscribe();
    };
  }, [user]);

  // Auto-save Standard & Section straight to Firestore as the teacher types
  // them into the register header, instead of requiring a trip through the
  // separate Settings form. Debounced so rapid keystrokes don't each trigger
  // a write, and gated on configLoadedRef so it can never fire with stale or
  // default values before this user's own config has finished loading.
  useEffect(() => {
    if (!user || !configLoadedRef.current) return;
    const userId = user.uid;
    const timer = setTimeout(() => {
      // institutionName is included because firestore.rules requires it to be
      // present and non-empty on every write to config/settings — a merge
      // write of just {standard, section} would fail rule validation the
      // first time it creates the document (e.g. brand-new account).
      setDoc(doc(db, `users/${userId}/config/settings`), { institutionName, standard, section }, { merge: true })
        .catch(error => handleFirestoreError(error, OperationType.WRITE, `users/${userId}/config/settings`));
    }, 600);
    return () => clearTimeout(timer);
  }, [standard, section, institutionName, user]);

  const days = getDaysInMonth(monthDate);
  const dayArray = Array.from({ length: days }, (_, i) => i + 1);

  const getAttKey = (studentId: string, d?: Date) => `${format(d || monthDate, 'yyyy-MM')}_${studentId}`;

  const saveHistory = () => {
    setAttendanceHistory(prev => [...prev, attendance].slice(-10));
  };

  const undoHistory = () => {
    if (attendanceHistory.length > 0) {
      const prev = attendanceHistory[attendanceHistory.length - 1];
      setAttendance(prev);
      setAttendanceHistory(history => history.slice(0, -1));
      // Note: This undo only affects local state until the next change is saved to Firestore.
      // For a robust undo, we'd need to save the previous state back to Firestore.
    }
  };

  const handleMonthChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.value) {
      setMonthDate(startOfMonth(new Date(e.target.value)));
    }
  };

  const updateAttendance = async (studentId: string, day: number, session: Session, value: AttendanceValue) => {
    if (!user) return;
    saveHistory();
    const attKey = getAttKey(studentId);
    const userId = user.uid;
    
    const studentAtt = attendance[attKey] || {};
    const dayAtt = studentAtt[day] || { fn: '', an: '' };
    const newDayAtt = { ...dayAtt, [session]: value };
    const newStudentAtt = { ...studentAtt, [day]: newDayAtt };

    try {
      await setDoc(doc(db, `users/${userId}/attendance/${attKey}`), {
        days: newStudentAtt
      }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `users/${userId}/attendance/${attKey}`);
    }
  };

  const handleCellClick = (studentId: string, day: number, session: Session, e: React.MouseEvent) => {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    setActiveCell({
      studentId,
      day,
      session,
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2
    });
  };

  const handleBulkHeaderClick = (day: number, session: Session, e: React.MouseEvent) => {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    setActiveBulkColumn({
      day,
      session,
      x: rect.left + rect.width / 2,
      y: rect.bottom
    });
  };

  const markWholeDay = async (day: number, session: Session, value: AttendanceValue) => {
    if (!user) return;
    saveHistory();
    const userId = user.uid;
    const batch = writeBatch(db);

    students.forEach(s => {
      const attKey = getAttKey(s.id);
      const studentAtt = attendance[attKey] || {};
      const currentDay = studentAtt[day] || { fn: '', an: '' };
      const newDayAtt = { ...currentDay, [session]: value };
      const newStudentAtt = { ...studentAtt, [day]: newDayAtt };
      
      batch.set(doc(db, `users/${userId}/attendance/${attKey}`), {
        days: newStudentAtt
      }, { merge: true });
    });

    try {
      await batch.commit();
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `users/${userId}/attendance`);
    }
  };

  const markStudentAllDays = async (studentId: string, value: AttendanceValue) => {
    if (!user) return;
    saveHistory();
    const userId = user.uid;
    const attKey = getAttKey(studentId);
    const studentAtt: Record<number, AttendanceDay> = {};
    for (let d = 1; d <= days; d++) {
      studentAtt[d] = { fn: value, an: value };
    }

    try {
      await setDoc(doc(db, `users/${userId}/attendance/${attKey}`), {
        days: studentAtt
      }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `users/${userId}/attendance/${attKey}`);
    }
  };

  const sortedStudentsForMonth = useMemo(() => {
    const monthKey = format(monthDate, 'yyyy-MM');
    const active = students.filter(s => isStudentActiveInMonth(s, monthKey, attendance));

    const boys = active.filter(s => (s.gender || 'Boy') === 'Boy')
      .sort((a, b) => a.name.trim().toLowerCase().localeCompare(b.name.trim().toLowerCase()));
    const girls = active.filter(s => s.gender === 'Girl')
      .sort((a, b) => a.name.trim().toLowerCase().localeCompare(b.name.trim().toLowerCase()));

    return [...boys, ...girls].map((student, idx) => ({
      ...student,
      sno: String(idx + 1)
    }));
  }, [students, monthDate, attendance]);

  const getWorkingDaysForMonth = (targetDate: Date) => {
    const override = getMonthWorkingDaysOverride(targetDate);
    if (override !== undefined) return override;

    const monthKey = format(targetDate, 'yyyy-MM');
    const targetDays = getDaysInMonth(targetDate);
    const activeStudentsInMonth = students.filter(s => isStudentActiveInMonth(s, monthKey, attendance));

    let holidaysCount = 0;
    for (let d = 1; d <= targetDays; d++) {
      if (activeStudentsInMonth.length === 0) continue;

      // A day counts as a holiday when MOST active students carry a holiday
      // code (WH/GH/OH/OT) in both sessions, rather than requiring every
      // single one. Requiring unanimity meant one student with a blank cell
      // for that day (e.g. a student admitted after the holiday was bulk
      // marked via Holiday Manager) silently turned a real holiday into a
      // counted working day.
      let holidayMarkedCount = 0;
      for (const s of activeStudentsInMonth) {
        const attKey = `${monthKey}_${s.id}`;
        const record = attendance[attKey]?.[d] || { fn: '', an: '' };
        if (HOLIDAY_TYPES.includes(record.fn) && HOLIDAY_TYPES.includes(record.an)) {
          holidayMarkedCount++;
        }
      }
      if (holidayMarkedCount * 2 > activeStudentsInMonth.length) holidaysCount++;
    }
    return targetDays - holidaysCount;
  };

  const globalWorkingDays = useMemo(() => getWorkingDaysForMonth(monthDate), [attendance, students, monthDate]);

  const calculateTotals = (studentId: string, specificDate?: Date) => {
    const targetDate = specificDate || monthDate;
    const targetDays = getDaysInMonth(targetDate);
    const attKey = getAttKey(studentId, targetDate);
    
    let present = 0;
    let absent = 0;
    let holidays = 0;

    for (let d = 1; d <= targetDays; d++) {
      const record = attendance[attKey]?.[d] || { fn: '', an: '' };
      const fn = record.fn;
      const an = record.an;

      const isHolidayFn = HOLIDAY_TYPES.includes(fn);
      const isHolidayAn = HOLIDAY_TYPES.includes(an);

      if (fn === 'P') present += 0.5;
      else if (fn === 'A') absent += 0.5;
      else if (isHolidayFn) holidays += 0.5;

      if (an === 'P') present += 0.5;
      else if (an === 'A') absent += 0.5;
      else if (isHolidayAn) holidays += 0.5;
    }

    const workingDays = getWorkingDaysForMonth(targetDate);
    const percentage = workingDays > 0 ? (present / workingDays) * 100 : 0;
    return { present, absent, workingDays, holidays, percentage };
  };

  const filteredStudents = useMemo(() => {
    if (!searchQuery.trim()) return sortedStudentsForMonth;
    const query = searchQuery.toLowerCase();
    return sortedStudentsForMonth.filter(s => 
      s.name.toLowerCase().includes(query) ||
      s.admNo.toLowerCase().includes(query) ||
      s.sno.toLowerCase().includes(query)
    );
  }, [sortedStudentsForMonth, searchQuery]);

  const todayDay = useMemo(() => {
    const now = new Date();
    if (now.getMonth() === monthDate.getMonth() && now.getFullYear() === monthDate.getFullYear()) {
      return Math.min(now.getDate(), days);
    }
    return Math.min(reportDay || 1, days);
  }, [monthDate, reportDay, days]);

  const todayLabel = useMemo(() => {
    try {
      const targetDate = new Date(monthDate.getFullYear(), monthDate.getMonth(), todayDay);
      return format(targetDate, 'dd MMMM yyyy');
    } catch (e) {
      return `${todayDay} ${format(monthDate, 'MMMM yyyy')}`;
    }
  }, [monthDate, todayDay]);

  const {
    totalPresent,
    totalAbsent,
    avgAttendance,
    fnPresent,
    fnAbsent,
    fnTotal,
    anPresent,
    anAbsent,
    anTotal,
    todayPresent,
    todayAbsent,
  } = useMemo(() => {
    if (filteredStudents.length === 0) {
      return {
        totalPresent: 0,
        totalAbsent: 0,
        avgAttendance: "0.0",
        fnPresent: 0,
        fnAbsent: 0,
        fnTotal: 0,
        anPresent: 0,
        anAbsent: 0,
        anTotal: 0,
        todayPresent: 0,
        todayAbsent: 0,
      };
    }

    let p = 0;
    let a = 0;
    let fnP = 0;
    let fnA = 0;
    let anP = 0;
    let anA = 0;

    filteredStudents.forEach(s => {
      const attKey = getAttKey(s.id);
      const record = attendance[attKey]?.[todayDay] || { fn: '', an: '' };

      const fn = record.fn;
      const an = record.an;

      if (fn === 'P') { p += 0.5; fnP += 1; }
      else if (fn === 'A') { a += 0.5; fnA += 1; }

      if (an === 'P') { p += 0.5; anP += 1; }
      else if (an === 'A') { a += 0.5; anA += 1; }
    });

    const avg = filteredStudents.length > 0
      ? ((p / filteredStudents.length) * 100).toFixed(1)
      : "0.0";

    return {
      totalPresent: p,
      totalAbsent: a,
      avgAttendance: avg,
      fnPresent: fnP,
      fnAbsent: fnA,
      fnTotal: fnP + fnA,
      anPresent: anP,
      anAbsent: anA,
      anTotal: anP + anA,
      todayPresent: fnP + anP,
      todayAbsent: fnA + anA,
    };
  }, [attendance, filteredStudents, todayDay, monthDate]);

  // ── Extra dashboard KPIs (derived from the same attendance data) ──
  const dashboardKPIs = useMemo(() => {
    const totalStudents = filteredStudents.length;
    // Working days for the currently selected month
    const workingDays = globalWorkingDays;
    // Present % / Absent % for the selected day (today)
    const presentPct = totalStudents > 0 ? (totalPresent / totalStudents) * 100 : 0;
    const absentPct = totalStudents > 0 ? (totalAbsent / totalStudents) * 100 : 0;
    // Holidays on the selected day (any student carrying a holiday code on that day)
    let holidayStudents = 0;
    filteredStudents.forEach(s => {
      const attKey = getAttKey(s.id);
      const record = attendance[attKey]?.[todayDay] || { fn: '', an: '' };
      if (HOLIDAY_TYPES.includes(record.fn) && HOLIDAY_TYPES.includes(record.an)) holidayStudents++;
    });
    // Boys / Girls split
    const boysCount = filteredStudents.filter(s => (s.gender || 'Boy') === 'Boy').length;
    const girlsCount = filteredStudents.filter(s => s.gender === 'Girl').length;
    // Top & bottom performers this month (by attendance %)
    const ranked = filteredStudents
      .map(s => ({ s, totals: calculateTotals(s.id) }))
      .filter(x => x.totals.workingDays > 0)
      .sort((a, b) => b.totals.percentage - a.totals.percentage);
    const topPerformer = ranked.length > 0 ? ranked[0] : null;
    const bottomPerformer = ranked.length > 0 ? ranked[ranked.length - 1] : null;
    // Month-to-date overall attendance %
    const monthAvgPct = totalStudents > 0 && workingDays > 0
      ? (filteredStudents.reduce((sum, s) => sum + calculateTotals(s.id).present, 0) / (totalStudents * workingDays)) * 100
      : 0;
    return {
      totalStudents,
      workingDays,
      presentPct,
      absentPct,
      holidayStudents,
      boysCount,
      girlsCount,
      topPerformer,
      bottomPerformer,
      monthAvgPct,
    };
  }, [attendance, filteredStudents, todayDay, globalWorkingDays, monthDate]);

  const getDailySummary = (day: number, session: Session) => {
    let p = 0;
    let a = 0;
    filteredStudents.forEach(s => {
      const attKey = getAttKey(s.id);
      const val = attendance[attKey]?.[day]?.[session];
      if (val === 'P') p++;
      if (val === 'A') a++;
    });
    return { p, a };
  };

  const getAttendanceStatus = (percentage: number) => {
    if (percentage >= 90) return 'Excellent';
    if (percentage >= 75) return 'Good';
    if (percentage >= 60) return 'Average';
    return 'Poor';
  };

  // Academic months calculation (Jun to Apr) - May is removed
  const academicMonths = useMemo(() => {
    const startYear = monthDate.getMonth() >= 5 ? monthDate.getFullYear() : monthDate.getFullYear() - 1;
    const monthsList: Date[] = [];
    for (let m = 0; m < 11; m++) {
      const monthIndex = (5 + m) % 12;
      const yearOffset = Math.floor((5 + m) / 12);
      const date = new Date(startYear + yearOffset, monthIndex, 1);
      monthsList.push(date);
    }
    return monthsList;
  }, [monthDate]);

  const academicYearLabel = useMemo(() => {
    const startYear = monthDate.getMonth() >= 5 ? monthDate.getFullYear() : monthDate.getFullYear() - 1;
    return `A.Y. ${startYear} - ${startYear + 1}`;
  }, [monthDate]);

  // Working days per month for the academic year
  const annualWorkingDays = useMemo(() => {
    const workingDaysMap: Record<string, number> = {};
    academicMonths.forEach(mDate => {
      const monthKey = format(mDate, 'yyyy-MM');
      workingDaysMap[monthKey] = getWorkingDaysForMonth(mDate);
    });
    return workingDaysMap;
  }, [academicMonths, students, attendance]);

  // Student monthly totals for the academic year
  const annualStudentAttendance = useMemo(() => {
    const studentAttendanceMap: Record<string, Record<string, number>> = {};
    
    students.forEach(s => {
      studentAttendanceMap[s.id] = {};
      
      academicMonths.forEach(mDate => {
        const monthKey = format(mDate, 'yyyy-MM');
        
        const isActive = isStudentActiveInMonth(s, monthKey, attendance);

        if (!isActive) {
          studentAttendanceMap[s.id][monthKey] = 0;
          return;
        }

        const targetDays = getDaysInMonth(mDate);
        const attKey = `${monthKey}_${s.id}`;
        
        let presentCount = 0;
        for (let d = 1; d <= targetDays; d++) {
          const record = attendance[attKey]?.[d] || { fn: '', an: '' };
          if (record.fn === 'P') presentCount += 0.5;
          if (record.an === 'P') presentCount += 0.5;
        }
        
        studentAttendanceMap[s.id][monthKey] = presentCount;
      });
    });
    return studentAttendanceMap;
  }, [academicMonths, students, attendance]);

  const annualSummary = useMemo(() => {
    let totalWorkingDays = 0;
    academicMonths.forEach(mDate => {
      const monthKey = format(mDate, 'yyyy-MM');
      totalWorkingDays += annualWorkingDays[monthKey] ?? 0;
    });
    
    const studentsSummary = students.map(s => {
      let totalPresent = 0;
      academicMonths.forEach(mDate => {
        const monthKey = format(mDate, 'yyyy-MM');
        totalPresent += annualStudentAttendance[s.id]?.[monthKey] || 0;
      });
      const percentage = totalWorkingDays > 0 ? (totalPresent / totalWorkingDays) * 100 : 0;
      return {
        studentId: s.id,
        totalPresent,
        totalWorkingDays,
        percentage
      };
    });
    
    return {
      totalWorkingDays,
      studentsSummary
    };
  }, [academicMonths, students, annualWorkingDays, annualStudentAttendance]);

  const genderwiseSummary = useMemo(() => {
    const boysList = students.filter(s => (s.gender || 'Boy') === 'Boy');
    const girlsList = students.filter(s => s.gender === 'Girl');
    
    const currentMonthKey = format(monthDate, 'yyyy-MM');
    const activeBoys = boysList.filter(s => isStudentActiveInMonth(s, currentMonthKey, attendance));
    const activeGirls = girlsList.filter(s => isStudentActiveInMonth(s, currentMonthKey, attendance));
    const numBoys = activeBoys.length;
    const numGirls = activeGirls.length;

    const monthlyData: Record<string, {
      boysPresent: number;
      boysAbsent: number;
      girlsPresent: number;
      girlsAbsent: number;
    }> = {};

    academicMonths.forEach(mDate => {
      const monthKey = format(mDate, 'yyyy-MM');
      const targetDays = getDaysInMonth(mDate);
      
      const activeBoysInMonth = boysList.filter(s => isStudentActiveInMonth(s, monthKey, attendance));
      const activeGirlsInMonth = girlsList.filter(s => isStudentActiveInMonth(s, monthKey, attendance));

      const numBoysInMonth = activeBoysInMonth.length;
      const numGirlsInMonth = activeGirlsInMonth.length;

      let boysPresSum = 0;
      let boysAbsSum = 0;
      let girlsPresSum = 0;
      let girlsAbsSum = 0;

      activeBoysInMonth.forEach(s => {
        const attKey = `${monthKey}_${s.id}`;
        let p = 0;
        let a = 0;
        for (let d = 1; d <= targetDays; d++) {
          const record = attendance[attKey]?.[d] || { fn: '', an: '' };
          const isHolidayFn = HOLIDAY_TYPES.includes(record.fn);
          const isHolidayAn = HOLIDAY_TYPES.includes(record.an);

          if (record.fn === 'P') p += 0.5;
          else if (record.fn === 'A') a += 0.5;
          else if (!record.fn && !isHolidayFn) a += 0.5;

          if (record.an === 'P') p += 0.5;
          else if (record.an === 'A') a += 0.5;
          else if (!record.an && !isHolidayAn) a += 0.5;
        }
        boysPresSum += p;
        boysAbsSum += a;
      });

      activeGirlsInMonth.forEach(s => {
        const attKey = `${monthKey}_${s.id}`;
        let p = 0;
        let a = 0;
        for (let d = 1; d <= targetDays; d++) {
          const record = attendance[attKey]?.[d] || { fn: '', an: '' };
          const isHolidayFn = HOLIDAY_TYPES.includes(record.fn);
          const isHolidayAn = HOLIDAY_TYPES.includes(record.an);

          if (record.fn === 'P') p += 0.5;
          else if (record.fn === 'A') a += 0.5;
          else if (!record.fn && !isHolidayFn) a += 0.5;

          if (record.an === 'P') p += 0.5;
          else if (record.an === 'A') a += 0.5;
          else if (!record.an && !isHolidayAn) a += 0.5;
        }
        girlsPresSum += p;
        girlsAbsSum += a;
      });

      const wd = annualWorkingDays[monthKey] ?? 0;
      monthlyData[monthKey] = {
        boysPresent: (numBoysInMonth > 0 && wd > 0) ? (boysPresSum / (numBoysInMonth * wd)) * 100 : 0,
        boysAbsent: (numBoysInMonth > 0 && wd > 0) ? (boysAbsSum / (numBoysInMonth * wd)) * 100 : 0,
        girlsPresent: (numGirlsInMonth > 0 && wd > 0) ? (girlsPresSum / (numGirlsInMonth * wd)) * 100 : 0,
        girlsAbsent: (numGirlsInMonth > 0 && wd > 0) ? (girlsAbsSum / (numGirlsInMonth * wd)) * 100 : 0,
      };
    });

    let boysSumOfMonthlyPercentages = 0;
    let girlsSumOfMonthlyPercentages = 0;
    
    academicMonths.forEach(mDate => {
      const monthKey = format(mDate, 'yyyy-MM');
      boysSumOfMonthlyPercentages += monthlyData[monthKey]?.boysPresent ?? 0;
      girlsSumOfMonthlyPercentages += monthlyData[monthKey]?.girlsPresent ?? 0;
    });

    const boysAvgAnnualPresence = academicMonths.length > 0 
      ? boysSumOfMonthlyPercentages / academicMonths.length 
      : 0;

    const girlsAvgAnnualPresence = academicMonths.length > 0
      ? girlsSumOfMonthlyPercentages / academicMonths.length
      : 0;

    return {
      monthlyData,
      boysCount: numBoys,
      girlsCount: numGirls,
      boysAvgAnnualPresence,
      girlsAvgAnnualPresence,
    };
  }, [students, attendance, academicMonths, annualWorkingDays, monthDate]);

  const monthwiseSummary = useMemo(() => {
    let priorWorkingDaysAccumulator = 0;
    
    let totalWorkingDays = 0;
    academicMonths.forEach(mDate => {
      const monthKey = format(mDate, 'yyyy-MM');
      totalWorkingDays += annualWorkingDays[monthKey] ?? 0;
    });

    const monthlyDataList = academicMonths.map(mDate => {
      const monthKey = format(mDate, 'yyyy-MM');
      const currentMonthWorkingDays = annualWorkingDays[monthKey] ?? 0;
      const priorWorkingDays = priorWorkingDaysAccumulator;
      priorWorkingDaysAccumulator += currentMonthWorkingDays;

      const totalSchoolWorkingDaysTillDate = currentMonthWorkingDays + priorWorkingDays;

      const activeStudentsInMonth = students.filter(s => isStudentActiveInMonth(s, monthKey, attendance));
      const studentsCountInMonth = activeStudentsInMonth.length;

      // Calculate total present student-days in current month
      let monthPresentSum = 0;
      activeStudentsInMonth.forEach(s => {
        monthPresentSum += annualStudentAttendance[s.id]?.[monthKey] ?? 0;
      });

      // Sum of all students' Total Days up to this month row
      let sumOfAllStudentsTotalDays = 0;
      activeStudentsInMonth.forEach(s => {
        let studentTotalDaysUpToCurrent = 0;
        const currentMonthIdx = academicMonths.findIndex(m => format(m, 'yyyy-MM') === monthKey);
        const monthsUpToCurrent = academicMonths.slice(0, currentMonthIdx + 1);
        monthsUpToCurrent.forEach(pDate => {
          const pKey = format(pDate, 'yyyy-MM');
          studentTotalDaysUpToCurrent += annualStudentAttendance[s.id]?.[pKey] ?? 0;
        });
        sumOfAllStudentsTotalDays += studentTotalDaysUpToCurrent;
      });

      // Average Daily Attendance = Total Present / Total Working Days
      const averageDailyAttendance = currentMonthWorkingDays > 0 ? (monthPresentSum / currentMonthWorkingDays) : 0;

      // Average Monthly Attendance = % Monthly = Total Present / (Total Students * Working Days) * 100
      const averageMonthlyAttendance = (studentsCountInMonth > 0 && currentMonthWorkingDays > 0)
        ? (monthPresentSum / (studentsCountInMonth * currentMonthWorkingDays)) * 100
        : 0;

      // Average Attendance Till Date = Total Days(sum all total days) / [Total Students * (Current Month + Cumulative Days)] * 100
      const averageAttendanceTillDate = (studentsCountInMonth > 0 && totalSchoolWorkingDaysTillDate > 0)
        ? (sumOfAllStudentsTotalDays / (studentsCountInMonth * totalSchoolWorkingDaysTillDate)) * 100
        : 0;

      return {
        monthLabel: format(mDate, 'MMMM yyyy'),
        monthShort: format(mDate, 'MMM'),
        monthKey,
        currentMonthWorkingDays,
        priorWorkingDays,
        totalSchoolWorkingDaysTillDate,
        averageDailyAttendance,
        averageMonthlyAttendance,
        averageAttendanceTillDate
      };
    });

    return {
      monthlyDataList,
      totalWorkingDays
    };
  }, [students, academicMonths, annualWorkingDays, annualStudentAttendance, attendance]);

  const overallSummary = useMemo(() => {
    let totalFnP = 0, totalFnA = 0, totalAnP = 0, totalAnA = 0;
    for (let d = 1; d <= days; d++) {
      const fn = getDailySummary(d, 'fn');
      const an = getDailySummary(d, 'an');
      totalFnP += fn.p;
      totalFnA += fn.a;
      totalAnP += an.p;
      totalAnA += an.a;
    }
    return { totalFnP, totalFnA, totalAnP, totalAnA };
  }, [attendance, filteredStudents, days]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!user) {
      alert("Please log in to import students.");
      return;
    }

    setIsImporting(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];

      if (!jsonData || jsonData.length === 0) {
        alert("The selected Excel file is empty.");
        return;
      }

      // Dynamic header search in top 20 rows
      let headerRowIndex = -1;
      let nameIdx = -1;
      let admNoIdx = -1;
      let snoIdx = -1;
      let genderIdx = -1;

      const namePatterns = ['NAME', 'STUDENT NAME', 'STUDENT\'S NAME', 'NAME OF THE STUDENT', 'STUDENT', 'FULL NAME'];
      const admNoPatterns = ['ADM NO', 'ADM.NO', 'ADM. NO.', 'ADM NO.', 'ADMISSION NO', 'ADMISSION NUMBER', 'ROLL NO', 'ROLL.NO', 'ROLL NO.', 'REG NO', 'REG.NO', 'REGISTRATION NO', 'STUDENT ID', 'ID', 'ADM_NO'];
      const snoPatterns = ['S.NO', 'S.NO.', 'SL.NO', 'SL NO', 'SNO', 'SERIAL NO', 'SR NO', 'NO', 'S NO'];
      const genderPatterns = ['GENDER', 'SEX'];

      for (let r = 0; r < Math.min(jsonData.length, 20); r++) {
        const rowCells = (jsonData[r] || []).map(cell => String(cell).trim().toUpperCase());
        
        const foundName = rowCells.findIndex(h => namePatterns.some(p => h === p || h.includes(p)));
        const foundAdm = rowCells.findIndex(h => admNoPatterns.some(p => h === p || h.includes(p)));
        
        if (foundName !== -1 && foundAdm !== -1) {
          headerRowIndex = r;
          nameIdx = foundName;
          admNoIdx = foundAdm;
          snoIdx = rowCells.findIndex(h => snoPatterns.some(p => h === p || h.includes(p)));
          genderIdx = rowCells.findIndex(h => genderPatterns.some(p => h === p || h.includes(p)));
          break;
        }
      }

      // Fallback if not found together
      if (headerRowIndex === -1) {
        for (let r = 0; r < Math.min(jsonData.length, 20); r++) {
          const rowCells = (jsonData[r] || []).map(cell => String(cell).trim().toUpperCase());
          const foundName = rowCells.findIndex(h => namePatterns.some(p => h.includes(p)));
          if (foundName !== -1) {
            headerRowIndex = r;
            nameIdx = foundName;
            admNoIdx = rowCells.findIndex(h => admNoPatterns.some(p => h.includes(p)));
            snoIdx = rowCells.findIndex(h => snoPatterns.some(p => h.includes(p)));
            genderIdx = rowCells.findIndex(h => genderPatterns.some(p => h.includes(p)));
            break;
          }
        }
      }

      if (headerRowIndex === -1 || nameIdx === -1) {
        alert(
          "Could not detect student columns in the uploaded file.\n\n" +
          "Please make sure your Excel sheet contains a header row with columns like 'NAME' (or 'Student Name') and 'ADM NO' (or 'Admission No' / 'Roll No')."
        );
        return;
      }

      const rows = jsonData.slice(headerRowIndex + 1);
      const userId = user.uid;
      const currentMonthKey = format(monthDate, 'yyyy-MM');
      const newStudentsToSave: Student[] = [];

      rows.forEach((row) => {
        const rawName = nameIdx >= 0 ? row[nameIdx] : '';
        const nameStr = String(rawName || '').trim();
        if (!nameStr) return;

        const rawAdmNo = admNoIdx >= 0 ? row[admNoIdx] : '';
        let admNoStr = String(rawAdmNo || '').trim();
        if (!admNoStr) {
          admNoStr = 'TMP-' + Math.floor(1000 + Math.random() * 9000);
        }

        const rawSno = snoIdx >= 0 ? row[snoIdx] : '';
        const snoStr = String(rawSno || '').trim() || String(students.length + newStudentsToSave.length + 1);

        let gender: 'Boy' | 'Girl' = 'Boy';
        if (genderIdx >= 0 && row[genderIdx]) {
          const val = String(row[genderIdx]).trim().toLowerCase();
          if (val.startsWith('g') || val.startsWith('f') || val.includes('girl') || val.includes('female')) {
            gender = 'Girl';
          }
        }

        const studentId = doc(collection(db, `users/${userId}/students`)).id;
        newStudentsToSave.push({
          id: studentId,
          sno: snoStr,
          admNo: admNoStr,
          name: nameStr.toUpperCase(),
          gender,
          createdMonth: currentMonthKey
        });
      });

      if (newStudentsToSave.length === 0) {
        alert("No valid student rows found in the uploaded file.");
        return;
      }

      // Chunk batch writes (max 400 per batch)
      const chunkSize = 400;
      const saveStudentChunks = async (useLegacyFields: boolean) => {
        for (let i = 0; i < newStudentsToSave.length; i += chunkSize) {
          const chunk = newStudentsToSave.slice(i, i + chunkSize);
          const batch = writeBatch(db);
          chunk.forEach(s => {
            batch.set(
              doc(db, `users/${userId}/students/${s.id}`),
              useLegacyFields ? getLegacyStudentData(s) : getFullStudentData(s)
            );
          });
          await batch.commit();
        }
      };

      try {
        await saveStudentChunks(false);
      } catch (error) {
        if (!isPermissionDeniedError(error)) throw error;
        console.warn("Import retried with legacy student fields because Firestore rejected optional metadata:", error);
        try {
          await saveStudentChunks(true);
        } catch (legacyError) {
          if (!isPermissionDeniedError(legacyError)) throw legacyError;
          console.warn("Import saved locally because Firestore rejected student writes:", legacyError);
          persistLocalStudents(current => [...current, ...newStudentsToSave]);
        }
      }

      try {
        await sortAndRenumberStudents([...students, ...newStudentsToSave]);
      } catch (e) {
        console.warn("Post-import renumbering skipped:", e);
      }
      alert(`Successfully imported ${newStudentsToSave.length} student(s)!`);
    } catch (error) {
      console.error("Error parsing or uploading Excel file:", error);
      alert("Error importing file: " + (error instanceof Error ? error.message : String(error)));
    } finally {
      setIsImporting(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  const validationErrors = useMemo(() => {
    const errors: Record<string, string> = {};
    const admNos = new Set<string>();
    
    sortedStudentsForMonth.forEach(s => {
      if (!s.name.trim()) errors[`${s.id}-name`] = 'Name is required';
      
      const admNoTrimmed = s.admNo.trim();
      if (!admNoTrimmed) {
        errors[`${s.id}-admNo`] = 'Admission No is required';
      } else if (admNos.has(admNoTrimmed)) {
        errors[`${s.id}-admNo`] = 'Duplicate Admission No';
      }
      admNos.add(admNoTrimmed);
    });
    
    return errors;
  }, [sortedStudentsForMonth]);

  const hasErrors = Object.keys(validationErrors).length > 0;

  const exportToExcel = () => {
    if (hasErrors) {
      alert("Please fix validation errors before exporting.");
      return;
    }
    // Create a simple export of the current view
    const exportData = filteredStudents.map(s => {
      const totals = calculateTotals(s.id);
      const attKey = getAttKey(s.id);
      const row: any = {
        'S.NO': s.sno,
        'ADM NO': s.admNo,
        'NAME': s.name,
        'GENDER': s.gender || 'Boy',
      };
      
      for (let d = 1; d <= days; d++) {
        const att = attendance[attKey]?.[d] || { fn: '', an: '' };
        row[`${d} FN`] = att.fn;
        row[`${d} AN`] = att.an;
      }
      
      const currentMonthIdx = academicMonths.findIndex(m => format(m, 'yyyy-MM') === format(monthDate, 'yyyy-MM'));
      const priorMonths = currentMonthIdx > 0 ? academicMonths.slice(0, currentMonthIdx) : [];
      let previousDays = 0;
      priorMonths.forEach(mDate => {
        const mKey = format(mDate, 'yyyy-MM');
        previousDays += annualStudentAttendance[s.id]?.[mKey] ?? 0;
      });
      const totalDays = totals.present + previousDays;

      row['Total Present'] = totals.present;
      row['Previous Days'] = previousDays;
      row['Total Absent'] = totals.absent;
      row['Total Days'] = totalDays;
      row['% Monthly'] = totals.percentage.toFixed(1) + '%';
      row['Remarks'] = getAttendanceStatus(totals.percentage);
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance");
    XLSX.writeFile(wb, `Attendance_${format(monthDate, 'MMM_yyyy')}.xlsx`);
  };

  const openAddStudentModal = () => {
    const autoAdmNo = 'TMP-' + Math.floor(1000 + Math.random() * 9000);
    setNewStudentName('');
    setNewStudentAdmNo(autoAdmNo);
    setNewStudentGender('Boy');
    setShowAddStudentModal(true);
  };

  const openExcelImportPicker = () => {
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
      fileInputRef.current.click();
    }
  };

  const removeStudent = async (id: string) => {
    if (!user) return;
    const userId = user.uid;
    const currentMonthKey = format(monthDate, 'yyyy-MM');
    const attKey = `${currentMonthKey}_${id}`;
    const batch = writeBatch(db);

    try {
      batch.set(doc(db, `users/${userId}/students/${id}`), {
        deletedMonths: arrayUnion(currentMonthKey)
      }, { merge: true });
      batch.delete(doc(db, `users/${userId}/attendance/${attKey}`));
      await batch.commit();

      console.log(`Removed student ${id} from ${currentMonthKey} only.`);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `users/${userId}/students/${id}`);
      if (!isPermissionDeniedError(error)) throw error;

      persistLocalStudents(current => {
        const existingLocal = current.find(student => student.id === id);
        const existingRemote = students.find(student => student.id === id);
        const target = existingLocal || existingRemote;
        if (!target) return current;

        const deletedMonths = Array.from(new Set([...(target.deletedMonths || []), currentMonthKey]));
        const updatedStudent = { ...target, deletedMonths };
        return existingLocal
          ? current.map(student => student.id === id ? updatedStudent : student)
          : [...current, updatedStudent];
      });
    }
  };

  const updateStudent = async (id: string, field: keyof Student, value: string) => {
    if (!user) return;
    const userId = user.uid;
    try {
      await updateDoc(doc(db, `users/${userId}/students/${id}`), {
        [field]: value
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `users/${userId}/students/${id}`);
    }
  };

  const sortAndRenumberStudents = async (studentList: Student[]) => {
    if (!user || studentList.length === 0) return;
    const userId = user.uid;
    
    // Sort Boys alphabetically, then Girls alphabetically
    const boys = studentList.filter(s => (s.gender || 'Boy') === 'Boy')
      .sort((a, b) => a.name.trim().toLowerCase().localeCompare(b.name.trim().toLowerCase()));
    const girls = studentList.filter(s => s.gender === 'Girl')
      .sort((a, b) => a.name.trim().toLowerCase().localeCompare(b.name.trim().toLowerCase()));
    
    const sorted = [...boys, ...girls];
    
    // Filter only students whose SNO has actually changed
    const changedStudents = sorted.filter((s, idx) => s.sno !== String(idx + 1));
    if (changedStudents.length === 0) return;

    const chunkSize = 400;
    for (let i = 0; i < changedStudents.length; i += chunkSize) {
      const chunk = changedStudents.slice(i, i + chunkSize);
      const batch = writeBatch(db);
      chunk.forEach((s) => {
        const targetIndex = sorted.findIndex(item => item.id === s.id);
        batch.set(
          doc(db, `users/${userId}/students/${s.id}`),
          getFullStudentData({ ...s, sno: String(targetIndex + 1) }),
          { merge: true }
        );
      });
      try {
        await batch.commit();
      } catch (error) {
        if (!isPermissionDeniedError(error)) {
          console.warn("Renumbering batch update warning:", error);
          continue;
        }

        console.warn("Renumbering retried with legacy student fields because Firestore rejected optional metadata:", error);
        const legacyBatch = writeBatch(db);
        chunk.forEach((s) => {
          const targetIndex = sorted.findIndex(item => item.id === s.id);
          legacyBatch.set(
            doc(db, `users/${userId}/students/${s.id}`),
            getLegacyStudentData({ ...s, sno: String(targetIndex + 1) }),
            { merge: true }
          );
        });
        try {
          await legacyBatch.commit();
        } catch (legacyError) {
          console.warn("Legacy renumbering batch update warning:", legacyError);
        }
      }
    }
  };

  // ── Add Student ──
  const handleAddStudentSave = async () => {
    if (!user) {
      alert("Please log in to add students.");
      return;
    }
    if (!newStudentName.trim() || isAddingStudent) return;

    setIsAddingStudent(true);
    const userId = user.uid;
    const studentRef = doc(collection(db, `users/${userId}/students`));
    const studentId = studentRef.id;
    const admNo = newStudentAdmNo.trim() || ('TMP-' + Math.floor(1000 + Math.random() * 9000));
    const savedName = newStudentName.trim().toUpperCase();
    const currentMonthKey = format(monthDate, 'yyyy-MM');

    try {
      const newStudent: Student = {
        id: studentId,
        sno: String(students.length + 1),
        admNo: admNo,
        name: savedName,
        gender: newStudentGender,
        createdMonth: currentMonthKey
      };

      try {
        await setDoc(studentRef, getFullStudentData(newStudent));
      } catch (error) {
        if (!isPermissionDeniedError(error)) throw error;
        console.warn("Add student retried with legacy student fields because Firestore rejected optional metadata:", error);
        try {
          await setDoc(studentRef, getLegacyStudentData(newStudent));
        } catch (legacyError) {
          if (!isPermissionDeniedError(legacyError)) throw legacyError;
          console.warn("Add student saved locally because Firestore rejected student writes:", legacyError);
          persistLocalStudents(current => [...current, newStudent]);
        }
      }

      try {
        await sortAndRenumberStudents([...students, newStudent]);
      } catch (e) {
        console.warn("Renumbering skipped:", e);
      }
      
      setShowAddStudentModal(false);
      setNewStudentName('');
      setNewStudentAdmNo('');
      setNewStudentGender('Boy');
      setSearchQuery('');
      setLastAddedStudentId(studentId);
    } catch (error) {
      console.error("Error adding student:", error);
      handleFirestoreError(error, OperationType.WRITE, `users/${userId}/students/${studentId}`);
      alert("Failed to add student: " + (error instanceof Error ? error.message : String(error)));
    } finally {
      setIsAddingStudent(false);
    }
  };

  // ── Edit Student ──
  const openEditModal = (student: Student) => {
    setEditStudentName(student.name);
    setEditStudentAdmNo(student.admNo);
    setEditStudentGender((student.gender || 'Boy') as 'Boy' | 'Girl');
    setShowEditStudentModal(student.id);
  };

  const handleEditStudentSave = async () => {
    if (!user || !showEditStudentModal || !editStudentName.trim()) return;
    const userId = user.uid;
    const studentId = showEditStudentModal;
    const originalStudent = students.find(s => s.id === studentId);
    if (!originalStudent) return;
 
    const savedName = editStudentName.trim().toUpperCase();
    const nameChanged = savedName !== originalStudent.name.trim().toUpperCase();
    const genderChanged = editStudentGender !== originalStudent.gender;
 
    try {
      await updateDoc(doc(db, `users/${userId}/students/${studentId}`), {
        name: savedName,
        admNo: editStudentAdmNo.trim(),
        gender: editStudentGender
      });
 
      if (nameChanged || genderChanged) {
        // Rebuild list with updated name and gender for sorting
        const updatedList = students.map(s =>
          s.id === studentId
            ? { ...s, name: savedName, admNo: editStudentAdmNo.trim(), gender: editStudentGender }
            : s
        );
        await sortAndRenumberStudents(updatedList);
      }
 
      setShowEditStudentModal(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `users/${userId}/students/${studentId}`);
    }
  };

  const openPrintModal = () => {
    if (hasErrors) return;
    setPrintOrientation((viewMode === 'register' || viewMode === 'annual' || viewMode === 'genderwise') ? 'l' : 'p');
    setPrintSize((viewMode === 'register' || viewMode === 'annual' || viewMode === 'genderwise') ? 'legal' : 'a4');
    setPrintMargin(14);
    setShowPrintModal(true);
  };

  const generatePDF = async () => {
    setIsExporting(true);
    setShowPrintModal(false);
    try {
      const doc = new jsPDF(printOrientation, 'mm', printSize);
      const pageWidth = doc.internal.pageSize.getWidth();
      
      const centerX = pageWidth / 2;
      const rightX = pageWidth - printMargin;

      doc.setFontSize(14);
      doc.text(institutionName, centerX, 15, { align: 'center' });
      
      if (viewMode === 'register') {
        doc.setFontSize(14);
        doc.text(`REGISTER OF ATTENDANCE FOR MONTH OF : ${format(monthDate, 'MMMM yyyy')}`, printMargin, 22);
        
        doc.setFontSize(10);
        doc.text(`Standard: ${standard}    Section: ${section}    Total Working Days: ${globalWorkingDays}`, printMargin, 28);
        
        const head: any[] = [
          [
            { content: 'S.NO', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'ADM NO', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'NAME', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'DATE', colSpan: days * 2, styles: { halign: 'center' } },
            { content: 'Present', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'Previous Days', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'Absent', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'Total Days', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: '% Monthly', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'Remarks', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } }
          ],
          dayArray.flatMap(d => [{ content: d.toString(), colSpan: 2, styles: { halign: 'center' } }])
        ];

        const subHead = [
          '', '', '', ...dayArray.flatMap(() => ['FN', 'AN']), '', '', '', '', '', ''
        ];
        head.push(subHead.map(text => ({ content: text, styles: { halign: 'center', fontSize: 6 } } as any)));

        const body = filteredStudents.map(s => {
          const totals = calculateTotals(s.id);
          const attKey = getAttKey(s.id);
          
          const currentMonthIdx = academicMonths.findIndex(m => format(m, 'yyyy-MM') === format(monthDate, 'yyyy-MM'));
          const priorMonths = currentMonthIdx > 0 ? academicMonths.slice(0, currentMonthIdx) : [];
          let previousDays = 0;
          priorMonths.forEach(mDate => {
            const mKey = format(mDate, 'yyyy-MM');
            previousDays += annualStudentAttendance[s.id]?.[mKey] ?? 0;
          });
          const totalDays = totals.present + previousDays;

          return [
            s.sno,
            s.admNo,
            s.name,
            ...dayArray.flatMap(d => [
              attendance[attKey]?.[d]?.fn || '',
              attendance[attKey]?.[d]?.an || ''
            ]),
            totals.present.toString(),
            previousDays.toString(),
            totals.absent.toString(),
            totalDays.toString(),
            `${totals.percentage.toFixed(1)}%`,
            getAttendanceStatus(totals.percentage)
          ];
        });

        autoTable(doc, {
          startY: 32,
          margin: { top: printMargin, right: printMargin, bottom: printMargin, left: printMargin },
          head: head,
          body: body,
          theme: 'grid',
          styles: { fontSize: 6, cellPadding: 0.5 },
          headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], lineWidth: 0.1, lineColor: [0, 0, 0] },
          bodyStyles: { lineWidth: 0.1, lineColor: [0, 0, 0] },
          columnStyles: {
            0: { halign: 'center' }, // S.NO
            1: { halign: 'center' }, // ADM NO
            // DATE columns are handled by subHead styles
            [days * 2 + 3]: { halign: 'center' }, // Present
            [days * 2 + 4]: { halign: 'center' }, // Previous Days
            [days * 2 + 5]: { halign: 'center' }, // Absent
            [days * 2 + 6]: { halign: 'center' }, // Total Days
            [days * 2 + 7]: { halign: 'center' }, // % Monthly
          },
          alternateRowStyles: { fillColor: [250, 250, 250] },
        });

        doc.save(`Attendance_${format(monthDate, 'MMM_yyyy')}.pdf`);
      } else if (viewMode === 'daily') {
        doc.setFontSize(12);
        doc.text(`DAILY ATTENDANCE REPORT`, centerX, 22, { align: 'center' });
        
        doc.setFontSize(10);
        doc.text(`Date: ${reportDay} ${format(monthDate, 'MMM yyyy')}`, printMargin, 30);
        doc.text(`Standard: ${standard}    Section: ${section}`, centerX, 30, { align: 'center' });
        doc.text(`Month: ${format(monthDate, 'MMMM-yyyy')}`, rightX, 30, { align: 'right' });

        const head = [['S.No', 'Student Name', 'FN Status(P/A)', 'AN Status(P/A)']];
        const body = filteredStudents.map(student => {
          const attKey = getAttKey(student.id);
          const fn = attendance[attKey]?.[reportDay]?.fn || '-';
          const an = attendance[attKey]?.[reportDay]?.an || '-';
          return [student.sno, student.name, fn, an];
        });

        autoTable(doc, {
          startY: 35,
          margin: { top: printMargin, right: printMargin, bottom: printMargin, left: printMargin },
          head: head,
          body: body,
          theme: 'grid',
          headStyles: { fillColor: [243, 244, 246], textColor: [0, 0, 0], fontStyle: 'bold', lineWidth: 0.1, lineColor: [0, 0, 0] },
          styles: { fontSize: 10, cellPadding: 2, textColor: [0, 0, 0] },
          bodyStyles: { fillColor: [255, 255, 255], lineWidth: 0.1, lineColor: [0, 0, 0] },
          columnStyles: {
            0: { cellWidth: 20, halign: 'center' },
            2: { cellWidth: 35, halign: 'center' },
            3: { cellWidth: 35, halign: 'center' }
          }
        });
        
        doc.save(`Daily_Report_${reportDay}_${format(monthDate, 'MMM_yyyy')}.pdf`);
        
      } else if (viewMode === 'monthly') {
        doc.setFontSize(12);
        doc.text(`MONTHLY ATTENDANCE REPORT : ${format(monthDate, 'MMMM yyyy')}`, centerX, 22, { align: 'center' });
        
        doc.setFontSize(10);
        doc.text(`Total Working Days: ${globalWorkingDays}    Month: ${format(monthDate, 'MMMM-yyyy')}`, printMargin, 30);
        doc.text(`Standard: ${standard}    Section: ${section}`, rightX, 30, { align: 'right' });

        const head = [['S.No', 'Student Name', 'Total Present', 'Total\nAbsent', 'Attendance %', 'Remarks']];
        const body = filteredStudents.map(student => {
          const totals = calculateTotals(student.id);
          const status = totals.percentage >= 75 ? 'Regular' : 'Irregular';
          return [
            student.sno,
            student.name,
            totals.present.toString(),
            totals.absent.toString(),
            `${totals.percentage.toFixed(1)}%`,
            status
          ];
        });

        autoTable(doc, {
          startY: 35,
          margin: { top: printMargin, right: printMargin, bottom: printMargin, left: printMargin },
          head: head,
          body: body,
          theme: 'grid',
          headStyles: { fillColor: [243, 244, 246], textColor: [0, 0, 0], fontStyle: 'bold', lineWidth: 0.1, lineColor: [0, 0, 0] },
          styles: { fontSize: 10, cellPadding: 2, textColor: [0, 0, 0] },
          bodyStyles: { fillColor: [255, 255, 255], lineWidth: 0.1, lineColor: [0, 0, 0] },
          columnStyles: {
            0: { cellWidth: 15, halign: 'center' },
            2: { cellWidth: 25, halign: 'center' },
            3: { cellWidth: 25, halign: 'center' },
            4: { cellWidth: 25, halign: 'center' },
            5: { cellWidth: 25, halign: 'center' }
          }
        });
        
        doc.save(`Monthly_Report_${format(monthDate, 'MMM_yyyy')}.pdf`);
      } else if (viewMode === 'annual') {
        doc.setFillColor(255, 255, 255);
        doc.rect(0, 0, pageWidth, 38, 'F');

        doc.setTextColor(0, 0, 0);
        doc.setFontSize(20);
        doc.setFont(undefined, 'bold');
        doc.text(institutionName.toUpperCase(), centerX, 15, { align: 'center' });
        
        doc.setTextColor(0, 0, 0);
        doc.setFontSize(13);
        doc.setFont(undefined, 'bold');
        doc.text(`ANNUAL ATTENDANCE REPORT - STUDENTWISE`, printMargin, 25);
        doc.text(academicYearLabel, rightX, 25, { align: 'right' });
        
        doc.setTextColor(0, 0, 0);
        doc.setFontSize(11);
        doc.setFont(undefined, 'bold');
        doc.text(`Class-Div: ${standard || '—'} ${section || ''}       Total Students: ${students.length}`, printMargin, 33);
        
        const head = [
          [
            { content: 'Roll No.', rowSpan: 2, styles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], valign: 'middle' } },
            { content: "Student's Name", rowSpan: 2, styles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], valign: 'middle', halign: 'left', overflow: 'visible' } },
            { content: 'Academic Months', colSpan: academicMonths.length, styles: { fillColor: [49, 46, 129], textColor: [255, 255, 255], halign: 'center' } },
            { content: 'Total Days\nPresent', rowSpan: 2, styles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], valign: 'middle' } },
            { content: 'Total Working\nDays', rowSpan: 2, styles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], valign: 'middle' } },
            { content: '%', rowSpan: 2, styles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], valign: 'middle' } }
          ],
          [
            ...academicMonths.map(mDate => ({ content: format(mDate, 'MMM'), styles: { fillColor: [30, 27, 75], textColor: [255, 255, 255], halign: 'center' } }))
          ]
        ];
        
        const body: any[] = [];
        
        // Month working days row
        const workingDaysRow = [
          {
            content: 'Total monthly working days',
            colSpan: 2,
            styles: {
              halign: 'left',
              valign: 'middle',
              fontStyle: 'bold',
              fontSize: 8.5,
              cellPadding: { top: 1.8, right: 2, bottom: 1.8, left: 2 }
            }
          },
          ...academicMonths.map(mDate => {
            const monthKey = format(mDate, 'yyyy-MM');
            return (annualWorkingDays[monthKey] ?? 0).toString();
          }),
          annualSummary.totalWorkingDays.toString(),
          annualSummary.totalWorkingDays.toString(),
          '100%'
        ];
        body.push(workingDaysRow);
        
        // Student rows
        filteredStudents.forEach(student => {
          let studentTotalPresent = 0;
          const row = [
            student.sno,
            student.name,
            ...academicMonths.map(mDate => {
              const monthKey = format(mDate, 'yyyy-MM');
              const presentCount = annualStudentAttendance[student.id]?.[monthKey] ?? 0;
              studentTotalPresent += presentCount;
              return presentCount > 0 ? presentCount.toString() : '-';
            }),
            studentTotalPresent.toFixed(1).replace('.0', ''),
            annualSummary.totalWorkingDays.toString(),
            `${(annualSummary.totalWorkingDays > 0 ? (studentTotalPresent / annualSummary.totalWorkingDays) * 100 : 0).toFixed(1)}%`
          ];
          body.push(row);
        });

        const colStyles: any = {
          0: { cellWidth: 12, halign: 'center' }, // Roll No.
          1: { cellWidth: 42, overflow: 'visible', halign: 'left', fontSize: 7.6 }, // Name
        };
        for (let i = 0; i < academicMonths.length; i++) {
          colStyles[2 + i] = { halign: 'center' };
        }
        colStyles[2 + academicMonths.length] = { cellWidth: 15, halign: 'center', fontStyle: 'bold' };
        colStyles[2 + academicMonths.length + 1] = { cellWidth: 15, halign: 'center' };
        colStyles[2 + academicMonths.length + 2] = { cellWidth: 12, halign: 'center', fontStyle: 'bold' };
        
        autoTable(doc, {
          startY: 39,
          margin: { top: printMargin, right: printMargin, bottom: printMargin, left: printMargin },
          head: head,
          body: body,
          theme: 'grid',
          headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8.5, lineWidth: 0.1, lineColor: [0, 0, 0], halign: 'center', valign: 'middle' },
          styles: { fontSize: 8.5, cellPadding: 1.35, textColor: [0, 0, 0], lineWidth: 0.1, lineColor: [0, 0, 0], valign: 'middle', overflow: 'linebreak' },
          bodyStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], lineWidth: 0.1, lineColor: [0, 0, 0] },
          alternateRowStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0] },
          columnStyles: colStyles,
          didParseCell: function(data) {
            if (data.section === 'head') {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fontSize = 8.5;
              data.cell.styles.halign = data.cell.styles.halign || 'center';
              data.cell.styles.valign = 'middle';
            }

            if (data.section === 'body' && data.row.index === 0) {
              data.cell.styles.fillColor = [255, 248, 225];
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fontSize = 8.5;

              if (data.column.index === 0) {
                data.cell.styles.overflow = 'visible';
                data.cell.styles.halign = 'left';
              }

              if (data.column.index >= 2 + academicMonths.length) {
                data.cell.styles.fillColor = [255, 236, 179];
                data.cell.styles.textColor = [30, 58, 138];
              }
            }
          }
        });
        
        doc.save(`Annual_Report_${academicYearLabel.replace(/ /g, '')}.pdf`);
      } else if (viewMode === 'genderwise') {
        doc.setFontSize(14);
        doc.text(institutionName, centerX, 15, { align: 'center' });
        
        doc.setFontSize(11);
        doc.text(`ANNUAL ATTENDANCE REPORT - GENDERWISE`, printMargin, 23);
        doc.text(academicYearLabel, rightX, 23, { align: 'right' });
        
        doc.setFontSize(10);
        doc.text(`Class-Div: ${standard || '—'} ${section || ''}       Boys: ${genderwiseSummary.boysCount}   Girls: ${genderwiseSummary.girlsCount}`, printMargin, 30);
        
        const head = [
          ['Month', ...academicMonths.map(mDate => format(mDate, 'MMM')), 'Average Annual Presence']
        ];
        
        const body = [
          [
            'Boys Present',
            ...academicMonths.map(mDate => {
              const val = genderwiseSummary.monthlyData[format(mDate, 'yyyy-MM')]?.boysPresent ?? 0;
              return val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-';
            }),
            genderwiseSummary.boysCount > 0 ? `${genderwiseSummary.boysAvgAnnualPresence.toFixed(1)}%` : '—'
          ],
          [
            'Girls Present',
            ...academicMonths.map(mDate => {
              const val = genderwiseSummary.monthlyData[format(mDate, 'yyyy-MM')]?.girlsPresent ?? 0;
              return val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-';
            }),
            genderwiseSummary.girlsCount > 0 ? `${genderwiseSummary.girlsAvgAnnualPresence.toFixed(1)}%` : '—'
          ],
          [
            'Boys Absent',
            ...academicMonths.map(mDate => {
              const val = genderwiseSummary.monthlyData[format(mDate, 'yyyy-MM')]?.boysAbsent ?? 0;
              return val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-';
            }),
            '—'
          ],
          [
            'Girls Absent',
            ...academicMonths.map(mDate => {
              const val = genderwiseSummary.monthlyData[format(mDate, 'yyyy-MM')]?.girlsAbsent ?? 0;
              return val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-';
            }),
            '—'
          ]
        ];

        const colStyles: any = {
          0: { cellWidth: 35, fontStyle: 'bold' }, // Label
        };
        for (let i = 0; i < academicMonths.length; i++) {
          colStyles[1 + i] = { halign: 'center' };
        }
        colStyles[1 + academicMonths.length] = { cellWidth: 35, halign: 'center', fontStyle: 'bold' };
        
        autoTable(doc, {
          startY: 35,
          margin: { top: printMargin, right: printMargin, bottom: printMargin, left: printMargin },
          head: head,
          body: body,
          theme: 'grid',
          headStyles: { fillColor: [30, 58, 138], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8, lineWidth: 0.1, lineColor: [0, 0, 0], halign: 'center', valign: 'middle' },
          styles: { fontSize: 8, cellPadding: 2, textColor: [0, 0, 0] },
          bodyStyles: { fillColor: [255, 255, 255], lineWidth: 0.1, lineColor: [0, 0, 0] },
          columnStyles: colStyles
        });
        
        doc.save(`Genderwise_Report_${academicYearLabel.replace(/ /g, '')}.pdf`);
      } else if (viewMode === 'monthwise') {
        doc.setFontSize(14);
        doc.text(institutionName, centerX, 15, { align: 'center' });
        
        doc.setFontSize(11);
        doc.text(`ATTENDANCE MONTHWISE REPORT`, printMargin, 23);
        doc.text(academicYearLabel, rightX, 23, { align: 'right' });
        
        doc.setFontSize(10);
        doc.text(`Class-Div: ${standard || '—'} ${section || ''}       Total Working Days: ${monthwiseSummary.totalWorkingDays}`, printMargin, 30);
        
        const head: any[] = [
          [
            { content: 'Months', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
            { content: 'Total Working Days', colSpan: 3, styles: { halign: 'center' } },
            { content: 'Different Averages Of Attendance', colSpan: 3, styles: { halign: 'center' } }
          ],
          [
            { content: 'Current Month', styles: { halign: 'center' } },
            { content: 'Cumulative Days', styles: { halign: 'center' } },
            { content: 'Total Days', styles: { halign: 'center' } },
            { content: 'Average Daily Attendance', styles: { halign: 'center' } },
            { content: 'Average Monthly Attendance', styles: { halign: 'center' } },
            { content: 'Average Attendance Till Date', styles: { halign: 'center' } }
          ]
        ];
        
        const body = monthwiseSummary.monthlyDataList.map(row => [
          row.monthLabel,
          row.currentMonthWorkingDays.toString(),
          row.priorWorkingDays.toString(),
          row.totalSchoolWorkingDaysTillDate.toString(),
          row.averageDailyAttendance.toFixed(1).replace('.0', ''),
          `${row.averageMonthlyAttendance.toFixed(1)}%`,
          `${row.averageAttendanceTillDate.toFixed(1)}%`
        ]);
        
        autoTable(doc, {
          startY: 35,
          margin: { top: printMargin, right: printMargin, bottom: printMargin, left: printMargin },
          head: head,
          body: body,
          theme: 'grid',
          headStyles: { fillColor: [30, 58, 138], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8, lineWidth: 0.1, lineColor: [0, 0, 0], halign: 'center', valign: 'middle' },
          styles: { fontSize: 8, cellPadding: 2, textColor: [0, 0, 0] },
          bodyStyles: { fillColor: [255, 255, 255], lineWidth: 0.1, lineColor: [0, 0, 0] },
          columnStyles: {
            0: { cellWidth: 40, fontStyle: 'bold' }, // Label
            1: { halign: 'center' },
            2: { halign: 'center' },
            3: { halign: 'center' },
            4: { halign: 'center', fontStyle: 'bold' },
            5: { halign: 'center', fontStyle: 'bold' },
            6: { halign: 'center', fontStyle: 'bold' }
          }
        });
        
        doc.save(`Monthwise_Report_${academicYearLabel.replace(/ /g, '')}.pdf`);
      }
      
    } catch (error) {
      console.error("Error generating PDF:", error);
    } finally {
      setIsExporting(false);
    }
  };

  const getAuthErrorMessage = (error: any) => {
    switch (error.code) {
      case 'auth/user-not-found':
        return 'No account found with this email. Please sign up first.';
      case 'auth/wrong-password':
      case 'auth/invalid-credential':
        return 'Incorrect email or password. Please try again.';
      case 'auth/invalid-email':
        return 'Please enter a valid email address.';
      case 'auth/email-already-in-use':
        return 'An account already exists with this email.';
      case 'auth/weak-password':
        return 'Password should be at least 6 characters.';
      case 'auth/operation-not-allowed':
        return 'Email/Password login is not enabled in the Firebase Console. Please enable it in the Authentication tab.';
      default:
        return error.message || 'An unexpected error occurred.';
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSubmitting(true);
    try {
      const credential = await signInWithEmailAndPassword(auth, email, password);
      setUser(credential.user);
    } catch (error: any) {
      console.error("Login Error:", error);
      alert(getAuthErrorMessage(error));
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSignUp = async (e: React.FormEvent) => {
    e.preventDefault();
    if (password.length < 6) {
      alert("Password must be at least 6 characters long.");
      return;
    }
    setIsSubmitting(true);
    try {
      const credential = await createUserWithEmailAndPassword(auth, email, password);
      setUser(credential.user);
      alert("Account created successfully! You are now logged in.");
    } catch (error: any) {
      console.error("Sign Up Error:", error);
      alert(getAuthErrorMessage(error));
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSaveSettings = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user) return;
    const userId = user.uid;
    
    const updates: any = {
      institutionName,
      standard,
      section,
      adminUsername,
      adminPassword
    };
    
    if (newUsername.trim()) updates.adminUsername = newUsername.trim();
    if (newPassword.trim()) updates.adminPassword = newPassword.trim();

    try {
      await setDoc(doc(db, `users/${userId}/config/settings`), updates, { merge: true });
      setNewUsername('');
      setNewPassword('');
      setShowSettings(false);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `users/${userId}/config/settings`);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error("Logout Error:", error);
    }
  };

  const confirmDelete = async (id: string) => {
    if (isDeletingStudent) return;
    setIsDeletingStudent(true);

    try {
      await removeStudent(id);
      setShowDeleteConfirm(null);
    } catch (error) {
      console.error("Could not remove student from selected month:", error);
      alert("Unable to delete student. Please check your connection and try again.");
    } finally {
      setIsDeletingStudent(false);
    }
  };

  if (isAuthLoading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <Loader2 className="w-12 h-12 text-blue-600 animate-spin" />
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{fontFamily: 'system-ui, -apple-system, sans-serif', background: 'linear-gradient(160deg, #f0f4ff 0%, #e8eeff 100%)'}}>

        {/* ── Centered login panel ── */}
        <div className="flex items-center justify-center p-6 w-full">
          <div className="w-full max-w-md">

            {/* Card */}
            <div className="rounded-3xl overflow-hidden" style={{boxShadow: '0 32px 80px rgba(30,64,175,0.18), 0 0 0 1px rgba(30,64,175,0.08)'}}>

              {/* card header */}
              <div className="px-8 pt-8 pb-7" style={{background: 'linear-gradient(135deg, #0c1e4a 0%, #1e40af 100%)'}}>
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{background: 'rgba(250,204,21,0.18)', border: '1px solid rgba(250,204,21,0.35)'}}>
                    <Lock className="w-3.5 h-3.5 text-yellow-400"/>
                  </div>
                  <span className="text-blue-300 text-[11px] font-bold tracking-widest uppercase">
                    {authMode === 'login' ? 'Secure Login' : 'Create Account'}
                  </span>
                </div>
                <h2 className="text-2xl font-black text-white mb-1" style={{fontFamily: 'Georgia, serif'}}>
                  {authMode === 'login' ? 'School Attendance Register' : 'Get Started'}
                </h2>
                <p className="text-blue-300 text-sm">
                  {authMode === 'login'
                    ? 'Sign in to access your attendance dashboard'
                    : 'Create your account to begin managing attendance'}
                </p>
              </div>

              {/* form body */}
              <div className="bg-white px-8 py-7">
                <form onSubmit={authMode === 'login' ? handleLogin : handleSignUp} className="space-y-5">
                  <div>
                    <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest mb-2">Email Address</label>
                    <input
                      type="email"
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      className="w-full px-4 py-3 rounded-xl text-sm text-gray-900 outline-none transition-all"
                      style={{background: '#f0f4ff', border: '2px solid #e0e7ff'}}
                      onFocus={e => {e.target.style.borderColor = '#2563eb'; e.target.style.background = '#fff';}}
                      onBlur={e => {e.target.style.borderColor = '#e0e7ff'; e.target.style.background = '#f0f4ff';}}
                      placeholder="your@email.com"
                      required
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest mb-2">Password</label>
                    <input
                      type="password"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      className="w-full px-4 py-3 rounded-xl text-sm text-gray-900 outline-none transition-all"
                      style={{background: '#f0f4ff', border: '2px solid #e0e7ff'}}
                      onFocus={e => {e.target.style.borderColor = '#2563eb'; e.target.style.background = '#fff';}}
                      onBlur={e => {e.target.style.borderColor = '#e0e7ff'; e.target.style.background = '#f0f4ff';}}
                      placeholder={authMode === 'login' ? '••••••••' : 'Min. 6 characters'}
                      required
                    />
                  </div>
                  <button
                    type="submit"
                    disabled={isSubmitting}
                    className="w-full text-white font-black py-3.5 rounded-xl text-sm tracking-wider flex items-center justify-center gap-2 transition-all mt-1"
                    style={{
                      background: isSubmitting ? '#93c5fd' : 'linear-gradient(135deg, #1e40af 0%, #2563eb 100%)',
                      boxShadow: isSubmitting ? 'none' : '0 6px 20px rgba(37,99,235,0.38)',
                    }}
                  >
                    {isSubmitting && <Loader2 className="w-4 h-4 animate-spin"/>}
                    {authMode === 'login' ? '⚡  SIGN IN TO DASHBOARD' : '🚀  CREATE ACCOUNT'}
                  </button>
                </form>

                <div className="mt-6 pt-5 border-t border-gray-100 text-center">
                  <button
                    onClick={() => setAuthMode(authMode === 'login' ? 'signup' : 'login')}
                    disabled={isSubmitting}
                    className="text-sm font-semibold text-blue-600 hover:text-blue-800 disabled:opacity-50 transition-colors"
                  >
                    {authMode === 'login'
                      ? "Don't have an account? " 
                      : 'Already have an account? '}
                    <span className="underline underline-offset-2">
                      {authMode === 'login' ? 'Sign Up' : 'Login'}
                    </span>
                  </button>
                </div>
              </div>
            </div>


          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen print:bg-white" style={{background: '#eef2ff', fontFamily: 'system-ui, -apple-system, sans-serif'}}>

      {/* ═══ TOP NAVIGATION BAR ═══ */}
      <div className="print:hidden sticky top-0 z-40" style={{background: 'linear-gradient(135deg, #0c1e4a 0%, #1a3580 55%, #1e56c8 100%)', boxShadow: '0 4px 24px rgba(12,30,74,0.45)'}}>
        <div className="max-w-screen-2xl mx-auto px-5 h-[60px] flex items-center justify-between gap-4">

          {/* Brand */}
          <div className="flex items-center gap-3 flex-shrink-0">
            <div className="w-9 h-9 rounded-xl flex items-center justify-center flex-shrink-0" style={{background: 'linear-gradient(135deg, #facc15 0%, #f59e0b 100%)'}}>
              <GraduationCap className="text-blue-950" style={{width: 20, height: 20}}/>
            </div>
            <div className="hidden sm:block">
              <p className="text-white font-black text-[15px] leading-tight" style={{fontFamily: 'Georgia, serif'}}>School Attendance Register</p>
              <p className="text-blue-300 text-[10px] leading-none truncate max-w-[260px]">{institutionName}</p>
            </div>
          </div>

          {/* Action buttons */}
          <div className="flex items-center gap-1.5 flex-shrink-0">
            <button onClick={() => setShowHolidayManager(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-bold text-blue-200 hover:text-white hover:bg-white/10 transition-all">
              <CalendarDays style={{width: 13, height: 13}}/><span className="hidden md:inline">Holidays</span>
            </button>
            <button onClick={openAddStudentModal}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-bold text-blue-200 hover:text-white hover:bg-white/10 transition-all">
              <Plus style={{width: 13, height: 13}}/><span className="hidden md:inline">Add Student</span>
            </button>
            <button onClick={() => setShowSettings(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-bold text-blue-200 hover:text-white hover:bg-white/10 transition-all">
              <Settings style={{width: 13, height: 13}}/><span className="hidden md:inline">Settings</span>
            </button>
            <div className="w-px h-5 bg-white/20 mx-1"/>
            <button onClick={handleLogout}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-bold text-red-300 hover:text-white hover:bg-red-500/25 border border-red-500/30 transition-all">
              <LogOut style={{width: 13, height: 13}}/><span className="hidden sm:inline">Logout</span>
            </button>
          </div>
        </div>
      </div>

      {/* ═══ DASHBOARD STATS CARDS ═══ */}
      <div className="print:hidden max-w-screen-2xl mx-auto px-5 pt-6 pb-2">
        <div className="mb-4 flex items-end justify-between">
          <div>
            <h2 className="text-[22px] font-black text-gray-800 leading-tight" style={{fontFamily: 'Georgia, serif'}}>Dashboard Overview</h2>
            <p className="text-xs text-gray-500 mt-0.5">Live analytics for <span className="font-bold text-blue-700">{todayLabel}</span></p>
          </div>
        </div>

        <div className="grid grid-cols-2 xl:grid-cols-4 gap-4 mb-5">
          {/* Students */}
          <div className="rounded-2xl p-5 relative overflow-hidden flex items-center gap-4"
            style={{background: 'linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)', boxShadow: '0 8px 24px rgba(37,99,235,0.28)'}}>
            <div className="absolute -right-4 -top-4 w-24 h-24 rounded-full bg-white/10"/>
            <div className="w-11 h-11 rounded-xl bg-white/20 flex items-center justify-center flex-shrink-0">
              <Users className="text-white" style={{width: 22, height: 22}}/>
            </div>
            <div>
              <p className="text-blue-200 text-xs font-semibold">Total Students</p>
              <p className="text-3xl font-black text-white leading-none mt-0.5">{filteredStudents.length}</p>
            </div>
          </div>

          {/* Morning Session */}
          <div className="rounded-2xl p-5 relative overflow-hidden flex items-center gap-4"
            style={{background: 'linear-gradient(135deg, #055160 0%, #0284c7 100%)', boxShadow: '0 8px 24px rgba(2,132,199,0.28)'}}>
            <div className="absolute -right-4 -top-4 w-24 h-24 rounded-full bg-white/10"/>
            <div className="w-11 h-11 rounded-xl bg-white/20 flex items-center justify-center flex-shrink-0">
              <CheckCircle2 className="text-white" style={{width: 22, height: 22}}/>
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-blue-200 text-xs font-semibold">Morning Session</p>
              <div className="mt-2 grid grid-cols-3 gap-3 text-white">
                <div>
                  <p className="text-[10px] font-semibold text-blue-100 uppercase tracking-wide">Present</p>
                  <p className="text-2xl font-black leading-none mt-0.5">{fnPresent}</p>
                </div>
                <div>
                  <p className="text-[10px] font-semibold text-blue-100 uppercase tracking-wide">Absent</p>
                  <p className="text-2xl font-black leading-none mt-0.5">{fnAbsent}</p>
                </div>
                <div>
                  <p className="text-[10px] font-semibold text-blue-100 uppercase tracking-wide">Total</p>
                  <p className="text-2xl font-black leading-none mt-0.5">{fnTotal}</p>
                </div>
              </div>
            </div>
          </div>

          {/* Evening Session */}
          <div className="rounded-2xl p-5 relative overflow-hidden flex items-center gap-4"
            style={{background: 'linear-gradient(135deg, #4338ca 0%, #7c3aed 100%)', boxShadow: '0 8px 24px rgba(124,58,237,0.28)'}}>
            <div className="absolute -right-4 -top-4 w-24 h-24 rounded-full bg-white/10"/>
            <div className="w-11 h-11 rounded-xl bg-white/20 flex items-center justify-center flex-shrink-0">
              <CheckCircle2 className="text-white" style={{width: 22, height: 22}}/>
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-purple-200 text-xs font-semibold">Evening Session</p>
              <div className="mt-2 grid grid-cols-3 gap-3 text-white">
                <div>
                  <p className="text-[10px] font-semibold text-purple-100 uppercase tracking-wide">Present</p>
                  <p className="text-2xl font-black leading-none mt-0.5">{anPresent}</p>
                </div>
                <div>
                  <p className="text-[10px] font-semibold text-purple-100 uppercase tracking-wide">Absent</p>
                  <p className="text-2xl font-black leading-none mt-0.5">{anAbsent}</p>
                </div>
                <div>
                  <p className="text-[10px] font-semibold text-purple-100 uppercase tracking-wide">Total</p>
                  <p className="text-2xl font-black leading-none mt-0.5">{anTotal}</p>
                </div>
              </div>
            </div>
          </div>

          {/* Today Present = Morning Present + Evening Present */}
          <div className="rounded-2xl p-5 relative overflow-hidden flex items-center gap-4"
            style={{background: 'linear-gradient(135deg, #065f46 0%, #10b981 100%)', boxShadow: '0 8px 24px rgba(16,185,129,0.28)'}}>
            <div className="absolute -right-4 -top-4 w-24 h-24 rounded-full bg-white/10"/>
            <div className="w-11 h-11 rounded-xl bg-white/20 flex items-center justify-center flex-shrink-0">
              <CheckCircle2 className="text-white" style={{width: 22, height: 22}}/>
            </div>
            <div>
              <p className="text-emerald-200 text-xs font-semibold">Today Present</p>
              <p className="text-3xl font-black text-white leading-none mt-0.5">{todayPresent}</p>
            </div>
          </div>

          {/* Today Absent = Morning Absent + Evening Absent */}
          <div className="rounded-2xl p-5 relative overflow-hidden flex items-center gap-4"
            style={{background: 'linear-gradient(135deg, #9a3412 0%, #ea580c 100%)', boxShadow: '0 8px 24px rgba(234,88,12,0.28)'}}>
            <div className="absolute -right-4 -top-4 w-24 h-24 rounded-full bg-white/10"/>
            <div className="w-11 h-11 rounded-xl bg-white/20 flex items-center justify-center flex-shrink-0">
              <XCircle className="text-white" style={{width: 22, height: 22}}/>
            </div>
            <div>
              <p className="text-orange-200 text-xs font-semibold">Today Absent</p>
              <p className="text-3xl font-black text-white leading-none mt-0.5">{todayAbsent}</p>
            </div>
          </div>

          {/* Avg */}
          <div className="rounded-2xl p-5 relative overflow-hidden flex items-center gap-4"
            style={{background: 'linear-gradient(135deg, #4c1d95 0%, #7c3aed 100%)', boxShadow: '0 8px 24px rgba(124,58,237,0.28)'}}>
            <div className="absolute -right-4 -top-4 w-24 h-24 rounded-full bg-white/10"/>
            <div className="w-11 h-11 rounded-xl bg-white/20 flex items-center justify-center flex-shrink-0">
              <Percent className="text-white" style={{width: 22, height: 22}}/>
            </div>
            <div>
              <p className="text-purple-200 text-xs font-semibold">Avg Attendance</p>
              <p className="text-3xl font-black text-white leading-none mt-0.5">{avgAttendance}%</p>
            </div>
          </div>
        </div>
      </div>

      {/* ═══ CONTROLS BAR ═══ */}
      <div className="print:hidden max-w-screen-2xl mx-auto px-5 mb-4">
        <div className="bg-white rounded-2xl px-5 py-3.5 flex flex-wrap gap-3 items-center justify-between"
          style={{boxShadow: '0 2px 16px rgba(30,64,175,0.08)', border: '1px solid rgba(30,64,175,0.10)'}}>

          <div className="flex items-center gap-3 flex-wrap">
            {/* Month picker */}
            <div className="flex items-center gap-2 px-3.5 py-2 rounded-xl text-sm font-bold text-blue-800"
              style={{background: '#eff6ff', border: '1.5px solid #bfdbfe'}}>
              <CalendarDays style={{width: 15, height: 15, color: '#2563eb'}}/>
              <input type="month" value={format(monthDate, 'yyyy-MM')} onChange={handleMonthChange}
                className="bg-transparent outline-none cursor-pointer text-blue-800"/>
            </div>

            {/* Search */}
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" style={{width: 14, height: 14}}/>
              <input type="text" placeholder="Search students…" value={searchQuery}
                onChange={e => setSearchQuery(e.target.value)}
                className="pl-8 pr-4 py-2 text-sm rounded-xl outline-none w-48"
                style={{background: '#f8faff', border: '1.5px solid #e0e7ff', color: '#1e3a8a'}}/>
            </div>

            <div className="w-px h-5 bg-gray-200"/>

            {/* Import */}
            <button type="button" onClick={openExcelImportPicker} disabled={isImporting}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold text-gray-700 hover:text-blue-700 border border-gray-200 hover:border-blue-300 hover:bg-blue-50 transition-all disabled:opacity-50">
              {isImporting ? <Loader2 style={{width: 13, height: 13}} className="animate-spin text-blue-600"/> : <Upload style={{width: 13, height: 13, color: '#2563eb'}}/>}
              {isImporting ? 'Importing…' : 'Import Excel'}
            </button>
            <input type="file" ref={fileInputRef} onChange={handleFileUpload} accept=".xlsx,.xls,.csv" className="hidden"/>

            {/* Export Excel */}
            <button onClick={exportToExcel}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold text-gray-700 hover:text-green-700 border border-gray-200 hover:border-green-300 hover:bg-green-50 transition-all">
              <Download style={{width: 13, height: 13, color: '#16a34a'}}/> Export Excel
            </button>

            {/* Export PDF */}
            <button onClick={openPrintModal} disabled={isExporting}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold text-gray-700 hover:text-purple-700 border border-gray-200 hover:border-purple-300 hover:bg-purple-50 transition-all disabled:opacity-50">
              {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <FileText style={{width: 13, height: 13, color: '#9333ea'}}/>}
              {isExporting ? 'Exporting…' : 'Export PDF'}
            </button>

            <div className="w-px h-5 bg-gray-200"/>

            {/* Undo */}
            <button onClick={undoHistory} disabled={attendanceHistory.length === 0}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold text-gray-700 border border-gray-200 hover:bg-orange-50 hover:border-orange-300 disabled:opacity-40 disabled:cursor-not-allowed transition-all">
              <Activity style={{width: 13, height: 13, color: '#ea580c'}}/> Undo Fill
            </button>

            {/* Print Register */}
            <button onClick={openPrintModal} disabled={isExporting}
              className="flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-black text-white disabled:opacity-50 transition-all"
              style={{background: 'linear-gradient(135deg, #4338ca 0%, #6366f1 100%)', boxShadow: '0 3px 10px rgba(99,102,241,0.35)'}}>
              {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <Printer style={{width: 13, height: 13}}/>}
              {isExporting ? 'Generating…' : 'Print Register'}
            </button>
          </div>
        </div>
      </div>

      {/* ═══ NAVIGATION TABS ═══ */}
      <div className="print:hidden max-w-screen-2xl mx-auto px-5 mb-5">
        <div className="bg-white rounded-2xl p-1.5 flex flex-wrap gap-1"
          style={{boxShadow: '0 2px 16px rgba(30,64,175,0.08)', border: '1px solid rgba(30,64,175,0.10)'}}>

          {/* Register */}
          <button onClick={() => setViewMode('register')}
            className={cn('flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black transition-all',
              viewMode === 'register' ? 'text-white shadow-lg' : 'text-gray-500 hover:text-blue-700 hover:bg-blue-50')}
            style={viewMode === 'register' ? {background: 'linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)', boxShadow: '0 4px 12px rgba(37,99,235,0.35)'} : {}}>
            <Home style={{width: 13, height: 13}}/> Attendance Register
          </button>

          {/* separator */}
          <div className="flex items-center"><span className="w-px h-5 bg-gray-200 mx-1"/></div>
          <span className="flex items-center px-2 text-[9px] font-black text-gray-400 uppercase tracking-[0.15em]">Reports</span>

          {[
            {key: 'daily',     label: 'Dailywise',     icon: <CalendarDays style={{width:13,height:13}}/>},
            {key: 'monthly',   label: 'Monthlywise',   icon: <BarChart3    style={{width:13,height:13}}/>},
            {key: 'annual',    label: 'Annual',         icon: <FileText     style={{width:13,height:13}}/>},
            {key: 'genderwise',label: 'Genderwise',    icon: <Users        style={{width:13,height:13}}/>},
            {key: 'monthwise', label: 'Monthwise',     icon: <Activity     style={{width:13,height:13}}/>},
          ].map(tab => (
            <button key={tab.key} onClick={() => setViewMode(tab.key as any)}
              className={cn('flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black transition-all',
                viewMode === tab.key ? 'text-white shadow-lg' : 'text-gray-500 hover:text-indigo-700 hover:bg-indigo-50')}
              style={viewMode === tab.key ? {background: 'linear-gradient(135deg, #4338ca 0%, #6366f1 100%)', boxShadow: '0 4px 12px rgba(99,102,241,0.35)'} : {}}>
              {tab.icon} {tab.label}
            </button>
          ))}

          {/* separator */}
          <div className="flex items-center"><span className="w-px h-5 bg-gray-200 mx-1"/></div>

          {/* About */}
          <button onClick={() => setViewMode('about')}
            className={cn('flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black transition-all',
              viewMode === 'about' ? 'text-white shadow-lg' : 'text-gray-500 hover:text-gray-700 hover:bg-gray-100')}
            style={viewMode === 'about' ? {background: 'linear-gradient(135deg, #374151 0%, #4b5563 100%)', boxShadow: '0 4px 12px rgba(75,85,99,0.35)'} : {}}>
            <Info style={{width: 13, height: 13}}/> About
          </button>
        </div>
      </div>

      {/* Main Register Container */}
      {viewMode === 'register' && (
      <div className="max-w-screen-2xl mx-auto px-5 pb-10 print:px-0 print:pb-0">
        {/* professional toolbar for register */}
        <div className="print:hidden flex items-center justify-between px-6 py-4 gap-4 flex-wrap mb-2 rounded-2xl"
          style={{background: 'linear-gradient(135deg, #0c1e4a 0%, #1a3580 55%, #1e56c8 100%)', boxShadow: '0 4px 20px rgba(12,30,74,0.30)'}}>
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center">
              <Home className="text-white" style={{width: 18, height: 18}}/>
            </div>
            <div>
              <p className="text-white font-black text-sm" style={{fontFamily: 'Georgia, serif'}}>Monthly Attendance Register</p>
              <p className="text-blue-300 text-[10px]">{format(monthDate, 'MMMM yyyy')} · {institutionName}</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={() => setShowHolidayManager(true)}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-[11px] font-bold text-blue-200 hover:text-white hover:bg-white/10 border border-white/20 transition-all">
              <CalendarDays style={{width:12,height:12}}/> Holidays
            </button>
            <button onClick={openAddStudentModal}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-[11px] font-bold text-blue-200 hover:text-white hover:bg-white/10 border border-white/20 transition-all">
              <Plus style={{width:12,height:12}}/> Add Student
            </button>
            <button onClick={openPrintModal} disabled={isExporting}
              className="flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black text-blue-950 disabled:opacity-50 transition-all"
              style={{background: 'linear-gradient(135deg, #facc15 0%, #f59e0b 100%)', boxShadow: '0 3px 10px rgba(250,204,21,0.35)'}}>
              {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <Printer style={{width:13,height:13}}/>}
              {isExporting ? 'Generating…' : 'Print Register'}
            </button>
          </div>
        </div>
        <div className="w-full overflow-x-auto print:overflow-visible rounded-2xl print:rounded-none overflow-hidden"
          style={{boxShadow: '0 4px 24px rgba(30,64,175,0.10)', border: '1px solid rgba(30,64,175,0.10)'}}>
        <div className="min-w-max bg-white border-2 border-black print:border-none print:w-full">
          
          {/* Table */}
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr className="border-b-2 border-black">
                <th colSpan={days * 2 + 11} className="p-0 border-none bg-white">
                  {/* Header 1 */}
                  <div className="flex items-stretch border-b-2 border-black">
                    <div className="flex-1 flex items-center justify-center font-bold text-lg p-2 uppercase border-r-2 border-black">
                      REGISTER OF ATTENDANCE FOR MONTH OF : {format(monthDate, 'MMMM yyyy')}
                    </div>
                    <div className="flex items-center px-4 gap-6 font-bold text-sm">
                      <div className="flex items-center gap-2">
                        <span>Standard</span>
                        <input 
                          value={standard} 
                          onChange={e => setStandard(e.target.value)} 
                          className="border-b border-black w-16 outline-none text-center uppercase" 
                        />
                      </div>
                      <div className="flex items-center gap-2">
                        <span>Section</span>
                        <input 
                          value={section} 
                          onChange={e => setSection(e.target.value)} 
                          className="border-b border-black w-16 outline-none text-center uppercase" 
                        />
                      </div>
                      <div className="flex items-center gap-2">
                        <span>Total Working Days</span>
                        <span className="border-b border-black w-16 inline-block text-center">{globalWorkingDays}</span>
                      </div>
                    </div>
                  </div>

                  {/* Header 2 */}
                  <div className="flex items-stretch">
                    <div className="w-48 flex items-center justify-center font-bold text-sm p-2 border-r-2 border-black">
                      Name of the Institution
                    </div>
                    <div className="flex-1 p-2">
                      <input 
                        value={institutionName} 
                        onChange={e => setInstitutionName(e.target.value)} 
                        className="w-full outline-none font-bold uppercase text-sm" 
                      />
                    </div>
                  </div>
                </th>
              </tr>
              <tr>
                <th rowSpan={3} className="border border-black p-1 w-10 bg-gray-50">S.NO</th>
                <th rowSpan={3} className="border border-black p-1 w-20 bg-gray-50">ADM NO</th>
                <th rowSpan={3} className="border border-black p-1 w-48 bg-gray-50">NAME</th>
                <th rowSpan={3} className="border border-black p-1 w-20 bg-gray-50">GENDER</th>
                <th colSpan={days * 2} className="border border-black p-1 bg-gray-50">DATE</th>
                <th rowSpan={3} className="border border-black p-1 w-16 bg-gray-50 text-[10px] leading-tight">No of Days<br/>Present</th>
                <th rowSpan={3} className="border border-black p-1 w-16 bg-gray-50 text-[10px] leading-tight text-blue-800">Previous<br/>Days</th>
                <th rowSpan={3} className="border border-black p-1 w-16 bg-gray-50 text-[10px] leading-tight">No of Days<br/>Absent</th>
                <th rowSpan={3} className="border border-black p-1 w-16 bg-gray-50 text-[10px] leading-tight text-indigo-900">Total<br/>Days</th>
                <th rowSpan={3} className="border border-black p-1 w-16 bg-gray-50 text-[10px] leading-tight font-bold">% Monthly</th>
                <th rowSpan={3} className="border border-black p-1 w-32 bg-gray-50 text-[10px] leading-tight">Remarks</th>
                <th rowSpan={3} className="border border-black p-1 w-16 bg-gray-50 print:hidden">Actions</th>
              </tr>
              <tr>
                {dayArray.map(d => (
                  <th key={d} colSpan={2} className="border border-black p-0 bg-gray-50 relative group">
                    <div className="p-1 cursor-default">{d}</div>
                  </th>
                ))}
              </tr>
              <tr>
                {dayArray.map(d => (
                  <React.Fragment key={d}>
                    <th 
                      onClick={(e) => handleBulkHeaderClick(d, 'fn', e)}
                      className="border border-black p-0 text-[9px] w-6 bg-gray-50 font-medium cursor-pointer hover:bg-blue-100 transition-colors"
                    >
                      FN
                    </th>
                    <th 
                      onClick={(e) => handleBulkHeaderClick(d, 'an', e)}
                      className="border border-black p-0 text-[9px] w-6 bg-gray-50 font-medium cursor-pointer hover:bg-blue-100 transition-colors"
                    >
                      AN
                    </th>
                  </React.Fragment>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredStudents.map((student, index) => {
                const totals = calculateTotals(student.id);
                const nameError = validationErrors[`${student.id}-name`];
                const admNoError = validationErrors[`${student.id}-admNo`];

                const currentMonthIdx = academicMonths.findIndex(m => format(m, 'yyyy-MM') === format(monthDate, 'yyyy-MM'));
                const priorMonths = currentMonthIdx > 0 ? academicMonths.slice(0, currentMonthIdx) : [];
                let previousDays = 0;
                priorMonths.forEach(mDate => {
                  const mKey = format(mDate, 'yyyy-MM');
                  previousDays += annualStudentAttendance[student.id]?.[mKey] ?? 0;
                });
                const totalDays = totals.present + previousDays;
                
                return (
                  <tr 
                    key={student.id} 
                    className={cn(
                      "hover:bg-blue-50/30 transition-colors",
                      lastAddedStudentId === student.id && "bg-green-50 animate-pulse"
                    )}
                  >
                    <td className="border border-black p-1 text-center font-medium">
                      {student.sno}
                    </td>
                    <td className="border border-black p-1 text-center font-medium">
                      {student.admNo}
                    </td>
                    <td className="border border-black p-1 px-2 font-medium uppercase">
                      {student.name}
                    </td>
                    <td className="border border-black p-1 text-center font-medium">
                      {student.gender || 'Boy'}
                    </td>
                    
                    {dayArray.map(d => {
                      const attKey = getAttKey(student.id);
                      const fnVal = attendance[attKey]?.[d]?.fn || '';
                      const anVal = attendance[attKey]?.[d]?.an || '';
                      
                      return (
                        <React.Fragment key={d}>
                          <td className="border border-black p-0 h-6">
                            <Cell 
                              value={fnVal} 
                              onClick={(e) => handleCellClick(student.id, d, 'fn', e)}
                            />
                          </td>
                          <td className="border border-black p-0 h-6">
                            <Cell 
                              value={anVal} 
                              onClick={(e) => handleCellClick(student.id, d, 'an', e)}
                            />
                          </td>
                        </React.Fragment>
                      );
                    })}

                    <td className={cn(
                      "border border-black p-1 text-center font-bold",
                      totals.present > 0 && totals.present % 1 !== 0 ? "bg-yellow-100" : ""
                    )}>
                      {totals.present}
                    </td>
                    <td className="border border-black p-1 text-center font-bold text-blue-800 bg-blue-50/25">
                      {previousDays}
                    </td>
                    <td className="border border-black p-1 text-center font-bold">
                      {totals.absent}
                    </td>
                    <td className="border border-black p-1 text-center font-extrabold text-indigo-900 bg-indigo-50/25">
                      {totalDays}
                    </td>
                    <td className={cn(
                      "border border-black p-1 text-center font-bold",
                      totals.percentage < 75 ? "text-red-600" : "text-green-600"
                    )}>
                      {totals.percentage.toFixed(1)}%
                    </td>
                    <td className={cn(
                      "border border-black p-1 text-center text-xs font-bold",
                      totals.percentage >= 90 ? "text-green-600" :
                      totals.percentage >= 75 ? "text-blue-600" :
                      totals.percentage >= 60 ? "text-yellow-600" : "text-red-600"
                    )}>
                      {getAttendanceStatus(totals.percentage)}
                    </td>
                    <td className="border border-black p-1 text-center print:hidden">
                      <div className="flex items-center justify-center gap-1.5">
                        <button 
                          onClick={() => setShowHistoryModal(student.id)} 
                          className="p-1 text-blue-500 hover:text-blue-700 hover:bg-blue-50 rounded transition-colors" 
                          title="View History"
                        >
                          <Activity className="w-4 h-4" />
                        </button>
                        <button 
                          onClick={() => openEditModal(student)} 
                          className="p-1 text-amber-500 hover:text-amber-700 hover:bg-amber-50 rounded transition-colors" 
                          title="Edit Student"
                        >
                          <Pencil className="w-4 h-4" />
                        </button>
                        <button 
                          onClick={() => setShowDeleteConfirm(student.id)} 
                          className="p-1 text-red-500 hover:text-red-700 hover:bg-red-50 rounded transition-colors" 
                          title="Remove from selected month"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}

              {/* Spacer Row */}
              <tr>
                <td colSpan={days * 2 + 11} className="border border-black h-4 bg-gray-100">
                  <div ref={tableEndRef} />
                </td>
              </tr>

              {/* No Present - Daily (FN) */}
              <tr>
                <td colSpan={3} rowSpan={2} className="border border-black p-1 font-bold text-right align-top">
                  No Present - Daily
                </td>
                <td className="border border-black p-1 font-bold text-right">
                  Morning (FN)
                </td>
                {dayArray.map(d => {
                  const summary = getDailySummary(d, 'fn');
                  return (
                    <React.Fragment key={d}>
                      <td className="border border-black p-0.5 text-center text-[9px] font-bold bg-blue-50/50">
                        {summary.p > 0 || summary.a > 0 ? `${summary.p}/${summary.a}` : ''}
                      </td>
                      <td className="border border-black p-0.5 bg-gray-100"></td>
                    </React.Fragment>
                  );
                })}
                    <td colSpan={7} rowSpan={2} className="border border-black p-1 text-center bg-gray-50/50">
                      <div className="flex flex-col items-center justify-center gap-1">
                        <span className="text-[10px] text-gray-500 font-bold">SUMMARY</span>
                        <div className="w-8 h-0.5 bg-blue-400 rounded-full"></div>
                      </div>
                    </td>
                  </tr>

              {/* No Present - Daily (AN) */}
              <tr>
                <td className="border border-black p-1 font-bold text-right">
                  Evening (AN)
                </td>
                {dayArray.map(d => {
                  const summary = getDailySummary(d, 'an');
                  return (
                    <React.Fragment key={d}>
                      <td className="border border-black p-0.5 bg-gray-100"></td>
                      <td className="border border-black p-0.5 text-center text-[9px] font-bold bg-blue-50/50">
                        {summary.p > 0 || summary.a > 0 ? `${summary.p}/${summary.a}` : ''}
                      </td>
                    </React.Fragment>
                  );
                })}
              </tr>
            </tbody>
          </table>
          
          {/* Rules and Summary Section */}
          <div className="p-4 grid grid-cols-2 gap-8 text-xs border-t-2 border-black">
            <div>
              <h3 className="font-bold mb-2 underline">RULES:</h3>
              <ul className="space-y-1 text-gray-700">
                <li><span className="font-bold w-8 inline-block">P</span> : Present</li>
                <li><span className="font-bold w-8 inline-block">A</span> : Absent</li>
                <li><span className="font-bold w-8 inline-block">WH</span> : Weekend Holiday</li>
                <li><span className="font-bold w-8 inline-block">GH</span> : Govt Holiday</li>
                <li><span className="font-bold w-8 inline-block">OH</span> : Other Holiday</li>
                <li className="mt-2 text-gray-500 italic">Holidays are excluded from working days calculation.</li>
              </ul>
            </div>
            
            <div className="border-l-2 border-black pl-8">
              <h3 className="font-bold mb-2 underline">DAILY SUMMARY (OVERALL):</h3>
              <div className="space-y-2 font-medium">
                <div className="flex justify-between max-w-xs">
                  <span>MORNINGS (FN):</span>
                  <span>TOTAL DAYS <span className="text-green-600">{overallSummary.totalFnP}</span> / <span className="text-red-600">{overallSummary.totalFnA}</span></span>
                </div>
                <div className="flex justify-between max-w-xs">
                  <span>EVENING (AN):</span>
                  <span>TOTAL DAYS <span className="text-green-600">{overallSummary.totalAnP}</span> / <span className="text-red-600">{overallSummary.totalAnA}</span></span>
                </div>
              </div>
            </div>
          </div>

        </div>
        </div>
      </div>
      )}

      {/* ═══ ABOUT PANEL ═══ */}
      {viewMode === 'about' && (
      <div className="max-w-screen-2xl mx-auto px-5 pb-10">
        <div className="max-w-2xl mx-auto">
          <div className="rounded-3xl overflow-hidden" style={{boxShadow: '0 24px 64px rgba(12,30,74,0.22)', background: 'linear-gradient(145deg, #0c1e4a 0%, #1a3580 55%, #1e56c8 100%)'}}>

            {/* hero */}
            <div className="relative px-10 pt-14 pb-10 text-center overflow-hidden">
              <div className="absolute -top-16 -right-16 w-64 h-64 rounded-full opacity-15" style={{background: 'radial-gradient(circle, #facc15, transparent)'}}/>
              <div className="absolute -bottom-10 -left-10 w-56 h-56 rounded-full opacity-12" style={{background: 'radial-gradient(circle, #38bdf8, transparent)'}}/>
              <div className="relative flex justify-center mb-6">
                <div className="w-22 h-22 rounded-2xl flex items-center justify-center shadow-2xl" style={{width: 88, height: 88, background: 'linear-gradient(135deg, #facc15 0%, #f59e0b 100%)'}}>
                  <GraduationCap className="text-blue-950" style={{width: 48, height: 48}}/>
                </div>
              </div>
              <h1 className="text-3xl font-black text-yellow-300 mb-1" style={{fontFamily: 'Georgia, serif', textShadow: '0 2px 8px rgba(0,0,0,0.4)'}}>
                App Developer Info
              </h1>
            </div>

            {/* developer card */}
            <div className="mx-6 mb-6 rounded-2xl p-6" style={{background: 'rgba(255,255,255,0.07)', border: '1px solid rgba(255,255,255,0.13)'}}>
              <p className="text-[10px] font-black text-blue-400 uppercase tracking-[0.15em] mb-4">Developed By</p>
              <div className="flex items-start gap-4">
                <div className="w-14 h-14 rounded-xl flex items-center justify-center flex-shrink-0" style={{background: 'rgba(250,204,21,0.14)', border: '1px solid rgba(250,204,21,0.30)'}}>
                  <BookOpen className="text-yellow-400" style={{width: 26, height: 26}}/>
                </div>
                <div>
                  <h2 className="text-xl font-black text-white" style={{fontFamily: 'Georgia, serif'}}>P. SIVA</h2>
                  <p className="text-yellow-300 text-sm font-bold mt-0.5">M.Sc., B.Ed., M.Phil., MCA.</p>
                  <p className="text-blue-200 text-sm mt-2">PG Computer Science Teacher</p>
                  <p className="text-blue-100 text-sm font-bold">Govt Hr Sec School, Pannaipuram</p>
                </div>
              </div>
              <div className="mt-5 pt-4 flex items-center gap-3 flex-wrap" style={{borderTop: '1px solid rgba(255,255,255,0.10)'}}>
                <div className="w-8 h-8 rounded-lg flex items-center justify-center" style={{background: 'rgba(250,204,21,0.14)'}}>
                  <Phone className="text-yellow-400" style={{width: 15, height: 15}}/>
                </div>
                <div className="mr-auto">
                  <p className="text-[10px] font-black text-blue-400 uppercase tracking-widest">Cell No</p>
                  <p className="text-white font-black text-sm">7603930445</p>
                </div>
                <button
                  type="button"
                  onClick={() => navigator.clipboard?.writeText('7603930445')}
                  className="flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-black text-white transition-all hover:scale-105 active:scale-95"
                  style={{background: 'linear-gradient(135deg, #1e3a8a 0%, #172554 100%)', boxShadow: '0 4px 12px rgba(30,58,138,0.35)', border: '1px solid rgba(255,255,255,0.12)'}}
                >
                  <Copy style={{width: 13, height: 13, color: '#facc15'}}/>
                  Copy
                </button>
                <a
                  href="https://wa.me/917603930445"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-black text-white transition-all hover:scale-105 active:scale-95"
                  style={{background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)', boxShadow: '0 4px 12px rgba(16,185,129,0.35)'}}
                >
                  <MessageCircle style={{width: 13, height: 13}}/>
                  WhatsApp
                </a>
              </div>
            </div>

          </div>
        </div>
      </div>
      )}

      {/* Daily Report Container */}
      {viewMode === 'daily' && (
      <div className="max-w-screen-2xl mx-auto px-5 pb-10 print:px-0 print:pb-0" id="daily-report">
        {/* Professional report wrapper */}
        <div className="rounded-2xl overflow-hidden print:rounded-none" style={{boxShadow: '0 4px 24px rgba(30,64,175,0.10)', border: '1px solid rgba(30,64,175,0.10)'}}>
          {/* toolbar */}
          <div className="print:hidden flex items-center justify-between px-6 py-4 gap-4 flex-wrap"
            style={{background: 'linear-gradient(135deg, #0c1e4a 0%, #1e40af 100%)'}}>
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center">
                <CalendarDays className="text-white" style={{width: 18, height: 18}}/>
              </div>
              <div>
                <p className="text-white font-black text-sm" style={{fontFamily: 'Georgia, serif'}}>Dailywise Attendance Report</p>
                <p className="text-blue-300 text-[10px]">{institutionName}</p>
              </div>
            </div>
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-2 px-3 py-2 rounded-xl" style={{background: 'rgba(255,255,255,0.12)', border: '1px solid rgba(255,255,255,0.18)'}}>
                <label className="text-blue-200 text-xs font-bold">Day:</label>
                <select value={reportDay} onChange={e => setReportDay(Number(e.target.value))}
                  className="bg-transparent outline-none text-white text-sm font-bold cursor-pointer">
                  {dayArray.map(d => <option key={d} value={d} className="text-gray-900">{d}</option>)}
                </select>
              </div>
              <button onClick={openPrintModal} disabled={isExporting}
                className="flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black text-blue-900 disabled:opacity-50 transition-all"
                style={{background: 'linear-gradient(135deg, #facc15 0%, #f59e0b 100%)', boxShadow: '0 3px 10px rgba(250,204,21,0.35)'}}>
                {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <Printer style={{width:13,height:13}}/>}
                {isExporting ? 'Generating…' : 'Print PDF'}
              </button>
            </div>
          </div>
          {/* content */}
          <div className="bg-white">
          <table className="w-full border-collapse text-sm">
            <thead>
              <tr className="border-b-2 border-black">
                <th colSpan={4} className="p-0 border-none bg-white">
                  <div className="flex flex-col items-center justify-center p-4">
                    <h1 className="font-bold text-xl uppercase mb-1">{institutionName}</h1>
                    <h2 className="font-bold text-lg uppercase mb-2">DAILY ATTENDANCE REPORT</h2>
                    <div className="flex w-full justify-between px-8 text-sm font-bold">
                      <span>Date: {reportDay} {format(monthDate, 'MMM yyyy')}</span>
                      <span>Standard: {standard}</span>
                      <span>Section: {section}</span>
                    </div>
                  </div>
                </th>
              </tr>
              <tr>
                <th className="border border-black p-2 w-16 bg-gray-50">S.No</th>
                <th className="border border-black p-2 text-left bg-gray-50">Student Name</th>
                <th className="border border-black p-2 w-32 bg-gray-50">FN Status(P/A)</th>
                <th className="border border-black p-2 w-32 bg-gray-50">AN Status(P/A)</th>
              </tr>
            </thead>
            <tbody>
              {filteredStudents.map(student => {
                const attKey = getAttKey(student.id);
                const fn = attendance[attKey]?.[reportDay]?.fn || '-';
                const an = attendance[attKey]?.[reportDay]?.an || '-';
                return (
                  <tr key={student.id} className="hover:bg-gray-50">
                    <td className="border border-black p-2 text-center">{student.sno}</td>
                    <td className="border border-black p-2 font-medium">{student.name}</td>
                    <td className={cn("border border-black p-2 text-center font-bold", fn === 'P' ? 'text-green-600' : fn === 'A' ? 'text-red-600' : 'text-gray-600')}>{fn}</td>
                    <td className={cn("border border-black p-2 text-center font-bold", an === 'P' ? 'text-green-600' : an === 'A' ? 'text-red-600' : 'text-gray-600')}>{an}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          </div>
        </div>
      </div>
      )}

      {/* Monthly Report Container */}
      {viewMode === 'monthly' && (
      <div className="max-w-screen-2xl mx-auto px-5 pb-10 print:px-0 print:pb-0" id="monthly-report">
        <div className="rounded-2xl overflow-hidden print:rounded-none" style={{boxShadow: '0 4px 24px rgba(30,64,175,0.10)', border: '1px solid rgba(30,64,175,0.10)'}}>
          {/* toolbar */}
          <div className="print:hidden flex items-center justify-between px-6 py-4 gap-4 flex-wrap"
            style={{background: 'linear-gradient(135deg, #065f46 0%, #059669 100%)'}}>
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center">
                <BarChart3 className="text-white" style={{width: 18, height: 18}}/>
              </div>
              <div>
                <p className="text-white font-black text-sm" style={{fontFamily: 'Georgia, serif'}}>Monthlywise Attendance Report</p>
                <p className="text-green-200 text-[10px]">{format(monthDate, 'MMMM yyyy')} · {institutionName}</p>
              </div>
            </div>
            <button onClick={openPrintModal} disabled={isExporting}
              className="flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black text-green-900 disabled:opacity-50 transition-all"
              style={{background: 'linear-gradient(135deg, #facc15 0%, #f59e0b 100%)', boxShadow: '0 3px 10px rgba(250,204,21,0.35)'}}>
              {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <Printer style={{width:13,height:13}}/>}
              {isExporting ? 'Generating…' : 'Print PDF'}
            </button>
          </div>
          {/* content */}
          <div className="bg-white">
          <table className="w-full border-collapse text-sm">
            <thead>
              <tr className="border-b-2 border-black">
                <th colSpan={7} className="p-0 border-none bg-white">
                  <div className="flex flex-col items-center justify-center p-4">
                    <h1 className="font-bold text-xl uppercase mb-1">{institutionName}</h1>
                    <h2 className="font-bold text-lg uppercase mb-2">MONTHLY ATTENDANCE REPORT : {format(monthDate, 'MMMM yyyy')}</h2>
                    <div className="flex w-full justify-between px-8 text-sm font-bold">
                      <span>Total Working Days: {globalWorkingDays}</span>
                      <span>Standard: {standard}</span>
                      <span>Section: {section}</span>
                    </div>
                  </div>
                </th>
              </tr>
              <tr>
                <th className="border border-black p-2 w-16 bg-gray-50">S.No</th>
                <th className="border border-black p-2 text-left bg-gray-50">Student Name</th>
                <th className="border border-black p-2 w-24 bg-gray-50">Days Present</th>
                <th className="border border-black p-2 w-24 bg-gray-50">Days Absent</th>
                <th className="border border-black p-2 w-24 bg-gray-50">Attendance %</th>
                <th className="border border-black p-2 w-32 bg-gray-50">Status</th>
                <th className="border border-black p-2 w-20 bg-gray-50 print:hidden">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredStudents.map(student => {
                const totals = calculateTotals(student.id);
                const status = totals.percentage >= 75 ? 'Regular' : 'Irregular';
                return (
                  <tr key={student.id} className="hover:bg-gray-50">
                    <td className="border border-black p-2 text-center">{student.sno}</td>
                    <td className="border border-black p-2 font-medium">{student.name}</td>
                    <td className="border border-black p-2 text-center font-bold text-green-600">{totals.present}</td>
                    <td className="border border-black p-2 text-center font-bold text-red-600">{totals.absent}</td>
                    <td className="border border-black p-2 text-center font-bold">{totals.percentage.toFixed(1)}%</td>
                    <td className={cn("border border-black p-2 text-center font-bold", status === 'Regular' ? 'text-green-600' : 'text-red-600')}>{status}</td>
                    <td className="border border-black p-2 text-center print:hidden">
                      <button 
                        onClick={() => setShowDeleteConfirm(student.id)} 
                        className="p-1 text-red-500 hover:text-red-700 hover:bg-red-50 rounded transition-colors" 
                        title="Remove from selected month"
                      >
                        <Trash2 className="w-5 h-5 mx-auto" />
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          </div>
        </div>
      </div>
      )}

      {/* Annual Report Container */}
      {viewMode === 'annual' && (
      <div className="max-w-screen-2xl mx-auto px-5 pb-10 print:px-0 print:pb-0 font-sans" id="annual-report">
        <div className="rounded-2xl overflow-hidden print:rounded-none" style={{boxShadow: '0 8px 32px rgba(15,23,42,0.18)', border: '1px solid rgba(99,102,241,0.18)'}}>
          {/* toolbar */}
          <div className="print:hidden flex items-center justify-between px-6 py-4 gap-4 flex-wrap"
            style={{background: 'linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #312e81 100%)'}}>
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background: 'linear-gradient(135deg, #6366f1 0%, #818cf8 100%)', boxShadow: '0 4px 12px rgba(99,102,241,0.45)'}}>
                <FileText className="text-white" style={{width: 18, height: 18}}/>
              </div>
              <div>
                <p className="text-white font-black text-sm" style={{fontFamily: 'Georgia, serif'}}>Annual Attendance Report — Studentwise</p>
                <p className="text-indigo-300 text-[10px]">{academicYearLabel} · {institutionName}</p>
              </div>
            </div>
            <button onClick={openPrintModal} disabled={isExporting}
              className="flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black text-indigo-950 disabled:opacity-50 transition-all"
              style={{background: 'linear-gradient(135deg, #a5b4fc 0%, #818cf8 100%)', boxShadow: '0 3px 10px rgba(99,102,241,0.45)'}}>
              {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <Printer style={{width:13,height:13}}/>}
              {isExporting ? 'Generating…' : 'Print PDF'}
            </button>
          </div>

          <div className="w-full overflow-x-auto print:overflow-visible">
            <div className="min-w-[1100px] print:min-w-0 print:w-full p-1 bg-white">
              {/* Header Banner */}
              <div className="text-white p-5 text-center border-b-2 border-black" style={{background: 'linear-gradient(135deg, #0f172a 0%, #1e1b4b 60%, #312e81 100%)'}}>
                <div className="flex justify-center mb-2">
                  <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background: 'linear-gradient(135deg, #6366f1, #818cf8)'}}>
                    <FileText style={{width: 20, height: 20, color: '#fff'}}/>
                  </div>
                </div>
                <h1 className="text-2xl font-extrabold tracking-wider uppercase" style={{color: '#a5b4fc', fontFamily: 'Georgia, serif'}}>
                  {institutionName}
                </h1>
                <h2 className="text-base font-bold mt-1 uppercase tracking-widest text-white/80">
                  Annual Attendance Report — Studentwise
                </h2>
                <div className="mt-2 inline-block px-4 py-1 rounded-full text-xs font-black" style={{background: 'rgba(99,102,241,0.25)', border: '1px solid rgba(165,180,252,0.35)', color: '#c7d2fe'}}>
                  {academicYearLabel}
                </div>
              </div>

              {/* Sub-Header metadata bar */}
              <div className="text-white px-4 py-3 grid grid-cols-1 md:grid-cols-4 gap-4 items-center border-b border-black text-sm font-semibold" style={{background: 'linear-gradient(90deg, #1e1b4b 0%, #312e81 100%)'}}>
                <div className="uppercase tracking-wider font-bold text-indigo-200">
                  ANNUAL ATTENDANCE REPORT - STUDENTWISE
                </div>
                <div className="md:text-center font-bold text-yellow-300">
                  {academicYearLabel}
                </div>
                <div className="md:col-span-2 flex flex-wrap gap-4 md:justify-end">
                  <div className="flex items-center gap-2">
                    <span className="text-indigo-300">Class-Div.:</span>
                    <span className="px-3 py-1 rounded text-white border" style={{background: 'rgba(99,102,241,0.25)', borderColor: 'rgba(165,180,252,0.35)'}}>
                      {standard || '—'} {section || ''}
                    </span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-indigo-300">Total Students:</span>
                    <span className="px-3 py-1 rounded text-white border" style={{background: 'rgba(99,102,241,0.25)', borderColor: 'rgba(165,180,252,0.35)'}}>
                      {students.length}
                    </span>
                  </div>
                </div>
              </div>

              {/* Attendance Table */}
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="text-white" style={{background: '#1e1b4b'}}>
                    <th className="border border-black p-2 w-[5%] text-center" rowSpan={2} style={{background: '#0f172a'}}>Roll No.</th>
                    <th className="border border-black p-2 w-[20%] text-left" rowSpan={2} style={{background: '#0f172a'}}>Student's Name</th>
                    <th className="border border-black p-1 text-center" colSpan={academicMonths.length} style={{background: '#312e81'}}>Academic Months</th>
                    <th className="border border-black p-1 w-[8%] text-center font-bold" rowSpan={2} style={{background: '#0f172a'}}>Total Days Present</th>
                    <th className="border border-black p-1 w-[8%] text-center font-bold" rowSpan={2} style={{background: '#0f172a'}}>Total Working Days</th>
                    <th className="border border-black p-2 w-[7%] text-center font-bold" rowSpan={2} style={{background: '#0f172a'}}>%</th>
                  </tr>
                  <tr className="text-white text-center" style={{background: '#0f172a'}}>
                    {academicMonths.map((mDate) => {
                      const label = format(mDate, 'MMM');
                      return (
                        <th key={label} className="border border-black p-2 w-[5%] font-semibold" style={{background: '#1e1b4b'}}>
                          {label}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {/* Total monthly working days row */}
                  <tr className="bg-amber-50 font-bold border-b-2 border-black text-center text-amber-950">
                    <td className="border border-black p-2 text-left font-bold" colSpan={2}>
                      <div className="flex items-center justify-between px-1">
                        <span className="uppercase text-[11px] tracking-wide">Total monthly working days</span>
                        <span className="text-amber-700 font-sans">➡️</span>
                      </div>
                    </td>
                    {academicMonths.map((mDate) => {
                      const monthKey = format(mDate, 'yyyy-MM');
                      const wd = annualWorkingDays[monthKey] ?? 0;
                      return (
                        <td key={monthKey} className="border border-black p-2 text-center bg-amber-100/60 text-amber-900">
                          {wd}
                        </td>
                      );
                    })}
                    <td className="border border-black p-2 text-center bg-amber-200/80 font-extrabold text-blue-900">
                      {annualSummary.totalWorkingDays}
                    </td>
                    <td className="border border-black p-2 text-center bg-amber-200/80 font-extrabold text-blue-900">
                      {annualSummary.totalWorkingDays}
                    </td>
                    <td className="border border-black p-2 text-center bg-amber-200/80 font-extrabold text-blue-900">
                      100%
                    </td>
                  </tr>

                  {/* Student Rows */}
                  {filteredStudents.map((student) => {
                    let studentTotalPresent = 0;
                    return (
                      <tr key={student.id} className="hover:bg-slate-50 transition-colors text-center border-b border-gray-400">
                        <td className="border border-gray-400 p-2 font-mono text-center bg-slate-50 font-semibold">{student.sno}</td>
                        <td className="border border-gray-400 p-2 text-left font-semibold text-gray-900 bg-white">{student.name}</td>
                        {academicMonths.map((mDate) => {
                          const monthKey = format(mDate, 'yyyy-MM');
                          const presentCount = annualStudentAttendance[student.id]?.[monthKey] ?? 0;
                          studentTotalPresent += presentCount;
                          return (
                            <td key={monthKey} className="border border-gray-400 p-2 text-center">
                              {presentCount > 0 ? presentCount.toFixed(1).replace('.0', '') : '-'}
                            </td>
                          );
                        })}
                        <td className="border border-black p-2 text-center font-bold text-green-700 bg-green-50/50">
                          {studentTotalPresent.toFixed(1).replace('.0', '')}
                        </td>
                        <td className="border border-black p-2 text-center text-gray-700 bg-slate-50 font-medium">
                          {annualSummary.totalWorkingDays}
                        </td>
                        <td className={cn(
                          "border border-black p-2 text-center font-extrabold",
                          (annualSummary.totalWorkingDays > 0 ? (studentTotalPresent / annualSummary.totalWorkingDays) * 100 : 0) >= 75 
                            ? "text-emerald-700 bg-emerald-50/20" 
                            : "text-rose-700 bg-rose-50/50"
                        )}>
                          {(annualSummary.totalWorkingDays > 0 ? (studentTotalPresent / annualSummary.totalWorkingDays) * 100 : 0).toFixed(1)}%
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>

            </div>
          </div>
        </div>
      </div>
      )}

      {/* Annual Report Genderwise Container */}
      {viewMode === 'genderwise' && (
      <div className="max-w-screen-2xl mx-auto px-5 pb-10 print:px-0 print:pb-0 font-sans" id="annual-report-genderwise">
        <div className="rounded-2xl overflow-hidden print:rounded-none" style={{boxShadow: '0 8px 32px rgba(15,23,42,0.18)', border: '1px solid rgba(20,184,166,0.20)'}}>
          {/* toolbar */}
          <div className="print:hidden flex items-center justify-between px-6 py-4 gap-4 flex-wrap"
            style={{background: 'linear-gradient(135deg, #042f2e 0%, #134e4a 50%, #0f766e 100%)'}}>
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background: 'linear-gradient(135deg, #14b8a6 0%, #2dd4bf 100%)', boxShadow: '0 4px 12px rgba(20,184,166,0.45)'}}>
                <Users className="text-white" style={{width: 18, height: 18}}/>
              </div>
              <div>
                <p className="text-white font-black text-sm" style={{fontFamily: 'Georgia, serif'}}>Annual Attendance Report — Genderwise</p>
                <p className="text-teal-300 text-[10px]">{academicYearLabel} · {institutionName}</p>
              </div>
            </div>
            <button onClick={openPrintModal} disabled={isExporting}
              className="flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black text-teal-950 disabled:opacity-50 transition-all"
              style={{background: 'linear-gradient(135deg, #5eead4 0%, #2dd4bf 100%)', boxShadow: '0 3px 10px rgba(20,184,166,0.45)'}}>
              {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <Printer style={{width:13,height:13}}/>}
              {isExporting ? 'Generating…' : 'Print PDF'}
            </button>
          </div>

          <div className="w-full overflow-x-auto print:overflow-visible">
            <div className="min-w-[1100px] print:min-w-0 print:w-full p-1 bg-white">
              {/* Header Banner */}
              <div className="text-white p-5 text-center border-b-2 border-black" style={{background: 'linear-gradient(135deg, #042f2e 0%, #134e4a 60%, #0f766e 100%)'}}>
                <div className="flex justify-center mb-2">
                  <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background: 'linear-gradient(135deg, #14b8a6, #2dd4bf)'}}>
                    <Users style={{width: 20, height: 20, color: '#fff'}}/>
                  </div>
                </div>
                <h1 className="text-2xl font-extrabold tracking-wider uppercase" style={{color: '#5eead4', fontFamily: 'Georgia, serif'}}>
                  {institutionName}
                </h1>
                <h2 className="text-base font-bold mt-1 uppercase tracking-widest text-white/80">
                  Annual Attendance Report — Genderwise
                </h2>
                <div className="mt-2 inline-block px-4 py-1 rounded-full text-xs font-black" style={{background: 'rgba(20,184,166,0.20)', border: '1px solid rgba(94,234,212,0.35)', color: '#99f6e4'}}>
                  {academicYearLabel}
                </div>
              </div>

              {/* Sub-Header metadata bar */}
              <div className="text-white px-4 py-3 grid grid-cols-1 md:grid-cols-4 gap-4 items-center border-b border-black text-sm font-semibold" style={{background: 'linear-gradient(90deg, #042f2e 0%, #134e4a 100%)'}}>
                <div className="flex items-center gap-2">
                  <span className="text-teal-300">Class:</span>
                  <span className="px-3 py-1 rounded text-white border" style={{background: 'rgba(20,184,166,0.20)', borderColor: 'rgba(94,234,212,0.35)'}}>
                    {standard || '—'} {section || ''}
                  </span>
                </div>
                <div className="md:text-center text-teal-300">
                  Boys: <span className="text-white px-2 py-0.5 rounded ml-1" style={{background: 'rgba(20,184,166,0.25)'}}>{genderwiseSummary.boysCount}</span>
                </div>
                <div className="md:text-center text-teal-300">
                  Girls: <span className="text-white px-2 py-0.5 rounded ml-1" style={{background: 'rgba(20,184,166,0.25)'}}>{genderwiseSummary.girlsCount}</span>
                </div>
                <div className="md:text-right font-bold text-yellow-300">
                  A.Y.: {academicYearLabel}
                </div>
              </div>

              {/* Attendance Table */}
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr className="text-white text-center" style={{background: '#042f2e'}}>
                    <th className="border border-black p-3 w-[15%] text-left font-bold" rowSpan={2} style={{background: '#042f2e'}}>Month</th>
                    <th className="border border-black p-1 text-center" colSpan={academicMonths.length} style={{background: '#134e4a'}}>Academic Months</th>
                    <th className="border border-black p-3 w-[18%] text-center font-bold" rowSpan={2} style={{background: '#042f2e'}}>Average Annual Presence</th>
                  </tr>
                  <tr className="text-white text-center text-xs" style={{background: '#0f172a'}}>
                    {academicMonths.map((mDate) => {
                      const label = format(mDate, 'MMM');
                      return (
                        <th key={label} className="border border-black p-2 font-semibold" style={{background: '#134e4a'}}>
                          {label}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {/* Boys Present */}
                  <tr className="hover:bg-slate-50 transition-colors border-b border-gray-400">
                    <td className="border border-gray-400 p-3 text-left font-extrabold text-blue-900 bg-slate-50">Boys Present</td>
                    {academicMonths.map((mDate) => {
                      const monthKey = format(mDate, 'yyyy-MM');
                      const val = genderwiseSummary.monthlyData[monthKey]?.boysPresent ?? 0;
                      return (
                        <td key={monthKey} className="border border-gray-400 p-3 text-center font-semibold text-gray-800">
                          {val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-'}
                        </td>
                      );
                    })}
                    <td className="border border-black p-3 text-center font-extrabold text-[#1e3a8a] bg-blue-50">
                      {genderwiseSummary.boysCount > 0 ? `${genderwiseSummary.boysAvgAnnualPresence.toFixed(1)}%` : '—'}
                    </td>
                  </tr>

                  {/* Girls Present */}
                  <tr className="hover:bg-slate-50 transition-colors border-b border-gray-400">
                    <td className="border border-gray-400 p-3 text-left font-extrabold text-blue-900 bg-slate-50">Girls Present</td>
                    {academicMonths.map((mDate) => {
                      const monthKey = format(mDate, 'yyyy-MM');
                      const val = genderwiseSummary.monthlyData[monthKey]?.girlsPresent ?? 0;
                      return (
                        <td key={monthKey} className="border border-gray-400 p-3 text-center font-semibold text-gray-800">
                          {val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-'}
                        </td>
                      );
                    })}
                    <td className="border border-black p-3 text-center font-extrabold text-[#1e3a8a] bg-blue-50">
                      {genderwiseSummary.girlsCount > 0 ? `${genderwiseSummary.girlsAvgAnnualPresence.toFixed(1)}%` : '—'}
                    </td>
                  </tr>

                  {/* Boys Absent */}
                  <tr className="hover:bg-slate-50 transition-colors border-b border-gray-400">
                    <td className="border border-gray-400 p-3 text-left font-extrabold text-[#b91c1c] bg-slate-50">Boys Absent</td>
                    {academicMonths.map((mDate) => {
                      const monthKey = format(mDate, 'yyyy-MM');
                      const val = genderwiseSummary.monthlyData[monthKey]?.boysAbsent ?? 0;
                      return (
                        <td key={monthKey} className="border border-gray-400 p-3 text-center text-red-600 font-medium font-mono">
                          {val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-'}
                        </td>
                      );
                    })}
                    <td className="border border-black p-3 text-center text-gray-400 bg-slate-100 font-bold">—</td>
                  </tr>

                  {/* Girls Absent */}
                  <tr className="hover:bg-slate-50 transition-colors border-b border-gray-400">
                    <td className="border border-gray-400 p-3 text-left font-extrabold text-[#b91c1c] bg-slate-50">Girls Absent</td>
                    {academicMonths.map((mDate) => {
                      const monthKey = format(mDate, 'yyyy-MM');
                      const val = genderwiseSummary.monthlyData[monthKey]?.girlsAbsent ?? 0;
                      return (
                        <td key={monthKey} className="border border-gray-400 p-3 text-center text-red-600 font-medium font-mono">
                          {val > 0 ? `${val.toFixed(1).replace('.0', '')}%` : '-'}
                        </td>
                      );
                    })}
                    <td className="border border-black p-3 text-center text-gray-400 bg-slate-100 font-bold">—</td>
                  </tr>
                </tbody>
              </table>

            </div>
          </div>
        </div>
      </div>
      )}

      {/* Attendance Monthwise Report Container */}
      {viewMode === 'monthwise' && (
      <div className="max-w-screen-2xl mx-auto px-5 pb-10 print:px-0 print:pb-0 font-sans" id="attendance-monthwise-report">
        <div className="rounded-2xl overflow-hidden print:rounded-none" style={{boxShadow: '0 8px 32px rgba(15,23,42,0.18)', border: '1px solid rgba(245,158,11,0.22)'}}>
          {/* toolbar */}
          <div className="print:hidden flex items-center justify-between px-6 py-4 gap-4 flex-wrap"
            style={{background: 'linear-gradient(135deg, #0c0a09 0%, #1c1008 50%, #451a03 100%)'}}>
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background: 'linear-gradient(135deg, #f59e0b 0%, #fbbf24 100%)', boxShadow: '0 4px 12px rgba(245,158,11,0.50)'}}>
                <Activity className="text-amber-950" style={{width: 18, height: 18}}/>
              </div>
              <div>
                <p className="text-white font-black text-sm" style={{fontFamily: 'Georgia, serif'}}>Attendance Monthwise Report</p>
                <p className="text-amber-300 text-[10px]">{academicYearLabel} · {institutionName}</p>
              </div>
            </div>
            <button onClick={openPrintModal} disabled={isExporting}
              className="flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-black text-amber-950 disabled:opacity-50 transition-all"
              style={{background: 'linear-gradient(135deg, #fcd34d 0%, #f59e0b 100%)', boxShadow: '0 3px 10px rgba(245,158,11,0.50)'}}>
              {isExporting ? <Loader2 style={{width:13,height:13}} className="animate-spin"/> : <Printer style={{width:13,height:13}}/>}
              {isExporting ? 'Generating…' : 'Print PDF'}
            </button>
          </div>

          <div className="w-full overflow-x-auto print:overflow-visible">
            <div className="min-w-[1100px] print:min-w-0 print:w-full p-1 bg-white">
              {/* Header Banner */}
              <div className="text-white p-5 text-center border-b-2 border-black" style={{background: 'linear-gradient(135deg, #0c0a09 0%, #1c1008 60%, #451a03 100%)'}}>
                <div className="flex justify-center mb-2">
                  <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background: 'linear-gradient(135deg, #f59e0b, #fbbf24)'}}>
                    <Activity style={{width: 20, height: 20, color: '#1c1008'}}/>
                  </div>
                </div>
                <h1 className="text-2xl font-extrabold tracking-wider uppercase" style={{color: '#fcd34d', fontFamily: 'Georgia, serif'}}>
                  {institutionName}
                </h1>
                <h2 className="text-base font-bold mt-1 uppercase tracking-widest text-white/80">
                  Attendance Monthwise Report
                </h2>
                <div className="mt-2 inline-block px-4 py-1 rounded-full text-xs font-black" style={{background: 'rgba(245,158,11,0.20)', border: '1px solid rgba(252,211,77,0.35)', color: '#fde68a'}}>
                  {academicYearLabel}
                </div>
              </div>

              {/* Sub-Header metadata bar */}
              <div className="text-white px-4 py-3 grid grid-cols-1 md:grid-cols-4 gap-4 items-center border-b border-black text-sm font-semibold" style={{background: 'linear-gradient(90deg, #1c1008 0%, #451a03 100%)'}}>
                <div className="flex items-center gap-2">
                  <span className="text-amber-300">Class:</span>
                  <span className="px-3 py-1 rounded text-white border" style={{background: 'rgba(245,158,11,0.18)', borderColor: 'rgba(252,211,77,0.35)'}}>
                    {standard || '—'} {section || ''}
                  </span>
                </div>
                <div className="md:col-span-2 md:text-center text-amber-200">
                  <span className="text-amber-300 font-bold uppercase mr-2">Total working days:</span>
                  <span className="border px-3 py-1 rounded text-white font-extrabold text-yellow-300" style={{background: 'rgba(245,158,11,0.20)', borderColor: 'rgba(252,211,77,0.35)'}}>
                    {monthwiseSummary.totalWorkingDays}
                  </span>
                </div>
                <div className="md:text-right font-bold text-yellow-300">
                  A.Y.: {academicYearLabel}
                </div>
              </div>

              {/* Attendance Table */}
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr className="text-white text-center" style={{background: '#1c1008'}}>
                    <th className="border border-black p-3 w-[15%] text-center font-bold" rowSpan={2} style={{background: '#0c0a09'}}>Months</th>
                    <th className="border border-black p-2 text-center font-bold" colSpan={3} style={{background: '#451a03'}}>Total Working Days</th>
                    <th className="border border-black p-2 text-center font-bold" colSpan={3} style={{background: '#451a03'}}>Different Averages Of Attendance</th>
                  </tr>
                  <tr className="text-white text-center text-xs" style={{background: '#0c0a09'}}>
                    <th className="border border-black p-2 font-semibold" style={{background: '#1c1008'}}>Current Month</th>
                    <th className="border border-black p-2 font-semibold" style={{background: '#1c1008'}}>Cumulative Days</th>
                    <th className="border border-black p-2 font-semibold" style={{background: '#1c1008'}}>Total Days</th>
                    <th className="border border-black p-2 font-semibold" style={{background: '#1c1008'}}>Average Daily Attendance</th>
                    <th className="border border-black p-2 font-semibold" style={{background: '#1c1008'}}>Average Monthly Attendance</th>
                    <th className="border border-black p-2 font-semibold" style={{background: '#1c1008'}}>Average Attendance Till Date</th>
                  </tr>
                </thead>
                <tbody>
                  {monthwiseSummary.monthlyDataList.map((row) => (
                    <tr key={row.monthKey} className="hover:bg-slate-50 transition-colors border-b border-gray-400">
                      <td className="border border-gray-400 p-3 text-left font-extrabold text-blue-900 bg-slate-50">
                        {row.monthLabel}
                      </td>
                      <td className="border border-gray-400 p-3 text-center font-semibold text-gray-800 bg-white">
                        {row.currentMonthWorkingDays}
                      </td>
                      <td className="border border-gray-400 p-3 text-center font-semibold text-gray-800 bg-white">
                        {row.priorWorkingDays}
                      </td>
                      <td className="border border-gray-400 p-3 text-center font-bold text-[#1e3a8a] bg-blue-50/10">
                        {row.totalSchoolWorkingDaysTillDate}
                      </td>
                      <td className="border border-gray-400 p-3 text-center font-bold text-slate-800 bg-emerald-50/20">
                        {row.averageDailyAttendance.toFixed(1).replace('.0', '')}
                      </td>
                      <td className="border border-gray-400 p-3 text-center font-bold text-emerald-700 bg-emerald-50/20">
                        {row.averageMonthlyAttendance.toFixed(1)}%
                      </td>
                      <td className="border border-black p-3 text-center font-extrabold text-blue-900 bg-blue-50/30">
                        {row.averageAttendanceTillDate.toFixed(1)}%
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

            </div>
          </div>
        </div>
      </div>
      )}

      {/* History Modal */}
      {showHistoryModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl max-w-2xl w-full p-6 max-h-[90vh] flex flex-col">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-gray-800">
                Attendance History: {students.find(s => s.id === showHistoryModal)?.name}
              </h2>
              <button onClick={() => setShowHistoryModal(null)} className="text-gray-500 hover:text-gray-700">✕</button>
            </div>
            <div className="overflow-y-auto flex-1">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr>
                    <th className="border p-2 bg-gray-50 text-left">Month</th>
                    <th className="border p-2 bg-gray-50 text-center">Working Days</th>
                    <th className="border p-2 bg-gray-50 text-center">Present</th>
                    <th className="border p-2 bg-gray-50 text-center">Absent</th>
                    <th className="border p-2 bg-gray-50 text-center">%</th>
                  </tr>
                </thead>
                <tbody>
                  {(() => {
                    // Find all months that have data for this student
                    const studentMonths = new Set<string>();
                    Object.keys(attendance).forEach(key => {
                      if (key.endsWith(`_${showHistoryModal}`)) {
                        studentMonths.add(key.split('_')[0]);
                      }
                    });
                    
                    // Always include current month
                    studentMonths.add(format(monthDate, 'yyyy-MM'));
                    
                    const sortedMonths = Array.from(studentMonths).sort((a, b) => b.localeCompare(a));
                    
                    return sortedMonths.map(monthStr => {
                      const d = new Date(monthStr + '-01T00:00:00');
                      const totals = calculateTotals(showHistoryModal, d);
                      return (
                        <tr key={monthStr} className="hover:bg-gray-50">
                          <td className="border p-2 font-medium">{format(d, 'MMMM yyyy')}</td>
                          <td className="border p-2 text-center">{totals.workingDays}</td>
                          <td className="border p-2 text-center text-green-600 font-bold">{totals.present}</td>
                          <td className="border p-2 text-center text-red-600 font-bold">{totals.absent}</td>
                          <td className="border p-2 text-center font-bold">{totals.percentage.toFixed(1)}%</td>
                        </tr>
                      );
                    });
                  })()}
                </tbody>
              </table>
            </div>
            <div className="mt-4 flex justify-end">
              <button onClick={() => setShowHistoryModal(null)} className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg">Close</button>
            </div>
          </div>
        </div>
      )}

      {/* Holiday Manager Modal */}
      {showHolidayManager && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[110] p-4 backdrop-blur-[2px]">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden border border-gray-100"
          >
            <div className="bg-gray-50 px-6 py-4 border-b border-gray-100 flex items-center gap-3">
              <div className="bg-indigo-100 p-2 rounded-lg text-indigo-600">
                <CalendarDays className="w-5 h-5" />
              </div>
              <h2 className="text-lg font-bold text-gray-800 tracking-tight">Holiday Manager</h2>
            </div>
            
            <div className="p-6 space-y-5">
              <div className="space-y-2">
                <label className="block text-sm font-semibold text-gray-700">Date</label>
                <input 
                  type="date"
                  className="w-full px-4 py-2.5 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all shadow-sm"
                  value={holidayDate}
                  onChange={(e) => setHolidayDate(e.target.value)}
                />
              </div>

              <div className="space-y-2">
                <label className="block text-sm font-semibold text-gray-700">Type</label>
                <div className="relative group">
                  <select 
                    className="w-full appearance-none px-4 py-2.5 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all shadow-sm pr-10 hover:border-indigo-300 cursor-pointer text-sm font-medium"
                    value={holidayType}
                    onChange={(e) => setHolidayType(e.target.value as AttendanceValue)}
                  >
                    <option value="GH">🏛️ Government Holiday (GH)</option>
                    <option value="WH">🏖️ Weekend Holiday (WH)</option>
                    <option value="OH">🥳 Other Holiday (OH)</option>
                    <option value="">❌ Remove Holiday</option>
                  </select>
                  <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-gray-400 group-hover:text-indigo-400 transition-colors">
                    <LogOut className="w-4 h-4 rotate-90" />
                  </div>
                </div>
              </div>
            </div>
            
            <div className="bg-gray-50 px-6 py-4 flex justify-end gap-3">
              <button 
                className="px-6 py-2 bg-green-600 hover:bg-green-700 text-white text-sm font-bold rounded-lg shadow-sm transition-all active:scale-95"
                onClick={async () => {
                  if (!user) return;
                  const date = new Date(holidayDate);
                  const day = date.getDate();
                  const monthYear = format(date, 'yyyy-MM');
                  
                  saveHistory();
                  const userId = user.uid;
                  const batch = writeBatch(db);
                  
                  students.forEach(s => {
                    const isActive = isStudentActiveInMonth(s, monthYear, attendance);
                    if (!isActive) return;

                    const attKey = `${monthYear}_${s.id}`;
                    batch.set(doc(db, `users/${userId}/attendance/${attKey}`), {
                      days: {
                        [day]: { fn: holidayType, an: holidayType }
                      }
                    }, { merge: true });
                  });

                  try {
                    await batch.commit();
                    setShowHolidayManager(false);
                  } catch (err) {
                    handleFirestoreError(err, OperationType.WRITE, `users/${userId}/attendance`);
                  }
                }}
              >
                Apply
              </button>
              <button 
                onClick={() => setShowHolidayManager(false)}
                className="px-5 py-2 text-sm font-bold text-gray-600 hover:text-gray-800 hover:bg-gray-200 rounded-lg transition-all"
              >
                Cancel
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Print Settings Modal */}
      {showPrintModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                <Printer className="w-5 h-5" />
                Print Settings (PDF)
              </h2>
              <button onClick={() => setShowPrintModal(false)} className="text-gray-500 hover:text-gray-700">✕</button>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Paper Size</label>
                <select 
                  value={printSize} 
                  onChange={(e) => setPrintSize(e.target.value as any)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                >
                  <option value="a4">A4</option>
                  <option value="legal">Legal</option>
                  <option value="letter">Letter</option>
                </select>
              </div>
              
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Orientation</label>
                <select 
                  value={printOrientation} 
                  onChange={(e) => setPrintOrientation(e.target.value as any)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                >
                  <option value="p">Portrait</option>
                  <option value="l">Landscape</option>
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Margin (mm)</label>
                <input 
                  type="number" 
                  value={printMargin} 
                  onChange={(e) => setPrintMargin(Number(e.target.value))}
                  min={0}
                  max={50}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
            </div>

            <div className="mt-6 flex justify-end gap-3">
              <button onClick={() => setShowPrintModal(false)} className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg font-medium">Cancel</button>
              <button onClick={generatePDF} className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg font-medium flex items-center gap-2">
                <FileText className="w-4 h-4" />
                Generate PDF
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Settings Modal */}
      {showSettings && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                <Settings className="w-5 h-5" />
                Admin Settings
              </h2>
              <button onClick={() => setShowSettings(false)} className="text-gray-500 hover:text-gray-700">✕</button>
            </div>
            
            <form onSubmit={handleSaveSettings} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Change Admin Username</label>
                <input
                  type="text"
                  value={newUsername}
                  onChange={(e) => setNewUsername(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all"
                  placeholder="Leave blank to keep current"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Change Admin Password</label>
                <input
                  type="password"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all"
                  placeholder="Leave blank to keep current"
                />
              </div>
              <div className="pt-4 flex justify-end gap-3">
                <button
                  type="button"
                  onClick={() => setShowSettings(false)}
                  className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors font-medium"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-colors font-medium"
                >
                  Save Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {showDeleteConfirm && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl max-w-sm w-full p-6">
            <div className="flex items-center gap-3 text-red-600 mb-4">
              <AlertTriangle className="w-6 h-6" />
              <h2 className="text-lg font-bold">Confirm Deletion</h2>
            </div>
            <div className="text-gray-600 mb-6">
              <p className="mb-4">
                Are you sure you want to remove this student from <strong>{format(monthDate, 'MMMM yyyy')}</strong>?
              </p>
              <div className="bg-red-50 p-3 rounded-lg border border-red-100 mb-4">
                <p className="font-bold text-red-700">{students.find(s => s.id === showDeleteConfirm)?.name}</p>
                <p className="text-xs text-red-600">Admin No: {students.find(s => s.id === showDeleteConfirm)?.admNo}</p>
              </div>
              <p className="text-sm">
                This will remove the student and attendance marks from the selected month only. Other months will stay unchanged.
              </p>
            </div>
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setShowDeleteConfirm(null)}
                disabled={isDeletingStudent}
                className="px-4 py-2 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors font-medium border border-gray-200"
              >
                No, Keep Record
              </button>
              <button
                onClick={() => confirmDelete(showDeleteConfirm)}
                disabled={isDeletingStudent}
                className="px-4 py-2 bg-red-600 hover:bg-red-700 disabled:opacity-60 disabled:cursor-not-allowed text-white rounded-lg transition-colors font-medium shadow-md shadow-red-100 uppercase text-xs tracking-wider font-bold flex items-center gap-2"
              >
                {isDeletingStudent && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
                {isDeletingStudent ? 'Removing...' : 'Remove from Month'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add Student Modal */}
      {showAddStudentModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[110] p-4 backdrop-blur-[2px]">
          <motion.div
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden border border-gray-100"
          >
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between"
              style={{ background: 'linear-gradient(135deg, #0c1e4a 0%, #1a3580 55%, #1e56c8 100%)' }}>
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center">
                  <Plus className="text-white" style={{ width: 18, height: 18 }} />
                </div>
                <div>
                  <h2 className="text-white font-black text-sm" style={{ fontFamily: 'Georgia, serif' }}>Add New Student</h2>
                  <p className="text-blue-300 text-[10px]">Fill in the student details below</p>
                </div>
              </div>
              <button onClick={() => setShowAddStudentModal(false)} className="text-blue-300 hover:text-white transition-colors text-xl leading-none">✕</button>
            </div>

            <form onSubmit={(e) => { e.preventDefault(); handleAddStudentSave(); }}>
              <div className="p-6 space-y-4">
                <div className="space-y-1.5">
                  <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest">Student Name <span className="text-red-500">*</span></label>
                  <input
                    type="text"
                    value={newStudentName}
                    onChange={e => setNewStudentName(e.target.value.toUpperCase())}
                    className="w-full px-4 py-2.5 rounded-xl text-sm font-medium outline-none transition-all uppercase"
                    style={{ background: '#f8faff', border: '2px solid #e0e7ff' }}
                    onFocus={e => { e.target.style.borderColor = '#2563eb'; e.target.style.background = '#fff'; }}
                    onBlur={e => { e.target.style.borderColor = '#e0e7ff'; e.target.style.background = '#f8faff'; }}
                    placeholder="STUDENT NAME"
                    autoFocus
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest">Admission Number</label>
                  <input
                    type="text"
                    value={newStudentAdmNo}
                    onChange={e => setNewStudentAdmNo(e.target.value)}
                    className="w-full px-4 py-2.5 rounded-xl text-sm font-medium outline-none transition-all"
                    style={{ background: '#f8faff', border: '2px solid #e0e7ff' }}
                    onFocus={e => { e.target.style.borderColor = '#2563eb'; e.target.style.background = '#fff'; }}
                    onBlur={e => { e.target.style.borderColor = '#e0e7ff'; e.target.style.background = '#f8faff'; }}
                    placeholder="Admission Number"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest">Gender</label>
                  <div className="flex gap-3">
                    {(['Boy', 'Girl'] as const).map(g => (
                      <button
                        key={g}
                        type="button"
                        onClick={() => setNewStudentGender(g)}
                        className={cn(
                          'flex-1 py-2.5 rounded-xl text-sm font-bold transition-all border-2',
                          newStudentGender === g
                            ? 'border-blue-600 bg-blue-600 text-white shadow-md'
                            : 'border-gray-200 bg-gray-50 text-gray-600 hover:border-blue-300'
                        )}
                      >
                        {g === 'Boy' ? '👦 Boy' : '👧 Girl'}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div className="bg-gray-50 px-6 py-4 flex justify-end gap-3 border-t border-gray-100">
                <button
                  type="button"
                  onClick={() => setShowAddStudentModal(false)}
                  className="px-5 py-2 text-sm font-bold text-gray-600 hover:text-gray-800 hover:bg-gray-200 rounded-lg transition-all"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={!newStudentName.trim() || isAddingStudent}
                  className="px-6 py-2 text-sm font-black text-white rounded-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2"
                  style={{ background: 'linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)', boxShadow: '0 4px 12px rgba(37,99,235,0.35)' }}
                >
                  {isAddingStudent ? <Loader2 style={{ width: 14, height: 14 }} className="animate-spin" /> : <Plus style={{ width: 14, height: 14 }} />}
                  {isAddingStudent ? 'Saving…' : 'Save Student'}
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}

      {/* Edit Student Modal */}
      {showEditStudentModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[110] p-4 backdrop-blur-[2px]">
          <motion.div
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden border border-gray-100"
          >
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between"
              style={{ background: 'linear-gradient(135deg, #0c1e4a 0%, #1a3580 55%, #1e56c8 100%)' }}>
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center">
                  <Pencil className="text-white" style={{ width: 18, height: 18 }} />
                </div>
                <div>
                  <h2 className="text-white font-black text-sm" style={{ fontFamily: 'Georgia, serif' }}>Edit Student</h2>
                  <p className="text-blue-300 text-[10px]">Update the student details below</p>
                </div>
              </div>
              <button onClick={() => setShowEditStudentModal(null)} className="text-blue-300 hover:text-white transition-colors text-xl leading-none">✕</button>
            </div>

            <div className="p-6 space-y-4">
              <div className="space-y-1.5">
                <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest">Student Name <span className="text-red-500">*</span></label>
                <input
                  type="text"
                  value={editStudentName}
                  onChange={e => setEditStudentName(e.target.value.toUpperCase())}
                  className="w-full px-4 py-2.5 rounded-xl text-sm font-medium outline-none transition-all uppercase"
                  style={{ background: '#f8faff', border: '2px solid #e0e7ff' }}
                  onFocus={e => { e.target.style.borderColor = '#2563eb'; e.target.style.background = '#fff'; }}
                  onBlur={e => { e.target.style.borderColor = '#e0e7ff'; e.target.style.background = '#f8faff'; }}
                  placeholder="STUDENT NAME"
                  autoFocus
                />
              </div>

              <div className="space-y-1.5">
                <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest">Admission Number</label>
                <input
                  type="text"
                  value={editStudentAdmNo}
                  onChange={e => setEditStudentAdmNo(e.target.value)}
                  className="w-full px-4 py-2.5 rounded-xl text-sm font-medium outline-none transition-all"
                  style={{ background: '#f8faff', border: '2px solid #e0e7ff' }}
                  onFocus={e => { e.target.style.borderColor = '#2563eb'; e.target.style.background = '#fff'; }}
                  onBlur={e => { e.target.style.borderColor = '#e0e7ff'; e.target.style.background = '#f8faff'; }}
                  placeholder="Admission Number"
                />
              </div>

              <div className="space-y-1.5">
                <label className="block text-[11px] font-black text-gray-400 uppercase tracking-widest">Gender</label>
                <div className="flex gap-3">
                  {(['Boy', 'Girl'] as const).map(g => (
                    <button
                      key={g}
                      type="button"
                      onClick={() => setEditStudentGender(g)}
                      className={cn(
                        'flex-1 py-2.5 rounded-xl text-sm font-bold transition-all border-2',
                        editStudentGender === g
                          ? 'border-blue-600 bg-blue-600 text-white shadow-md'
                          : 'border-gray-200 bg-gray-50 text-gray-600 hover:border-blue-300'
                      )}
                    >
                      {g === 'Boy' ? '👦 Boy' : '👧 Girl'}
                    </button>
                  ))}
                </div>
              </div>

              <div className="rounded-xl p-3 text-xs text-amber-700 font-medium" style={{ background: '#fffbeb', border: '1px solid #fde68a' }}>
                ⚠️ Changing the <strong>Student Name</strong> will automatically reposition this student alphabetically and regenerate S.No for all students.
              </div>
            </div>

            <div className="bg-gray-50 px-6 py-4 flex justify-end gap-3 border-t border-gray-100">
              <button
                onClick={() => setShowEditStudentModal(null)}
                className="px-5 py-2 text-sm font-bold text-gray-600 hover:text-gray-800 hover:bg-gray-200 rounded-lg transition-all"
              >
                Cancel
              </button>
              <button
                onClick={handleEditStudentSave}
                disabled={!editStudentName.trim()}
                className="px-6 py-2 text-sm font-black text-white rounded-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2"
                style={{ background: 'linear-gradient(135deg, #065f46 0%, #059669 100%)', boxShadow: '0 4px 12px rgba(5,150,105,0.35)' }}
              >
                <Pencil style={{ width: 14, height: 14 }} />
                Update Student
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Selector for Individual Attendance */}
      <AnimatePresence>
        {activeCell && (
          <AttendanceSelector
            x={activeCell.x}
            y={activeCell.y}
            onSelect={(val) => {
              updateAttendance(activeCell.studentId, activeCell.day, activeCell.session, val);
              setActiveCell(null);
            }}
            onClose={() => setActiveCell(null)}
          />
        )}
      </AnimatePresence>

      {/* Bulk Marking Popup */}
      <AnimatePresence>
        {activeBulkColumn && (
          <BulkMarkingPopup
            x={activeBulkColumn.x}
            y={activeBulkColumn.y}
            day={activeBulkColumn.day}
            session={activeBulkColumn.session}
            onSelect={(val) => {
              markWholeDay(activeBulkColumn.day, activeBulkColumn.session, val);
              setActiveBulkColumn(null);
            }}
            onClose={() => setActiveBulkColumn(null)}
          />
        )}
      </AnimatePresence>
    </div>
  );
}

const Cell = ({ value, onClick }: { 
  value: AttendanceValue, 
  onClick: (e: React.MouseEvent) => void
}) => {
  return (
    <div
      onClick={onClick}
      className={cn(
        "w-full h-full min-w-[20px] flex items-center justify-center text-[9px] font-bold cursor-pointer transition-colors",
        value === 'P' && "bg-green-100 text-green-800",
        value === 'A' && "bg-red-100 text-red-800",
        value === 'GH' && "bg-orange-100 text-[#8B4513]", // Reddish-brown/Burnt orange
        value === 'WH' && "bg-purple-100 text-[#8A2BE2]", // Purple/Violet
        value === 'OH' && "bg-orange-50 text-[#FF8C00]", // Bright orange
        value === 'OT' && "bg-gray-200 text-gray-600",
        value === '' && "bg-transparent hover:bg-gray-100"
      )}
    >
      {value}
    </div>
  );
};
