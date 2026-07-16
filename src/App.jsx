import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { db, auth } from "./firebase";
import {
  collection,
  addDoc,
  deleteDoc,
  doc,
  setDoc,
  onSnapshot,
  query,
  where,
  getDocs,
  Timestamp,
  serverTimestamp,
} from "firebase/firestore";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
} from "firebase/auth";

/* ============================================================
   학원 출결 앱 · Firebase(Firestore) · 요약문서(집계 최적화) 버전
   컬렉션:
     students        : { seat, name, phone, studentPhone, createdAt }
     records         : { studentId, seat, name, type, ts }    ← 원본(영구보관, 상세 엑셀용)
     dailySummaries  : id=`${영업일}_${studentId}` { date, month, studentId, seat, name, stayMin, outCount, arrive, leave }
                       ← 출결 찍힐 때 자동 갱신(추가 읽기 0). 월별 집계/대시보드는 이걸로 가볍게.
                       ← studentId 기준이라 좌석번호가 바뀌어도 누적 집계가 끊기지 않음.
   영업일(비즈니스 데이)은 자정이 아니라 새벽 1시에 넘어감 — 자정을 넘겨도 새벽 1시 전까지는
   어제 영업일 그대로 이어지고, 새벽 1시가 지나면 아직 하원 안 찍은 학생은 자동으로 하원 처리됨.
   ============================================================ */

const BRAND = "#4F46E5";
const INK = "#1A2233";
const MUTED = "#6B7688";
const BG = "#F5F7FA";
const SURFACE = "#FFFFFF";
const LINE = "#E5E9F0";

const STATE = {
  등원: { label: "등원중", color: "#16A34A", soft: "#E7F6EC" },
  외출: { label: "외출중", color: "#F59E0B", soft: "#FEF4E2" },
  외출복귀: { label: "등원중", color: "#16A34A", soft: "#E7F6EC" },
  하원: { label: "하원", color: "#64748B", soft: "#EEF1F6" },
  none: { label: "미등원", color: "#B4BCCA", soft: "#F2F4F8" },
};

const pad = (n) => String(n).padStart(2, "0");
const fmtTime = (d) => (d ? `${pad(d.getHours())}:${pad(d.getMinutes())}` : "—");
const fmtDur = (min) => {
  if (min <= 0) return "0분";
  const h = Math.floor(min / 60);
  const m = Math.round(min % 60);
  return h > 0 ? `${h}시간 ${m}분` : `${m}분`;
};
const dayKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const startOfMonth = () => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), 1); };

/* 영업일 기준 새벽 1시 — 자정이 지나도 새벽 1시 전까지는 "어제 영업일"로 취급해
   등원 상태·누적 기록이 끊기지 않게 함. 새벽 1시가 지나면 새 영업일로 넘어감. */
const BIZ_CUTOFF_HOUR = 1;
const startOfBizDay = (now = new Date()) => {
  const d = new Date(now);
  d.setHours(BIZ_CUTOFF_HOUR, 0, 0, 0);
  if (d > now) d.setDate(d.getDate() - 1);
  return d;
};
const bizKey = (d) => dayKey(startOfBizDay(d));

/* 하루치 기록 → 순 체류시간(분), 외출횟수/시간, 등·하원 시각 */
function computeDay(records) {
  const sorted = [...records].sort((a, b) => a.ts - b.ts);
  let presentSince = null, stay = 0, outCount = 0, outMin = 0;
  let arrive = null, leave = null, lastOut = null;
  for (const r of sorted) {
    if (r.type === "등원") {
      presentSince = r.ts;
      if (!arrive) arrive = r.ts;
    } else if (r.type === "외출") {
      if (presentSince) stay += (r.ts - presentSince) / 60000;
      presentSince = null; lastOut = r.ts; outCount++;
    } else if (r.type === "외출복귀") {
      presentSince = r.ts;
      if (lastOut) outMin += (r.ts - lastOut) / 60000;
      lastOut = null;
    } else if (r.type === "하원") {
      if (presentSince) stay += (r.ts - presentSince) / 60000;
      presentSince = null; leave = r.ts;
    }
  }
  if (presentSince) stay += (new Date() - presentSince) / 60000;
  return { stay, outCount, outMin, arrive, leave };
}

/* 출결 이벤트 발생 시, 그 학생의 '영업일' 요약문서를 갱신 (추가 읽기 0 — 이미 로드된 데이터로 계산)
   studentId를 문서 id로 써서, 이후 좌석번호가 바뀌어도 같은 학생의 누적 집계가 이어짐.
   dayStr은 영업일 기준(bizKey) 날짜 — 자정을 넘겨도 새벽 1시 전이면 어제 영업일로 계속 누적됨 */
async function upsertDailySummary(studentId, seat, name, recordsForDay, dayStr) {
  const c = computeDay(recordsForDay);
  await setDoc(
    doc(db, "dailySummaries", `${dayStr}_${studentId}`),
    {
      date: dayStr,
      month: dayStr.slice(0, 7),
      studentId,
      seat,
      name,
      stayMin: Math.round(c.stay),
      outCount: c.outCount,
      arrive: c.arrive ? Timestamp.fromDate(c.arrive) : null,
      leave: c.leave ? Timestamp.fromDate(c.leave) : null,
      updatedAt: serverTimestamp(),
    },
    { merge: true }
  );
}

/* 새벽 1시가 지났는데 아직 하원을 안 찍은 학생을, 어제 영업일 마감 시각(새벽 1시)으로 자동 하원 처리.
   태블릿이 꺼져 있다가 다음날 켜져도, 켜지는 시점에 밀린 영업일을 찾아 자동으로 정리함(멱등적으로 재실행 가능). */
async function autoCloseStaleSessions(students) {
  const curStart = startOfBizDay(new Date());
  const prevStart = new Date(curStart);
  prevStart.setDate(prevStart.getDate() - 1);
  const prevKey = dayKey(prevStart);
  // 마감 시각을 정각(curStart)보다 1분 앞당겨, bizKey로 다시 분류될 때도 확실히 어제 영업일에 속하게 함
  const closeTs = new Date(curStart.getTime() - 60000);

  const snap = await getDocs(
    query(
      collection(db, "records"),
      where("ts", ">=", Timestamp.fromDate(prevStart)),
      where("ts", "<", Timestamp.fromDate(curStart))
    )
  );
  const recs = snap.docs.map((d) => { const v = d.data(); return { ...v, ts: v.ts.toDate() }; });
  const byStudent = {};
  recs.forEach((r) => { if (r.studentId) (byStudent[r.studentId] ||= []).push(r); });

  for (const stu of students) {
    const rs = (byStudent[stu.id] || []).slice().sort((a, b) => a.ts - b.ts);
    if (!rs.length) continue;
    const lastType = rs[rs.length - 1].type;
    if (lastType === "등원" || lastType === "외출복귀") {
      await addDoc(collection(db, "records"), {
        studentId: stu.id, seat: stu.seat, name: stu.name, type: "하원", ts: Timestamp.fromDate(closeTs), auto: true,
      });
      await upsertDailySummary(stu.id, stu.seat, stu.name, [...rs, { type: "하원", ts: closeTs }], prevKey);
    }
  }
}

export default function App() {
  const [user, setUser] = useState(undefined); // undefined=확인중, null=미로그인, 객체=로그인됨

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return () => unsub();
  }, []);

  if (user === undefined) {
    return (
      <div style={{ minHeight: "100vh", display: "grid", placeItems: "center", background: BG,
        fontFamily: "'Pretendard', -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif" }}>
        <div style={{ color: MUTED, fontSize: 14 }}>불러오는 중…</div>
      </div>
    );
  }
  if (!user) return <Login />;
  return <AttendanceApp />;
}

function Login() {
  const [mode, setMode] = useState("login"); // "login" | "signup"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const errText = (code) => ({
    "auth/invalid-email": "이메일 형식이 올바르지 않아요",
    "auth/user-not-found": "등록되지 않은 계정이에요",
    "auth/wrong-password": "비밀번호가 틀렸어요",
    "auth/invalid-credential": "이메일 또는 비밀번호가 올바르지 않아요",
    "auth/email-already-in-use": "이미 가입된 이메일이에요",
    "auth/weak-password": "비밀번호는 6자 이상이어야 해요",
  }[code] || "처리 중 오류가 발생했어요");

  const submit = async () => {
    if (!email.trim() || !password) { setError("이메일과 비밀번호를 입력해 주세요"); return; }
    setBusy(true); setError("");
    try {
      if (mode === "login") await signInWithEmailAndPassword(auth, email.trim(), password);
      else await createUserWithEmailAndPassword(auth, email.trim(), password);
    } catch (e) {
      setError(errText(e.code));
    }
    setBusy(false);
  };

  return (
    <div style={{ minHeight: "100vh", display: "grid", placeItems: "center", background: BG, padding: 20,
      fontFamily: "'Pretendard', -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif" }}>
      <div style={{ width: "100%", maxWidth: 360, background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 20,
        boxShadow: "0 10px 30px rgba(20,34,51,.06)", padding: "32px 26px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, justifyContent: "center", marginBottom: 22 }}>
          <img src="/pharos-icon.png" alt="파로스스터디카페 로고" style={{ width: 30, height: 30, objectFit: "contain" }} />
          <div style={{ fontWeight: 800, fontSize: 17 }}>파로스스터디카페</div>
        </div>
        <div style={{ fontSize: 13, color: MUTED, textAlign: "center", marginBottom: 20 }}>
          {mode === "login" ? "직원 계정으로 로그인해 주세요" : "새 직원 계정을 만들어 주세요"}
        </div>
        <input type="email" placeholder="이메일" value={email} onChange={(e) => setEmail(e.target.value)}
          style={{ width: "100%", border: `1px solid ${LINE}`, borderRadius: 10, padding: "12px 14px", fontSize: 15, outline: "none", marginBottom: 10 }} />
        <input type="password" placeholder="비밀번호" value={password}
          onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && submit()}
          style={{ width: "100%", border: `1px solid ${LINE}`, borderRadius: 10, padding: "12px 14px", fontSize: 15, outline: "none" }} />
        {error && <div style={{ color: "#B42318", fontSize: 12.5, marginTop: 10 }}>{error}</div>}
        <button className="abtn" onClick={submit} disabled={busy}
          style={{ width: "100%", marginTop: 16, border: "none", background: BRAND, color: "#fff", borderRadius: 10,
            padding: "13px 0", fontSize: 15, fontWeight: 800, cursor: "pointer" }}>
          {busy ? "처리 중…" : mode === "login" ? "로그인" : "계정 만들기"}
        </button>
        <button onClick={() => { setMode(mode === "login" ? "signup" : "login"); setError(""); }}
          style={{ width: "100%", marginTop: 12, border: "none", background: "transparent", color: MUTED,
            fontSize: 12.5, cursor: "pointer", textDecoration: "underline" }}>
          {mode === "login" ? "처음이신가요? 계정 만들기" : "이미 계정이 있으신가요? 로그인"}
        </button>
      </div>
    </div>
  );
}

function AttendanceApp() {
  const [tab, setTab] = useState("kiosk");
  const [students, setStudents] = useState([]);
  const [todayRecords, setTodayRecords] = useState([]);
  const [entry, setEntry] = useState("");
  const [toast, setToast] = useState(null);
  const [clock, setClock] = useState(new Date());
  const [settings, setSettings] = useState({ kakaoEnabled: false });

  useEffect(() => {
    const t = setInterval(() => setClock(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  /* 원생 실시간 구독 */
  useEffect(() => {
    const unsub = onSnapshot(collection(db, "students"), (snap) => {
      setStudents(
        snap.docs.map((d) => ({ id: d.id, ...d.data() })).sort((a, b) => a.seat.localeCompare(b.seat))
      );
    });
    return () => unsub();
  }, []);

  /* 앱 설정(카톡 자동 발송 on/off 등) 실시간 구독 */
  useEffect(() => {
    const unsub = onSnapshot(doc(db, "settings", "app"), (snap) => {
      setSettings(snap.exists() ? snap.data() : { kakaoEnabled: false });
    });
    return () => unsub();
  }, []);

  const toggleKakao = async () => {
    try {
      await setDoc(doc(db, "settings", "app"), { kakaoEnabled: !settings.kakaoEnabled }, { merge: true });
    } catch {
      setToast({ kind: "error", text: "설정 변경 실패 — 네트워크를 확인해 주세요" });
    }
  };

  const businessKey = bizKey(clock);

  /* 영업일(새벽 1시 기준) 출결 실시간 구독 — 자정을 넘겨도 새벽 1시 전까지는 어제 영업일 데이터를 계속 봄 */
  useEffect(() => {
    const q = query(collection(db, "records"), where("ts", ">=", Timestamp.fromDate(startOfBizDay(new Date()))));
    const unsub = onSnapshot(q, (snap) => {
      setTodayRecords(snap.docs.map((d) => { const v = d.data(); return { id: d.id, ...v, ts: v.ts.toDate() }; }));
    });
    return () => unsub();
  }, [businessKey]);

  /* 새벽 1시가 지나면(=영업일이 바뀌면), 그리고 앱을 처음 열 때, 전날 밀린 미하원 학생을 자동 하원 처리 */
  useEffect(() => {
    if (students.length) autoCloseStaleSessions(students);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [businessKey, students.length]);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 2600);
    return () => clearTimeout(t);
  }, [toast]);

  const todayStateOf = (studentId) => {
    const rs = todayRecords.filter((r) => r.studentId === studentId && bizKey(r.ts) === businessKey).sort((a, b) => a.ts - b.ts);
    return rs.length ? rs[rs.length - 1].type : "none";
  };

  const doAction = async (type, extra = {}) => {
    const seat = entry.padStart(2, "0");
    const stu = students.find((s) => s.seat === seat);
    if (!stu) { setToast({ kind: "error", text: `좌석번호 ${entry || "—"} 는 등록되지 않았어요` }); return; }
    const cur = todayStateOf(stu.id);
    const allowed = {
      등원: cur === "none" || cur === "하원",
      외출: cur === "등원" || cur === "외출복귀",
      외출복귀: cur === "외출",
      하원: cur === "등원" || cur === "외출복귀",
    };
    if (!allowed[type]) { setToast({ kind: "error", text: `${stu.name} 님은 지금 '${STATE[cur].label}' 상태라 ${type} 처리할 수 없어요` }); return; }

    const now = new Date();
    setEntry("");
    let toastText = `${stu.name} · ${type} ${fmtTime(now)}`;
    if (type === "외출") {
      toastText += ` (${extra.reason || "기타"})`;
      toastText += extra.expectedReturn ? ` · 복귀예정 ${fmtTime(extra.expectedReturn)}` : " · 복귀 미정";
    }
    setToast({ kind: type, text: toastText });
    try {
      // 1) 원본 기록 저장 (상세 엑셀 · 영구보관) — studentId를 같이 저장해 좌석 이동에도 이력이 이어지게 함
      const record = { studentId: stu.id, seat, name: stu.name, type, ts: Timestamp.fromDate(now) };
      if (type === "외출") {
        record.reason = extra.reason || "기타";
        record.expectedReturn = extra.expectedReturn ? Timestamp.fromDate(extra.expectedReturn) : null;
      }
      await addDoc(collection(db, "records"), record);
      // 2) 영업일 요약문서 갱신 (이미 로드된 데이터 + 이번 이벤트로 계산 → 추가 읽기 없음)
      const todaysForStudent = todayRecords
        .filter((r) => r.studentId === stu.id && bizKey(r.ts) === businessKey)
        .map((r) => ({ type: r.type, ts: r.ts }));
      todaysForStudent.push({ type, ts: now });
      await upsertDailySummary(stu.id, seat, stu.name, todaysForStudent, businessKey);
      // 3) 카톡 자동 발송 (설정 켜짐 + 부모님 전화번호 있을 때만) — 서버(/api/notify)에서 솔라피 키로 발송, 실패해도 출결 기록엔 영향 없음
      if (settings.kakaoEnabled && stu.phone) {
        fetch("/api/notify", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            phone: stu.phone,
            type,
            name: stu.name,
            time: fmtTime(now),
            reason: type === "외출" ? (extra.reason || "기타") : undefined,
            expectedReturn: type === "외출" ? (extra.expectedReturn ? fmtTime(extra.expectedReturn) : "미정") : undefined,
          }),
        }).catch(() => {});
      }
    } catch (e) {
      setToast({ kind: "error", text: "저장 실패 — 네트워크를 확인해 주세요" });
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: BG, color: INK,
      fontFamily: "'Pretendard', -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif" }}>
      <style>{CSS}</style>

      <div style={{ background: SURFACE, borderBottom: `1px solid ${LINE}`, padding: "14px 20px",
        display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <img src="/pharos-icon.png" alt="파로스스터디카페 로고" style={{ width: 30, height: 30, objectFit: "contain" }} />
          <div style={{ fontWeight: 800, fontSize: 17, letterSpacing: -0.3 }}>파로스스터디카페</div>
        </div>
        <div style={{ display: "flex", background: BG, borderRadius: 10, padding: 4, flexWrap: "wrap" }}>
          {[["kiosk", "출결"], ["admin", "대시보드"], ["manage", "원생 관리"]].map(([k, l]) => (
            <button key={k} className="tab" onClick={() => setTab(k)}
              style={{ border: "none", cursor: "pointer", padding: "7px 13px", borderRadius: 8, fontSize: 13.5,
                fontWeight: 700, background: tab === k ? SURFACE : "transparent",
                color: tab === k ? INK : MUTED, boxShadow: tab === k ? "0 1px 3px rgba(20,34,51,.08)" : "none" }}>
              {l}
            </button>
          ))}
          <button onClick={() => signOut(auth)}
            style={{ border: "none", cursor: "pointer", padding: "7px 13px", borderRadius: 8, fontSize: 13.5,
              fontWeight: 700, background: "transparent", color: MUTED }}>
            로그아웃
          </button>
        </div>
      </div>

      {tab === "kiosk" && (
        <Kiosk clock={clock} entry={entry} students={students} settings={settings}
          press={(n) => entry.length < 2 && setEntry((e) => e + n)}
          back={() => setEntry((e) => e.slice(0, -1))}
          clearEntry={() => setEntry("")} doAction={doAction} toast={toast} />
      )}
      {tab === "admin" && (
        <Admin clock={clock} students={students} todayRecords={todayRecords} todayStateOf={todayStateOf}
          settings={settings} toggleKakao={toggleKakao} />
      )}
      {tab === "manage" && <Manage students={students} setToast={setToast} toast={toast} />}
    </div>
  );
}

/* =================== 키오스크 =================== */
const OUT_REASONS = ["학원", "식사", "편의점", "기타"];
const OUT_RETURN_OPTIONS = [
  { label: "30분 후", minutes: 30 },
  { label: "1시간 후", minutes: 60 },
  { label: "2시간 후", minutes: 120 },
  { label: "미정", minutes: null },
];

function Kiosk({ clock, entry, students, settings, press, back, clearEntry, doAction, toast }) {
  const [outStep, setOutStep] = useState(null); // null | "reason" | "return"
  const [outReason, setOutReason] = useState(null);
  const [outOther, setOutOther] = useState("");
  const [outExpectedReturn, setOutExpectedReturn] = useState(null);
  const [confirmType, setConfirmType] = useState(null); // 오작동 방지 — 실제 처리 전 마지막 확인 단계

  const dateStr = `${clock.getFullYear()}년 ${clock.getMonth() + 1}월 ${clock.getDate()}일 ` +
    `${["일", "월", "화", "수", "목", "금", "토"][clock.getDay()]}요일`;
  const keys = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "다시입력", "0", "back"];
  const actions = [["등원", "#16A34A"], ["외출", "#F59E0B"], ["외출복귀", "#2563EB"], ["하원", "#64748B"]];

  const stu = students.find((s) => s.seat === entry.padStart(2, "0"));

  const resetAll = () => {
    setOutStep(null); setOutReason(null); setOutOther("");
    setOutExpectedReturn(null); setConfirmType(null);
  };
  const cancelOuting = () => { setOutStep(null); setOutReason(null); setOutOther(""); setOutExpectedReturn(null); };
  const pickReason = (r) => {
    if (r === "기타") { setOutReason("기타"); return; }
    setOutReason(r);
    setOutStep("return");
  };
  const confirmOther = () => {
    const text = outOther.trim();
    setOutReason(text ? `기타(${text})` : "기타");
    setOutStep("return");
  };
  const finishOuting = (minutes) => {
    setOutExpectedReturn(minutes != null ? new Date(Date.now() + minutes * 60000) : null);
    setConfirmType("외출");
  };
  const runConfirmed = () => {
    const extra = confirmType === "외출" ? { reason: outReason, expectedReturn: outExpectedReturn } : {};
    doAction(confirmType, extra);
    resetAll();
  };

  const previewText = (() => {
    if (!confirmType || !stu) return "";
    const t = fmtTime(clock);
    if (confirmType === "등원") return `${stu.name} 학생이 ${t}에 등원하였습니다.`;
    if (confirmType === "하원") return `${stu.name} 학생이 ${t}에 하원하였습니다.`;
    if (confirmType === "외출복귀") return `${stu.name} 학생이 ${t}에 외출에서 복귀(재등원)하였습니다.`;
    if (confirmType === "외출")
      return `${stu.name} 학생이 ${t}에 외출하였습니다. (사유: ${outReason} · 복귀예정: ${outExpectedReturn ? fmtTime(outExpectedReturn) : "미정"})`;
    return "";
  })();

  const ghostBtn = { border: `1px solid ${LINE}`, background: SURFACE, color: INK, borderRadius: 14,
    padding: "14px 0", fontSize: 15, fontWeight: 700, cursor: "pointer" };

  return (
    <div style={{ maxWidth: 460, margin: "0 auto", padding: "28px 20px 48px" }}>
      <div style={{ background: SURFACE, borderRadius: 20, border: `1px solid ${LINE}`,
        boxShadow: "0 10px 30px rgba(20,34,51,.06)", overflow: "hidden" }}>
        <div style={{ background: "linear-gradient(135deg,#4F46E5,#6366F1)", color: "#fff", padding: "22px 22px 20px", textAlign: "center" }}>
          <div style={{ fontSize: 13, opacity: 0.9 }}>{dateStr}</div>
          <div style={{ fontSize: 34, fontWeight: 800, letterSpacing: 1, marginTop: 4 }}>
            {pad(clock.getHours())}:{pad(clock.getMinutes())}
          </div>
        </div>
        <div style={{ padding: "22px 22px 8px", textAlign: "center" }}>
          <div style={{ fontSize: 13, color: MUTED, fontWeight: 600 }}>
            {confirmType ? "내용을 확인해 주세요" : outStep ? "외출 처리 중" : "좌석번호를 입력해 주세요"}
          </div>
          <div style={{ fontSize: 40, fontWeight: 800, letterSpacing: 6, marginTop: 6, minHeight: 48,
            color: entry ? INK : "#CBD2DE" }}>{entry ? entry.padStart(2, "0") : "––"}</div>
        </div>

        {!outStep && !confirmType && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, padding: "8px 22px 4px" }}>
              {keys.map((k) => {
                const isUtil = k === "다시입력" || k === "back";
                return (
                  <button key={k} className="kbtn"
                    onClick={() => (k === "다시입력" ? clearEntry() : k === "back" ? back() : press(k))}
                    style={{ border: `1px solid ${LINE}`, background: isUtil ? "#F7F9FC" : SURFACE,
                      borderRadius: 14, padding: "16px 0", fontSize: isUtil ? 15 : 24, fontWeight: 700,
                      color: isUtil ? MUTED : INK, cursor: "pointer" }}>
                    {k === "back" ? "←" : k}
                  </button>
                );
              })}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, padding: "14px 22px 24px" }}>
              {actions.map(([a, c]) => (
                <button key={a} className="abtn"
                  onClick={() => (a === "외출" ? setOutStep("reason") : setConfirmType(a))}
                  style={{ border: "none", background: c, color: "#fff", borderRadius: 14, padding: "16px 0",
                    fontSize: 17, fontWeight: 800, cursor: "pointer" }}>{a}</button>
              ))}
            </div>
          </>
        )}

        {confirmType && (
          <div style={{ padding: "8px 22px 24px" }}>
            <div style={{ fontSize: 14.5, fontWeight: 800, textAlign: "center", marginBottom: 4 }}>
              {stu?.name} 학생 · {confirmType}
            </div>
            <div style={{ fontSize: 12.5, color: MUTED, textAlign: "center", marginBottom: 14 }}>
              아래 내용으로 처리할까요?
            </div>
            <div style={{ background: "#F7F9FC", border: `1px solid ${LINE}`, borderRadius: 12,
              padding: "14px 16px", fontSize: 13.5, color: INK, lineHeight: 1.6 }}>
              {previewText}
            </div>
            {settings?.kakaoEnabled && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10, fontSize: 12.5,
                color: stu?.phone ? "#16A34A" : "#B45309", fontWeight: 700 }}>
                <span style={{ width: 6, height: 6, borderRadius: "50%",
                  background: stu?.phone ? "#16A34A" : "#F59E0B", display: "inline-block" }} />
                {stu?.phone
                  ? "알림톡 기능이 켜져있습니다 · 부모님께 카톡이 발송돼요"
                  : "알림톡 기능이 켜져있지만 부모님 번호가 없어 발송되지 않아요"}
              </div>
            )}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 16 }}>
              <button className="abtn" onClick={resetAll}
                style={{ border: `1px solid ${LINE}`, background: "#F7F9FC", color: MUTED, borderRadius: 12,
                  padding: "13px 0", fontWeight: 700, cursor: "pointer" }}>
                아니오
              </button>
              <button className="abtn" onClick={runConfirmed}
                style={{ border: "none", background: BRAND, color: "#fff", borderRadius: 12,
                  padding: "13px 0", fontWeight: 800, cursor: "pointer" }}>
                예, 맞아요
              </button>
            </div>
          </div>
        )}

        {outStep === "reason" && !confirmType && (
          <div style={{ padding: "8px 22px 24px" }}>
            <div style={{ fontSize: 14.5, fontWeight: 800, textAlign: "center", marginBottom: 12 }}>외출 사유를 선택해 주세요</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              {OUT_REASONS.map((r) => (
                <button key={r} className="abtn" onClick={() => pickReason(r)}
                  style={{ ...ghostBtn, background: outReason === r || (r === "기타" && outReason?.startsWith("기타")) ? "#FEF4E2" : SURFACE,
                    borderColor: outReason === r || (r === "기타" && outReason?.startsWith("기타")) ? "#F59E0B" : LINE }}>
                  {r}
                </button>
              ))}
            </div>
            {outReason && outReason.startsWith("기타") && (
              <div style={{ marginTop: 12, display: "flex", gap: 8 }}>
                <input value={outOther} onChange={(e) => setOutOther(e.target.value)}
                  placeholder="사유 직접 입력 (선택)"
                  onKeyDown={(e) => e.key === "Enter" && confirmOther()}
                  style={{ flex: 1, border: `1px solid ${LINE}`, borderRadius: 10, padding: "10px 12px", fontSize: 14, outline: "none" }} />
                <button className="abtn" onClick={confirmOther}
                  style={{ border: "none", background: "#F59E0B", color: "#fff", borderRadius: 10, padding: "0 18px", fontWeight: 800, cursor: "pointer" }}>
                  다음
                </button>
              </div>
            )}
            <button className="abtn" onClick={cancelOuting}
              style={{ marginTop: 14, width: "100%", border: `1px solid ${LINE}`, background: "#F7F9FC", color: MUTED,
                borderRadius: 12, padding: "12px 0", fontWeight: 700, cursor: "pointer" }}>
              취소
            </button>
          </div>
        )}

        {outStep === "return" && !confirmType && (
          <div style={{ padding: "8px 22px 24px" }}>
            <div style={{ textAlign: "center", marginBottom: 12 }}>
              <div style={{ fontSize: 14.5, fontWeight: 800 }}>복귀 예정 시간을 선택해 주세요</div>
              <div style={{ fontSize: 12.5, color: MUTED, marginTop: 4 }}>사유 · {outReason}</div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              {OUT_RETURN_OPTIONS.map((opt) => (
                <button key={opt.label} className="abtn" onClick={() => finishOuting(opt.minutes)} style={ghostBtn}>
                  {opt.label}
                </button>
              ))}
            </div>
            <button className="abtn" onClick={cancelOuting}
              style={{ marginTop: 14, width: "100%", border: `1px solid ${LINE}`, background: "#F7F9FC", color: MUTED,
                borderRadius: 12, padding: "12px 0", fontWeight: 700, cursor: "pointer" }}>
              취소
            </button>
          </div>
        )}
      </div>
      <p style={{ textAlign: "center", color: MUTED, fontSize: 12.5, marginTop: 16, lineHeight: 1.6 }}>
        {students.length === 0
          ? "먼저 '원생 관리' 탭에서 학생을 등록해 주세요."
          : `등록된 원생 ${students.length}명 · 번호 입력 후 등원·하원·외출·복귀를 누르세요.`}
      </p>
      {toast && <Toast toast={toast} />}
    </div>
  );
}

/* =================== 대시보드 =================== */
function Admin({ clock, students, todayRecords, todayStateOf, settings, toggleKakao }) {
  const [summaryRows, setSummaryRows] = useState(null); // dailySummaries 집계
  const [loading, setLoading] = useState(false);
  const businessKey = bizKey(clock);
  const monthLabel = `${clock.getFullYear()}년 ${clock.getMonth() + 1}월`;

  /* 월별 집계: dailySummaries만 읽음 (원본 대비 훨씬 가벼움)
     studentId 기준으로 묶어서 좌석이 중간에 바뀌어도 누적이 이어짐.
     (studentId 없는 옛 기록은 seat로 대신 묶어 하위호환) */
  const loadMonth = async () => {
    setLoading(true);
    const mk = businessKey.slice(0, 7);
    const snap = await getDocs(query(collection(db, "dailySummaries"), where("month", "==", mk)));
    const byStudent = {};
    snap.docs.forEach((d) => {
      const v = d.data();
      const key = v.studentId || `seat:${v.seat}`;
      (byStudent[key] ||= { days: 0, totalStay: 0, totalOut: 0 });
      byStudent[key].days += 1;
      byStudent[key].totalStay += v.stayMin || 0;
      byStudent[key].totalOut += v.outCount || 0;
    });
    setSummaryRows(byStudent);
    setLoading(false);
  };
  useEffect(() => { loadMonth(); /* eslint-disable-next-line */ }, []);

  const todayRows = students.map((s) => {
    const recs = todayRecords.filter((r) => r.studentId === s.id && bizKey(r.ts) === businessKey);
    const state = todayStateOf(s.id);
    const lastOut = state === "외출"
      ? recs.filter((r) => r.type === "외출").sort((a, b) => a.ts - b.ts).slice(-1)[0]
      : null;
    return { ...s, state, outInfo: lastOut, ...computeDay(recs) };
  });
  const present = todayRows.filter((r) => r.state === "등원" || r.state === "외출복귀").length;
  const out = todayRows.filter((r) => r.state === "외출").length;
  const left = todayRows.filter((r) => r.state === "하원").length;

  const monthRows = useMemo(() => {
    if (!summaryRows) return [];
    return students.map((s) => {
      const r = summaryRows[s.id] || summaryRows[`seat:${s.seat}`] || { days: 0, totalStay: 0, totalOut: 0 };
      return { ...s, days: r.days, totalStay: r.totalStay, avg: r.days ? r.totalStay / r.days : 0, totalOut: r.totalOut };
    });
  }, [summaryRows, students]);

  /* 엑셀: 월별집계(요약) + 일별상세(원본을 이 시점에만 조회) */
  const exportExcel = async () => {
    // 월별집계 시트
    const summary = [["이름", "좌석", "출석일수", "누적 체류시간", "평균 체류", "외출 총횟수"]];
    monthRows.forEach((m) => summary.push([m.name, m.seat, m.days, fmtDur(m.totalStay), fmtDur(m.avg), m.totalOut]));

    // 일별상세 시트 — 원본은 내보내기 누를 때만 읽음 (studentId 기준 그룹핑 → 좌석 이동해도 이력 유지)
    const snap = await getDocs(query(collection(db, "records"), where("ts", ">=", Timestamp.fromDate(startOfMonth()))));
    const recs = snap.docs.map((d) => { const v = d.data(); return { ...v, ts: v.ts.toDate() }; });
    const grouped = {};
    recs.forEach((r) => { (grouped[`${bizKey(r.ts)}__${r.studentId || r.seat}`] ||= []).push(r); });
    const detail = [["날짜", "좌석", "이름", "등원", "하원", "외출횟수", "외출시간(분)", "순체류시간"]];
    Object.entries(grouped).sort().forEach(([k, rs]) => {
      const [dk] = k.split("__");
      const c = computeDay(rs);
      const last = rs[rs.length - 1];
      const stu = students.find((s) => s.id === last.studentId);
      const seat = stu?.seat || last.seat;
      const name = stu?.name || last.name;
      detail.push([dk, seat, name, fmtTime(c.arrive), fmtTime(c.leave), c.outCount, Math.round(c.outMin), fmtDur(c.stay)]);
    });

    const wb = XLSX.utils.book_new();
    const ws2 = XLSX.utils.aoa_to_sheet(summary);
    const ws1 = XLSX.utils.aoa_to_sheet(detail);
    ws2["!cols"] = [{ wch: 8 }, { wch: 6 }, { wch: 9 }, { wch: 14 }, { wch: 12 }, { wch: 11 }];
    ws1["!cols"] = [{ wch: 12 }, { wch: 6 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }, { wch: 12 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, ws2, "월별집계");
    XLSX.utils.book_append_sheet(wb, ws1, "일별상세");
    XLSX.writeFile(wb, `출결집계_${businessKey.slice(0, 7)}.xlsx`);
  };

  const Stat = ({ label, value, color }) => (
    <div style={{ background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, padding: "16px 18px", flex: 1, minWidth: 120 }}>
      <div style={{ fontSize: 12.5, color: MUTED, fontWeight: 600 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 800, marginTop: 4, color: color || INK }}>{value}</div>
    </div>
  );

  return (
    <div style={{ maxWidth: 860, margin: "0 auto", padding: "26px 20px 60px" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ fontSize: 20, fontWeight: 800 }}>오늘 출결 현황</div>
          <div style={{ fontSize: 13, color: MUTED, marginTop: 2 }}>실시간 · 태블릿에서 체크하면 여기 바로 반영돼요</div>
        </div>
        <button className="abtn" onClick={exportExcel}
          style={{ border: "none", background: "#16A34A", color: "#fff", borderRadius: 11,
            padding: "12px 18px", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
          ⬇ {monthLabel} 엑셀 다운로드
        </button>
      </div>

      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12,
        background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, padding: "14px 18px", marginTop: 16 }}>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700 }}>카톡 자동 발송</div>
          <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>
            {settings?.kakaoEnabled ? "켜짐 · 등원·외출·하원 시 부모님께 카톡이 발송돼요" : "꺼짐 · 발송 연동 준비 전에는 꺼두세요"}
          </div>
        </div>
        <button className="abtn" onClick={toggleKakao}
          aria-pressed={!!settings?.kakaoEnabled}
          style={{ border: "none", cursor: "pointer", width: 52, height: 30, borderRadius: 999, padding: 3,
            background: settings?.kakaoEnabled ? "#16A34A" : "#D8DEE7", display: "flex",
            justifyContent: settings?.kakaoEnabled ? "flex-end" : "flex-start" }}>
          <span style={{ width: 24, height: 24, borderRadius: "50%", background: "#fff", display: "block",
            boxShadow: "0 1px 3px rgba(0,0,0,.25)" }} />
        </button>
      </div>

      <div style={{ display: "flex", gap: 12, marginTop: 18, flexWrap: "wrap" }}>
        <Stat label="등원중" value={`${present}명`} color="#16A34A" />
        <Stat label="외출중" value={`${out}명`} color="#F59E0B" />
        <Stat label="하원" value={`${left}명`} color="#64748B" />
        <Stat label="전체 원생" value={`${students.length}명`} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(190px,1fr))", gap: 12, marginTop: 18 }}>
        {todayRows.map((r) => {
          const s = STATE[r.state];
          return (
            <div key={r.id} style={{ background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, padding: 15 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ fontWeight: 800, fontSize: 15.5 }}>
                  <span style={{ color: MUTED, fontWeight: 700, fontSize: 13 }}>{r.seat}</span> {r.name}
                </div>
                <span style={{ background: s.soft, color: s.color, fontSize: 11.5, fontWeight: 800, padding: "3px 9px", borderRadius: 999 }}>{s.label}</span>
              </div>
              <div style={{ fontSize: 12.5, color: MUTED, marginTop: 10, lineHeight: 1.7 }}>
                등원 {fmtTime(r.arrive)}{r.outCount > 0 ? ` · 외출 ${r.outCount}회` : ""}<br />
                오늘 체류 <b style={{ color: INK }}>{fmtDur(r.stay)}</b>
                {r.state === "외출" && r.outInfo && (
                  <><br />외출 사유 {r.outInfo.reason || "기타"} · 복귀예정 {r.outInfo.expectedReturn ? fmtTime(r.outInfo.expectedReturn.toDate ? r.outInfo.expectedReturn.toDate() : r.outInfo.expectedReturn) : "미정"}</>
                )}
              </div>
            </div>
          );
        })}
        {students.length === 0 && (
          <div style={{ gridColumn: "1/-1", color: MUTED, fontSize: 14, padding: "20px 4px" }}>
            원생 관리 탭에서 학생을 먼저 등록해 주세요.
          </div>
        )}
      </div>

      <div style={{ fontSize: 18, fontWeight: 800, marginTop: 34 }}>{monthLabel} 누적 집계</div>
      <div style={{ fontSize: 13, color: MUTED, marginTop: 2, marginBottom: 14 }}>순 체류시간은 외출 시간을 뺀 실제 학원 이용 시간이에요. (요약문서 기반 · 집계 최적화)</div>
      <div style={{ background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5 }}>
          <thead>
            <tr style={{ background: "#F7F9FC", color: MUTED }}>
              {["이름", "출석일수", "누적 체류시간", "평균 체류", "외출 횟수"].map((h, i) => (
                <th key={h} style={{ padding: "12px 16px", fontWeight: 700, textAlign: i === 0 ? "left" : "right" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {loading && <tr><td colSpan={5} style={{ padding: 20, textAlign: "center", color: MUTED }}>불러오는 중…</td></tr>}
            {!loading && monthRows.map((m) => (
              <tr key={m.id} className="row" style={{ borderTop: `1px solid ${LINE}` }}>
                <td style={{ padding: "12px 16px", fontWeight: 700 }}>
                  <span style={{ color: MUTED, fontWeight: 600, marginRight: 6 }}>{m.seat}</span>{m.name}
                </td>
                <td style={{ padding: "12px 16px", textAlign: "right" }}>{m.days}일</td>
                <td style={{ padding: "12px 16px", textAlign: "right", fontWeight: 700 }}>{fmtDur(m.totalStay)}</td>
                <td style={{ padding: "12px 16px", textAlign: "right", color: MUTED }}>{fmtDur(m.avg)}</td>
                <td style={{ padding: "12px 16px", textAlign: "right", color: MUTED }}>{m.totalOut}회</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* =================== 원생 관리 =================== */
function Manage({ students, setToast, toast }) {
  const [name, setName] = useState("");
  const [seat, setSeat] = useState("");
  const [phone, setPhone] = useState("");
  const [studentPhone, setStudentPhone] = useState("");
  const [busy, setBusy] = useState(false);
  const [moveId, setMoveId] = useState(null);
  const [moveSeat, setMoveSeat] = useState("");

  const add = async () => {
    const nm = name.trim();
    const st = seat.trim().padStart(2, "0");
    const ph = phone.trim();
    const sph = studentPhone.trim();
    if (!nm || !seat.trim()) { setToast({ kind: "error", text: "이름과 좌석번호를 모두 입력해 주세요" }); return; }
    if (students.some((s) => s.seat === st)) { setToast({ kind: "error", text: `좌석번호 ${st} 는 이미 사용 중이에요` }); return; }
    setBusy(true);
    try {
      await addDoc(collection(db, "students"), { seat: st, name: nm, phone: ph, studentPhone: sph, createdAt: serverTimestamp() });
      setName(""); setSeat(""); setPhone(""); setStudentPhone("");
      setToast({ kind: "등원", text: `${nm} 등록 완료` });
    } catch { setToast({ kind: "error", text: "등록 실패 — 네트워크를 확인해 주세요" }); }
    setBusy(false);
  };
  const remove = async (s) => {
    if (!window.confirm(`${s.name}(좌석 ${s.seat}) 원생을 삭제할까요?`)) return;
    try { await deleteDoc(doc(db, "students", s.id)); setToast({ kind: "하원", text: `${s.name} 삭제됨` }); }
    catch { setToast({ kind: "error", text: "삭제 실패 — 네트워크를 확인해 주세요" }); }
  };
  const startMove = (s) => { setMoveId(s.id); setMoveSeat(s.seat); };
  const cancelMove = () => { setMoveId(null); setMoveSeat(""); };
  const confirmMove = async (s) => {
    const ns = moveSeat.trim().padStart(2, "0");
    if (!moveSeat.trim()) { setToast({ kind: "error", text: "새 좌석번호를 입력해 주세요" }); return; }
    if (ns === s.seat) { cancelMove(); return; }
    if (students.some((x) => x.id !== s.id && x.seat === ns)) { setToast({ kind: "error", text: `좌석번호 ${ns} 는 이미 사용 중이에요` }); return; }
    try {
      await setDoc(doc(db, "students", s.id), { seat: ns }, { merge: true });
      setToast({ kind: "등원", text: `${s.name} 좌석 ${s.seat} → ${ns} 이동 완료 (누적 기록 유지됨)` });
      cancelMove();
    } catch { setToast({ kind: "error", text: "좌석 이동 실패 — 네트워크를 확인해 주세요" }); }
  };

  const inputStyle = { flex: 1, minWidth: 0, border: `1px solid ${LINE}`, borderRadius: 10, padding: "12px 14px", fontSize: 15, outline: "none" };

  return (
    <div style={{ maxWidth: 560, margin: "0 auto", padding: "26px 20px 60px" }}>
      <div style={{ fontSize: 20, fontWeight: 800 }}>원생 관리</div>
      <div style={{ fontSize: 13, color: MUTED, marginTop: 2, marginBottom: 18 }}>좌석번호는 출결 키오스크에서 학생이 입력하는 번호예요. 좌석은 나중에 이동해도 출결 기록이 그대로 유지돼요.</div>

      <div style={{ background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, padding: 16, display: "flex", gap: 8, flexWrap: "wrap" }}>
        <input style={{ ...inputStyle, maxWidth: 100 }} placeholder="좌석번호" inputMode="numeric"
          value={seat} onChange={(e) => setSeat(e.target.value.replace(/\D/g, "").slice(0, 2))} />
        <input style={inputStyle} placeholder="학생 이름" value={name}
          onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
        <input style={{ ...inputStyle, maxWidth: 160 }} placeholder="학생 본인 전화번호" inputMode="tel"
          value={studentPhone} onChange={(e) => setStudentPhone(e.target.value.replace(/[^0-9-]/g, ""))} onKeyDown={(e) => e.key === "Enter" && add()} />
        <input style={{ ...inputStyle, maxWidth: 160 }} placeholder="부모님 전화번호" inputMode="tel"
          value={phone} onChange={(e) => setPhone(e.target.value.replace(/[^0-9-]/g, ""))} onKeyDown={(e) => e.key === "Enter" && add()} />
        <button className="abtn" onClick={add} disabled={busy}
          style={{ border: "none", background: BRAND, color: "#fff", borderRadius: 10, padding: "0 20px", fontSize: 15, fontWeight: 800, cursor: "pointer" }}>
          등록
        </button>
      </div>

      <div style={{ marginTop: 16, background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, overflow: "hidden" }}>
        {students.length === 0 && <div style={{ padding: 24, textAlign: "center", color: MUTED, fontSize: 14 }}>아직 등록된 원생이 없어요. 위에서 추가해 주세요.</div>}
        {students.map((s) => (
          <div key={s.id} className="row" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "13px 16px", borderTop: `1px solid ${LINE}`, gap: 10, flexWrap: "wrap" }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 15 }}>
                <span style={{ display: "inline-block", minWidth: 34, color: BRAND, fontWeight: 800 }}>{s.seat}</span> {s.name}
              </div>
              {(s.studentPhone || s.phone) && (
                <div style={{ fontSize: 12.5, color: MUTED, marginTop: 3, marginLeft: 34, lineHeight: 1.6 }}>
                  {s.studentPhone && <>학생 {s.studentPhone}</>}
                  {s.studentPhone && s.phone && <> · </>}
                  {s.phone && <>부모님 {s.phone}</>}
                </div>
              )}
            </div>
            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              {moveId === s.id ? (
                <>
                  <input style={{ width: 64, border: `1px solid ${LINE}`, borderRadius: 8, padding: "6px 8px", fontSize: 13 }}
                    inputMode="numeric" autoFocus value={moveSeat}
                    onChange={(e) => setMoveSeat(e.target.value.replace(/\D/g, "").slice(0, 2))}
                    onKeyDown={(e) => e.key === "Enter" && confirmMove(s)} />
                  <button onClick={() => confirmMove(s)} style={{ border: "none", background: BRAND, color: "#fff", borderRadius: 8, padding: "6px 10px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>확인</button>
                  <button onClick={cancelMove} style={{ border: `1px solid ${LINE}`, background: SURFACE, color: MUTED, borderRadius: 8, padding: "6px 10px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>취소</button>
                </>
              ) : (
                <>
                  <button onClick={() => startMove(s)} style={{ border: `1px solid ${LINE}`, background: SURFACE, color: BRAND, borderRadius: 8, padding: "6px 12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>좌석 이동</button>
                  <button onClick={() => remove(s)} style={{ border: `1px solid ${LINE}`, background: SURFACE, color: "#B42318", borderRadius: 8, padding: "6px 12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>삭제</button>
                </>
              )}
            </div>
          </div>
        ))}
      </div>
      {toast && <Toast toast={toast} />}
    </div>
  );
}

/* =================== 공통 =================== */
function Toast({ toast }) {
  const isErr = toast.kind === "error";
  const c = STATE[toast.kind];
  return (
    <div style={{ position: "fixed", left: "50%", bottom: 30, transform: "translateX(-50%)",
      background: isErr ? "#FBEAEA" : (c?.soft || "#EEF1F6"),
      color: isErr ? "#B42318" : (c?.color || INK),
      border: `1px solid ${isErr ? "#F3C9C4" : "transparent"}`,
      padding: "13px 20px", borderRadius: 12, fontWeight: 700, fontSize: 14.5,
      boxShadow: "0 8px 24px rgba(20,34,51,.14)", animation: "pop .18s ease", zIndex: 50 }}>
      {toast.text}
    </div>
  );
}

const CSS = `
  * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  .kbtn { transition: transform .05s ease, background .12s ease; user-select: none; }
  .kbtn:active { transform: scale(.96); background: #EEF1F8 !important; }
  .abtn { transition: transform .05s ease, filter .12s ease; }
  .abtn:active { transform: scale(.97); filter: brightness(.93); }
  .row:hover { background: #FAFBFD; }
  @keyframes pop { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }
`;
