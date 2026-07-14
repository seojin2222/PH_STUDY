import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
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

/* ============================================================
   학원 출결 앱 · Firebase(Firestore) · 요약문서(집계 최적화) 버전
   컬렉션:
     students        : { seat, name, createdAt }
     records         : { seat, name, type, ts }              ← 원본(영구보관, 상세 엑셀용)
     dailySummaries  : id=`${날짜}_${좌석}` { date, month, seat, name, stayMin, outCount, arrive, leave }
                       ← 출결 찍힐 때 자동 갱신(추가 읽기 0). 월별 집계/대시보드는 이걸로 가볍게.
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
const monthKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
const startOfToday = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };
const startOfMonth = () => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), 1); };

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

/* 출결 이벤트 발생 시, 그 학생의 '오늘' 요약문서를 갱신 (추가 읽기 0 — 이미 로드된 today 데이터로 계산) */
async function upsertDailySummary(seat, name, todaysForStudent) {
  const now = new Date();
  const c = computeDay(todaysForStudent);
  await setDoc(
    doc(db, "dailySummaries", `${dayKey(now)}_${seat}`),
    {
      date: dayKey(now),
      month: monthKey(now),
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

export default function App() {
  const [tab, setTab] = useState("kiosk");
  const [students, setStudents] = useState([]);
  const [todayRecords, setTodayRecords] = useState([]);
  const [entry, setEntry] = useState("");
  const [toast, setToast] = useState(null);
  const [clock, setClock] = useState(new Date());

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

  /* 오늘 출결 실시간 구독 (오늘 것만 — 읽기 최소화) */
  useEffect(() => {
    const q = query(collection(db, "records"), where("ts", ">=", Timestamp.fromDate(startOfToday())));
    const unsub = onSnapshot(q, (snap) => {
      setTodayRecords(snap.docs.map((d) => { const v = d.data(); return { id: d.id, ...v, ts: v.ts.toDate() }; }));
    });
    return () => unsub();
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 2600);
    return () => clearTimeout(t);
  }, [toast]);

  const todayKey = dayKey(new Date());
  const todayStateOf = (seat) => {
    const rs = todayRecords.filter((r) => r.seat === seat && dayKey(r.ts) === todayKey).sort((a, b) => a.ts - b.ts);
    return rs.length ? rs[rs.length - 1].type : "none";
  };

  const doAction = async (type) => {
    const seat = entry.padStart(2, "0");
    const stu = students.find((s) => s.seat === seat);
    if (!stu) { setToast({ kind: "error", text: `좌석번호 ${entry || "—"} 는 등록되지 않았어요` }); return; }
    const cur = todayStateOf(seat);
    const allowed = {
      등원: cur === "none" || cur === "하원",
      외출: cur === "등원" || cur === "외출복귀",
      외출복귀: cur === "외출",
      하원: cur === "등원" || cur === "외출복귀",
    };
    if (!allowed[type]) { setToast({ kind: "error", text: `${stu.name} 님은 지금 '${STATE[cur].label}' 상태라 ${type} 처리할 수 없어요` }); return; }

    const now = new Date();
    setEntry("");
    setToast({ kind: type, text: `${stu.name} · ${type} ${fmtTime(now)}` });
    try {
      // 1) 원본 기록 저장 (상세 엑셀 · 영구보관)
      await addDoc(collection(db, "records"), { seat, name: stu.name, type, ts: Timestamp.fromDate(now) });
      // 2) 오늘 요약문서 갱신 (이미 로드된 today 데이터 + 이번 이벤트로 계산 → 추가 읽기 없음)
      const todaysForStudent = todayRecords
        .filter((r) => r.seat === seat && dayKey(r.ts) === todayKey)
        .map((r) => ({ type: r.type, ts: r.ts }));
      todaysForStudent.push({ type, ts: now });
      await upsertDailySummary(seat, stu.name, todaysForStudent);
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
          <div style={{ width: 30, height: 30, borderRadius: 8, background: BRAND, color: "#fff",
            display: "grid", placeItems: "center", fontWeight: 800, fontSize: 15 }}>P</div>
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
        </div>
      </div>

      {tab === "kiosk" && (
        <Kiosk clock={clock} entry={entry} students={students}
          press={(n) => entry.length < 2 && setEntry((e) => e + n)}
          back={() => setEntry((e) => e.slice(0, -1))}
          clearEntry={() => setEntry("")} doAction={doAction} toast={toast} />
      )}
      {tab === "admin" && (
        <Admin clock={clock} students={students} todayRecords={todayRecords} todayStateOf={todayStateOf} />
      )}
      {tab === "manage" && <Manage students={students} setToast={setToast} toast={toast} />}
    </div>
  );
}

/* =================== 키오스크 =================== */
function Kiosk({ clock, entry, students, press, back, clearEntry, doAction, toast }) {
  const dateStr = `${clock.getFullYear()}년 ${clock.getMonth() + 1}월 ${clock.getDate()}일 ` +
    `${["일", "월", "화", "수", "목", "금", "토"][clock.getDay()]}요일`;
  const keys = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "다시입력", "0", "back"];
  const actions = [["등원", "#16A34A"], ["외출", "#F59E0B"], ["외출복귀", "#2563EB"], ["하원", "#64748B"]];
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
          <div style={{ fontSize: 13, color: MUTED, fontWeight: 600 }}>좌석번호를 입력해 주세요</div>
          <div style={{ fontSize: 40, fontWeight: 800, letterSpacing: 6, marginTop: 6, minHeight: 48,
            color: entry ? INK : "#CBD2DE" }}>{entry ? entry.padStart(2, "0") : "––"}</div>
        </div>
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
            <button key={a} className="abtn" onClick={() => doAction(a)}
              style={{ border: "none", background: c, color: "#fff", borderRadius: 14, padding: "16px 0",
                fontSize: 17, fontWeight: 800, cursor: "pointer" }}>{a}</button>
          ))}
        </div>
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
function Admin({ clock, students, todayRecords, todayStateOf }) {
  const [summaryRows, setSummaryRows] = useState(null); // dailySummaries 집계
  const [loading, setLoading] = useState(false);
  const todayKey = dayKey(new Date());
  const monthLabel = `${clock.getFullYear()}년 ${clock.getMonth() + 1}월`;

  /* 월별 집계: dailySummaries만 읽음 (원본 대비 훨씬 가벼움) */
  const loadMonth = async () => {
    setLoading(true);
    const mk = monthKey(new Date());
    const snap = await getDocs(query(collection(db, "dailySummaries"), where("month", "==", mk)));
    const bySeat = {};
    snap.docs.forEach((d) => {
      const v = d.data();
      (bySeat[v.seat] ||= { seat: v.seat, name: v.name, days: 0, totalStay: 0, totalOut: 0 });
      bySeat[v.seat].days += 1;
      bySeat[v.seat].totalStay += v.stayMin || 0;
      bySeat[v.seat].totalOut += v.outCount || 0;
    });
    setSummaryRows(bySeat);
    setLoading(false);
  };
  useEffect(() => { loadMonth(); /* eslint-disable-next-line */ }, []);

  const todayRows = students.map((s) => {
    const recs = todayRecords.filter((r) => r.seat === s.seat && dayKey(r.ts) === todayKey);
    return { ...s, state: todayStateOf(s.seat), ...computeDay(recs) };
  });
  const present = todayRows.filter((r) => r.state === "등원" || r.state === "외출복귀").length;
  const out = todayRows.filter((r) => r.state === "외출").length;
  const left = todayRows.filter((r) => r.state === "하원").length;

  const monthRows = useMemo(() => {
    if (!summaryRows) return [];
    return students.map((s) => {
      const r = summaryRows[s.seat] || { days: 0, totalStay: 0, totalOut: 0 };
      return { ...s, days: r.days, totalStay: r.totalStay, avg: r.days ? r.totalStay / r.days : 0, totalOut: r.totalOut };
    });
  }, [summaryRows, students]);

  /* 엑셀: 월별집계(요약) + 일별상세(원본을 이 시점에만 조회) */
  const exportExcel = async () => {
    // 월별집계 시트
    const summary = [["이름", "좌석", "출석일수", "누적 체류시간", "평균 체류", "외출 총횟수"]];
    monthRows.forEach((m) => summary.push([m.name, m.seat, m.days, fmtDur(m.totalStay), fmtDur(m.avg), m.totalOut]));

    // 일별상세 시트 — 원본은 내보내기 누를 때만 읽음
    const snap = await getDocs(query(collection(db, "records"), where("ts", ">=", Timestamp.fromDate(startOfMonth()))));
    const recs = snap.docs.map((d) => { const v = d.data(); return { ...v, ts: v.ts.toDate() }; });
    const grouped = {};
    recs.forEach((r) => { (grouped[`${dayKey(r.ts)}__${r.seat}`] ||= []).push(r); });
    const detail = [["날짜", "좌석", "이름", "등원", "하원", "외출횟수", "외출시간(분)", "순체류시간"]];
    Object.entries(grouped).sort().forEach(([k, rs]) => {
      const [dk, seat] = k.split("__");
      const c = computeDay(rs);
      const name = students.find((s) => s.seat === seat)?.name || rs[0]?.name || "";
      detail.push([dk, seat, name, fmtTime(c.arrive), fmtTime(c.leave), c.outCount, Math.round(c.outMin), fmtDur(c.stay)]);
    });

    const wb = XLSX.utils.book_new();
    const ws2 = XLSX.utils.aoa_to_sheet(summary);
    const ws1 = XLSX.utils.aoa_to_sheet(detail);
    ws2["!cols"] = [{ wch: 8 }, { wch: 6 }, { wch: 9 }, { wch: 14 }, { wch: 12 }, { wch: 11 }];
    ws1["!cols"] = [{ wch: 12 }, { wch: 6 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }, { wch: 12 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, ws2, "월별집계");
    XLSX.utils.book_append_sheet(wb, ws1, "일별상세");
    XLSX.writeFile(wb, `출결집계_${monthKey(new Date())}.xlsx`);
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
            <div key={r.seat} style={{ background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, padding: 15 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ fontWeight: 800, fontSize: 15.5 }}>
                  <span style={{ color: MUTED, fontWeight: 700, fontSize: 13 }}>{r.seat}</span> {r.name}
                </div>
                <span style={{ background: s.soft, color: s.color, fontSize: 11.5, fontWeight: 800, padding: "3px 9px", borderRadius: 999 }}>{s.label}</span>
              </div>
              <div style={{ fontSize: 12.5, color: MUTED, marginTop: 10, lineHeight: 1.7 }}>
                등원 {fmtTime(r.arrive)}{r.outCount > 0 ? ` · 외출 ${r.outCount}회` : ""}<br />
                오늘 체류 <b style={{ color: INK }}>{fmtDur(r.stay)}</b>
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
              <tr key={m.seat} className="row" style={{ borderTop: `1px solid ${LINE}` }}>
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
  const [busy, setBusy] = useState(false);

  const add = async () => {
    const nm = name.trim();
    const st = seat.trim().padStart(2, "0");
    const ph = phone.trim();
    if (!nm || !seat.trim()) { setToast({ kind: "error", text: "이름과 좌석번호를 모두 입력해 주세요" }); return; }
    if (students.some((s) => s.seat === st)) { setToast({ kind: "error", text: `좌석번호 ${st} 는 이미 사용 중이에요` }); return; }
    setBusy(true);
    try {
      await addDoc(collection(db, "students"), { seat: st, name: nm, phone: ph, createdAt: serverTimestamp() });
      setName(""); setSeat(""); setPhone("");
      setToast({ kind: "등원", text: `${nm} 등록 완료` });
    } catch { setToast({ kind: "error", text: "등록 실패 — 네트워크를 확인해 주세요" }); }
    setBusy(false);
  };
  const remove = async (s) => {
    if (!window.confirm(`${s.name}(좌석 ${s.seat}) 원생을 삭제할까요?`)) return;
    try { await deleteDoc(doc(db, "students", s.id)); setToast({ kind: "하원", text: `${s.name} 삭제됨` }); }
    catch { setToast({ kind: "error", text: "삭제 실패 — 네트워크를 확인해 주세요" }); }
  };

  const inputStyle = { flex: 1, minWidth: 0, border: `1px solid ${LINE}`, borderRadius: 10, padding: "12px 14px", fontSize: 15, outline: "none" };

  return (
    <div style={{ maxWidth: 560, margin: "0 auto", padding: "26px 20px 60px" }}>
      <div style={{ fontSize: 20, fontWeight: 800 }}>원생 관리</div>
      <div style={{ fontSize: 13, color: MUTED, marginTop: 2, marginBottom: 18 }}>좌석번호는 출결 키오스크에서 학생이 입력하는 번호예요.</div>

      <div style={{ background: SURFACE, border: `1px solid ${LINE}`, borderRadius: 14, padding: 16, display: "flex", gap: 8, flexWrap: "wrap" }}>
        <input style={{ ...inputStyle, maxWidth: 100 }} placeholder="좌석번호" inputMode="numeric"
          value={seat} onChange={(e) => setSeat(e.target.value.replace(/\D/g, "").slice(0, 2))} />
        <input style={inputStyle} placeholder="학생 이름" value={name}
          onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
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
          <div key={s.id} className="row" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "13px 16px", borderTop: `1px solid ${LINE}` }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 15 }}>
                <span style={{ display: "inline-block", minWidth: 34, color: BRAND, fontWeight: 800 }}>{s.seat}</span> {s.name}
              </div>
              {s.phone && <div style={{ fontSize: 12.5, color: MUTED, marginTop: 3, marginLeft: 34 }}>{s.phone}</div>}
            </div>
            <button onClick={() => remove(s)} style={{ border: `1px solid ${LINE}`, background: SURFACE, color: "#B42318", borderRadius: 8, padding: "6px 12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>삭제</button>
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
