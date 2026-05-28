// ─────────────────────────────────────────────────────────────────────────────
//  Smart Records Hub — School Systems Edition
//  Phase 4 · File Ingestion · Consolidation · Charts · AI Analyze
//  Color theme: Deep Cobalt Blue
// ─────────────────────────────────────────────────────────────────────────────

import { useState, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from "recharts";

// ── Color tokens ──────────────────────────────────────────────────────────────
const C = {
  headerBg:    "#06102e",
  navBg:       "#0a1840",
  subnavBg:    "#0f1f4a",
  accent:      "#1d6ef5",
  accentHover: "#1558d6",
  accentLight: "#dbeafe",
  accentMid:   "#3b82f6",
  white:       "#ffffff",
  dim:         "rgba(255,255,255,0.55)",
  dimmer:      "rgba(255,255,255,0.28)",
  border:      "rgba(255,255,255,0.10)",
  borderBlue:  "rgba(29,110,245,0.35)",
  contentBg:   "#f0f5ff",
  cardBg:      "#ffffff",
  textPrimary: "#0d1b3e",
  textSec:     "#4a6080",
  textTertiary:"#8faac8",
  rowHover:    "#eef4ff",
  danger:      "#dc2626",
  success:     "#16a34a",
  warning:     "#d97706",
};

const BLUE_GRADIENT = `linear-gradient(135deg, #06102e 0%, #0d2060 40%, #1640c2 100%)`;
const PIE_COLORS = ["#1d6ef5","#3b82f6","#60a5fa","#93c5fd","#bfdbfe"];

// ── Sub-navigation categories ─────────────────────────────────────────────────
const SUB_CATS = [
  "Student Records", "Enrollment", "Grade Reports",
  "Attendance", "Staff Certifications", "IEP / 504",
];

// ── Action tabs ───────────────────────────────────────────────────────────────
const ACTION_TABS = [
  { id: "about",      label: "About",      icon: "ℹ️" },
  { id: "records",    label: "Records",    icon: "📋" },
  { id: "ai",         label: "AI Analyze", icon: "🤖" },
  { id: "charts",     label: "Charts",     icon: "📊" },
  { id: "mapper",     label: "Mapper",     icon: "🔀" },
  { id: "dupes",      label: "Dupes",      icon: "⚠️" },
  { id: "schema",     label: "Schema",     icon: "🔍" },
];

// ── Sub-nav → template mapping ────────────────────────────────────────────────
const SUBNAV_TEMPLATE_MAP = {
  "Student Records":     "enrollment",
  "Enrollment":          "enrollment",
  "Grade Reports":       "grades",
  "Attendance":          "attendance",
  "Staff Certifications":"staff",
  "IEP / 504":           "iep",
};

// ── Templates ─────────────────────────────────────────────────────────────────
const TEMPLATES = [
  { id: "enrollment", icon: "🎓", label: "Student Enrollment", desc: "Registration & demographics" },
  { id: "attendance", icon: "📅", label: "Attendance",         desc: "Daily/period records" },
  { id: "staff",      icon: "📋", label: "Staff Certifications", desc: "Credentials & renewals" },
  { id: "grades",     icon: "📊", label: "Grade Reports",      desc: "Academic performance" },
  { id: "iep",        icon: "📝", label: "IEP / 504 Plans",    desc: "Special education records" },
  { id: "health",     icon: "💉", label: "Health & Immunization", desc: "Compliance tracking" },
];

// ── Template sample data generators ──────────────────────────────────────────
const SAMPLE_DATA = {
  enrollment: {
    headers: ["Student ID","Last Name","First Name","Grade","DOB","Enrollment Date","Guardian","Phone","Status","IEP"],
    rows: [
      ["S-1001","Brooks","Aaliyah","10","2009-03-14","2023-08-21","Denise Brooks","(301)555-0112","Active","No"],
      ["S-1002","Rivera","Marcus","11","2008-07-22","2022-08-19","Carlos Rivera","(301)555-0247","Active","Yes"],
      ["S-1003","Nair","Priya","9","2010-01-05","2024-08-20","Sunita Nair","(301)555-0389","Active","No"],
      ["S-1004","Wells","Jordan","12","2007-11-30","2021-08-23","Tamara Wells","(301)555-0455","Active","No"],
      ["S-1005","Kowalski","Ethan","10","2009-06-17","2023-08-21","Brian Kowalski","(301)555-0578","Suspended","Yes"],
      ["S-1006","Mendez","Sofia","11","2008-09-03","2022-08-19","Maria Mendez","(301)555-0621","Active","No"],
      ["S-1007","Okafor","Damien","9","2010-04-25","2024-08-20","Chioma Okafor","(301)555-0744","Active","No"],
      ["S-1008","Zhang","Lily","12","2007-02-14","2021-08-23","Wei Zhang","(301)555-0832","Active","No"],
      ["S-1009","Tran","Noah","10","2009-08-09","2023-08-21","Linh Tran","(301)555-0966","Withdrawn","No"],
      ["S-1010","Jackson","Imani","11","2008-12-20","2022-08-19","Renee Jackson","(301)555-0103","Active","Yes"],
    ],
  },
  attendance: {
    headers: ["Student ID","Name","Date","Period","Status","Teacher","Reason"],
    rows: [
      ["S-1001","Aaliyah Brooks","2026-05-27","1","Present","Mr. Kim",""],
      ["S-1002","Marcus Rivera","2026-05-27","1","Absent","Mr. Kim","Illness"],
      ["S-1003","Priya Nair","2026-05-27","1","Present","Mr. Kim",""],
      ["S-1004","Jordan Wells","2026-05-27","1","Tardy","Mr. Kim","Bus delay"],
      ["S-1005","Ethan Kowalski","2026-05-27","1","Absent","Mr. Kim","Suspension"],
      ["S-1006","Sofia Mendez","2026-05-27","1","Present","Mr. Kim",""],
      ["S-1001","Aaliyah Brooks","2026-05-26","1","Present","Mr. Kim",""],
      ["S-1002","Marcus Rivera","2026-05-26","1","Absent","Mr. Kim","Illness"],
      ["S-1007","Damien Okafor","2026-05-27","1","Present","Mr. Kim",""],
      ["S-1008","Lily Zhang","2026-05-27","1","Present","Mr. Kim",""],
    ],
  },
  staff: {
    headers: ["Staff ID","Name","Role","Department","Hire Date","Certification","Cert Expiry","Background Check","Email"],
    rows: [
      ["T-201","Dr. Angela Torres","Principal","Administration","2015-07-01","Valid","2027-06-30","Clear","atorres@school.edu"],
      ["T-202","Mr. David Kim","Teacher","Mathematics","2019-08-12","Valid","2026-12-31","Clear","dkim@school.edu"],
      ["T-203","Ms. Patricia Dunn","Counselor","Student Services","2018-01-15","Valid","2027-01-14","Clear","pdunn@school.edu"],
      ["T-204","Mr. James Holloway","Teacher","English","2021-08-10","Expiring","2026-07-01","Clear","jholloway@school.edu"],
      ["T-205","Ms. Yuki Tanaka","Nurse","Health Services","2020-03-01","Valid","2028-02-28","Clear","ytanaka@school.edu"],
      ["T-206","Mr. Raymond Bell","Coach","Athletics","2017-08-15","Valid","2027-08-14","Pending","rbell@school.edu"],
    ],
  },
  grades: {
    headers: ["Student ID","Name","Grade","Subject","Teacher","Q1","Q2","Q3","Q4","Final GPA","Status"],
    rows: [
      ["S-1001","Aaliyah Brooks","10","Mathematics","Mr. Kim","92","88","95","91","3.8","Passing"],
      ["S-1001","Aaliyah Brooks","10","English","Mr. Holloway","90","87","93","89","3.8","Passing"],
      ["S-1002","Marcus Rivera","11","Mathematics","Mr. Kim","80","78","84","82","3.2","Passing"],
      ["S-1002","Marcus Rivera","11","English","Mr. Holloway","75","79","81","78","3.2","Passing"],
      ["S-1003","Priya Nair","9","Mathematics","Mr. Kim","98","100","97","99","4.0","Passing"],
      ["S-1003","Priya Nair","9","English","Mr. Holloway","96","98","100","97","4.0","Passing"],
      ["S-1004","Jordan Wells","12","Mathematics","Mr. Kim","72","70","75","74","2.9","Passing"],
      ["S-1004","Jordan Wells","12","English","Mr. Holloway","68","71","73","70","2.9","At Risk"],
      ["S-1005","Ethan Kowalski","10","Mathematics","Mr. Kim","60","58","65","62","2.4","At Risk"],
      ["S-1005","Ethan Kowalski","10","English","Mr. Holloway","55","60","63","58","2.4","At Risk"],
      ["S-1006","Sofia Mendez","11","Mathematics","Mr. Kim","88","90","87","91","3.6","Passing"],
      ["S-1006","Sofia Mendez","11","English","Mr. Holloway","85","89","92","88","3.6","Passing"],
      ["S-1007","Damien Okafor","9","Mathematics","Mr. Kim","78","80","76","79","3.1","Passing"],
      ["S-1008","Lily Zhang","12","Mathematics","Mr. Kim","99","100","98","100","4.0","Passing"],
      ["S-1008","Lily Zhang","12","English","Mr. Holloway","97","100","99","98","4.0","Passing"],
    ],
  },
  iep: {
    headers: ["Student ID","Name","Grade","Disability Category","Plan Type","Service Hours/Wk","Case Manager","Annual Review Date","Next Meeting","Accommodations","Status"],
    rows: [
      ["S-1002","Marcus Rivera","11","Specific Learning Disability","IEP","8","Ms. Dunn","2026-08-19","2026-06-15","Extended time, preferential seating, read-aloud","Active"],
      ["S-1005","Ethan Kowalski","10","Emotional Behavioral Disorder","IEP","10","Ms. Dunn","2026-08-21","2026-06-10","Behavior intervention plan, reduced workload, check-in/check-out","Active"],
      ["S-1010","Imani Jackson","11","Other Health Impairment (ADHD)","IEP","6","Ms. Dunn","2026-08-19","2026-06-18","Extended time, breaks, organizational supports","Active"],
      ["S-1003","Priya Nair","9","Speech/Language Impairment","504","2","Ms. Dunn","2026-08-20","2026-07-01","Speech therapy twice weekly, presentation accommodations","Active"],
      ["S-1007","Damien Okafor","9","Anxiety Disorder","504","0","Ms. Dunn","2026-08-20","2026-07-05","Testing accommodations, flexible deadlines, counselor access","Active"],
      ["S-1004","Jordan Wells","12","Dyslexia","504","0","Ms. Dunn","2026-08-23","2026-06-20","Audio textbooks, extended time, spell-check permitted","Active - Senior"],
    ],
  },
  health: {
    headers: ["Student ID","Name","Grade","MMR","DTaP","Varicella","Hepatitis B","Meningococcal","TB Test","Exemption Type","Last Updated","Status"],
    rows: [
      ["S-1001","Aaliyah Brooks","10","Complete","Complete","Complete","Complete","Complete","Negative","None","2023-08-21","Compliant"],
      ["S-1002","Marcus Rivera","11","Complete","Complete","Complete","Complete","Complete","Negative","None","2022-08-19","Compliant"],
      ["S-1003","Priya Nair","9","Complete","Incomplete","Complete","Complete","Pending","Not Done","None","2024-08-20","Incomplete"],
      ["S-1004","Jordan Wells","12","Complete","Complete","Complete","Complete","Complete","Negative","None","2021-08-23","Compliant"],
      ["S-1005","Ethan Kowalski","10","Complete","Complete","Complete","Complete","Complete","Negative","None","2023-08-21","Compliant"],
      ["S-1006","Sofia Mendez","11","Complete","Complete","Complete","Complete","Complete","Negative","None","2022-08-19","Compliant"],
      ["S-1007","Damien Okafor","9","Complete","Complete","Waiver","Complete","Pending","Not Done","Religious","2024-08-20","Waiver on File"],
      ["S-1008","Lily Zhang","12","Complete","Complete","Complete","Complete","Complete","Negative","None","2021-08-23","Compliant"],
      ["S-1009","Noah Tran","10","Complete","Complete","Complete","Complete","Complete","Negative","None","2023-08-21","Compliant"],
      ["S-1010","Imani Jackson","11","Complete","Complete","Complete","Incomplete","Complete","Pending","None","2022-08-19","Incomplete"],
    ],
  },
};

// ── Utility helpers ───────────────────────────────────────────────────────────
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const sheets = [];
        wb.SheetNames.forEach((name) => {
          const ws = wb.Sheets[name];
          const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
          if (json.length > 1) {
            sheets.push({ sheet: name, file: file.name, headers: json[0], rows: json.slice(1) });
          }
        });
        resolve(sheets);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

function detectDuplicates(rows, headers) {
  const seen = new Map();
  const dupes = [];
  rows.forEach((row, i) => {
    const key = row.slice(0, 3).join("|").toLowerCase().trim();
    if (seen.has(key)) dupes.push({ row: i + 1, original: seen.get(key) + 1, key });
    else seen.set(key, i);
  });
  return dupes;
}

function computeHealth(rows, headers) {
  if (!rows.length) return [];
  return headers.map((h, ci) => {
    const filled = rows.filter(r => r[ci] !== "" && r[ci] !== null && r[ci] !== undefined).length;
    return { col: h, pct: Math.round((filled / rows.length) * 100) };
  });
}

// ── Small UI components ───────────────────────────────────────────────────────
function Pill({ label, active, onClick }) {
  return (
    <button onClick={onClick} style={{
      padding: "5px 14px", borderRadius: 99, fontSize: 12, fontWeight: 600,
      cursor: "pointer", border: active ? "none" : `1px solid ${C.border}`,
      background: active ? C.accent : "transparent",
      color: active ? C.white : C.dim,
      transition: "all .15s", whiteSpace: "nowrap",
      fontFamily: "inherit",
    }}>
      {label}
    </button>
  );
}

function StatCard({ label, value, color }) {
  return (
    <div style={{
      background: C.cardBg, border: `1px solid ${C.accentLight}`,
      borderRadius: 10, padding: "12px 16px", borderLeft: `4px solid ${color || C.accent}`,
    }}>
      <div style={{ fontSize: 11, color: C.textSec, fontWeight: 700, letterSpacing: 0.5, textTransform: "uppercase", marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 800, color: color || C.accent }}>{value}</div>
    </div>
  );
}

function Badge({ text, type }) {
  const map = {
    success: { bg: "#dcfce7", color: "#166534", border: "#bbf7d0" },
    warning: { bg: "#fef9c3", color: "#854d0e", border: "#fde047" },
    danger:  { bg: "#fee2e2", color: "#991b1b", border: "#fca5a5" },
    info:    { bg: "#dbeafe", color: "#1e3a8a", border: "#93c5fd" },
    neutral: { bg: "#f1f5f9", color: "#475569", border: "#cbd5e1" },
  };
  const s = map[type] || map.neutral;
  return (
    <span style={{
      padding: "2px 8px", borderRadius: 99, fontSize: 11, fontWeight: 700,
      background: s.bg, color: s.color, border: `1px solid ${s.border}`,
    }}>{text}</span>
  );
}

// ── Records Tab ───────────────────────────────────────────────────────────────
function RecordsTab({ rows, headers, sources }) {
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const PER_PAGE = 15;

  const filtered = useMemo(() =>
    rows.filter(r => r.some(cell => String(cell).toLowerCase().includes(search.toLowerCase()))),
    [rows, search]
  );
  const total = filtered.length;
  const paged = filtered.slice(page * PER_PAGE, (page + 1) * PER_PAGE);
  const pages = Math.max(1, Math.ceil(total / PER_PAGE));

  if (!rows.length) return <EmptyState />;

  return (
    <div>
      {/* Stats */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 18 }}>
        <StatCard label="Total Records" value={rows.length} />
        <StatCard label="Columns" value={headers.length} color="#7c3aed" />
        <StatCard label="Source Files" value={sources} color="#0891b2" />
        <StatCard label="Showing" value={total} color="#059669" />
      </div>

      {/* Search + export */}
      <div style={{ display: "flex", gap: 10, marginBottom: 14, alignItems: "center" }}>
        <input
          value={search} onChange={e => { setSearch(e.target.value); setPage(0); }}
          placeholder="Search records..."
          style={{
            flex: 1, padding: "8px 12px", borderRadius: 8,
            border: `1px solid ${C.accentLight}`, fontSize: 13,
            fontFamily: "inherit", color: C.textPrimary, outline: "none",
          }}
        />
        <button
          onClick={() => {
            const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Records");
            XLSX.writeFile(wb, "smart_records_export.xlsx");
          }}
          style={{
            padding: "8px 16px", borderRadius: 8, fontSize: 13, fontWeight: 600,
            background: C.accent, color: C.white, border: "none", cursor: "pointer", fontFamily: "inherit",
          }}>
          ↓ Export
        </button>
      </div>

      {/* Table */}
      <div style={{ overflowX: "auto", borderRadius: 10, border: `1px solid ${C.accentLight}` }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#eff6ff" }}>
              {headers.map(h => (
                <th key={h} style={{ padding: "9px 12px", textAlign: "left", fontWeight: 700, color: "#1e3a8a", fontSize: 11, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${C.accentLight}`, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {paged.map((row, ri) => (
              <tr key={ri}
                style={{ background: ri % 2 === 0 ? C.cardBg : "#f8faff" }}
                onMouseEnter={e => e.currentTarget.style.background = C.rowHover}
                onMouseLeave={e => e.currentTarget.style.background = ri % 2 === 0 ? C.cardBg : "#f8faff"}
              >
                {headers.map((_, ci) => (
                  <td key={ci} style={{ padding: "8px 12px", color: C.textPrimary, borderBottom: `1px solid #f0f4fc`, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {String(row[ci] ?? "")}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 12 }}>
        <span style={{ fontSize: 12, color: C.textSec }}>Page {page + 1} of {pages}</span>
        <div style={{ display: "flex", gap: 6 }}>
          <button disabled={page === 0} onClick={() => setPage(p => p - 1)}
            style={{ padding: "5px 12px", borderRadius: 6, border: `1px solid ${C.accentLight}`, background: page === 0 ? "#f8faff" : C.accent, color: page === 0 ? C.textSec : C.white, cursor: page === 0 ? "default" : "pointer", fontSize: 12, fontFamily: "inherit" }}>← Prev</button>
          <button disabled={page >= pages - 1} onClick={() => setPage(p => p + 1)}
            style={{ padding: "5px 12px", borderRadius: 6, border: `1px solid ${C.accentLight}`, background: page >= pages - 1 ? "#f8faff" : C.accent, color: page >= pages - 1 ? C.textSec : C.white, cursor: page >= pages - 1 ? "default" : "pointer", fontSize: 12, fontFamily: "inherit" }}>Next →</button>
        </div>
      </div>
    </div>
  );
}

// ── AI Analyze Tab ────────────────────────────────────────────────────────────
function AITab({ rows, headers }) {
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const chatRef = useRef(null);

  const health = useMemo(() => computeHealth(rows, headers), [rows, headers]);
  const avgHealth = health.length ? Math.round(health.reduce((a, b) => a + b.pct, 0) / health.length) : 0;
  const healthLabel = avgHealth >= 90 ? "Excellent" : avgHealth >= 75 ? "Good" : avgHealth >= 55 ? "Fair" : "Needs Attention";
  const healthColor = avgHealth >= 90 ? C.success : avgHealth >= 75 ? C.accent : avgHealth >= 55 ? C.warning : C.danger;

  const QUICK_PROMPTS = [
    "Which students have incomplete immunization records?",
    "Show a summary of IEP students by grade level",
    "Flag any staff certifications expiring within 6 months",
    "Identify attendance patterns suggesting at-risk students",
  ];

  async function sendMessage(text) {
    if (!text.trim()) return;
    const userMsg = { role: "user", content: text };
    const next = [...messages, userMsg];
    setMessages(next);
    setInput("");
    setLoading(true);

    const dataContext = rows.length
      ? `Schema: ${headers.join(", ")}\nRows (sample): ${rows.slice(0, 40).map(r => r.join(" | ")).join("\n")}\nTotal records: ${rows.length}`
      : "No data loaded yet.";

    const systemPrompt = `You are an AI analyst embedded in the Smart Records Hub — a school records consolidation tool. You help school administrators analyze student, staff, and compliance data. You are direct, concise, and highlight actionable findings. When analyzing data: flag compliance gaps, surface at-risk patterns, and note any FERPA-relevant concerns. Data health score: ${avgHealth}%.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          system: systemPrompt,
          messages: [
            { role: "user", content: `Current data:\n${dataContext}\n\nQuestion: ${text}` },
          ],
        }),
      });
      const data = await res.json();
      const reply = data.content?.[0]?.text || "No response received.";
      setMessages([...next, { role: "assistant", content: reply }]);
    } catch {
      setMessages([...next, { role: "assistant", content: "⚠️ Unable to reach AI service. Check your connection." }]);
    } finally {
      setLoading(false);
      setTimeout(() => chatRef.current?.scrollTo({ top: 9999, behavior: "smooth" }), 50);
    }
  }

  return (
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 20 }}>
      {/* Left: Data Health */}
      <div>
        <div style={{ background: C.cardBg, border: `1px solid ${C.accentLight}`, borderRadius: 12, padding: "16px" }}>
          <div style={{ fontWeight: 700, color: C.textPrimary, fontSize: 14, marginBottom: 12 }}>📊 Data Health Score</div>
          {rows.length === 0 ? (
            <p style={{ color: C.textSec, fontSize: 13 }}>Load data to see health metrics.</p>
          ) : (
            <>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
                <div style={{ width: 52, height: 52, borderRadius: "50%", background: `conic-gradient(${healthColor} ${avgHealth}%, #e2e8f0 0)`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <div style={{ width: 36, height: 36, borderRadius: "50%", background: C.cardBg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800, color: healthColor }}>{avgHealth}%</div>
                </div>
                <div>
                  <div style={{ fontWeight: 700, color: healthColor, fontSize: 15 }}>{healthLabel}</div>
                  <div style={{ fontSize: 12, color: C.textSec }}>{rows.length} records · {headers.length} columns</div>
                </div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                {health.map(h => (
                  <div key={h.col}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 2 }}>
                      <span style={{ color: C.textSec, fontWeight: 600 }}>{h.col}</span>
                      <span style={{ color: h.pct < 60 ? C.danger : C.textSec }}>{h.pct}%</span>
                    </div>
                    <div style={{ height: 5, borderRadius: 99, background: "#e2e8f0" }}>
                      <div style={{ height: 5, borderRadius: 99, width: `${h.pct}%`, background: h.pct >= 90 ? C.success : h.pct >= 65 ? C.accent : C.warning }} />
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </div>

      {/* Right: Chat — DEMO GATE */}
      {/* AI analysis is disabled on the public preview instance. The Data Health Score
          above runs entirely on client-side sample data. To run AI queries against
          your own records, request a custom deployment via the link below. */}
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <div style={{
          background: C.cardBg, border: `1px solid ${C.accentLight}`,
          borderRadius: 12, padding: "24px 20px", flex: 1,
          display: "flex", flexDirection: "column", alignItems: "center",
          justifyContent: "center", textAlign: "center",
        }}>
          <div style={{ fontSize: 32, marginBottom: 12 }}>🔒</div>
          <div style={{ fontWeight: 700, color: C.textPrimary, fontSize: 15, marginBottom: 10 }}>
            AI Analyzer is disabled in this demo
          </div>
          <div style={{ color: C.textSec, fontSize: 13, lineHeight: 1.7, maxWidth: 320, marginBottom: 20 }}>
            This is a public preview instance. The AI Analyzer — which queries your
            records for compliance gaps, at-risk patterns, and FERPA concerns — is
            available in custom deployments only.
          </div>
          <div style={{ color: C.textSec, fontSize: 13, lineHeight: 1.7, maxWidth: 320, marginBottom: 24 }}>
            In your own deployment, your data never leaves your environment. Only a
            capped 40-row sample is sent to the AI — never the full dataset.
          </div>
          <a
            href="https://rocklinconsult.netlify.app/#contact"
            target="_blank"
            rel="noopener noreferrer"
            style={{
              display: "inline-block", padding: "10px 22px", borderRadius: 8,
              background: C.accent, color: C.white, fontSize: 13,
              fontWeight: 700, textDecoration: "none", letterSpacing: "0.01em",
            }}
          >
            Request Your Instance →
          </a>
          <div style={{ color: C.textSec, fontSize: 11, marginTop: 14, lineHeight: 1.5 }}>
            Custom deployments include AI analysis, file upload,<br />and a private URL for your team.
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Charts Tab ────────────────────────────────────────────────────────────────
function ChartsTab({ rows, headers }) {
  if (!rows.length) return <EmptyState />;

  // Build numeric column candidates for a bar chart
  const numericCols = headers.filter((_, ci) =>
    rows.some(r => !isNaN(parseFloat(r[ci])) && r[ci] !== "")
  );
  const firstNumCol = numericCols[0];
  const ci = headers.indexOf(firstNumCol);

  // Distribution of first text column values
  const textCol = headers.find((_, i) => rows.some(r => isNaN(parseFloat(r[i]))));
  const tci = headers.indexOf(textCol);
  const distMap = {};
  rows.forEach(r => { const v = String(r[tci] || "Unknown"); distMap[v] = (distMap[v] || 0) + 1; });
  const distData = Object.entries(distMap).slice(0, 8).map(([name, value]) => ({ name, value }));

  // Numeric bar data
  const barData = firstNumCol
    ? rows.slice(0, 12).map((r, i) => ({ name: String(r[0]).slice(0, 10), value: parseFloat(r[ci]) || 0 }))
    : [];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        {/* Pie: distribution */}
        <div style={{ background: C.cardBg, border: `1px solid ${C.accentLight}`, borderRadius: 12, padding: "16px" }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: C.textPrimary, marginBottom: 12 }}>
            Distribution — {textCol || "Field 1"}
          </div>
          <ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie data={distData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={75} label={({ name, percent }) => `${name} ${Math.round(percent * 100)}%`} labelLine={false}>
                {distData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>

        {/* Bar: numeric or record count */}
        <div style={{ background: C.cardBg, border: `1px solid ${C.accentLight}`, borderRadius: 12, padding: "16px" }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: C.textPrimary, marginBottom: 12 }}>
            {firstNumCol ? `Values — ${firstNumCol}` : "Record Count by Group"}
          </div>
          <ResponsiveContainer width="100%" height={200}>
            {firstNumCol ? (
              <BarChart data={barData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e0e9f8" />
                <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                <YAxis tick={{ fontSize: 10 }} />
                <Tooltip />
                <Bar dataKey="value" fill={C.accent} radius={[4,4,0,0]} />
              </BarChart>
            ) : (
              <BarChart data={distData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e0e9f8" />
                <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                <YAxis tick={{ fontSize: 10 }} />
                <Tooltip />
                <Bar dataKey="value" fill={C.accentMid} radius={[4,4,0,0]} />
              </BarChart>
            )}
          </ResponsiveContainer>
        </div>
      </div>

      {/* Line: row index trend */}
      {firstNumCol && (
        <div style={{ background: C.cardBg, border: `1px solid ${C.accentLight}`, borderRadius: 12, padding: "16px" }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: C.textPrimary, marginBottom: 12 }}>Trend — {firstNumCol}</div>
          <ResponsiveContainer width="100%" height={160}>
            <LineChart data={rows.slice(0, 20).map((r, i) => ({ i: i + 1, value: parseFloat(r[ci]) || 0 }))}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e0e9f8" />
              <XAxis dataKey="i" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} />
              <Tooltip />
              <Line type="monotone" dataKey="value" stroke={C.accent} strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

// ── Field Mapper Tab ──────────────────────────────────────────────────────────
function MapperTab({ allSheets }) {
  if (!allSheets.length) return <EmptyState />;

  const schoolFields = [
    "Student ID", "First Name", "Last Name", "Grade", "DOB", "Enrollment Date",
    "Status", "Guardian", "Phone", "Email", "Teacher", "Dept", "Certification",
    "Hire Date", "IEP", "Immunization", "GPA", "Attendance %",
  ];

  const allCols = [...new Set(allSheets.flatMap(s => s.headers))];
  const [mapping, setMapping] = useState(() => {
    const m = {};
    allCols.forEach(c => {
      const match = schoolFields.find(sf => sf.toLowerCase().replace(/\s/g,"") === c.toLowerCase().replace(/\s/g,"") || sf.toLowerCase().includes(c.toLowerCase().split(" ")[0]));
      m[c] = match || "";
    });
    return m;
  });

  return (
    <div>
      <p style={{ fontSize: 13, color: C.textSec, marginBottom: 14 }}>
        Map your uploaded column names to standardized school record fields. Auto-matched where possible.
      </p>
      <div style={{ overflowX: "auto", borderRadius: 10, border: `1px solid ${C.accentLight}` }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#eff6ff" }}>
              {["Source Column", "From File", "Maps To", "Status"].map(h => (
                <th key={h} style={{ padding: "9px 14px", textAlign: "left", fontSize: 11, fontWeight: 700, color: "#1e3a8a", letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${C.accentLight}` }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {allCols.map((col, i) => {
              const src = allSheets.find(s => s.headers.includes(col));
              const mapped = mapping[col];
              return (
                <tr key={col} style={{ background: i % 2 === 0 ? C.cardBg : "#f8faff" }}>
                  <td style={{ padding: "8px 14px", fontWeight: 600, color: C.textPrimary }}>{col}</td>
                  <td style={{ padding: "8px 14px", color: C.textSec, fontSize: 12 }}>{src?.file || "—"}</td>
                  <td style={{ padding: "8px 14px" }}>
                    <select value={mapped} onChange={e => setMapping({ ...mapping, [col]: e.target.value })}
                      style={{ padding: "4px 8px", borderRadius: 6, border: `1px solid ${C.accentLight}`, fontSize: 12, fontFamily: "inherit", color: C.textPrimary, background: C.cardBg, outline: "none" }}>
                      <option value="">— Unmapped —</option>
                      {schoolFields.map(f => <option key={f} value={f}>{f}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: "8px 14px" }}>
                    <Badge text={mapped ? "Mapped" : "Unmapped"} type={mapped ? "success" : "warning"} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: 14, display: "flex", gap: 10 }}>
        <button style={{ padding: "8px 20px", borderRadius: 8, background: C.accent, color: C.white, border: "none", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>
          Apply Mapping
        </button>
        <button style={{ padding: "8px 20px", borderRadius: 8, background: "#f0f5ff", color: C.accent, border: `1px solid ${C.accentLight}`, fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>
          Reset
        </button>
      </div>
    </div>
  );
}

// ── Dupes Tab ─────────────────────────────────────────────────────────────────
function DupesTab({ rows, headers }) {
  if (!rows.length) return <EmptyState />;
  const dupes = useMemo(() => detectDuplicates(rows, headers), [rows, headers]);

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
        <div style={{ padding: "8px 18px", borderRadius: 10, background: dupes.length ? "#fef2f2" : "#f0fdf4", border: `1px solid ${dupes.length ? "#fca5a5" : "#86efac"}`, fontSize: 14, fontWeight: 700, color: dupes.length ? C.danger : C.success }}>
          {dupes.length ? `⚠️ ${dupes.length} Potential Duplicate${dupes.length !== 1 ? "s" : ""} Found` : "✅ No Duplicates Detected"}
        </div>
        <span style={{ fontSize: 12, color: C.textSec }}>Matching on first 3 columns</span>
      </div>
      {dupes.length > 0 && (
        <div style={{ overflowX: "auto", borderRadius: 10, border: `1px solid #fca5a5` }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#fef2f2" }}>
                {["Row #", "Matches Row", "Key Values", "Action"].map(h => (
                  <th key={h} style={{ padding: "9px 14px", textAlign: "left", fontSize: 11, fontWeight: 700, color: "#991b1b", letterSpacing: 0.4, textTransform: "uppercase", borderBottom: "1px solid #fca5a5" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {dupes.map((d, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? C.cardBg : "#fff5f5" }}>
                  <td style={{ padding: "8px 14px", fontWeight: 700, color: C.danger }}>Row {d.row}</td>
                  <td style={{ padding: "8px 14px", color: C.textSec }}>Row {d.original}</td>
                  <td style={{ padding: "8px 14px", color: C.textPrimary, fontFamily: "monospace", fontSize: 12 }}>{d.key.slice(0, 60)}</td>
                  <td style={{ padding: "8px 14px" }}>
                    <button style={{ padding: "3px 10px", borderRadius: 6, fontSize: 11, background: "#fee2e2", color: "#991b1b", border: "1px solid #fca5a5", cursor: "pointer", fontFamily: "inherit", fontWeight: 600 }}>Review</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── Schema Tab ────────────────────────────────────────────────────────────────
function SchemaTab({ rows, headers }) {
  if (!rows.length) return <EmptyState />;

  const colMeta = headers.map((h, ci) => {
    const vals = rows.map(r => r[ci]).filter(v => v !== "" && v !== null && v !== undefined);
    const isNum = vals.every(v => !isNaN(parseFloat(v)));
    const unique = new Set(vals).size;
    const fill = Math.round((vals.length / rows.length) * 100);
    const sample = [...new Set(vals)].slice(0, 3).join(", ");
    return { name: h, type: isNum ? "Numeric" : "Text", unique, fill, sample };
  });

  return (
    <div style={{ overflowX: "auto", borderRadius: 10, border: `1px solid ${C.accentLight}` }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead>
          <tr style={{ background: "#eff6ff" }}>
            {["Column Name", "Type", "Unique Values", "Fill Rate", "Sample Values"].map(h => (
              <th key={h} style={{ padding: "9px 14px", textAlign: "left", fontSize: 11, fontWeight: 700, color: "#1e3a8a", letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${C.accentLight}` }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {colMeta.map((c, i) => (
            <tr key={c.name} style={{ background: i % 2 === 0 ? C.cardBg : "#f8faff" }}>
              <td style={{ padding: "9px 14px", fontWeight: 700, color: C.textPrimary }}>{c.name}</td>
              <td style={{ padding: "9px 14px" }}>
                <Badge text={c.type} type={c.type === "Numeric" ? "info" : "neutral"} />
              </td>
              <td style={{ padding: "9px 14px", color: C.textSec }}>{c.unique}</td>
              <td style={{ padding: "9px 14px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ width: 70, height: 6, borderRadius: 99, background: "#e2e8f0" }}>
                    <div style={{ height: 6, borderRadius: 99, width: `${c.fill}%`, background: c.fill >= 90 ? C.success : c.fill >= 65 ? C.accent : C.warning }} />
                  </div>
                  <span style={{ fontSize: 12, color: c.fill < 65 ? C.warning : C.textSec }}>{c.fill}%</span>
                </div>
              </td>
              <td style={{ padding: "9px 14px", color: C.textSec, fontSize: 12, fontFamily: "monospace" }}>{c.sample || "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ── Empty State ───────────────────────────────────────────────────────────────
function EmptyState() {
  return (
    <div style={{ textAlign: "center", padding: "60px 20px", color: C.textSec }}>
      <div style={{ fontSize: 44, marginBottom: 12 }}>📋</div>
      <div style={{ fontSize: 16, fontWeight: 700, color: C.textPrimary, marginBottom: 6 }}>No records loaded yet</div>
      <div style={{ fontSize: 13, color: C.textSec }}>Upload Excel or CSV files using the panel on the left<br />to see your unified records table here.</div>
      <div style={{ marginTop: 16, fontSize: 13, color: C.accent, fontWeight: 600 }}>← Drop files to begin</div>
    </div>
  );
}

// ── About Tab ─────────────────────────────────────────────────────────────────
function AboutTab({ onGetStarted }) {
  const features = [
    { icon: "📁", title: "Multi-File Ingestion", desc: "Upload any number of .xlsx, .xls, .csv, or .tsv files. Multi-sheet workbooks are parsed automatically. All sources merge into a single unified table — no manual copy-pasting required." },
    { icon: "📋", title: "Records View", desc: "Browse, search, and paginate your consolidated records. Export the full merged dataset as a clean .xlsx file at any time." },
    { icon: "🤖", title: "AI Analyzer", desc: "Ask plain-language questions about your data — 'Which students have incomplete immunizations?' or 'Flag staff certifications expiring within 6 months.' Powered by Claude AI, grounded in your actual uploaded records." },
    { icon: "📊", title: "Auto Charts", desc: "Instant bar, pie, and line charts generated from your data the moment files are loaded. No configuration needed." },
    { icon: "🔀", title: "Field Mapper", desc: "Different files using different column names? The Mapper normalizes them to standard school record fields automatically, with manual override available." },
    { icon: "⚠️", title: "Duplicate Detector", desc: "Scans all loaded records for potential duplicates across source files so you can review and clean before exporting or reporting." },
    { icon: "🔍", title: "Schema Inspector", desc: "See column-by-column data types, fill rates, and sample values. Identifies data quality gaps before you run analysis." },
    { icon: "📘", title: "Study Corpus", desc: "A built-in reference library of school compliance frameworks — FERPA, IDEA, Section 504, immunization standards, and more — that the AI Analyzer draws from when answering compliance questions." },
  ];

  const workflow = [
    { step: "1", label: "Load Data", desc: "Drag in your Excel/CSV files or click a template to explore with sample data" },
    { step: "2", label: "Review & Clean", desc: "Use Mapper to normalize columns, Dupes to flag issues, Schema to check quality" },
    { step: "3", label: "Analyze", desc: "Ask the AI questions or explore auto-generated charts" },
    { step: "4", label: "Export", desc: "Download your clean, consolidated records as a single .xlsx file" },
  ];

  return (
    <div style={{ maxWidth: 860, margin: "0 auto" }}>
      {/* Hero */}
      <div style={{
        background: BLUE_GRADIENT, borderRadius: 14, padding: "32px 36px",
        marginBottom: 28, position: "relative", overflow: "hidden",
      }}>
        <div style={{ position: "absolute", top: -40, right: -40, width: 200, height: 200, borderRadius: "50%", background: "rgba(29,110,245,0.2)", pointerEvents: "none" }} />
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 16 }}>
          <div style={{ width: 52, height: 52, borderRadius: 12, background: "linear-gradient(135deg,#1d6ef5,#60a5fa)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 26 }}>🏫</div>
          <div>
            <div style={{ fontWeight: 800, fontSize: 22, color: C.white }}>Smart Records Hub</div>
            <div style={{ fontSize: 13, color: C.dim }}>School Systems Edition · Phase 4</div>
          </div>
        </div>
        <p style={{ fontSize: 15, color: "rgba(255,255,255,0.82)", lineHeight: 1.7, margin: "0 0 22px", maxWidth: 620 }}>
          A browser-based records consolidation tool built for school administrators, registrars, and compliance officers. It ingests messy multi-file spreadsheet data, unifies it into a searchable hub, and layers AI-powered analysis on top — no formulas, no pivot tables, no IT ticket required.
        </p>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {["FERPA-Aware","AI-Powered","No Installation","Works Offline","Privacy-First"].map(tag => (
            <span key={tag} style={{ padding: "4px 12px", borderRadius: 99, background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.2)", color: "rgba(255,255,255,0.85)", fontSize: 11, fontWeight: 700 }}>{tag}</span>
          ))}
        </div>
      </div>

      {/* Who it's for */}
      <div style={{ background: C.cardBg, border: `1px solid ${C.accentLight}`, borderRadius: 12, padding: "20px 24px", marginBottom: 22 }}>
        <div style={{ fontWeight: 800, fontSize: 15, color: C.textPrimary, marginBottom: 14 }}>👥 Who This Tool Is For</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
          {[
            { role: "School Registrars", desc: "Consolidate enrollment, withdrawal, and transfer records from multiple spreadsheets into one clean view." },
            { role: "Compliance Officers", desc: "Track immunization status, IEP deadlines, staff certification renewals, and FERPA-related documentation gaps." },
            { role: "Administrators", desc: "Get instant AI-generated summaries and charts from attendance, grade, and staffing data without touching a formula." },
          ].map(u => (
            <div key={u.role} style={{ background: "#f8faff", borderRadius: 10, padding: "14px 16px", border: `1px solid ${C.accentLight}` }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: C.accent, marginBottom: 6 }}>{u.role}</div>
              <div style={{ fontSize: 12, color: C.textSec, lineHeight: 1.55 }}>{u.desc}</div>
            </div>
          ))}
        </div>
      </div>

      {/* How it works */}
      <div style={{ background: C.cardBg, border: `1px solid ${C.accentLight}`, borderRadius: 12, padding: "20px 24px", marginBottom: 22 }}>
        <div style={{ fontWeight: 800, fontSize: 15, color: C.textPrimary, marginBottom: 16 }}>⚡ How It Works</div>
        <div style={{ display: "flex", gap: 0, position: "relative" }}>
          {workflow.map((w, i) => (
            <div key={w.step} style={{ flex: 1, textAlign: "center", position: "relative" }}>
              {i < workflow.length - 1 && (
                <div style={{ position: "absolute", top: 18, left: "60%", width: "80%", height: 2, background: C.accentLight, zIndex: 0 }} />
              )}
              <div style={{ width: 38, height: 38, borderRadius: "50%", background: C.accent, color: C.white, fontWeight: 800, fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 10px", position: "relative", zIndex: 1 }}>{w.step}</div>
              <div style={{ fontWeight: 700, fontSize: 13, color: C.textPrimary, marginBottom: 4 }}>{w.label}</div>
              <div style={{ fontSize: 11, color: C.textSec, lineHeight: 1.5, padding: "0 8px" }}>{w.desc}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Feature grid */}
      <div style={{ background: C.cardBg, border: `1px solid ${C.accentLight}`, borderRadius: 12, padding: "20px 24px", marginBottom: 22 }}>
        <div style={{ fontWeight: 800, fontSize: 15, color: C.textPrimary, marginBottom: 14 }}>🛠 Features at a Glance</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          {features.map(f => (
            <div key={f.title} style={{ display: "flex", gap: 12, padding: "12px 14px", borderRadius: 10, background: "#f8faff", border: `1px solid ${C.accentLight}` }}>
              <span style={{ fontSize: 22, flexShrink: 0, lineHeight: 1.3 }}>{f.icon}</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 13, color: C.textPrimary, marginBottom: 3 }}>{f.title}</div>
                <div style={{ fontSize: 11, color: C.textSec, lineHeight: 1.55 }}>{f.desc}</div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Privacy note */}
      <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 12, padding: "16px 20px", marginBottom: 22, display: "flex", gap: 14, alignItems: "flex-start" }}>
        <span style={{ fontSize: 24, flexShrink: 0 }}>🔒</span>
        <div>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#166534", marginBottom: 4 }}>Privacy-First by Design</div>
          <div style={{ fontSize: 12, color: "#166534", lineHeight: 1.6 }}>All file parsing happens in your browser. Your raw data never leaves your device unless you explicitly use the AI Analyze feature, which sends only a capped 40-row sample and column schema to the AI — never your full dataset. No accounts, no cloud storage, no data retention.</div>
        </div>
      </div>

      {/* CTA */}
      <div style={{ textAlign: "center", padding: "10px 0 20px" }}>
        <button onClick={onGetStarted} style={{
          padding: "12px 36px", borderRadius: 10, background: C.accent,
          color: C.white, border: "none", fontWeight: 800, fontSize: 15,
          cursor: "pointer", fontFamily: "inherit",
          boxShadow: "0 4px 14px rgba(29,110,245,0.35)",
        }}>
          Get Started → Load a Template
        </button>
        <div style={{ fontSize: 12, color: C.textSec, marginTop: 10 }}>Or drop your own files in the left panel at any time</div>
      </div>
    </div>
  );
}

// ── Empty State ───────────────────────────────────────────────────────────────
function DropZone({ onFiles }) {
  // ── DEMO GATE ─────────────────────────────────────────────────────────────────
  // File upload is disabled on the public preview instance. Visitors can explore
  // all features using the sample data templates below. To upload real school data,
  // a custom deployment is required — request one via the link below.
  const CONTACT_URL = "https://rocklinconsult.netlify.app/#contact";

  return (
    <div style={{
      border: `2px dashed ${C.borderBlue}`,
      borderRadius: 12, padding: "22px 16px", textAlign: "center",
      background: "rgba(255,255,255,0.02)",
    }}>
      <div style={{ fontSize: 26, marginBottom: 8 }}>🔒</div>
      <div style={{ color: C.white, fontWeight: 700, fontSize: 13, marginBottom: 6 }}>
        File upload is disabled in this demo
      </div>
      <div style={{ color: C.dim, fontSize: 11, lineHeight: 1.6, marginBottom: 14 }}>
        This is a public preview. To use Smart Records Hub with your own school data,
        request a custom deployment — your data stays in your environment, not ours.
      </div>
      <a
        href={CONTACT_URL}
        target="_blank"
        rel="noopener noreferrer"
        style={{
          display: "inline-block", padding: "7px 16px", borderRadius: 8,
          background: C.accent, color: C.white, fontSize: 12,
          fontWeight: 700, textDecoration: "none", letterSpacing: "0.01em",
        }}
      >
        Request Your Instance →
      </a>
      <div style={{ color: C.dimmer, fontSize: 10, marginTop: 10 }}>
        Full support for .xlsx · .xls · .csv · .tsv in your deployment
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
// ── Settings Modal ────────────────────────────────────────────────────────────
function SettingsModal({ onClose }) {
  const [schoolName, setSchoolName] = useState("Lincoln High School");
  const [district, setDistrict] = useState("Prince George's County Public Schools");
  const [rowsPerPage, setRowsPerPage] = useState("15");
  const [saved, setSaved] = useState(false);

  function save() {
    setSaved(true);
    setTimeout(() => { setSaved(false); onClose(); }, 900);
  }

  const inputStyle = {
    width: "100%", padding: "8px 11px", borderRadius: 8,
    border: `1px solid ${C.accentLight}`, fontSize: 13,
    fontFamily: "inherit", color: C.textPrimary, outline: "none",
    background: "#f8faff",
  };
  const labelStyle = { fontSize: 11, fontWeight: 700, color: C.textSec, letterSpacing: 0.4, textTransform: "uppercase", marginBottom: 5, display: "block" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(5,14,40,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 500 }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{ background: C.cardBg, borderRadius: 16, padding: "28px", width: 420, boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
          <div style={{ fontWeight: 800, fontSize: 17, color: C.textPrimary }}>⚙️ Settings</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: C.textSec, cursor: "pointer", lineHeight: 1 }}>✕</button>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div>
            <label style={labelStyle}>School Name</label>
            <input style={inputStyle} value={schoolName} onChange={e => setSchoolName(e.target.value)} />
          </div>
          <div>
            <label style={labelStyle}>District</label>
            <input style={inputStyle} value={district} onChange={e => setDistrict(e.target.value)} />
          </div>
          <div>
            <label style={labelStyle}>Rows Per Page</label>
            <select value={rowsPerPage} onChange={e => setRowsPerPage(e.target.value)}
              style={{ ...inputStyle, cursor: "pointer" }}>
              {["10","15","25","50","100"].map(n => <option key={n} value={n}>{n} rows</option>)}
            </select>
          </div>
          <div style={{ borderTop: `1px solid ${C.accentLight}`, paddingTop: 14, display: "flex", gap: 10 }}>
            <button onClick={save} style={{ flex: 1, padding: "9px 0", borderRadius: 8, background: saved ? C.success : C.accent, color: C.white, border: "none", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit", transition: "background .2s" }}>
              {saved ? "✓ Saved!" : "Save Settings"}
            </button>
            <button onClick={onClose} style={{ flex: 1, padding: "9px 0", borderRadius: 8, background: "#f0f4fa", color: C.textSec, border: "none", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>Cancel</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Help Modal ────────────────────────────────────────────────────────────────
function HelpModal({ onClose }) {
  const sections = [
    { icon: "📁", title: "Upload Files", body: "Drag and drop .xlsx, .xls, .csv, or .tsv files into the left panel, or click to browse. Multiple files merge into one unified table automatically." },
    { icon: "📋", title: "Records Tab", body: "View and search your consolidated data. Use the Export button to download a clean .xlsx file of all records." },
    { icon: "🤖", title: "AI Analyze", body: "Ask natural-language questions about your data — e.g. 'Which students have incomplete immunizations?' Use the quick-prompt chips for common school record queries." },
    { icon: "📊", title: "Charts", body: "Auto-generated bar, pie, and line charts based on your loaded data. Charts update whenever you upload new files." },
    { icon: "🔀", title: "Mapper", body: "Map inconsistent column names from different source files to standardized school record fields. Auto-matches where possible." },
    { icon: "⚠️", title: "Dupes", body: "Automatically detects potential duplicate records by matching the first 3 columns. Review flagged rows before exporting." },
    { icon: "🔍", title: "Schema", body: "Inspect column types, fill rates, and sample values. Use this to identify data quality issues before analysis." },
    { icon: "🎓", title: "Templates", body: "Click any template in the left sidebar to load sample school records data instantly — great for exploring the tool before uploading real files." },
  ];
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(5,14,40,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 500 }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{ background: C.cardBg, borderRadius: 16, padding: "28px", width: 500, maxHeight: "80vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.25)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontWeight: 800, fontSize: 17, color: C.textPrimary }}>❓ Help & Guide</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, color: C.textSec, cursor: "pointer", lineHeight: 1 }}>✕</button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {sections.map(s => (
            <div key={s.title} style={{ display: "flex", gap: 12, padding: "12px 14px", borderRadius: 10, background: "#f8faff", border: `1px solid ${C.accentLight}` }}>
              <span style={{ fontSize: 20, lineHeight: 1.4 }}>{s.icon}</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 13, color: C.textPrimary, marginBottom: 3 }}>{s.title}</div>
                <div style={{ fontSize: 12, color: C.textSec, lineHeight: 1.55 }}>{s.body}</div>
              </div>
            </div>
          ))}
        </div>
        <button onClick={onClose} style={{ marginTop: 18, width: "100%", padding: "9px 0", borderRadius: 8, background: C.accent, color: C.white, border: "none", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>Got it</button>
      </div>
    </div>
  );
}

// ── Confirm Modal ─────────────────────────────────────────────────────────────
function ConfirmModal({ message, onConfirm, onCancel }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(5,14,40,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 500 }} onClick={onCancel}>
      <div onClick={e => e.stopPropagation()} style={{ background: C.cardBg, borderRadius: 16, padding: "28px", width: 360, boxShadow: "0 24px 60px rgba(0,0,0,0.25)", textAlign: "center" }}>
        <div style={{ fontSize: 36, marginBottom: 12 }}>⚠️</div>
        <div style={{ fontWeight: 700, fontSize: 16, color: C.textPrimary, marginBottom: 8 }}>Are you sure?</div>
        <div style={{ fontSize: 13, color: C.textSec, marginBottom: 22, lineHeight: 1.5 }}>{message}</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onConfirm} style={{ flex: 1, padding: "9px 0", borderRadius: 8, background: C.danger, color: C.white, border: "none", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>Yes, Clear</button>
          <button onClick={onCancel} style={{ flex: 1, padding: "9px 0", borderRadius: 8, background: "#f0f4fa", color: C.textSec, border: "none", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ── More Menu (··· button) ────────────────────────────────────────────────────
function MoreMenu({ onSettings, onExportAll, onClearData, onHelp }) {
  const [open, setOpen] = useState(false);
  const [hovered, setHovered] = useState(false);

  const items = [
    { label: "⚙️  Settings",   action: onSettings   },
    { label: "↓  Export All",  action: onExportAll  },
    { label: "🗑  Clear Data",  action: onClearData  },
    { label: "❓  Help",        action: onHelp       },
  ];

  function fire(fn) { setOpen(false); fn && fn(); }

  return (
    <div style={{ position: "relative" }}>
      <button
        onClick={() => setOpen(v => !v)}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        style={{
          width: 32, height: 32, borderRadius: 8,
          background: open ? "rgba(29,110,245,0.28)" : hovered ? "rgba(255,255,255,0.08)" : "rgba(255,255,255,0.05)",
          border: open ? `1px solid rgba(29,110,245,0.55)` : `1px solid ${C.border}`,
          color: open || hovered ? C.white : C.dim,
          fontSize: 15, cursor: "pointer", transition: "all .15s",
          display: "flex", alignItems: "center", justifyContent: "center", letterSpacing: 1,
        }}>···</button>
      {open && (
        <div style={{
          position: "absolute", right: 0, top: 38, zIndex: 400,
          background: "#0d1f50", border: `1px solid rgba(29,110,245,0.35)`,
          borderRadius: 10, padding: "6px 0", minWidth: 160,
          boxShadow: "0 8px 24px rgba(0,0,0,0.4)",
        }}>
          {items.map(item => (
            <button key={item.label} onClick={() => fire(item.action)}
              style={{
                display: "block", width: "100%", padding: "9px 16px",
                background: "none", border: "none", color: C.dim,
                fontSize: 13, textAlign: "left", cursor: "pointer",
                fontFamily: "inherit", transition: "all .1s",
              }}
              onMouseEnter={e => { e.currentTarget.style.background = "rgba(29,110,245,0.2)"; e.currentTarget.style.color = C.white; }}
              onMouseLeave={e => { e.currentTarget.style.background = "none"; e.currentTarget.style.color = C.dim; }}
            >{item.label}</button>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Study Corpus Panel ────────────────────────────────────────────────────────
const CORPUS_DOCS = [
  {
    category: "Federal Law",
    items: [
      { title: "FERPA — Family Educational Rights and Privacy Act", desc: "Governs access to student education records. Schools must obtain written consent before disclosing personally identifiable information.", tag: "Privacy", icon: "🔒" },
      { title: "IDEA — Individuals with Disabilities Education Act", desc: "Requires schools to provide Free Appropriate Public Education (FAPE) to eligible students with disabilities via individualized IEPs.", tag: "Special Ed", icon: "📝" },
      { title: "Section 504 — Rehabilitation Act", desc: "Prohibits discrimination against students with disabilities. 504 Plans provide accommodations without requiring special education services.", tag: "Special Ed", icon: "♿" },
      { title: "PPRA — Protection of Pupil Rights Amendment", desc: "Requires parental consent before students participate in surveys, analyses, or evaluations funded by the Department of Education.", tag: "Privacy", icon: "📋" },
    ],
  },
  {
    category: "Health & Safety",
    items: [
      { title: "Immunization Requirements (State Compliance)", desc: "Most states require documentation of vaccines including MMR, Varicella, DTaP, and Hepatitis B before school enrollment. Exemptions must be documented.", tag: "Health", icon: "💉" },
      { title: "Mandatory Reporting Obligations", desc: "School staff are mandated reporters under state law. Suspected abuse or neglect must be reported to child protective services immediately.", tag: "Safety", icon: "🚨" },
      { title: "HIPAA vs. FERPA in School Settings", desc: "Student health records maintained by schools are covered by FERPA, not HIPAA — except for records held by a school nurse providing non-school health services.", tag: "Compliance", icon: "⚖️" },
    ],
  },
  {
    category: "Records Management",
    items: [
      { title: "Student Record Retention Schedule", desc: "Permanent records (transcripts, enrollment) must be kept indefinitely. Temporary records (discipline, health) typically retained 5–7 years after graduation.", tag: "Records", icon: "🗂️" },
      { title: "IEP Documentation Standards", desc: "IEPs must be reviewed annually. Records must include present levels of performance, measurable annual goals, services provided, and progress reports.", tag: "Special Ed", icon: "📊" },
      { title: "Attendance Record Requirements", desc: "Daily attendance records are required by state law for funding calculations. Unexcused absence thresholds trigger mandatory intervention protocols.", tag: "Compliance", icon: "📅" },
    ],
  },
];

const TAG_COLORS = {
  Privacy:    { bg: "#ede9fe", color: "#5b21b6", border: "#c4b5fd" },
  "Special Ed":{ bg: "#dbeafe", color: "#1e40af", border: "#93c5fd" },
  Health:     { bg: "#dcfce7", color: "#166534", border: "#86efac" },
  Safety:     { bg: "#fee2e2", color: "#991b1b", border: "#fca5a5" },
  Compliance: { bg: "#fef9c3", color: "#854d0e", border: "#fde047" },
  Records:    { bg: "#f0fdf4", color: "#166534", border: "#86efac" },
};

function StudyCorpusPanel({ onClose }) {
  const [search, setSearch] = useState("");
  const [expanded, setExpanded] = useState(null);

  const filtered = CORPUS_DOCS.map(section => ({
    ...section,
    items: section.items.filter(item =>
      item.title.toLowerCase().includes(search.toLowerCase()) ||
      item.desc.toLowerCase().includes(search.toLowerCase()) ||
      item.tag.toLowerCase().includes(search.toLowerCase())
    ),
  })).filter(s => s.items.length > 0);

  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(5,14,40,0.55)",
      display: "flex", justifyContent: "flex-end", zIndex: 500,
    }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{
        width: 480, height: "100%", background: C.cardBg,
        overflowY: "auto", display: "flex", flexDirection: "column",
        boxShadow: "-8px 0 40px rgba(0,0,0,0.3)",
      }}>
        {/* Header */}
        <div style={{ background: BLUE_GRADIENT, padding: "20px 24px", flexShrink: 0 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
            <div>
              <div style={{ fontWeight: 800, fontSize: 17, color: C.white }}>📘 Study Corpus</div>
              <div style={{ fontSize: 12, color: C.dim, marginTop: 3 }}>Reference library for the AI Analyzer</div>
            </div>
            <button onClick={onClose} style={{ background: "rgba(255,255,255,0.1)", border: "none", color: C.white, width: 28, height: 28, borderRadius: 6, cursor: "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
          </div>
          <input
            value={search} onChange={e => setSearch(e.target.value)}
            placeholder="Search regulations, policies..."
            style={{ width: "100%", padding: "8px 12px", borderRadius: 8, border: "1px solid rgba(255,255,255,0.2)", background: "rgba(255,255,255,0.1)", color: C.white, fontSize: 13, fontFamily: "inherit", outline: "none", boxSizing: "border-box" }}
          />
          <div style={{ marginTop: 10, fontSize: 11, color: C.dimmer }}>
            {CORPUS_DOCS.reduce((a, s) => a + s.items.length, 0)} reference documents · Used by AI Analyze to ground responses in school compliance standards
          </div>
        </div>

        {/* Document list */}
        <div style={{ padding: "20px 24px", flex: 1 }}>
          {filtered.map(section => (
            <div key={section.category} style={{ marginBottom: 24 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: C.textSec, letterSpacing: 0.8, textTransform: "uppercase", marginBottom: 10, borderBottom: `1px solid ${C.accentLight}`, paddingBottom: 6 }}>
                {section.category}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {section.items.map(doc => {
                  const isOpen = expanded === doc.title;
                  const tc = TAG_COLORS[doc.tag] || TAG_COLORS.Records;
                  return (
                    <div key={doc.title}
                      onClick={() => setExpanded(isOpen ? null : doc.title)}
                      style={{ background: isOpen ? "#f0f6ff" : "#f8faff", border: `1px solid ${isOpen ? C.accentLight : "#e2eaf6"}`, borderRadius: 10, padding: "12px 14px", cursor: "pointer", transition: "all .15s" }}
                    >
                      <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                        <span style={{ fontSize: 20, flexShrink: 0 }}>{doc.icon}</span>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                            <div style={{ fontWeight: 700, fontSize: 13, color: C.textPrimary, lineHeight: 1.4 }}>{doc.title}</div>
                            <span style={{ fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 99, background: tc.bg, color: tc.color, border: `1px solid ${tc.border}`, whiteSpace: "nowrap", flexShrink: 0 }}>{doc.tag}</span>
                          </div>
                          {isOpen && (
                            <div style={{ fontSize: 12, color: C.textSec, marginTop: 8, lineHeight: 1.6 }}>{doc.desc}</div>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
          {filtered.length === 0 && (
            <div style={{ textAlign: "center", padding: "40px 0", color: C.textSec, fontSize: 13 }}>No documents match your search.</div>
          )}
        </div>

        {/* Footer note */}
        <div style={{ padding: "14px 24px", borderTop: `1px solid ${C.accentLight}`, background: "#f8faff", fontSize: 11, color: C.textSec }}>
          💡 The AI Analyzer references this corpus when answering questions about your school records. Future versions will allow you to upload custom policy documents.
        </div>
      </div>
    </div>
  );
}

// ── Header Nav Button ─────────────────────────────────────────────────────────
function HeaderNavBtn({ icon, label, active, onClick }) {
  const [hovered, setHovered] = useState(false);
  const isLit = active || hovered;
  return (
    <button
      onClick={onClick}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        display: "flex", alignItems: "center", gap: 7,
        padding: "6px 12px", borderRadius: 8,
        background: active
          ? "rgba(29,110,245,0.28)"
          : hovered ? "rgba(255,255,255,0.08)" : "transparent",
        border: active
          ? `1px solid rgba(29,110,245,0.55)`
          : hovered ? `1px solid rgba(255,255,255,0.15)` : "1px solid transparent",
        color: isLit ? C.white : C.dim,
        fontSize: 12, fontWeight: active ? 700 : 500,
        cursor: "pointer", fontFamily: "inherit",
        transition: "all .15s",
      }}
    >
      <span style={{ fontSize: 15, lineHeight: 1 }}>{icon}</span>
      <span>{label}</span>
    </button>
  );
}

export default function SmartRecordsHub() {
  const [allSheets, setAllSheets] = useState([]);
  const [activeSubCat, setActiveSubCat] = useState("Student Records");
  const [activeAction, setActiveAction] = useState("about");
  const [activeNav, setActiveNav] = useState(null);
  const [modal, setModal] = useState(null); // "settings" | "help" | "confirm-clear"
  const [showCorpus, setShowCorpus] = useState(false);

  function handleExportAll() {
    if (!rows.length) return alert("No data loaded to export.");
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "All Records");
    XLSX.writeFile(wb, "smart_records_hub_export.xlsx");
  }

  function handleClearConfirmed() {
    setAllSheets([]);
    setActiveAction("records");
    setModal(null);
  }

  // Merge all sheets into one flat table
  const { headers, rows } = useMemo(() => {
    if (!allSheets.length) return { headers: [], rows: [] };
    const allHeaders = [...new Set(allSheets.flatMap(s => s.headers))];
    const allRows = allSheets.flatMap(s =>
      s.rows.map(r => allHeaders.map(h => {
        const idx = s.headers.indexOf(h);
        return idx >= 0 ? r[idx] : "";
      }))
    );
    return { headers: allHeaders, rows: allRows };
  }, [allSheets]);

  // Load template sample data
  function loadTemplate(id) {
    const t = SAMPLE_DATA[id];
    if (!t) return;
    setAllSheets([{ sheet: id, file: `${id}_template.xlsx`, headers: t.headers, rows: t.rows }]);
  }

  function handleFiles(newSheets) {
    setAllSheets(prev => [...prev, ...newSheets]);
    setActiveAction("records");
  }

  return (
    <div style={{ fontFamily: "'Inter', 'Segoe UI', sans-serif", minHeight: "100vh", background: C.contentBg, display: "flex", flexDirection: "column" }}>
      <style>{`* { box-sizing: border-box; } body { margin: 0; }`}</style>

      {/* ── Modals ──────────────────────────────────────────────────────── */}
      {showCorpus                && <StudyCorpusPanel onClose={() => { setShowCorpus(false); setActiveNav(null); }} />}
      {modal === "settings"      && <SettingsModal onClose={() => setModal(null)} />}
      {modal === "help"          && <HelpModal onClose={() => setModal(null)} />}
      {modal === "confirm-clear" && (
        <ConfirmModal
          message="This will remove all loaded files and data from the current session. This cannot be undone."
          onConfirm={handleClearConfirmed}
          onCancel={() => setModal(null)}
        />
      )}

      {/* ── Top Header ──────────────────────────────────────────────────── */}
      <div style={{ background: BLUE_GRADIENT, padding: "0 24px", position: "sticky", top: 0, zIndex: 300 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 0", borderBottom: `1px solid ${C.border}` }}>
          {/* Logo + Title */}
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{
              width: 38, height: 38, borderRadius: 9,
              background: "linear-gradient(135deg, #1d6ef5 0%, #60a5fa 100%)",
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 18, boxShadow: "0 2px 8px rgba(29,110,245,0.5)",
            }}>🏫</div>
            <div>
              <div style={{ fontWeight: 800, fontSize: 16, color: C.white, letterSpacing: 0.2 }}>Smart Records Hub</div>
              <div style={{ fontSize: 11, color: C.dim }}>Records Consolidation · AI Analyzer</div>
            </div>
            <div style={{ padding: "2px 10px", borderRadius: 99, background: C.accent, color: C.white, fontSize: 10, fontWeight: 800, letterSpacing: 1, marginLeft: 4 }}>PHASE 4</div>
          </div>

          {/* Right side nav */}
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <HeaderNavBtn icon="📘" label="Study Corpus" active={showCorpus} onClick={() => { setShowCorpus(v => !v); setActiveNav(null); }} />
            <HeaderNavBtn icon="📁" label="Records" active={activeNav === "records-nav"} onClick={() => { setActiveNav("records-nav"); setActiveAction("records"); setShowCorpus(false); }} />
            <MoreMenu
              onSettings={() => setModal("settings")}
              onExportAll={handleExportAll}
              onClearData={() => setModal("confirm-clear")}
              onHelp={() => setModal("help")}
            />
          </div>
        </div>

        {/* ── Sub-navigation ───────────────────────────────────────────── */}
        <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "12px 0", flexWrap: "wrap" }}>
          <span style={{ color: C.accent, fontWeight: 800, fontSize: 13, marginRight: 8, whiteSpace: "nowrap" }}>School Systems</span>
          <span style={{ color: C.dimmer, fontSize: 13 }}>|</span>
          {SUB_CATS.map(cat => (
            <Pill key={cat} label={cat} active={activeSubCat === cat} onClick={() => {
              setActiveSubCat(cat);
              const templateId = SUBNAV_TEMPLATE_MAP[cat];
              if (templateId) {
                const t = SAMPLE_DATA[templateId];
                if (t) {
                  setAllSheets([{ sheet: templateId, file: `${templateId}_template.xlsx`, headers: t.headers, rows: t.rows }]);
                  setActiveAction("records");
                }
              }
            }} />
          ))}
        </div>
      </div>

      {/* ── Body ────────────────────────────────────────────────────────── */}
      <div style={{ display: "grid", gridTemplateColumns: "260px 1fr", flex: 1, gap: 0, minHeight: 0 }}>

        {/* ── Left Sidebar ──────────────────────────────────────────────── */}
        <div style={{ background: C.navBg, padding: "20px 16px", display: "flex", flexDirection: "column", gap: 20, overflowY: "auto", minHeight: "calc(100vh - 110px)" }}>

          {/* Drop zone */}
          <DropZone onFiles={handleFiles} />

          {/* Loaded files */}
          {allSheets.length > 0 && (
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: C.dimmer, letterSpacing: 1, textTransform: "uppercase", marginBottom: 8 }}>Loaded ({allSheets.length} sheet{allSheets.length !== 1 ? "s" : ""})</div>
              {allSheets.map((s, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 8px", borderRadius: 7, background: "rgba(29,110,245,0.12)", border: `1px solid ${C.borderBlue}`, marginBottom: 5 }}>
                  <div>
                    <div style={{ fontSize: 11, color: C.white, fontWeight: 600 }}>{s.file}</div>
                    <div style={{ fontSize: 10, color: C.dimmer }}>{s.sheet} · {s.rows.length} rows</div>
                  </div>
                  <button onClick={() => setAllSheets(prev => prev.filter((_, j) => j !== i))}
                    style={{ background: "none", border: "none", color: C.dimmer, cursor: "pointer", fontSize: 14, padding: 2 }}>✕</button>
                </div>
              ))}
            </div>
          )}

          {/* Templates */}
          <div>
            <div style={{ fontSize: 10, fontWeight: 700, color: C.dimmer, letterSpacing: 1, textTransform: "uppercase", marginBottom: 10 }}>Templates</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              {TEMPLATES.map(t => (
                <button key={t.id} onClick={() => { loadTemplate(t.id); setActiveAction("records"); }}
                  style={{ background: "rgba(255,255,255,0.04)", border: `1px solid ${C.border}`, borderRadius: 9, padding: "10px 8px", cursor: "pointer", textAlign: "left", transition: "all .15s", fontFamily: "inherit" }}
                  onMouseEnter={e => e.currentTarget.style.background = "rgba(29,110,245,0.15)"}
                  onMouseLeave={e => e.currentTarget.style.background = "rgba(255,255,255,0.04)"}
                >
                  <div style={{ fontSize: 20, marginBottom: 5 }}>{t.icon}</div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.white, lineHeight: 1.3 }}>{t.label}</div>
                  <div style={{ fontSize: 10, color: C.dimmer, marginTop: 2, lineHeight: 1.3 }}>{t.desc}</div>
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* ── Right Content ─────────────────────────────────────────────── */}
        <div style={{ display: "flex", flexDirection: "column" }}>
          {/* Action tab bar */}
          <div style={{ background: "#e8f0fe", borderBottom: `1px solid ${C.accentLight}`, display: "flex", gap: 6, padding: "10px 20px", flexWrap: "wrap" }}>
            {ACTION_TABS.map(tab => {
              const active = activeAction === tab.id;
              return (
                <button key={tab.id} onClick={() => setActiveAction(tab.id)}
                  style={{
                    padding: "7px 16px", borderRadius: 8, border: active ? "none" : `1px solid ${C.accentLight}`,
                    background: active ? C.accent : C.cardBg,
                    color: active ? C.white : C.textSec,
                    fontWeight: 700, fontSize: 13, cursor: "pointer",
                    display: "flex", alignItems: "center", gap: 6,
                    fontFamily: "inherit", transition: "all .15s",
                    boxShadow: active ? "0 2px 8px rgba(29,110,245,0.3)" : "none",
                  }}>
                  <span>{tab.icon}</span>{tab.label}
                </button>
              );
            })}
          </div>

          {/* Content area */}
          <div style={{ padding: "24px 24px", flex: 1, overflowY: "auto" }}>
            {activeAction === "about"   && <AboutTab onGetStarted={() => { loadTemplate("enrollment"); setActiveAction("records"); }} />}
            {activeAction === "records" && <RecordsTab rows={rows} headers={headers} sources={new Set(allSheets.map(s => s.file)).size} />}
            {activeAction === "ai"      && <AITab rows={rows} headers={headers} />}
            {activeAction === "charts"  && <ChartsTab rows={rows} headers={headers} />}
            {activeAction === "mapper"  && <MapperTab allSheets={allSheets} />}
            {activeAction === "dupes"   && <DupesTab rows={rows} headers={headers} />}
            {activeAction === "schema"  && <SchemaTab rows={rows} headers={headers} />}
          </div>
        </div>
      </div>
    </div>
  );
}
