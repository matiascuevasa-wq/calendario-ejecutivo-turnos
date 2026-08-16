import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Calendar as CalendarIcon, Users, Shuffle, BarChart3, ShieldCheck, LogIn, LogOut,
  Plus, X, Trash2, Pencil, Check, ChevronLeft, ChevronRight, Copy, AlertTriangle,
  Building2, Eye, RefreshCw, Loader2, KeyRound, Upload, Download, FileSpreadsheet
} from "lucide-react";
import { db } from "./firebase";
import { doc, getDoc, setDoc } from "firebase/firestore";

/* ------------------------------------------------------------------ */
/* CONSTANTES                                                          */
/* ------------------------------------------------------------------ */

const MASTER_USER = "Excelencia Operacional";
const MASTER_PASS = "Excelencia OEMS";
const TARGET_SHIFTS = 2;
const START_HOUR = 7;
const END_HOUR = 22;
const ROW_H = 26;
const WEEKDAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const WEEKDAYS_SHORT = ["LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB", "DOM"];
const MONTHS = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];

const PALETTE = [
  "#2F6F5E", "#3E6C9E", "#8A5FB0", "#B0562F", "#4A7A2A", "#2F8A8A", "#9E3E6C", "#6C6C2F"
];

const DEFAULT_CATEGORIES = [
  { id: "cat-comite", name: "Reunión Agenda Comité Ejecutivo", color: "#3C8A3E" },
  { id: "cat-auditorias", name: "Auditorías", color: "#3563A6" },
  { id: "cat-publica", name: "Actividad Pública", color: "#D4A017" },
  { id: "cat-lomasbayas", name: "Actividad Lomas Bayas", color: "#D2691E" },
];

const STORAGE_KEYS = {
  gerencias: "gerencias",
  executives: "executives",
  accounts: "accounts",
  activities: "activities",
  shiftDraws: "shift_draws",
  auditLog: "audit_log",
  categories: "activity_categories",
};

/* ------------------------------------------------------------------ */
/* HELPERS DE FECHA                                                    */
/* ------------------------------------------------------------------ */

function pad(n) { return String(n).padStart(2, "0"); }
function fmtDate(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function parseDate(s) { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); }
function getMonday(d) {
  const date = new Date(d);
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  date.setHours(0, 0, 0, 0);
  return date;
}
function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function isoWeekNumber(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}
function timeToMinutes(t) { const [h, m] = t.split(":").map(Number); return h * 60 + m; }
function minutesToTime(min) { return `${pad(Math.floor(min / 60))}:${pad(min % 60)}`; }
function timeSlots() {
  const slots = [];
  for (let h = START_HOUR; h < END_HOUR; h++) { slots.push(`${pad(h)}:00`); slots.push(`${pad(h)}:30`); }
  return slots;
}
function endTimeOptions(start) {
  const startMin = timeToMinutes(start);
  const opts = [];
  for (let m = startMin + 30; m <= END_HOUR * 60; m += 30) opts.push(minutesToTime(m));
  return opts;
}
function weekendsOfYear(year) {
  const weekends = [];
  let d = new Date(year, 0, 1);
  while (d.getDay() !== 4) d = addDays(d, 1); // primer jueves
  while (d.getFullYear() === year) {
    weekends.push({ thu: fmtDate(d), fri: fmtDate(addDays(d, 1)), sat: fmtDate(addDays(d, 2)), sun: fmtDate(addDays(d, 3)) });
    d = addDays(d, 7);
  }
  return weekends;
}
function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }

/* ---- recurrencia de actividades (similar a Outlook) ---- */
const WEEKDAY_NAMES = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
const NTH_NAMES = { 1: "primer", 2: "segundo", 3: "tercer", 4: "cuarto", 5: "último" };

function nthWeekdayOfMonth(d) {
  const n = Math.ceil(d.getDate() / 7);
  const next = addDays(d, 7);
  if (next.getMonth() !== d.getMonth()) return 5; // es el último de ese día en el mes
  return n;
}
function seriesMatchesDate(series, dateStr) {
  const r = series.recurrence;
  if (!r) return false;
  const d = parseDate(dateStr);
  const start = parseDate(series.date);
  if (d < start) return false;
  if (r.until && dateStr > r.until) return false;
  if (r.freq === "weekly") return d.getDay() === r.weekday;
  if (r.freq === "monthly") {
    if (d.getDay() !== r.weekday) return false;
    const n = nthWeekdayOfMonth(d);
    if (r.nth === 5) return n === 5;
    return n === r.nth;
  }
  return false;
}
function describeRecurrence(recurrence) {
  if (!recurrence) return "No se repite";
  const day = WEEKDAY_NAMES[recurrence.weekday];
  if (recurrence.freq === "weekly") return `Cada ${day}`;
  if (recurrence.freq === "monthly") return `El ${NTH_NAMES[recurrence.nth]} ${day} de cada mes`;
  return "No se repite";
}
// Expande actividades (simples + series recurrentes con excepciones) para un set de fechas visibles
function expandOccurrences(activities, dateStrs) {
  const dateSet = new Set(dateStrs);
  const result = [];
  activities.forEach(act => {
    if (!act.recurrence) {
      if (dateSet.has(act.date)) result.push({ ...act, occurrenceDate: act.date, originalDate: act.date, seriesId: act.id, isRecurring: false });
      return;
    }
    const exceptions = act.exceptions || {};
    dateStrs.forEach(dateStr => {
      if (!seriesMatchesDate(act, dateStr)) return;
      const exc = exceptions[dateStr];
      if (exc && exc.cancelled) return;
      if (exc && exc.override && exc.override.date && exc.override.date !== dateStr) return; // se movió a otra fecha
      const base = exc && exc.override ? { ...act, ...exc.override } : act;
      result.push({ ...base, id: act.id, occurrenceDate: exc?.override?.date || dateStr, originalDate: dateStr, seriesId: act.id, isRecurring: true });
    });
    // ocurrencias reprogramadas que "entran" a este rango desde otra fecha original
    Object.entries(exceptions).forEach(([origDate, exc]) => {
      if (exc.override && exc.override.date && dateSet.has(exc.override.date) && exc.override.date !== origDate) {
        result.push({ ...act, ...exc.override, id: act.id, occurrenceDate: exc.override.date, originalDate: origDate, seriesId: act.id, isRecurring: true });
      }
    });
  });
  return result;
}


/* ------------------------------------------------------------------ */
/* STORAGE HELPERS                                                     */
/* ------------------------------------------------------------------ */

async function loadKey(key, fallback) {
  try {
    const snap = await getDoc(doc(db, "app_data", key));
    if (!snap.exists()) return fallback;
    return JSON.parse(snap.data().value);
  } catch (e) {
    console.error("Error cargando", key, e);
    return fallback;
  }
}
async function saveKey(key, value) {
  try {
    await setDoc(doc(db, "app_data", key), { value: JSON.stringify(value) });
    return true;
  } catch (e) {
    console.error("Error guardando", key, e);
    return false;
  }
}

function gerenciaColor(gerencia, gerencias) {
  const idx = gerencias.indexOf(gerencia);
  if (idx === -1) return "#5B6B76";
  return PALETTE[idx % PALETTE.length];
}

function categoryById(categories, id) {
  return categories.find(c => c.id === id) || null;
}

/* ------------------------------------------------------------------ */
/* APP                                                                  */
/* ------------------------------------------------------------------ */

export default function App() {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [gerencias, setGerencias] = useState([]);
  const [executives, setExecutives] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [activities, setActivities] = useState([]);
  const [shiftDraws, setShiftDraws] = useState({});
  const [auditLog, setAuditLog] = useState([]);
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES);

  const [session, setSession] = useState(null); // {username, role:'master'|'gerente', gerencia}
  const [activeTab, setActiveTab] = useState("calendario");
  const [currentWeekStart, setCurrentWeekStart] = useState(getMonday(new Date()));

  const [showLogin, setShowLogin] = useState(false);
  const [showActivityModal, setShowActivityModal] = useState(null); // {mode, activity, date}
  const [showEditChoice, setShowEditChoice] = useState(null); // {occurrence, action}
  const [showExecModal, setShowExecModal] = useState(null);
  const [showAccountModal, setShowAccountModal] = useState(null);
  const [showCategoryModal, setShowCategoryModal] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);

  const [sorteoYear, setSorteoYear] = useState(new Date().getFullYear());
  const [sorteoDraft, setSorteoDraft] = useState(null);

  const thisYear = new Date().getFullYear();

  /* ---------------- carga inicial ---------------- */
  useEffect(() => {
    (async () => {
      const [g, e, a, act, sd, log, cats] = await Promise.all([
        loadKey(STORAGE_KEYS.gerencias, []),
        loadKey(STORAGE_KEYS.executives, []),
        loadKey(STORAGE_KEYS.accounts, []),
        loadKey(STORAGE_KEYS.activities, []),
        loadKey(STORAGE_KEYS.shiftDraws, {}),
        loadKey(STORAGE_KEYS.auditLog, []),
        loadKey(STORAGE_KEYS.categories, DEFAULT_CATEGORIES),
      ]);
      setGerencias(g); setExecutives(e); setAccounts(a); setActivities(act); setShiftDraws(sd); setAuditLog(log);
      setCategories(cats && cats.length ? cats : DEFAULT_CATEGORIES);
      setLoading(false);
    })();
  }, []);

  const pushAudit = useCallback(async (action, detail) => {
    const entry = { ts: new Date().toISOString(), user: session ? session.username : "Anónimo", action, detail };
    setAuditLog(prev => {
      const next = [entry, ...prev].slice(0, 500);
      saveKey(STORAGE_KEYS.auditLog, next);
      return next;
    });
  }, [session]);

  const canEditGerencia = (gerencia) => session && (session.role === "master" || (session.role === "gerente" && session.gerencia === gerencia));
  const isStaff = !!session && session.role !== "invitado";

  function loginAsGuest() {
    setSession({ username: "Invitado", role: "invitado", gerencia: null });
    setShowLogin(false);
  }

  /* ---------------- login ---------------- */
  function tryLogin(username, password) {
    if (username.trim().toLowerCase() === MASTER_USER.toLowerCase() && password === MASTER_PASS) {
      setSession({ username: MASTER_USER, role: "master", gerencia: null });
      setShowLogin(false);
      return null;
    }
    const acc = accounts.find(a => a.username.toLowerCase() === username.trim().toLowerCase() && a.password === password);
    if (acc) {
      setSession({ username: acc.username, role: "gerente", gerencia: acc.gerencia });
      setShowLogin(false);
      return null;
    }
    return "Usuario o clave incorrectos.";
  }
  function logout() { setSession(null); setActiveTab("calendario"); }

  /* ---------------- gerencias ---------------- */
  async function addGerencia(name) {
    const clean = name.trim();
    if (!clean || gerencias.includes(clean)) return;
    const next = [...gerencias, clean];
    setGerencias(next); await saveKey(STORAGE_KEYS.gerencias, next);
    pushAudit("Gerencia creada", clean);
  }
  async function deleteGerencia(name) {
    const nextG = gerencias.filter(g => g !== name);
    const nextE = executives.filter(e => e.gerencia !== name);
    setGerencias(nextG); await saveKey(STORAGE_KEYS.gerencias, nextG);
    setExecutives(nextE); await saveKey(STORAGE_KEYS.executives, nextE);
    pushAudit("Gerencia eliminada", name);
    setConfirmDialog(null);
  }

  /* ---------------- categorías de actividad ---------------- */
  async function saveCategory(cat) {
    const isEdit = cat.id && categories.some(c => c.id === cat.id);
    const next = isEdit ? categories.map(c => c.id === cat.id ? cat : c) : [...categories, { ...cat, id: uid() }];
    setCategories(next); await saveKey(STORAGE_KEYS.categories, next);
    pushAudit(isEdit ? "Categoría editada" : "Categoría creada", cat.name);
    setShowCategoryModal(null);
  }
  async function deleteCategory(cat) {
    const next = categories.filter(c => c.id !== cat.id);
    setCategories(next); await saveKey(STORAGE_KEYS.categories, next);
    pushAudit("Categoría eliminada", cat.name);
    setConfirmDialog(null);
  }

  /* ---------------- ejecutivos ---------------- */
  async function saveExecutive(exec) {
    let next;
    if (exec.id) next = executives.map(e => e.id === exec.id ? exec : e);
    else next = [...executives, { ...exec, id: uid() }];
    setExecutives(next); await saveKey(STORAGE_KEYS.executives, next);
    pushAudit(exec.id ? "Ejecutivo editado" : "Ejecutivo agregado", `${exec.name} (${exec.gerencia})`);
    setShowExecModal(null);
  }
  async function deleteExecutive(exec) {
    const next = executives.filter(e => e.id !== exec.id);
    setExecutives(next); await saveKey(STORAGE_KEYS.executives, next);
    pushAudit("Ejecutivo eliminado", `${exec.name} (${exec.gerencia})`);
    setConfirmDialog(null);
  }

  async function bulkImportExecutives(rows) {
    // rows: [{name, gerencia}]
    let gList = [...gerencias];
    let newGerenciasCount = 0;
    const toAdd = [];
    let skipped = 0;
    rows.forEach(r => {
      const name = (r.name || "").trim();
      const gerencia = (r.gerencia || "").trim();
      if (!name || !gerencia) { skipped++; return; }
      if (!gList.includes(gerencia)) { gList.push(gerencia); newGerenciasCount++; }
      toAdd.push({ id: uid(), name, gerencia });
    });
    if (toAdd.length === 0) return { added: 0, newGerenciasCount: 0, skipped };
    const nextExecs = [...executives, ...toAdd];
    setExecutives(nextExecs); await saveKey(STORAGE_KEYS.executives, nextExecs);
    if (newGerenciasCount > 0) { setGerencias(gList); await saveKey(STORAGE_KEYS.gerencias, gList); }
    pushAudit("Importación desde Excel", `${toAdd.length} ejecutivo(s) agregado(s)${newGerenciasCount ? `, ${newGerenciasCount} gerencia(s) nueva(s)` : ""}`);
    return { added: toAdd.length, newGerenciasCount, skipped };
  }

  /* ---------------- actividades ---------------- */
  async function saveActivity(act) {
    let next;
    if (act.id) next = activities.map(a => a.id === act.id ? { ...a, ...act } : a);
    else next = [...activities, { ...act, id: uid(), createdBy: session.username, exceptions: {} }];
    setActivities(next); await saveKey(STORAGE_KEYS.activities, next);
    pushAudit(act.id ? "Actividad/serie editada" : "Actividad agregada", `${act.title} — ${act.date} ${act.start}-${act.end}${act.recurrence ? ` (${describeRecurrence(act.recurrence)})` : ""}`);
    setShowActivityModal(null);
  }
  async function deleteActivitySeries(act) {
    const next = activities.filter(a => a.id !== act.seriesId);
    setActivities(next); await saveKey(STORAGE_KEYS.activities, next);
    pushAudit("Serie eliminada", `${act.title} (todas las ocurrencias)`);
    setConfirmDialog(null);
  }
  async function deleteActivity(act) {
    if (act.isRecurring) return deleteActivitySeries(act);
    const next = activities.filter(a => a.id !== act.id);
    setActivities(next); await saveKey(STORAGE_KEYS.activities, next);
    pushAudit("Actividad eliminada", `${act.title} — ${act.date}`);
    setConfirmDialog(null);
  }
  async function saveOccurrenceOverride(occurrence, fields) {
    const next = activities.map(a => {
      if (a.id !== occurrence.seriesId) return a;
      const exceptions = { ...(a.exceptions || {}) };
      exceptions[occurrence.originalDate] = { override: fields };
      return { ...a, exceptions };
    });
    setActivities(next); await saveKey(STORAGE_KEYS.activities, next);
    pushAudit("Ocurrencia modificada", `${fields.title} — ${occurrence.originalDate} → ${fields.date}`);
    setShowActivityModal(null);
  }
  async function cancelOccurrence(occurrence) {
    const next = activities.map(a => {
      if (a.id !== occurrence.seriesId) return a;
      const exceptions = { ...(a.exceptions || {}) };
      exceptions[occurrence.originalDate] = { cancelled: true };
      return { ...a, exceptions };
    });
    setActivities(next); await saveKey(STORAGE_KEYS.activities, next);
    pushAudit("Ocurrencia eliminada", `${occurrence.title} — ${occurrence.originalDate}`);
    setConfirmDialog(null);
  }

  /* ---------------- cuentas ---------------- */
  async function saveAccount(acc) {
    let next;
    if (acc.editingUsername) next = accounts.map(a => a.username === acc.editingUsername ? { username: acc.username, password: acc.password, gerencia: acc.gerencia } : a);
    else next = [...accounts, { username: acc.username, password: acc.password, gerencia: acc.gerencia }];
    setAccounts(next); await saveKey(STORAGE_KEYS.accounts, next);
    pushAudit(acc.editingUsername ? "Cuenta editada" : "Cuenta creada", `${acc.username} (${acc.gerencia})`);
    setShowAccountModal(null);
  }
  async function deleteAccount(acc) {
    const next = accounts.filter(a => a.username !== acc.username);
    setAccounts(next); await saveKey(STORAGE_KEYS.accounts, next);
    pushAudit("Cuenta eliminada", acc.username);
    setConfirmDialog(null);
  }

  /* ---------------- sorteo ---------------- */
  function generateDraft(year) {
    const weekends = weekendsOfYear(year);
    if (executives.length === 0) { setSorteoDraft({ error: "No hay ejecutivos de turno cargados." }); return; }
    const counts = {}; executives.forEach(e => counts[e.id] = 0);
    let prevId = null;
    const draws = weekends.map(w => {
      let pool = executives.filter(e => e.id !== prevId);
      if (pool.length === 0) pool = executives;
      const minCount = Math.min(...pool.map(e => counts[e.id]));
      const candidates = pool.filter(e => counts[e.id] === minCount);
      const chosen = candidates[Math.floor(Math.random() * candidates.length)];
      counts[chosen.id]++;
      prevId = chosen.id;
      return { id: uid(), thu: w.thu, fri: w.fri, sat: w.sat, sun: w.sun, executiveId: chosen.id, status: "draft" };
    });
    setSorteoDraft({ year, draws });
  }
  function rerollRow(index) {
    setSorteoDraft(prev => {
      if (!prev) return prev;
      const draws = [...prev.draws];
      const others = executives.filter(e => e.id !== draws[index].executiveId);
      const pick = (others.length ? others : executives)[Math.floor(Math.random() * (others.length ? others.length : executives.length))];
      draws[index] = { ...draws[index], executiveId: pick.id };
      return { ...prev, draws };
    });
  }
  async function validateDraft() {
    if (!sorteoDraft || sorteoDraft.error) return;
    const validated = sorteoDraft.draws.map(d => ({ ...d, status: "validated" }));
    const next = { ...shiftDraws, [sorteoDraft.year]: validated };
    setShiftDraws(next); await saveKey(STORAGE_KEYS.shiftDraws, next);
    pushAudit("Sorteo validado", `Año ${sorteoDraft.year} — ${validated.length} fines de semana`);
    setSorteoDraft(null);
  }
  async function manualSetShift(year, weekend, executiveId) {
    const current = shiftDraws[year] || [];
    let nextYear;
    if (!executiveId) {
      nextYear = current.filter(d => d.thu !== weekend.thu);
    } else {
      const idx = current.findIndex(d => d.thu === weekend.thu);
      const entry = { id: idx >= 0 ? current[idx].id : uid(), thu: weekend.thu, fri: weekend.fri, sat: weekend.sat, sun: weekend.sun, executiveId, status: "validated" };
      if (idx >= 0) { nextYear = [...current]; nextYear[idx] = entry; }
      else nextYear = [...current, entry].sort((a, b) => a.thu.localeCompare(b.thu));
    }
    const next = { ...shiftDraws, [year]: nextYear };
    setShiftDraws(next); await saveKey(STORAGE_KEYS.shiftDraws, next);
    const execName = executiveId ? (executives.find(e => e.id === executiveId)?.name || executiveId) : "sin asignar";
    pushAudit("Turno asignado manualmente", `${weekend.thu} (${year}) → ${execName}`);
  }

  /* ---------------- semana visible ---------------- */
  const weekDays = useMemo(() => Array.from({ length: 7 }, (_, i) => addDays(currentWeekStart, i)), [currentWeekStart]);
  const yearMinMonday = getMonday(new Date(thisYear, 0, 1));
  const yearMaxMonday = getMonday(new Date(thisYear, 11, 31));
  const canGoPrev = currentWeekStart > yearMinMonday;
  const canGoNext = currentWeekStart < yearMaxMonday;

  const validatedDrawsThisYear = shiftDraws[thisYear] || [];
  function turnoForDate(dateStr) {
    for (const d of validatedDrawsThisYear) {
      if (d.thu === dateStr || d.fri === dateStr || d.sat === dateStr || d.sun === dateStr) return d;
    }
    return null;
  }

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#F5F6F4]">
        <div className="flex flex-col items-center gap-3 text-[#5B6B76]">
          <Loader2 className="animate-spin" size={28} />
          <span className="text-sm font-medium">Cargando calendario de turnos…</span>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#F5F6F4] text-[#1B2733] font-sans" style={{ fontFamily: "'Inter', ui-sans-serif, system-ui" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        .font-display { font-family: 'Space Grotesk', ui-sans-serif, system-ui; }
        .font-mono2 { font-family: 'IBM Plex Mono', ui-monospace, monospace; }
      `}</style>

      <Header
        session={session}
        onLogin={() => setShowLogin(true)}
        onLogout={logout}
      />

      <nav className="sticky top-[60px] z-30 bg-[#1B2733] border-t border-[#2A3946] px-4 md:px-6 overflow-x-auto">
        <div className="flex gap-1 max-w-[1400px] mx-auto">
          <TabButton icon={<CalendarIcon size={15} />} label="Calendario" active={activeTab === "calendario"} onClick={() => setActiveTab("calendario")} />
          <TabButton icon={<Users size={15} />} label="Ejecutivos de turno" active={activeTab === "ejecutivos"} onClick={() => setActiveTab("ejecutivos")} />
          <TabButton icon={<Shuffle size={15} />} label="Sorteo de turnos" active={activeTab === "sorteo"} onClick={() => setActiveTab("sorteo")} />
          <TabButton icon={<BarChart3 size={15} />} label="Estadísticas" active={activeTab === "estadisticas"} onClick={() => setActiveTab("estadisticas")} />
          {isStaff && <TabButton icon={<ShieldCheck size={15} />} label="Administración" active={activeTab === "admin"} onClick={() => setActiveTab("admin")} />}
        </div>
      </nav>

      <main className="max-w-[1400px] mx-auto px-4 md:px-6 py-6 pb-16">
        {activeTab === "calendario" && (
          <CalendarView
            weekDays={weekDays}
            currentWeekStart={currentWeekStart}
            setCurrentWeekStart={setCurrentWeekStart}
            canGoPrev={canGoPrev} canGoNext={canGoNext}
            yearMinMonday={yearMinMonday} yearMaxMonday={yearMaxMonday}
            activities={activities}
            gerencias={gerencias}
            categories={categories}
            session={session}
            isStaff={isStaff}
            turnoForDate={turnoForDate}
            executives={executives}
            canEditGerencia={canEditGerencia}
            onAddActivity={(date, start) => setShowActivityModal({ mode: "add", date, start })}
            onEditActivity={(occ) => {
              if (occ.isRecurring) setShowEditChoice({ occurrence: occ, action: "edit" });
              else setShowActivityModal({ mode: "edit", activity: occ });
            }}
            onDeleteActivity={(occ) => {
              if (occ.isRecurring) setShowEditChoice({ occurrence: occ, action: "delete" });
              else setConfirmDialog({
                title: "Eliminar actividad",
                message: `¿Eliminar "${occ.title}" del ${occ.date}?`,
                onConfirm: () => deleteActivity(occ)
              });
            }}
            thisYear={thisYear}
          />
        )}

        {activeTab === "ejecutivos" && (
          <ExecutivosView
            executives={executives} gerencias={gerencias} session={session} isStaff={isStaff}
            canEditGerencia={canEditGerencia}
            onBulkImport={bulkImportExecutives}
            onAdd={() => setShowExecModal({ mode: "add" })}
            onEdit={(e) => setShowExecModal({ mode: "edit", executive: e })}
            onDelete={(e) => setConfirmDialog({
              title: "Eliminar ejecutivo",
              message: `¿Eliminar a ${e.name} del pool de ejecutivos de turno?`,
              onConfirm: () => deleteExecutive(e)
            })}
            onAddGerencia={addGerencia}
            onDeleteGerencia={(g) => {
              const count = executives.filter(e => e.gerencia === g).length;
              setConfirmDialog({
                title: "Eliminar gerencia",
                message: count > 0
                  ? `¿Eliminar "${g}"? Se eliminarán también los ${count} ejecutivo(s) de turno asociados a esta gerencia. Las actividades ya cargadas en el calendario con esta gerencia no se borran, pero quedarán con una gerencia que ya no existe.`
                  : `¿Eliminar la gerencia "${g}"?`,
                onConfirm: () => deleteGerencia(g)
              });
            }}
          />
        )}

        {activeTab === "sorteo" && (
          <SorteoView
            session={session}
            executives={executives}
            gerencias={gerencias}
            sorteoYear={sorteoYear} setSorteoYear={setSorteoYear}
            sorteoDraft={sorteoDraft}
            shiftDraws={shiftDraws}
            onGenerate={() => generateDraft(sorteoYear)}
            onReroll={rerollRow}
            onValidate={() => setConfirmDialog({
              title: "Validar sorteo",
              message: `Esto agregará ${sorteoDraft?.draws?.length || 0} turnos de fin de semana al calendario del año ${sorteoDraft?.year}. Esta acción reemplaza cualquier sorteo previo validado para ese año.`,
              onConfirm: validateDraft
            })}
            onDiscard={() => setSorteoDraft(null)}
            onManualAssign={manualSetShift}
          />
        )}

        {activeTab === "estadisticas" && (
          <EstadisticasView
            executives={executives} gerencias={gerencias} shiftDraws={shiftDraws} thisYear={thisYear}
          />
        )}

        {activeTab === "admin" && isStaff && (
          <AdminView
            session={session}
            accounts={accounts} gerencias={gerencias} auditLog={auditLog}
            categories={categories}
            onAddAccount={() => setShowAccountModal({ mode: "add" })}
            onEditAccount={(a) => setShowAccountModal({ mode: "edit", account: a })}
            onDeleteAccount={(a) => setConfirmDialog({
              title: "Eliminar cuenta",
              message: `¿Eliminar el acceso de ${a.username}?`,
              onConfirm: () => deleteAccount(a)
            })}
            onAddCategory={() => setShowCategoryModal({ mode: "add" })}
            onEditCategory={(c) => setShowCategoryModal({ mode: "edit", category: c })}
            onDeleteCategory={(c) => setConfirmDialog({
              title: "Eliminar categoría",
              message: `¿Eliminar la categoría "${c.name}"? Las actividades ya creadas con esta categoría quedarán sin clasificación.`,
              onConfirm: () => deleteCategory(c)
            })}
          />
        )}
      </main>

      {showLogin && <LoginModal accounts={accounts} onClose={() => setShowLogin(false)} onSubmit={tryLogin} onGuestLogin={loginAsGuest} />}

      {showActivityModal && (
        <ActivityModal
          data={showActivityModal}
          gerencias={gerencias}
          categories={categories}
          session={session}
          onClose={() => setShowActivityModal(null)}
          onSave={saveActivity}
          onSaveOccurrence={saveOccurrenceOverride}
        />
      )}

      {showEditChoice && (
        <EditChoiceModal
          occurrence={showEditChoice.occurrence}
          action={showEditChoice.action}
          onClose={() => setShowEditChoice(null)}
          onChooseOccurrence={() => {
            const occ = showEditChoice.occurrence;
            const action = showEditChoice.action;
            setShowEditChoice(null);
            if (action === "edit") {
              setShowActivityModal({ mode: "edit-occurrence", occurrence: occ });
            } else {
              setConfirmDialog({
                title: "Eliminar esta actividad",
                message: `¿Eliminar solo la actividad del ${occ.occurrenceDate}? Las demás fechas de la serie no se ven afectadas.`,
                onConfirm: () => cancelOccurrence(occ)
              });
            }
          }}
          onChooseSeries={() => {
            const occ = showEditChoice.occurrence;
            const action = showEditChoice.action;
            const seriesRoot = activities.find(a => a.id === occ.seriesId);
            setShowEditChoice(null);
            if (action === "edit") {
              setShowActivityModal({ mode: "edit", activity: seriesRoot });
            } else {
              setConfirmDialog({
                title: "Eliminar toda la serie",
                message: `¿Eliminar "${occ.title}" y TODAS sus repeticiones (${describeRecurrence(occ.recurrence)})? Esta acción no se puede deshacer.`,
                onConfirm: () => deleteActivitySeries(occ)
              });
            }
          }}
        />
      )}

      {showCategoryModal && (
        <CategoryModal
          data={showCategoryModal}
          onClose={() => setShowCategoryModal(null)}
          onSave={saveCategory}
        />
      )}

      {showExecModal && (
        <ExecutivoModal
          data={showExecModal}
          gerencias={gerencias}
          session={session}
          onClose={() => setShowExecModal(null)}
          onSave={saveExecutive}
          onAddGerencia={addGerencia}
        />
      )}

      {showAccountModal && (
        <AccountModal
          data={showAccountModal}
          gerencias={gerencias}
          onClose={() => setShowAccountModal(null)}
          onSave={saveAccount}
        />
      )}

      {confirmDialog && (
        <ConfirmModal {...confirmDialog} onCancel={() => setConfirmDialog(null)} />
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* HEADER + TABS                                                       */
/* ------------------------------------------------------------------ */

function Header({ session, onLogin, onLogout }) {
  return (
    <header className="sticky top-0 z-40 h-[60px] bg-[#1B2733] flex items-center justify-between px-4 md:px-6 border-b border-[#2A3946]">
      <div className="flex items-center gap-2.5">
        <div className="w-8 h-8 rounded-md bg-[#E8A33D] flex items-center justify-center shrink-0">
          <CalendarIcon size={17} className="text-[#1B2733]" strokeWidth={2.5} />
        </div>
        <div className="leading-tight">
          <div className="font-display font-semibold text-white text-[15px] tracking-tight">Calendario de Turnos</div>
          <div className="text-[11px] text-[#8CA0AC] font-mono2">Excelencia Operacional · OEMS</div>
        </div>
      </div>
      <div>
        {session ? (
          <div className="flex items-center gap-2.5">
            <div className="hidden sm:flex flex-col items-end leading-tight">
              <span className="text-white text-[13px] font-medium">{session.username}</span>
              <span className="text-[11px] text-[#8CA0AC]">{session.role === "master" ? "Acceso maestro" : session.role === "gerente" ? session.gerencia : "Solo lectura"}</span>
            </div>
            <span className={`w-2 h-2 rounded-full ${session.role === "master" ? "bg-[#E8A33D]" : session.role === "gerente" ? "bg-[#2F6F5E]" : "bg-[#9AA8AF]"}`} />
            <button onClick={onLogout} className="flex items-center gap-1.5 text-[13px] text-[#C7D2D8] hover:text-white bg-white/5 hover:bg-white/10 px-3 py-1.5 rounded-md transition-colors">
              <LogOut size={14} /> <span className="hidden sm:inline">Salir</span>
            </button>
          </div>
        ) : (
          <button onClick={onLogin} className="flex items-center gap-1.5 text-[13px] font-medium text-[#1B2733] bg-[#E8A33D] hover:bg-[#D6922E] px-3.5 py-1.5 rounded-md transition-colors">
            <LogIn size={14} /> Iniciar sesión
          </button>
        )}
      </div>
    </header>
  );
}

function TabButton({ icon, label, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className={`flex items-center gap-1.5 px-3.5 py-2.5 text-[13px] font-medium whitespace-nowrap border-b-2 transition-colors ${
        active ? "border-[#E8A33D] text-white" : "border-transparent text-[#8CA0AC] hover:text-[#C7D2D8]"
      }`}
    >
      {icon} {label}
    </button>
  );
}

/* ------------------------------------------------------------------ */
/* VISTA CALENDARIO                                                     */
/* ------------------------------------------------------------------ */

function CalendarView({
  weekDays, currentWeekStart, setCurrentWeekStart, canGoPrev, canGoNext, yearMinMonday, yearMaxMonday,
  activities, gerencias, categories, session, isStaff, turnoForDate, executives, canEditGerencia, onAddActivity, onEditActivity, onDeleteActivity, thisYear
}) {
  const weekDateStrs = useMemo(() => weekDays.map(fmtDate), [weekDays]);
  const occurrences = useMemo(() => expandOccurrences(activities, weekDateStrs), [activities, weekDateStrs]);
  const slots = timeSlots();
  const weekNum = isoWeekNumber(weekDays[0]);
  const monday = weekDays[0], sunday = weekDays[6];
  const rangeLabel = monday.getMonth() === sunday.getMonth()
    ? `${monday.getDate()} – ${sunday.getDate()} de ${MONTHS[monday.getMonth()]} ${monday.getFullYear()}`
    : `${monday.getDate()} de ${MONTHS[monday.getMonth()]} – ${sunday.getDate()} de ${MONTHS[sunday.getMonth()]} ${sunday.getFullYear()}`;

  function goto(delta) {
    let next = addDays(currentWeekStart, delta * 7);
    if (next < yearMinMonday) next = yearMinMonday;
    if (next > yearMaxMonday) next = yearMaxMonday;
    setCurrentWeekStart(next);
  }
  function gotoToday() { setCurrentWeekStart(getMonday(new Date())); }
  function onPickWeek(e) {
    const d = parseDate(e.target.value);
    setCurrentWeekStart(getMonday(d));
  }

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div>
          <h1 className="font-display text-xl font-semibold text-[#1B2733]">Semana {weekNum} · {rangeLabel}</h1>
          <p className="text-[12px] text-[#5B6B76] mt-0.5">Bloques de 30 minutos, de {START_HOUR}:00 a {END_HOUR}:00. Haz clic en un bloque para agregar una actividad.</p>
        </div>
        <div className="flex items-center gap-2">
          <input type="date" onChange={onPickWeek} className="text-[12px] border border-[#D8DEE1] rounded-md px-2 py-1.5 bg-white text-[#1B2733] font-mono2" />
          <button onClick={gotoToday} className="text-[12px] font-medium px-3 py-1.5 rounded-md border border-[#D8DEE1] bg-white hover:bg-[#F0F2F1] text-[#1B2733]">Hoy</button>
          <div className="flex items-center rounded-md border border-[#D8DEE1] bg-white overflow-hidden">
            <button disabled={!canGoPrev} onClick={() => goto(-1)} className="p-1.5 disabled:opacity-30 hover:bg-[#F0F2F1]"><ChevronLeft size={16} /></button>
            <button disabled={!canGoNext} onClick={() => goto(1)} className="p-1.5 disabled:opacity-30 hover:bg-[#F0F2F1] border-l border-[#D8DEE1]"><ChevronRight size={16} /></button>
          </div>
        </div>
      </div>

      {!isStaff && (
        <div className="mb-3 text-[12px] text-[#8A6A2A] bg-[#FDF3E1] border border-[#F0D9A8] rounded-md px-3 py-2 flex items-center gap-2">
          <Eye size={14} /> {session ? "Sesión de invitado: solo puedes visualizar el calendario." : "Estás viendo el calendario en modo lectura. Inicia sesión para agregar o editar actividades."}
        </div>
      )}

      {categories.length > 0 && (
        <div className="mb-3 flex flex-wrap items-center gap-x-4 gap-y-1.5">
          {categories.map(c => (
            <div key={c.id} className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-sm shrink-0" style={{ backgroundColor: c.color }} />
              <span className="text-[11.5px] text-[#5B6B76]">{c.name}</span>
            </div>
          ))}
        </div>
      )}

      <div className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden">
        {/* cabecera dias */}
        <div className="grid" style={{ gridTemplateColumns: "56px repeat(7, 1fr)" }}>
          <div className="border-b border-r border-[#E3E7E5]" />
          {weekDays.map((d, i) => {
            const dateStr = fmtDate(d);
            const isToday = fmtDate(new Date()) === dateStr;
            const draw = turnoForDate(dateStr);
            const execTurno = draw ? executives.find(e => e.id === draw.executiveId) : null;
            return (
              <div key={i} className={`border-b border-r border-[#E3E7E5] last:border-r-0 px-2 py-2 ${isToday ? "bg-[#FDF6E9]" : ""}`}>
                <div className="flex items-center justify-between">
                  <div>
                    <div className="text-[10px] font-semibold text-[#8CA0AC] tracking-wide">{WEEKDAYS_SHORT[i]}</div>
                    <div className={`text-[15px] font-display font-semibold ${isToday ? "text-[#B0562F]" : "text-[#1B2733]"}`}>{d.getDate()}</div>
                  </div>
                  {isStaff && (
                    <button onClick={() => onAddActivity(dateStr, "09:00")} className="p-1 rounded hover:bg-[#EFF1EF] text-[#8CA0AC] hover:text-[#1B2733]">
                      <Plus size={13} />
                    </button>
                  )}
                </div>
                {execTurno && (
                  <div className="mt-1.5 flex items-center gap-1 bg-[#FDF3E1] border border-[#F0D9A8] rounded px-1.5 py-1" title={`Turno de fin de semana — ${execTurno.gerencia}`}>
                    <span className="w-1.5 h-1.5 rounded-full bg-[#E8A33D] shrink-0" />
                    <span className="text-[10px] font-medium text-[#8A6A2A] truncate">Turno: {execTurno.name}</span>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* grilla horaria */}
        <div className="grid" style={{ gridTemplateColumns: "56px repeat(7, 1fr)" }}>
          <div>
            {slots.map((t, i) => (
              <div key={t} className="border-r border-[#E3E7E5] text-right pr-1.5 text-[10px] font-mono2 text-[#9AA8AF]" style={{ height: ROW_H }}>
                {t.endsWith(":00") ? <span className="relative -top-1.5">{t}</span> : null}
              </div>
            ))}
          </div>
          {weekDays.map((d, colIdx) => {
            const dateStr = fmtDate(d);
            const dayActs = occurrences.filter(a => a.occurrenceDate === dateStr);
            return (
              <div key={colIdx} className="relative border-r border-[#E3E7E5] last:border-r-0" style={{ height: slots.length * ROW_H }}>
                {slots.map((t, i) => (
                  <div
                    key={t}
                    onClick={() => isStaff && onAddActivity(dateStr, t)}
                    className={`border-b border-[#EEF1F0] ${isStaff ? "hover:bg-[#F5F8F1] cursor-pointer" : ""} ${t.endsWith(":00") ? "border-t border-t-[#E3E7E5]" : ""}`}
                    style={{ height: ROW_H }}
                  />
                ))}
                {dayActs.map(act => {
                  const startMin = timeToMinutes(act.start) - START_HOUR * 60;
                  const endMin = timeToMinutes(act.end) - START_HOUR * 60;
                  const top = (startMin / 30) * ROW_H;
                  const height = Math.max(((endMin - startMin) / 30) * ROW_H - 2, ROW_H - 4);
                  const cat = categoryById(categories, act.categoryId);
                  const color = cat ? cat.color : "#8CA0AC";
                  const editable = canEditGerencia(act.gerencia);
                  return (
                    <div
                      key={act.seriesId + "-" + act.occurrenceDate}
                      onClick={(e) => { e.stopPropagation(); if (editable) onEditActivity(act); }}
                      className={`absolute left-0.5 right-0.5 rounded-[4px] px-1.5 py-0.5 overflow-hidden ${editable ? "cursor-pointer" : ""}`}
                      style={{ top, height, backgroundColor: color + "1F", borderLeft: `3px solid ${color}` }}
                      title={`${act.title} · ${act.start}-${act.end} · ${act.gerencia}${cat ? " · " + cat.name : ""}${act.isRecurring ? " · " + describeRecurrence(act.recurrence) : ""}`}
                    >
                      <div className="text-[10.5px] font-semibold leading-tight truncate flex items-center gap-1" style={{ color }}>
                        {act.isRecurring && <RefreshCw size={9} className="shrink-0" />}
                        {act.title}
                      </div>
                      <div className="text-[9px] text-[#5B6B76] leading-tight font-mono2 truncate">{act.start}–{act.end} · {act.gerencia}</div>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* VISTA EJECUTIVOS                                                     */
/* ------------------------------------------------------------------ */

function ExecutivosView({ executives, gerencias, session, isStaff, canEditGerencia, onAdd, onEdit, onDelete, onAddGerencia, onDeleteGerencia, onBulkImport }) {
  const [newGerencia, setNewGerencia] = useState("");
  const [importMsg, setImportMsg] = useState(null);
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const isMaster = session && session.role === "master";

  const grouped = useMemo(() => {
    const map = {};
    gerencias.forEach(g => map[g] = []);
    executives.forEach(e => { if (!map[e.gerencia]) map[e.gerencia] = []; map[e.gerencia].push(e); });
    return map;
  }, [executives, gerencias]);

  function downloadTemplate() {
    const wsData = [
      ["Nombre", "Gerencia"],
      ["Juan Pérez", gerencias[0] || "Operaciones"],
      ["María González", gerencias[1] || "Mantenimiento"],
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{ wch: 28 }, { wch: 24 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ejecutivos");
    XLSX.writeFile(wb, "plantilla_ejecutivos_de_turno.xlsx");
  }

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImporting(true);
    setImportMsg(null);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const rows = json.map(r => {
          const keys = Object.keys(r);
          const nameKey = keys.find(k => k.trim().toLowerCase().startsWith("nombre"));
          const gerKey = keys.find(k => k.trim().toLowerCase().startsWith("gerencia"));
          return { name: nameKey ? String(r[nameKey]) : "", gerencia: gerKey ? String(r[gerKey]) : "" };
        });
        const result = await onBulkImport(rows);
        if (result.added === 0) {
          setImportMsg({ type: "error", text: "No se encontraron filas válidas. Verifica que el Excel tenga columnas 'Nombre' y 'Gerencia'." });
        } else {
          setImportMsg({
            type: "ok",
            text: `Se importaron ${result.added} ejecutivo(s)${result.newGerenciasCount ? ` y se crearon ${result.newGerenciasCount} gerencia(s) nueva(s)` : ""}.${result.skipped ? ` ${result.skipped} fila(s) se omitieron por datos incompletos.` : ""}`
          });
        }
      } catch (err) {
        setImportMsg({ type: "error", text: "No se pudo leer el archivo. Asegúrate de que sea un Excel (.xlsx) válido." });
      } finally {
        setImporting(false);
        if (fileInputRef.current) fileInputRef.current.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div>
          <h1 className="font-display text-xl font-semibold text-[#1B2733]">Ejecutivos de turno</h1>
          <p className="text-[12px] text-[#5B6B76] mt-0.5">Pool de ejecutivos disponibles para el sorteo de turnos de fin de semana, por gerencia.</p>
        </div>
        {isMaster && (
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={downloadTemplate} className="flex items-center gap-1.5 text-[12px] font-medium text-[#1B2733] bg-white border border-[#D8DEE1] hover:bg-[#F0F2F1] px-3 py-2 rounded-md">
              <Download size={14} /> Plantilla Excel
            </button>
            <button onClick={() => fileInputRef.current?.click()} disabled={importing} className="flex items-center gap-1.5 text-[12px] font-medium text-[#1B2733] bg-white border border-[#D8DEE1] hover:bg-[#F0F2F1] disabled:opacity-50 px-3 py-2 rounded-md">
              {importing ? <Loader2 size={14} className="animate-spin" /> : <Upload size={14} />} Importar Excel
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFile} className="hidden" />
            <button onClick={onAdd} className="flex items-center gap-1.5 text-[13px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] px-3.5 py-2 rounded-md">
              <Plus size={15} /> Agregar ejecutivo
            </button>
          </div>
        )}
      </div>

      {isMaster && (
        <div className="mb-4 flex items-start gap-2 bg-[#EFF6F2] border border-[#CFE3D8] rounded-md px-3 py-2.5 text-[12px] text-[#2F6F5E]">
          <FileSpreadsheet size={15} className="shrink-0 mt-0.5" />
          <span>Descarga la plantilla, complétala con columnas <strong>Nombre</strong> y <strong>Gerencia</strong> (una fila por ejecutivo), y luego súbela con "Importar Excel". Si escribes una gerencia que no existe, se crea automáticamente.</span>
        </div>
      )}

      {importMsg && (
        <div className={`mb-4 text-[12px] rounded-md px-3 py-2 flex items-center gap-2 ${importMsg.type === "ok" ? "text-[#2F6F5E] bg-[#EAF3EF] border border-[#CFE3D8]" : "text-[#C1443B] bg-[#FCEDEB] border border-[#F3C9C3]"}`}>
          {importMsg.type === "ok" ? <Check size={14} /> : <AlertTriangle size={14} />} {importMsg.text}
        </div>
      )}

      {isMaster && (
        <div className="mb-4 flex items-center gap-2 bg-white border border-[#E3E7E5] rounded-lg px-3 py-2.5 w-fit">
          <Building2 size={15} className="text-[#8CA0AC]" />
          <input
            value={newGerencia} onChange={e => setNewGerencia(e.target.value)}
            placeholder="Nombre de nueva gerencia"
            className="text-[13px] border-none outline-none w-56"
          />
          <button
            onClick={() => { onAddGerencia(newGerencia); setNewGerencia(""); }}
            className="text-[12px] font-medium text-[#2F6F5E] hover:text-[#1F4E42] px-2 py-1"
          >Agregar gerencia</button>
        </div>
      )}

      {gerencias.length === 0 && (
        <div className="text-[13px] text-[#5B6B76] bg-white border border-dashed border-[#D8DEE1] rounded-lg p-6 text-center">
          Aún no hay gerencias registradas. {isMaster ? "Agrega la primera gerencia arriba, o impórtalas junto con los ejecutivos desde Excel." : "Excelencia Operacional aún no ha configurado las gerencias."}
        </div>
      )}

      <div className="grid md:grid-cols-2 gap-4">
        {gerencias.map(g => {
          const list = grouped[g] || [];
          const editable = isMaster;
          const color = gerenciaColor(g, gerencias);
          return (
            <div key={g} className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden">
              <div className="px-4 py-2.5 border-b border-[#E3E7E5] flex items-center justify-between" style={{ backgroundColor: color + "12" }}>
                <span className="font-medium text-[13px]" style={{ color }}>{g}</span>
                <div className="flex items-center gap-2">
                  <span className="text-[11px] font-mono2 text-[#5B6B76]">{list.length} ejecutivo{list.length !== 1 ? "s" : ""}</span>
                  {isMaster && (
                    <button onClick={() => onDeleteGerencia(g)} title="Eliminar gerencia" className="p-1 rounded hover:bg-white/60 text-[#5B6B76] hover:text-[#C1443B]">
                      <Trash2 size={13} />
                    </button>
                  )}
                </div>
              </div>
              <div className="divide-y divide-[#F0F2F1]">
                {list.length === 0 && <div className="px-4 py-3 text-[12px] text-[#9AA8AF]">Sin ejecutivos cargados.</div>}
                {list.map(e => (
                  <div key={e.id} className="px-4 py-2.5 flex items-center justify-between">
                    <span className="text-[13px] text-[#1B2733]">{e.name}</span>
                    {editable && (
                      <div className="flex items-center gap-1">
                        <button onClick={() => onEdit(e)} className="p-1.5 rounded hover:bg-[#F0F2F1] text-[#5B6B76]"><Pencil size={13} /></button>
                        <button onClick={() => onDelete(e)} className="p-1.5 rounded hover:bg-[#FCEDEB] text-[#C1443B]"><Trash2 size={13} /></button>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* VISTA SORTEO                                                         */
/* ------------------------------------------------------------------ */

function SorteoView({ session, executives, gerencias, sorteoYear, setSorteoYear, sorteoDraft, shiftDraws, onGenerate, onReroll, onValidate, onDiscard, onManualAssign }) {
  const isMaster = session && session.role === "master";
  const validated = shiftDraws[sorteoYear] || [];
  const thisYr = new Date().getFullYear();
  const yearsAvailable = Array.from({ length: 7 }, (_, i) => thisYr - 1 + i);

  function nameFor(id) { return executives.find(e => e.id === id); }

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div>
          <h1 className="font-display text-xl font-semibold text-[#1B2733]">Sorteo de turnos de fin de semana</h1>
          <p className="text-[12px] text-[#5B6B76] mt-0.5">Asigna aleatoriamente un ejecutivo de turno a cada bloque jueves–domingo del año.</p>
        </div>
        {isMaster && (
          <div className="flex items-center gap-2">
            <select value={sorteoYear} onChange={e => setSorteoYear(Number(e.target.value))} className="text-[13px] border border-[#D8DEE1] rounded-md px-2.5 py-2 bg-white">
              {yearsAvailable.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
            <button onClick={onGenerate} className="flex items-center gap-1.5 text-[13px] font-medium text-white bg-[#E8A33D] hover:bg-[#D6922E] px-3.5 py-2 rounded-md">
              <Shuffle size={14} /> Generar sorteo
            </button>
          </div>
        )}
      </div>

      {!isMaster && (
        <div className="mb-4 text-[12px] text-[#8A6A2A] bg-[#FDF3E1] border border-[#F0D9A8] rounded-md px-3 py-2 flex items-center gap-2">
          <Eye size={14} /> Solo la sesión de Excelencia Operacional puede generar y validar el sorteo. Aquí puedes ver los turnos ya validados.
        </div>
      )}

      {sorteoDraft && sorteoDraft.error && (
        <div className="mb-4 text-[13px] text-[#C1443B] bg-[#FCEDEB] border border-[#F3C9C3] rounded-md px-3 py-2 flex items-center gap-2">
          <AlertTriangle size={14} /> {sorteoDraft.error}
        </div>
      )}

      {sorteoDraft && !sorteoDraft.error && (
        <div className="mb-6 bg-white border-2 border-[#E8A33D] rounded-lg overflow-hidden">
          <div className="px-4 py-2.5 bg-[#FDF3E1] flex items-center justify-between">
            <span className="text-[13px] font-semibold text-[#8A6A2A]">Borrador del sorteo {sorteoDraft.year} — sin validar</span>
            <div className="flex gap-2">
              <button onClick={onDiscard} className="text-[12px] font-medium text-[#8A6A2A] hover:underline px-2 py-1">Descartar</button>
              <button onClick={onValidate} className="flex items-center gap-1.5 text-[12px] font-semibold text-white bg-[#2F6F5E] hover:bg-[#285F51] px-3 py-1.5 rounded-md">
                <Check size={13} /> Validar y agregar al calendario
              </button>
            </div>
          </div>
          <DrawTable draws={sorteoDraft.draws} executives={executives} gerencias={gerencias} onReroll={onReroll} editable />
        </div>
      )}

      <div className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden">
        <div className="px-4 py-2.5 border-b border-[#E3E7E5] flex items-center justify-between">
          <span className="text-[13px] font-semibold text-[#1B2733]">Turnos de fin de semana — {sorteoYear}</span>
          <span className="text-[11px] font-mono2 text-[#5B6B76]">{validated.length}/{weekendsOfYear(sorteoYear).length} asignados</span>
        </div>
        {isMaster && (
          <div className="px-4 py-2 text-[11px] text-[#5B6B76] border-b border-[#F0F2F1] bg-[#FAFBFA]">
            Puedes asignar cada fin de semana manualmente desde el selector de la tabla, o usar "Generar sorteo" arriba para asignarlos aleatoriamente.
          </div>
        )}
        <ValidatedShiftsTable year={sorteoYear} validated={validated} executives={executives} gerencias={gerencias} isMaster={isMaster} onManualAssign={onManualAssign} />
      </div>
    </div>
  );
}

function ValidatedShiftsTable({ year, validated, executives, gerencias, isMaster, onManualAssign }) {
  const weekends = weekendsOfYear(year);
  const byThu = {};
  validated.forEach(d => { byThu[d.thu] = d; });
  return (
    <div className="overflow-x-auto max-h-[520px] overflow-y-auto">
      <table className="w-full text-[13px]">
        <thead>
          <tr className="text-left text-[11px] text-[#8CA0AC] border-b border-[#EEF1F0] sticky top-0 bg-white">
            <th className="px-4 py-2 font-medium">Fin de semana</th>
            <th className="px-4 py-2 font-medium">Ejecutivo asignado</th>
            <th className="px-4 py-2 font-medium">Gerencia</th>
          </tr>
        </thead>
        <tbody>
          {weekends.map(w => {
            const d = byThu[w.thu];
            const exec = d ? executives.find(e => e.id === d.executiveId) : null;
            const color = exec ? gerenciaColor(exec.gerencia, gerencias) : "#9AA8AF";
            const thu = parseDate(w.thu);
            return (
              <tr key={w.thu} className="border-b border-[#F5F6F5] last:border-b-0">
                <td className="px-4 py-2 font-mono2 text-[12px] text-[#5B6B76] whitespace-nowrap">{thu.getDate()} {MONTHS[thu.getMonth()].slice(0, 3)} – {parseDate(w.sun).getDate()} {MONTHS[parseDate(w.sun).getMonth()].slice(0, 3)}</td>
                <td className="px-4 py-2">
                  {isMaster ? (
                    <select
                      value={d ? d.executiveId : ""}
                      onChange={e => onManualAssign(year, w, e.target.value || null)}
                      className="text-[12px] border border-[#D8DEE1] rounded-md px-2 py-1.5 bg-white w-full max-w-[280px]"
                    >
                      <option value="">Sin asignar</option>
                      {executives.map(ex => <option key={ex.id} value={ex.id}>{ex.name} — {ex.gerencia}</option>)}
                    </select>
                  ) : (
                    <span className="font-medium">{exec ? exec.name : "Sin asignar"}</span>
                  )}
                </td>
                <td className="px-4 py-2">
                  {exec ? <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ backgroundColor: color + "1A", color }}>{exec.gerencia}</span> : <span className="text-[11px] text-[#9AA8AF]">—</span>}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function DrawTable({ draws, executives, gerencias, onReroll, editable }) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-[13px]">
        <thead>
          <tr className="text-left text-[11px] text-[#8CA0AC] border-b border-[#EEF1F0]">
            <th className="px-4 py-2 font-medium">Fin de semana</th>
            <th className="px-4 py-2 font-medium">Ejecutivo</th>
            <th className="px-4 py-2 font-medium">Gerencia</th>
            {editable && <th className="px-4 py-2 font-medium"></th>}
          </tr>
        </thead>
        <tbody>
          {draws.map((d, i) => {
            const exec = executives.find(e => e.id === d.executiveId);
            const color = exec ? gerenciaColor(exec.gerencia, gerencias) : "#5B6B76";
            const thu = parseDate(d.thu);
            return (
              <tr key={d.id} className="border-b border-[#F5F6F5] last:border-b-0">
                <td className="px-4 py-2 font-mono2 text-[12px] text-[#5B6B76]">{thu.getDate()} {MONTHS[thu.getMonth()].slice(0,3)} – {parseDate(d.sun).getDate()} {MONTHS[parseDate(d.sun).getMonth()].slice(0,3)}</td>
                <td className="px-4 py-2 font-medium">{exec ? exec.name : "—"}</td>
                <td className="px-4 py-2">
                  <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ backgroundColor: color + "1A", color }}>{exec ? exec.gerencia : "—"}</span>
                </td>
                {editable && (
                  <td className="px-4 py-2">
                    <button onClick={() => onReroll(i)} className="flex items-center gap-1 text-[11px] text-[#5B6B76] hover:text-[#1B2733]">
                      <RefreshCw size={12} /> Rehacer
                    </button>
                  </td>
                )}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* VISTA ESTADISTICAS                                                    */
/* ------------------------------------------------------------------ */

function EstadisticasView({ executives, gerencias, shiftDraws, thisYear }) {
  const [year, setYear] = useState(thisYear);
  const draws = shiftDraws[year] || [];

  const perGerencia = gerencias.map(g => {
    const execsInPool = executives.filter(e => e.gerencia === g);
    const shiftsAssigned = draws.filter(d => {
      const e = executives.find(ex => ex.id === d.executiveId);
      return e && e.gerencia === g;
    }).length;
    return { gerencia: g, execCount: execsInPool.length, shiftsAssigned };
  });

  const perExecutive = executives.map(e => {
    const count = draws.filter(d => d.executiveId === e.id).length;
    return { ...e, count };
  }).sort((a, b) => b.count - a.count);

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-5">
        <div>
          <h1 className="font-display text-xl font-semibold text-[#1B2733]">Estadísticas de turnos</h1>
          <p className="text-[12px] text-[#5B6B76] mt-0.5">Objetivo: máximo {TARGET_SHIFTS} turnos de fin de semana por ejecutivo al año.</p>
        </div>
        <select value={year} onChange={e => setYear(Number(e.target.value))} className="text-[13px] border border-[#D8DEE1] rounded-md px-2.5 py-2 bg-white">
          {[thisYear, thisYear + 1].map(y => <option key={y} value={y}>{y}</option>)}
        </select>
      </div>

      <div className="grid lg:grid-cols-2 gap-4 mb-6">
        <div className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden">
          <div className="px-4 py-2.5 border-b border-[#E3E7E5]"><span className="text-[13px] font-semibold">Por gerencia</span></div>
          <table className="w-full text-[13px]">
            <thead>
              <tr className="text-left text-[11px] text-[#8CA0AC] border-b border-[#EEF1F0]">
                <th className="px-4 py-2 font-medium">Gerencia</th>
                <th className="px-4 py-2 font-medium text-right">Ejecutivos en pool</th>
                <th className="px-4 py-2 font-medium text-right">Turnos asignados {year}</th>
              </tr>
            </thead>
            <tbody>
              {perGerencia.map(row => (
                <tr key={row.gerencia} className="border-b border-[#F5F6F5] last:border-b-0">
                  <td className="px-4 py-2 font-medium" style={{ color: gerenciaColor(row.gerencia, gerencias) }}>{row.gerencia}</td>
                  <td className="px-4 py-2 text-right font-mono2">{row.execCount}</td>
                  <td className="px-4 py-2 text-right font-mono2">{row.shiftsAssigned}</td>
                </tr>
              ))}
              {perGerencia.length === 0 && <tr><td colSpan={3} className="px-4 py-4 text-center text-[12px] text-[#9AA8AF]">Sin datos.</td></tr>}
            </tbody>
          </table>
        </div>

        <div className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden">
          <div className="px-4 py-2.5 border-b border-[#E3E7E5] flex items-center justify-between">
            <span className="text-[13px] font-semibold">Por ejecutivo</span>
            <span className="text-[11px] text-[#9AA8AF]">Meta ≤ {TARGET_SHIFTS}</span>
          </div>
          <table className="w-full text-[13px]">
            <thead>
              <tr className="text-left text-[11px] text-[#8CA0AC] border-b border-[#EEF1F0]">
                <th className="px-4 py-2 font-medium">Ejecutivo</th>
                <th className="px-4 py-2 font-medium">Gerencia</th>
                <th className="px-4 py-2 font-medium text-right">Turnos</th>
                <th className="px-4 py-2 font-medium">Estado</th>
              </tr>
            </thead>
            <tbody>
              {perExecutive.map(e => {
                const over = e.count > TARGET_SHIFTS;
                return (
                  <tr key={e.id} className="border-b border-[#F5F6F5] last:border-b-0">
                    <td className="px-4 py-2 font-medium">{e.name}</td>
                    <td className="px-4 py-2 text-[12px] text-[#5B6B76]">{e.gerencia}</td>
                    <td className="px-4 py-2 text-right font-mono2 font-semibold">{e.count}</td>
                    <td className="px-4 py-2">
                      {over ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-medium text-[#C1443B] bg-[#FCEDEB] px-2 py-0.5 rounded-full">
                          <AlertTriangle size={11} /> Sobre target
                        </span>
                      ) : (
                        <span className="text-[11px] font-medium text-[#2F6F5E] bg-[#EAF3EF] px-2 py-0.5 rounded-full">En target</span>
                      )}
                    </td>
                  </tr>
                );
              })}
              {perExecutive.length === 0 && <tr><td colSpan={4} className="px-4 py-4 text-center text-[12px] text-[#9AA8AF]">Sin ejecutivos cargados.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* VISTA ADMINISTRACION                                                  */
/* ------------------------------------------------------------------ */

function AdminView({ session, accounts, gerencias, auditLog, categories, onAddAccount, onEditAccount, onDeleteAccount, onAddCategory, onEditCategory, onDeleteCategory }) {
  const isMaster = session.role === "master";
  const [revealed, setRevealed] = useState({});
  const visibleLog = isMaster ? auditLog : auditLog.filter(l => l.user === session.username);

  return (
    <div>
      <h1 className="font-display text-xl font-semibold text-[#1B2733] mb-1">Administración</h1>
      <p className="text-[12px] text-[#5B6B76] mb-5">
        {isMaster ? "Gestiona los accesos por gerencia, la clasificación de actividades y revisa el historial de cambios." : "Tu cuenta y tu historial de cambios."}
      </p>

      {isMaster && (
        <div className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden mb-6">
          <div className="px-4 py-2.5 border-b border-[#E3E7E5] flex items-center justify-between">
            <span className="text-[13px] font-semibold">Clasificación de actividades</span>
            <button onClick={onAddCategory} className="flex items-center gap-1.5 text-[12px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] px-3 py-1.5 rounded-md">
              <Plus size={13} /> Nueva clasificación
            </button>
          </div>
          <div className="divide-y divide-[#F5F6F5]">
            {categories.map(c => (
              <div key={c.id} className="px-4 py-2.5 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <span className="w-3.5 h-3.5 rounded-sm shrink-0" style={{ backgroundColor: c.color }} />
                  <span className="text-[13px] text-[#1B2733]">{c.name}</span>
                  <span className="text-[11px] font-mono2 text-[#9AA8AF]">{c.color}</span>
                </div>
                <div className="flex items-center gap-1">
                  <button onClick={() => onEditCategory(c)} className="p-1.5 rounded hover:bg-[#F0F2F1] text-[#5B6B76]"><Pencil size={13} /></button>
                  <button onClick={() => onDeleteCategory(c)} className="p-1.5 rounded hover:bg-[#FCEDEB] text-[#C1443B]"><Trash2 size={13} /></button>
                </div>
              </div>
            ))}
            {categories.length === 0 && <div className="px-4 py-4 text-center text-[12px] text-[#9AA8AF]">Sin clasificaciones creadas.</div>}
          </div>
        </div>
      )}

      {isMaster && (
        <div className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden mb-6">
          <div className="px-4 py-2.5 border-b border-[#E3E7E5] flex items-center justify-between">
            <span className="text-[13px] font-semibold flex items-center gap-1.5"><KeyRound size={14} /> Cuentas Comité Ejecutivo</span>
            <button onClick={onAddAccount} className="flex items-center gap-1.5 text-[12px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] px-3 py-1.5 rounded-md">
              <Plus size={13} /> Crear acceso
            </button>
          </div>
          <table className="w-full text-[13px]">
            <thead>
              <tr className="text-left text-[11px] text-[#8CA0AC] border-b border-[#EEF1F0]">
                <th className="px-4 py-2 font-medium">Usuario</th>
                <th className="px-4 py-2 font-medium">Gerencia</th>
                <th className="px-4 py-2 font-medium">Clave</th>
                <th className="px-4 py-2 font-medium"></th>
              </tr>
            </thead>
            <tbody>
              {accounts.map(a => (
                <tr key={a.username} className="border-b border-[#F5F6F5] last:border-b-0">
                  <td className="px-4 py-2 font-medium">{a.username}</td>
                  <td className="px-4 py-2 text-[12px]" style={{ color: gerenciaColor(a.gerencia, gerencias) }}>{a.gerencia}</td>
                  <td className="px-4 py-2 font-mono2 text-[12px]">
                    <div className="flex items-center gap-2">
                      <span>{revealed[a.username] ? a.password : "••••••••"}</span>
                      <button onClick={() => setRevealed(r => ({ ...r, [a.username]: !r[a.username] }))} className="text-[#8CA0AC] hover:text-[#1B2733]"><Eye size={13} /></button>
                      <button onClick={() => navigator.clipboard?.writeText(a.password)} className="text-[#8CA0AC] hover:text-[#1B2733]"><Copy size={13} /></button>
                    </div>
                  </td>
                  <td className="px-4 py-2">
                    <div className="flex items-center gap-1 justify-end">
                      <button onClick={() => onEditAccount(a)} className="p-1.5 rounded hover:bg-[#F0F2F1] text-[#5B6B76]"><Pencil size={13} /></button>
                      <button onClick={() => onDeleteAccount(a)} className="p-1.5 rounded hover:bg-[#FCEDEB] text-[#C1443B]"><Trash2 size={13} /></button>
                    </div>
                  </td>
                </tr>
              ))}
              {accounts.length === 0 && <tr><td colSpan={4} className="px-4 py-4 text-center text-[12px] text-[#9AA8AF]">Aún no se han creado accesos del Comité Ejecutivo.</td></tr>}
            </tbody>
          </table>
        </div>
      )}

      <div className="bg-white border border-[#E3E7E5] rounded-lg overflow-hidden">
        <div className="px-4 py-2.5 border-b border-[#E3E7E5]">
          <span className="text-[13px] font-semibold">{isMaster ? "Historial de cambios" : "Mis cambios"}</span>
        </div>
        <div className="max-h-[420px] overflow-y-auto divide-y divide-[#F5F6F5]">
          {visibleLog.map((l, i) => (
            <div key={i} className="px-4 py-2.5 flex items-start justify-between gap-3">
              <div>
                <div className="text-[12.5px]"><span className="font-semibold">{l.user}</span> — {l.action}</div>
                <div className="text-[11.5px] text-[#8CA0AC]">{l.detail}</div>
              </div>
              <div className="text-[11px] font-mono2 text-[#9AA8AF] whitespace-nowrap">{new Date(l.ts).toLocaleString("es-CL")}</div>
            </div>
          ))}
          {visibleLog.length === 0 && <div className="px-4 py-5 text-center text-[12px] text-[#9AA8AF]">Sin registros aún.</div>}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* MODALES                                                              */
/* ------------------------------------------------------------------ */

function ModalShell({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 bg-[#1B2733]/50 flex items-center justify-center p-4" onClick={onClose}>
      <div onClick={e => e.stopPropagation()} className={`bg-white rounded-lg shadow-xl w-full ${wide ? "max-w-lg" : "max-w-sm"} max-h-[90vh] overflow-y-auto`}>
        <div className="flex items-center justify-between px-5 py-3.5 border-b border-[#EEF1F0]">
          <h2 className="font-display font-semibold text-[15px] text-[#1B2733]">{title}</h2>
          <button onClick={onClose} className="p-1 rounded hover:bg-[#F0F2F1] text-[#5B6B76]"><X size={16} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div className="mb-3.5">
      <label className="block text-[11.5px] font-medium text-[#5B6B76] mb-1">{label}</label>
      {children}
    </div>
  );
}
const inputCls = "w-full text-[13px] border border-[#D8DEE1] rounded-md px-2.5 py-2 outline-none focus:border-[#2F6F5E] bg-white";

function LoginModal({ accounts, onClose, onSubmit, onGuestLogin }) {
  const grouped = [...accounts].sort((a, b) => a.gerencia.localeCompare(b.gerencia) || a.username.localeCompare(b.username));
  const [selected, setSelected] = useState("GUEST");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  function submit(e) {
    if (e && e.preventDefault) e.preventDefault();
    if (selected === "GUEST") { onGuestLogin(); return; }
    const err = onSubmit(selected, password);
    if (err) setError(err);
  }

  return (
    <ModalShell title="Iniciar sesión" onClose={onClose}>
      <div onKeyDown={e => { if (e.key === "Enter") submit(e); }}>
        <Field label="Usuario">
          <select
            autoFocus
            value={selected}
            onChange={e => { setSelected(e.target.value); setPassword(""); setError(""); }}
            className={inputCls}
          >
            <option value="GUEST">👁 Invitado — solo visualización</option>
            <option value={MASTER_USER}>🛡 {MASTER_USER}</option>
            {grouped.length > 0 && (
              <optgroup label="Comité Ejecutivo">
                {grouped.map(a => <option key={a.username} value={a.username}>{a.username} — {a.gerencia}</option>)}
              </optgroup>
            )}
          </select>
        </Field>

        {selected === "GUEST" ? (
          <p className="text-[12px] text-[#5B6B76] mb-4">Entrarás como invitado: puedes ver el calendario, ejecutivos y estadísticas, pero no podrás agregar ni modificar nada.</p>
        ) : (
          <Field label="Clave">
            <input autoFocus type="password" value={password} onChange={e => setPassword(e.target.value)} className={inputCls} />
          </Field>
        )}

        {error && <p className="text-[12px] text-[#C1443B] mb-3">{error}</p>}
        <button type="button" onClick={submit} className="w-full text-[13px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] py-2.5 rounded-md">
          {selected === "GUEST" ? "Entrar como invitado" : "Entrar"}
        </button>
      </div>
    </ModalShell>
  );
}

function ActivityModal({ data, gerencias, categories, session, onClose, onSave, onSaveOccurrence }) {
  const mode = data.mode; // "add" | "edit" | "edit-occurrence"
  const isOccurrenceMode = mode === "edit-occurrence";
  const act = mode === "edit" ? data.activity : (isOccurrenceMode ? data.occurrence : null);
  const editing = mode !== "add";

  const [title, setTitle] = useState(act?.title || "");
  const [date, setDate] = useState((isOccurrenceMode ? act?.occurrenceDate : act?.date) || data.date);
  const [start, setStart] = useState(act?.start || data.start || "09:00");
  const [end, setEnd] = useState(act?.end || minutesToTime(timeToMinutes(data.start || "09:00") + 60));
  const [gerencia, setGerencia] = useState(act?.gerencia || (session.role === "gerente" ? session.gerencia : gerencias[0] || ""));
  const [categoryId, setCategoryId] = useState(act?.categoryId || "");
  const [note, setNote] = useState(act?.note || "");
  const [error, setError] = useState("");
  const gerenciaLocked = session.role === "gerente";

  // repetición: solo se elige/edita a nivel de "add" o "edit" de serie, no en edit-occurrence
  const [repeatType, setRepeatType] = useState(act?.recurrence ? act.recurrence.freq : "none");
  const [repeatUntil, setRepeatUntil] = useState(act?.recurrence?.until || "");

  const refDate = date ? parseDate(date) : new Date();
  const weekdayPreview = WEEKDAY_NAMES[refDate.getDay()];
  const nthPreview = NTH_NAMES[nthWeekdayOfMonth(refDate)] || "último";

  function submit(e) {
    if (e && e.preventDefault) e.preventDefault();
    if (!title.trim()) { setError("Ingresa un título para la actividad."); return; }
    if (!gerencia) { setError("Selecciona una gerencia."); return; }
    if (!categoryId) { setError("Selecciona una clasificación."); return; }
    if (timeToMinutes(end) - timeToMinutes(start) < 30) { setError("El bloque debe durar al menos 30 minutos."); return; }

    if (isOccurrenceMode) {
      onSaveOccurrence(act, { title: title.trim(), date, start, end, gerencia, categoryId, note: note.trim() });
      return;
    }

    let recurrence = null;
    if (repeatType === "weekly") recurrence = { freq: "weekly", weekday: refDate.getDay(), until: repeatUntil || null };
    if (repeatType === "monthly") recurrence = { freq: "monthly", weekday: refDate.getDay(), nth: nthWeekdayOfMonth(refDate), until: repeatUntil || null };

    onSave({ id: act?.id, title: title.trim(), date, start, end, gerencia, categoryId, note: note.trim(), recurrence });
  }

  return (
    <ModalShell title={isOccurrenceMode ? "Editar esta actividad" : editing ? "Editar actividad" : "Nueva actividad"} onClose={onClose}>
      <div onKeyDown={e => { if (e.key === "Enter") submit(e); }}>
        {isOccurrenceMode && (
          <div className="mb-3.5 text-[12px] text-[#8A6A2A] bg-[#FDF3E1] border border-[#F0D9A8] rounded-md px-3 py-2">
            Estás editando solo esta fecha ({act.occurrenceDate}) de la serie "{act.title}". El resto de las repeticiones no se modifican.
          </div>
        )}
        <Field label="Título"><input autoFocus value={title} onChange={e => setTitle(e.target.value)} className={inputCls} placeholder="Reunión de coordinación" /></Field>
        <div className="grid grid-cols-3 gap-2">
          <Field label="Fecha"><input type="date" value={date} onChange={e => setDate(e.target.value)} className={inputCls} /></Field>
          <Field label="Inicio">
            <select value={start} onChange={e => { setStart(e.target.value); if (timeToMinutes(end) <= timeToMinutes(e.target.value)) setEnd(minutesToTime(timeToMinutes(e.target.value) + 60)); }} className={inputCls}>
              {timeSlots().map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="Término">
            <select value={end} onChange={e => setEnd(e.target.value)} className={inputCls}>
              {endTimeOptions(start).map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-2">
          <Field label="Gerencia">
            <select value={gerencia} onChange={e => setGerencia(e.target.value)} className={inputCls} disabled={gerenciaLocked}>
              {gerencias.map(g => <option key={g} value={g}>{g}</option>)}
            </select>
          </Field>
          <Field label="Clasificación">
            <select value={categoryId} onChange={e => setCategoryId(e.target.value)} className={inputCls} required>
              <option value="" disabled>Selecciona un tipo de actividad…</option>
              {categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
        </div>
        {categoryId ? (
          <div className="flex items-center gap-1.5 -mt-2.5 mb-3.5">
            <span className="w-2.5 h-2.5 rounded-sm shrink-0" style={{ backgroundColor: categoryById(categories, categoryId)?.color }} />
            <span className="text-[11px] text-[#9AA8AF]">Color asignado a esta clasificación</span>
          </div>
        ) : (
          <div className="flex items-center gap-1.5 -mt-2.5 mb-3.5">
            <AlertTriangle size={12} className="text-[#C1443B] shrink-0" />
            <span className="text-[11px] text-[#C1443B]">Debes elegir un tipo de actividad para poder guardar.</span>
          </div>
        )}
        {categories.length === 0 && <p className="text-[12px] text-[#C1443B] mb-3">No hay clasificaciones creadas. Pídele a Excelencia Operacional que cree una en Administración.</p>}

        {!isOccurrenceMode && (
          <>
            <Field label="Repetición">
              <select value={repeatType} onChange={e => setRepeatType(e.target.value)} className={inputCls}>
                <option value="none">No se repite (única vez)</option>
                <option value="weekly">Semanalmente — cada {weekdayPreview}</option>
                <option value="monthly">Mensualmente — el {nthPreview} {weekdayPreview} de cada mes</option>
              </select>
            </Field>
            {repeatType !== "none" && (
              <Field label="Repetir hasta (opcional)">
                <input type="date" value={repeatUntil} onChange={e => setRepeatUntil(e.target.value)} className={inputCls} />
                <p className="text-[11px] text-[#9AA8AF] mt-1">Si lo dejas vacío, se repite indefinidamente. Puedes editar o eliminar fechas puntuales más adelante sin afectar el resto.</p>
              </Field>
            )}
          </>
        )}

        <Field label="Notas (opcional)"><textarea value={note} onChange={e => setNote(e.target.value)} rows={2} className={inputCls} /></Field>
        {error && <p className="text-[12px] text-[#C1443B] mb-3">{error}</p>}
        <div className="flex gap-2 mt-4">
          <button type="button" onClick={submit} disabled={!categoryId} className="flex-1 text-[13px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] disabled:opacity-40 disabled:cursor-not-allowed py-2.5 rounded-md">Guardar</button>
        </div>
      </div>
    </ModalShell>
  );
}

function EditChoiceModal({ occurrence, action, onClose, onChooseOccurrence, onChooseSeries }) {
  return (
    <ModalShell title={action === "edit" ? "Editar actividad recurrente" : "Eliminar actividad recurrente"} onClose={onClose}>
      <p className="text-[13px] text-[#3D4C56] mb-1">"{occurrence.title}" se repite: <strong>{describeRecurrence(occurrence.recurrence)}</strong>.</p>
      <p className="text-[12px] text-[#5B6B76] mb-4">¿Qué quieres {action === "edit" ? "editar" : "eliminar"}?</p>
      <div className="flex flex-col gap-2">
        <button onClick={onChooseOccurrence} className="text-left text-[13px] font-medium border border-[#D8DEE1] rounded-md px-3 py-2.5 hover:bg-[#F0F2F1]">
          Solo esta actividad <span className="text-[11px] text-[#9AA8AF] font-mono2">({occurrence.occurrenceDate})</span>
        </button>
        <button onClick={onChooseSeries} className="text-left text-[13px] font-medium border border-[#D8DEE1] rounded-md px-3 py-2.5 hover:bg-[#F0F2F1]">
          Toda la serie <span className="text-[11px] text-[#9AA8AF]">({describeRecurrence(occurrence.recurrence)})</span>
        </button>
      </div>
    </ModalShell>
  );
}

function ExecutivoModal({ data, gerencias, session, onClose, onSave, onAddGerencia }) {
  const editing = data.mode === "edit";
  const exec = editing ? data.executive : null;
  const [name, setName] = useState(exec?.name || "");
  const [gerencia, setGerencia] = useState(exec?.gerencia || (session.role === "gerente" ? session.gerencia : gerencias[0] || ""));
  const [newG, setNewG] = useState("");
  const [error, setError] = useState("");
  const gerenciaLocked = session.role === "gerente";

  function submit(e) {
    if (e && e.preventDefault) e.preventDefault();
    if (!name.trim()) { setError("Ingresa el nombre del ejecutivo."); return; }
    if (!gerencia) { setError("Selecciona una gerencia."); return; }
    onSave({ id: exec?.id, name: name.trim(), gerencia });
  }

  return (
    <ModalShell title={editing ? "Editar ejecutivo" : "Agregar ejecutivo de turno"} onClose={onClose}>
      <div onKeyDown={e => { if (e.key === "Enter") submit(e); }}>
        <Field label="Nombre completo"><input autoFocus value={name} onChange={e => setName(e.target.value)} className={inputCls} /></Field>
        <Field label="Gerencia">
          <select value={gerencia} onChange={e => setGerencia(e.target.value)} className={inputCls} disabled={gerenciaLocked}>
            {gerencias.map(g => <option key={g} value={g}>{g}</option>)}
          </select>
        </Field>
        {!gerenciaLocked && (
          <div className="flex items-center gap-2 mb-3.5 -mt-2">
            <input value={newG} onChange={e => setNewG(e.target.value)} placeholder="…o crea una gerencia nueva" className={inputCls} />
            <button type="button" onClick={() => { if (newG.trim()) { onAddGerencia(newG.trim()); setGerencia(newG.trim()); setNewG(""); } }} className="text-[12px] font-medium text-[#2F6F5E] whitespace-nowrap">Crear</button>
          </div>
        )}
        {error && <p className="text-[12px] text-[#C1443B] mb-3">{error}</p>}
        <button type="button" onClick={submit} className="w-full text-[13px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] py-2.5 rounded-md">Guardar</button>
      </div>
    </ModalShell>
  );
}

function AccountModal({ data, gerencias, onClose, onSave }) {
  const editing = data.mode === "edit";
  const acc = editing ? data.account : null;
  const [username, setUsername] = useState(acc?.username || "");
  const [password, setPassword] = useState(acc?.password || "");
  const [gerencia, setGerencia] = useState(acc?.gerencia || gerencias[0] || "");
  const [error, setError] = useState("");

  function submit(e) {
    if (e && e.preventDefault) e.preventDefault();
    if (!username.trim() || !password.trim() || !gerencia) { setError("Completa todos los campos."); return; }
    onSave({ username: username.trim(), password: password.trim(), gerencia, editingUsername: acc?.username });
  }

  return (
    <ModalShell title={editing ? "Editar acceso" : "Crear acceso Comité Ejecutivo"} onClose={onClose}>
      <div onKeyDown={e => { if (e.key === "Enter") submit(e); }}>
        <Field label="Nombre (usuario de acceso)"><input autoFocus value={username} onChange={e => setUsername(e.target.value)} className={inputCls} placeholder="Nombre del integrante" /></Field>
        <Field label="Clave a enviar">
          <div className="flex gap-2">
            <input value={password} onChange={e => setPassword(e.target.value)} className={inputCls} />
            <button type="button" onClick={() => setPassword(Math.random().toString(36).slice(-8))} className="text-[11px] font-medium text-[#2F6F5E] whitespace-nowrap">Generar</button>
          </div>
        </Field>
        <Field label="Gerencia">
          <select value={gerencia} onChange={e => setGerencia(e.target.value)} className={inputCls}>
            {gerencias.map(g => <option key={g} value={g}>{g}</option>)}
          </select>
        </Field>
        {gerencias.length === 0 && <p className="text-[12px] text-[#C1443B] mb-3">Primero crea al menos una gerencia en la pestaña "Ejecutivos de turno".</p>}
        {error && <p className="text-[12px] text-[#C1443B] mb-3">{error}</p>}
        <button type="button" onClick={submit} disabled={gerencias.length === 0} className="w-full text-[13px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] disabled:opacity-40 py-2.5 rounded-md">Guardar acceso</button>
      </div>
    </ModalShell>
  );
}

const COLOR_PRESETS = ["#3C8A3E", "#3563A6", "#D4A017", "#D2691E", "#8A5FB0", "#2F8A8A", "#9E3E6C", "#C1443B", "#5B6B76"];

function CategoryModal({ data, onClose, onSave }) {
  const editing = data.mode === "edit";
  const cat = editing ? data.category : null;
  const [name, setName] = useState(cat?.name || "");
  const [color, setColor] = useState(cat?.color || COLOR_PRESETS[0]);
  const [error, setError] = useState("");

  function submit(e) {
    if (e && e.preventDefault) e.preventDefault();
    if (!name.trim()) { setError("Ingresa el nombre de la clasificación."); return; }
    onSave({ id: cat?.id, name: name.trim(), color });
  }

  return (
    <ModalShell title={editing ? "Editar clasificación" : "Nueva clasificación de actividad"} onClose={onClose}>
      <div onKeyDown={e => { if (e.key === "Enter") submit(e); }}>
        <Field label="Nombre"><input autoFocus value={name} onChange={e => setName(e.target.value)} className={inputCls} placeholder="Ej: Auditorías" /></Field>
        <Field label="Color">
          <div className="flex items-center gap-2 mb-2">
            <input type="color" value={color} onChange={e => setColor(e.target.value)} className="w-9 h-9 rounded border border-[#D8DEE1] cursor-pointer p-0.5 bg-white" />
            <span className="text-[12px] font-mono2 text-[#5B6B76]">{color}</span>
          </div>
          <div className="flex flex-wrap gap-1.5">
            {COLOR_PRESETS.map(p => (
              <button type="button" key={p} onClick={() => setColor(p)} className="w-6 h-6 rounded-full border-2" style={{ backgroundColor: p, borderColor: color === p ? "#1B2733" : "transparent" }} />
            ))}
          </div>
        </Field>
        <div className="flex items-center gap-1.5 mb-4 px-2.5 py-2 rounded-md" style={{ backgroundColor: color + "1A" }}>
          <span className="w-2.5 h-2.5 rounded-sm shrink-0" style={{ backgroundColor: color }} />
          <span className="text-[12px] font-medium" style={{ color }}>{name || "Vista previa"}</span>
        </div>
        {error && <p className="text-[12px] text-[#C1443B] mb-3">{error}</p>}
        <button type="button" onClick={submit} className="w-full text-[13px] font-medium text-white bg-[#2F6F5E] hover:bg-[#285F51] py-2.5 rounded-md">Guardar clasificación</button>
      </div>
    </ModalShell>
  );
}

function ConfirmModal({ title, message, onConfirm, onCancel }) {
  return (
    <ModalShell title={title} onClose={onCancel}>
      <p className="text-[13px] text-[#3D4C56] mb-5">{message}</p>
      <div className="flex gap-2">
        <button onClick={onCancel} className="flex-1 text-[13px] font-medium text-[#3D4C56] bg-[#F0F2F1] hover:bg-[#E5E8E6] py-2.5 rounded-md">Cancelar</button>
        <button onClick={() => { onConfirm(); }} className="flex-1 text-[13px] font-medium text-white bg-[#C1443B] hover:bg-[#A93A32] py-2.5 rounded-md">Confirmar</button>
      </div>
    </ModalShell>
  );
}
