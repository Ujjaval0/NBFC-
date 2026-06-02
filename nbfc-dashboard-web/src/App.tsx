import Papa from "papaparse";
import {
  Activity,
  AlertTriangle,
  Clock,
  FilterX,
  LayoutDashboard,
  Landmark,
  PanelLeftClose,
  PanelLeft,
  ShieldCheck,
  TrendingUp,
  Download,
  ChevronDown,
  MoreHorizontal,
  ShoppingBag,
  Users,
  HelpCircle
} from "lucide-react";
import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

/* ═══════════════════════════════════════════
   Types
   ═══════════════════════════════════════════ */
type CsvRow = Record<string, string>;

type Application = {
  applicationId: string;
  customerName: string;
  age: number;
  employmentType: string;
  monthlyIncome: number;
  cibilScore: number;
  branch: string;
  loanType: string;
  requestedAmount: number;
  sanctionedAmount: number;
  status: string;
  applicationDate: Date;
  dateKey: string;
  monthKey: string;
  ticketSize: string;
  customerSegment: string;
  cibilBand: string;
  incomeBand: string;
  assignedUsers: string[];
};

type Operation = {
  logId: string;
  applicationId: string;
  stage: string;
  assignedTo: string;
  stageStart: Date;
  stageEnd: Date | null;
  stageDateKey: string;
  status: string;
  tatHours: number | null;
  branch: string;
  loanType: string;
  ticketSize: string;
  customerSegment: string;
};

type Repayment = {
  loanId: string;
  applicationId: string;
  tenureMonths: number;
  roi: number;
  principalOutstanding: number;
  amountDue: number;
  amountPaid: number;
  dpd: number;
  npaStatus: string;
  npaFlag: boolean;
  collectionStatus: string;
};

type LoanModel = Application & {
  loanId: string;
  tenureMonths: number;
  roi: number;
  principalOutstanding: number;
  amountDue: number;
  amountPaid: number;
  dpd: number;
  npaStatus: string;
  npaFlag: boolean;
  collectionStatus: string;
  hasLoan: boolean;
};

type DataState = {
  applications: Application[];
  operations: Operation[];
  repayments: Repayment[];
  model: LoanModel[];
};

type Filters = {
  startDate: string;
  endDate: string;
  branch: string;
  loanType: string;
  ticketSize: string;
  user: string;
  customerSegment: string;
};

/* ═══════════════════════════════════════════
   Helpers
   ═══════════════════════════════════════════ */
const initialData: DataState = { applications: [], operations: [], repayments: [], model: [] };

function parseNumber(v: string | undefined): number {
  const n = Number(v ?? 0);
  return Number.isFinite(n) ? n : 0;
}
function toDateKey(d: Date) { return d.toISOString().slice(0, 10); }
function ticketBucket(a: number) { return a < 300_000 ? "Small" : a < 750_000 ? "Medium" : a < 1_500_000 ? "Large" : "High Value"; }
function cibilBandFn(s: number) { return s < 650 ? "<650" : s < 700 ? "650-699" : s < 750 ? "700-749" : "750+"; }
function incomeBandFn(i: number) { return i < 50_000 ? "<50k" : i < 100_000 ? "50k-99k" : i < 150_000 ? "100k-149k" : "150k+"; }
function segmentFn(c: number, i: number) { return c >= 750 && i >= 100_000 ? "Prime" : c >= 700 && i >= 50_000 ? "Mass Affluent" : c < 650 ? "Risk Watch" : "Mass Market"; }
function collStatusFn(d: number, p: number) { return d === 0 ? "No Current Due" : p >= d ? "Paid" : p > 0 ? "Partially Paid" : "Unpaid"; }

function fmtAmt(v: number) {
  if (Math.abs(v) >= 1e7) return `₹${(v / 1e7).toFixed(2)} Cr`;
  if (Math.abs(v) >= 1e5) return `₹${(v / 1e5).toFixed(2)} L`;
  return `₹${Math.round(v).toLocaleString("en-IN")}`;
}
function fmtPct(v: number) { return `${(v * 100).toFixed(1)}%`; }
function monthLabel(d: Date) { return new Intl.DateTimeFormat("en", { month: "short", year: "2-digit" }).format(d); }

async function loadCsv(path: string): Promise<CsvRow[]> {
  const r = await fetch(path);
  if (!r.ok) throw new Error(`Could not load ${path}`);
  const t = await r.text();
  const p = Papa.parse<CsvRow>(t, { header: true, skipEmptyLines: true });
  if (p.errors.length > 0) throw new Error(p.errors[0].message);
  return p.data;
}

function avg(vs: number[]) { const f = vs.filter(Number.isFinite); return f.length ? f.reduce((s, v) => s + v, 0) / f.length : 0; }
function sum(vs: number[]) { return vs.reduce((s, v) => s + v, 0); }
function uniq(vs: string[]) { return Array.from(new Set(vs.filter(Boolean))).sort(); }
function stageLabel(s: string) { return ({ Lead_Generation: "Lead Gen", Credit_Review: "Credit Review", Document_Verification: "Doc Verify", Disbursal_Queue: "Disbursal" }[s] ?? s.replaceAll("_", " ")); }
function groupCount<T>(rows: T[], fn: (r: T) => string) {
  const m = new Map<string, number>();
  rows.forEach((r) => { const k = fn(r); m.set(k, (m.get(k) ?? 0) + 1); });
  return Array.from(m, ([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
}

/* ─── Color Palette ─── */
const C = {
  blue: "#2e90fa",
  teal: "#0d9488",
  slate: "#475467",
  indigo: "#6366f1",
  sage: "#667085",
  coral: "#f04438",
  amber: "#f79009",
  chart: ["#6366f1", "#2e90fa", "#0d9488", "#475467", "#f04438", "#f79009", "#8b5cf6"],
  pie: ["#6366f1", "#2e90fa", "#0d9488", "#667085", "#f04438", "#f79009"],
};

/* ─── Navigation ─── */
type PageId = "overview" | "sales" | "credit" | "operations" | "portfolio" | "queue";
const NAV: { id: PageId; label: string; icon: typeof LayoutDashboard }[] = [
  { id: "overview", label: "Overview", icon: LayoutDashboard },
  { id: "sales", label: "Sales & Pipeline", icon: TrendingUp },
  { id: "credit", label: "Credit Quality", icon: ShieldCheck },
  { id: "operations", label: "Operations", icon: Clock },
  { id: "portfolio", label: "Portfolio Health", icon: Landmark },
  { id: "queue", label: "Pending Queue", icon: Activity },
];

/* ═══════════════════════════════════════════
   Shared chart defaults
   ═══════════════════════════════════════════ */
const GRID = { strokeDasharray: "3 3", stroke: "#eaecf0", vertical: false };
const AXIS = { tickLine: false, axisLine: false };

/* ═══════════════════════════════════════════
   App Component
   ═══════════════════════════════════════════ */
function App() {
  const [data, setData] = useState<DataState>(initialData);
  const [filters, setFilters] = useState<Filters>({ startDate: "", endDate: "", branch: "All", loanType: "All", ticketSize: "All", user: "All", customerSegment: "All" });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [page, setPage] = useState<PageId>("overview");
  const [sidebar, setSidebar] = useState(true);
  const [showSecondaryMetrics, setShowSecondaryMetrics] = useState(false);

  /* Load data */
  useEffect(() => {
    (async () => {
      try {
        const [appRows, opRows, repRows] = await Promise.all([
          loadCsv("/data/Applications_and_Loans.csv"),
          loadCsv("/data/Daily_Operational_Logs.csv"),
          loadCsv("/data/Repayments_and_Performance.csv"),
        ]);
        const apps: Application[] = appRows.map((r) => {
          const d = new Date(r.Application_Date), ra = parseNumber(r.Requested_Amount), mi = parseNumber(r.Monthly_Income), cs = parseNumber(r.CIBIL_Score);
          return { applicationId: r.Application_ID, customerName: r.Customer_Name, age: parseNumber(r.Age), employmentType: r.Employment_Type, monthlyIncome: mi, cibilScore: cs, branch: r.Branch, loanType: r.Loan_Type, requestedAmount: ra, sanctionedAmount: parseNumber(r.Sanctioned_Amount), status: r.Application_Status, applicationDate: d, dateKey: toDateKey(d), monthKey: monthLabel(d), ticketSize: ticketBucket(ra), customerSegment: segmentFn(cs, mi), cibilBand: cibilBandFn(cs), incomeBand: incomeBandFn(mi), assignedUsers: [] };
        });
        const appById = new Map(apps.map((a) => [a.applicationId, a]));
        const usersByApp = new Map<string, Set<string>>();
        opRows.forEach((r) => { if (!usersByApp.has(r.Application_ID)) usersByApp.set(r.Application_ID, new Set()); usersByApp.get(r.Application_ID)!.add(r.Assigned_To); });
        apps.forEach((a) => { a.assignedUsers = Array.from(usersByApp.get(a.applicationId) ?? []); });

        const ops: Operation[] = opRows.map((r) => {
          const ss = new Date(r.Stage_Start_Timestamp), se = r.Stage_End_Timestamp ? new Date(r.Stage_End_Timestamp) : null, ap = appById.get(r.Application_ID);
          return { logId: r.Log_ID, applicationId: r.Application_ID, stage: r.Stage, assignedTo: r.Assigned_To, stageStart: ss, stageEnd: se, stageDateKey: toDateKey(ss), status: r.Status, tatHours: se ? (se.getTime() - ss.getTime()) / 36e5 : null, branch: ap?.branch ?? "?", loanType: ap?.loanType ?? "?", ticketSize: ap?.ticketSize ?? "?", customerSegment: ap?.customerSegment ?? "?" };
        });

        const reps: Repayment[] = repRows.map((r) => {
          const ad = parseNumber(r.Amount_Due_Current_Month), ap = parseNumber(r.Amount_Paid_Current_Month);
          return { loanId: r.Loan_ID, applicationId: r.Application_ID, tenureMonths: parseNumber(r.Total_Tenure_Months), roi: parseNumber(r.Current_ROI), principalOutstanding: parseNumber(r.Principal_Outstanding), amountDue: ad, amountPaid: ap, dpd: parseNumber(r.DPD), npaStatus: r.NPA_Status, npaFlag: r.NPA_Status.toLowerCase().includes("npa"), collectionStatus: collStatusFn(ad, ap) };
        });

        const repByApp = new Map(reps.map((r) => [r.applicationId, r]));
        const model: LoanModel[] = apps.map((a) => { const r = repByApp.get(a.applicationId); return { ...a, loanId: r?.loanId ?? "", tenureMonths: r?.tenureMonths ?? 0, roi: r?.roi ?? 0, principalOutstanding: r?.principalOutstanding ?? 0, amountDue: r?.amountDue ?? 0, amountPaid: r?.amountPaid ?? 0, dpd: r?.dpd ?? 0, npaStatus: r?.npaStatus ?? "No Active Loan", npaFlag: r?.npaFlag ?? false, collectionStatus: r?.collectionStatus ?? "No Active Loan", hasLoan: Boolean(r) }; });

        const dk = apps.map((a) => a.dateKey).sort();
        setData({ applications: apps, operations: ops, repayments: reps, model });
        setFilters((f) => ({ ...f, startDate: dk[0] ?? "", endDate: dk.at(-1) ?? "" }));
      } catch (e) { setError(e instanceof Error ? e.message : "Failed"); }
      finally { setLoading(false); }
    })();
  }, []);

  /* Filter options */
  const opts = useMemo(() => ({
    branches: ["All", ...uniq(data.applications.map((r) => r.branch))],
    loanTypes: ["All", ...uniq(data.applications.map((r) => r.loanType))],
    ticketSizes: ["All", "Small", "Medium", "Large", "High Value"],
    users: ["All", ...uniq(data.operations.map((r) => r.assignedTo))],
    segments: ["All", ...uniq(data.applications.map((r) => r.customerSegment))],
  }), [data]);

  /* Filtered data */
  const fm = useMemo(() => data.model.filter((r) => {
    const um = filters.user === "All" || r.assignedUsers.includes(filters.user);
    return (!filters.startDate || r.dateKey >= filters.startDate) && (!filters.endDate || r.dateKey <= filters.endDate) && (filters.branch === "All" || r.branch === filters.branch) && (filters.loanType === "All" || r.loanType === filters.loanType) && (filters.ticketSize === "All" || r.ticketSize === filters.ticketSize) && (filters.customerSegment === "All" || r.customerSegment === filters.customerSegment) && um;
  }), [data.model, filters]);

  const fo = useMemo(() => data.operations.filter((r) => (!filters.startDate || r.stageDateKey >= filters.startDate) && (!filters.endDate || r.stageDateKey <= filters.endDate) && (filters.branch === "All" || r.branch === filters.branch) && (filters.loanType === "All" || r.loanType === filters.loanType) && (filters.ticketSize === "All" || r.ticketSize === filters.ticketSize) && (filters.user === "All" || r.assignedTo === filters.user) && (filters.customerSegment === "All" || r.customerSegment === filters.customerSegment)), [data.operations, filters]);

  /* Dashboard metrics */
  const d = useMemo(() => {
    const total = fm.length;
    const approved = fm.filter((r) => r.status === "Sanctioned" || r.status === "Disbursed");
    const disbursed = fm.filter((r) => r.status === "Disbursed");
    const rejected = fm.filter((r) => r.status === "Rejected");
    const pending = fm.filter((r) => r.status === "Pending" || r.status === "Under Review");
    const active = fm.filter((r) => r.hasLoan);
    const npa = active.filter((r) => r.npaFlag);
    const sanctPos = fm.filter((r) => r.sanctionedAmount > 0);
    const due = sum(fm.map((r) => r.amountDue));
    const paid = sum(fm.map((r) => r.amountPaid));

    const monthOrd = Array.from(new Set(data.model.map((r) => r.monthKey)));
    const monthly = monthOrd.map((m) => { const rows = fm.filter((r) => r.monthKey === m); return { name: m, applications: rows.length, disbursed: rows.filter((r) => r.status === "Disbursed").length, sanctionedAmt: sum(rows.map((r) => r.sanctionedAmount)) / 1e5 }; }).filter((r) => r.applications > 0 || r.disbursed > 0);

    const stages = ["Lead_Generation", "Credit_Review", "Document_Verification", "Disbursal_Queue"];
    const stageTat = stages.map((s) => { const sr = fo.filter((r) => r.stage === s && r.tatHours !== null); return { name: stageLabel(s), hours: Number(avg(sr.map((r) => r.tatHours ?? 0)).toFixed(1)), pending: fo.filter((r) => r.stage === s && r.status === "Pending").length }; });

    const cibilBands = ["<650", "650-699", "700-749", "750+"].map((b) => { const br = fm.filter((r) => r.cibilBand === b); const ba = br.filter((r) => r.status === "Sanctioned" || r.status === "Disbursed"); return { name: b, count: br.length, approvalRate: br.length ? Number(((ba.length / br.length) * 100).toFixed(1)) : 0 }; });

    const segs = ["Prime", "Mass Affluent", "Mass Market", "Risk Watch"].map((seg) => { const sr = fm.filter((r) => r.customerSegment === seg); const sa = sr.filter((r) => r.hasLoan); const sn = sa.filter((r) => r.npaFlag); const sd = sum(sr.map((r) => r.amountDue)); const sp = sum(sr.map((r) => r.amountPaid)); return { name: seg, count: sr.length, npaRate: sa.length ? Number(((sn.length / sa.length) * 100).toFixed(1)) : 0, collEff: sd ? Number(((sp / sd) * 100).toFixed(1)) : 100 }; });

    const dpdBuckets = [{ label: "0 DPD", min: 0, max: 0 }, { label: "1-30", min: 1, max: 30 }, { label: "31-60", min: 31, max: 60 }, { label: "61-90", min: 61, max: 90 }, { label: "90+", min: 91, max: Infinity }];
    const dpdDist = dpdBuckets.map((b) => ({ name: b.label, value: active.filter((r) => r.dpd >= b.min && r.dpd <= b.max).length }));

    const pendingOps = fo.filter((r) => r.status === "Pending").sort((a, b) => a.stageStart.getTime() - b.stageStart.getTime());

    return {
      total, approvedCount: approved.length, disbursedCount: disbursed.length, rejectedCount: rejected.length, pendingCount: pending.length,
      sanctionedAmt: sum(fm.map((r) => r.sanctionedAmount)), approvalRate: total ? approved.length / total : 0, rejectionRate: total ? rejected.length / total : 0,
      avgCibil: avg(fm.map((r) => r.cibilScore)), avgTicket: avg(sanctPos.map((r) => r.sanctionedAmount)),
      principalOut: sum(fm.map((r) => r.principalOutstanding)), collEff: due ? paid / due : 0,
      npaRate: active.length ? npa.length / active.length : 0, activeCount: active.length, npaCount: npa.length,
      totalDue: due, totalPaid: paid,
      monthly, branchApps: groupCount(fm, (r) => r.branch), loanMix: groupCount(fm, (r) => r.loanType),
      segs, stageTat, cibilBands, dpdDist, pendingOps,
    };
  }, [data.model, fm, fo]);

  const resetFilters = useCallback(() => {
    const dk = data.applications.map((a) => a.dateKey).sort();
    setFilters({ startDate: dk[0] ?? "", endDate: dk.at(-1) ?? "", branch: "All", loanType: "All", ticketSize: "All", user: "All", customerSegment: "All" });
  }, [data.applications]);



  const pageLabel = NAV.find((n) => n.id === page)?.label ?? "Overview";

  if (loading) return <main className="loading-shell"><div className="loader" /><p>Loading dashboard…</p></main>;
  if (error) return <main className="loading-shell"><AlertTriangle size={28} /><p>{error}</p></main>;

  /* ═══════════════════════════════════════════
     Render Main Shell
     ═══════════════════════════════════════════ */
  return (
    <div className={`dashboard-shell${sidebar ? "" : " collapsed"}`}>

      {/* ─── Sidebar ─── */}
      <aside className={`sidebar${sidebar ? "" : " hidden"}`}>
        <div className="sidebar-brand">
          <div className="sidebar-brand-text">
            <h1>NBFC Command</h1>
            <span>Analytics Platform</span>
          </div>
          <button className="sidebar-collapse-btn" type="button" onClick={() => setSidebar(false)} title="Collapse">
            <PanelLeftClose size={14} />
          </button>
        </div>
        
        <nav className="sidebar-nav">
          <div className="sidebar-section-label">Dashboard</div>
          {NAV.map((n) => (
            <button 
              key={n.id} 
              className={`sidebar-link${page === n.id ? " active" : ""}`} 
              onClick={() => setPage(n.id)} 
              type="button"
            >
              <span className="sidebar-link-content">
                <n.icon size={18} />
                <span>{n.label}</span>
              </span>
              {n.id === "queue" && <span className="sidebar-badge">{d.pendingOps.length}</span>}
            </button>
          ))}
        </nav>

        {/* Original statistical footer preserved & styled */}
        <div className="sidebar-footer">
          <div className="sidebar-stat" style={{ padding: "6px 0", display: "flex", justifyContent: "space-between", fontSize: "11px", color: "var(--text-tertiary)" }}>
            <span>Records</span>
            <strong style={{ color: "var(--text-secondary)", fontWeight: 600 }}>{data.applications.length}</strong>
          </div>
          <div className="sidebar-stat" style={{ padding: "6px 0", display: "flex", justifyContent: "space-between", fontSize: "11px", color: "var(--text-tertiary)" }}>
            <span>Op Logs</span>
            <strong style={{ color: "var(--text-secondary)", fontWeight: 600 }}>{data.operations.length}</strong>
          </div>
          <div className="sidebar-stat" style={{ padding: "6px 0", display: "flex", justifyContent: "space-between", fontSize: "11px", color: "var(--text-tertiary)" }}>
            <span>Active Loans</span>
            <strong style={{ color: "var(--text-secondary)", fontWeight: 600 }}>{data.repayments.length}</strong>
          </div>
        </div>
      </aside>

      {/* ─── Main Content Container ─── */}
      <div style={{ display: "flex", flexDirection: "column", flex: 1, minWidth: 0 }}>
        
        {/* ─── Top Bar (Seamless header) ─── */}
        <header className="topbar">
          <div className="topbar-left">
            {!sidebar && <button className="expand-btn" type="button" onClick={() => setSidebar(true)} title="Expand"><PanelLeft size={16} /></button>}
            <h2>{pageLabel}</h2>
            <div className="topbar-badge"><span className="live-dot" />Live</div>
          </div>
          <div className="topbar-right">
            <span className="topbar-tagline">
              Real-time lending analytics
            </span>

            {page !== "overview" && (
              <button className="action-pill" type="button" onClick={() => alert("Exporting page data...")}>
                <Download size={14} />
                <span>Export</span>
              </button>
            )}
            <button className="reset-btn" type="button" onClick={resetFilters}><FilterX size={13} /> Reset</button>
          </div>
        </header>

        {/* ─── Content ─── */}
        <main className="main-content">
          
          {/* Shared Filters Panel */}
          <div className="filters-row">
            <div className="filter-item"><label>Start Date</label><input type="date" value={filters.startDate} onChange={(e) => setFilters((f) => ({ ...f, startDate: e.target.value }))} /></div>
            <div className="filter-item"><label>End Date</label><input type="date" value={filters.endDate} onChange={(e) => setFilters((f) => ({ ...f, endDate: e.target.value }))} /></div>
            <div className="filter-item"><label>Branch</label><select value={filters.branch} onChange={(e) => setFilters((f) => ({ ...f, branch: e.target.value }))}>{opts.branches.map((o) => <option key={o}>{o}</option>)}</select></div>
            <div className="filter-item"><label>Loan Type</label><select value={filters.loanType} onChange={(e) => setFilters((f) => ({ ...f, loanType: e.target.value }))}>{opts.loanTypes.map((o) => <option key={o}>{o}</option>)}</select></div>
            <div className="filter-item"><label>Ticket Size</label><select value={filters.ticketSize} onChange={(e) => setFilters((f) => ({ ...f, ticketSize: e.target.value }))}>{opts.ticketSizes.map((o) => <option key={o}>{o}</option>)}</select></div>
            <div className="filter-item"><label>Assigned User</label><select value={filters.user} onChange={(e) => setFilters((f) => ({ ...f, user: e.target.value }))}>{opts.users.map((o) => <option key={o}>{o}</option>)}</select></div>
            <div className="filter-item"><label>Segment</label><select value={filters.customerSegment} onChange={(e) => setFilters((f) => ({ ...f, customerSegment: e.target.value }))}>{opts.segments.map((o) => <option key={o}>{o}</option>)}</select></div>
          </div>

          {/* ════════════════════ OVERVIEW PAGE ════════════════════ */}
          {page === "overview" && (
            <>
              {/* Row 1 of 5 original KPI cards */}
              <div className="kpi-row">
                
                {/* 1. Applications */}
                <div className="kpi">
                  <div className="kpi-header">
                    <div className="kpi-icon-wrap indigo"><ShoppingBag size={16} /></div>
                    <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                  </div>
                  <div className="kpi-body">
                    <span className="kpi-label">Applications</span>
                    <div className="kpi-value">{d.total.toLocaleString("en-IN")}</div>
                    <div className="kpi-sub"><span>Pipeline demand</span></div>
                  </div>
                </div>

                {/* 2. Approved */}
                <div className="kpi">
                  <div className="kpi-header">
                    <div className="kpi-icon-wrap green"><ShieldCheck size={16} /></div>
                    <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                  </div>
                  <div className="kpi-body">
                    <span className="kpi-label">Approved</span>
                    <div className="kpi-value">{d.approvedCount.toLocaleString("en-IN")}</div>
                    <div className="kpi-sub">
                      <span className="kpi-change-tag up">{fmtPct(d.approvalRate)}</span>
                      <span>rate</span>
                    </div>
                  </div>
                </div>

                {/* 3. Disbursed */}
                <div className="kpi">
                  <div className="kpi-header">
                    <div className="kpi-icon-wrap teal"><Activity size={16} /></div>
                    <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                  </div>
                  <div className="kpi-body">
                    <span className="kpi-label">Disbursed</span>
                    <div className="kpi-value">{d.disbursedCount.toLocaleString("en-IN")}</div>
                    <div className="kpi-sub"><span>Conversions</span></div>
                  </div>
                </div>

                {/* 4. Sanctioned */}
                <div className="kpi">
                  <div className="kpi-header">
                    <div className="kpi-icon-wrap blue"><Landmark size={16} /></div>
                    <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                  </div>
                  <div className="kpi-body">
                    <span className="kpi-label">Sanctioned</span>
                    <div className="kpi-value">{fmtAmt(d.sanctionedAmt)}</div>
                    <div className="kpi-sub"><span>Approved value</span></div>
                  </div>
                </div>

                {/* 5. Avg Ticket */}
                <div className="kpi">
                  <div className="kpi-header">
                    <div className="kpi-icon-wrap teal"><TrendingUp size={16} /></div>
                    <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                  </div>
                  <div className="kpi-body">
                    <span className="kpi-label">Avg Ticket</span>
                    <div className="kpi-value">{fmtAmt(d.avgTicket)}</div>
                    <div className="kpi-sub"><span>Per approved loan</span></div>
                  </div>
                </div>

              </div>

              {/* Secondary KPIs Row - 5 remaining indicators (Rearranged for 5-column grid) */}
              {showSecondaryMetrics && (
                <div className="secondary-metrics-wrap">
                  <div className="kpi-row">
                    
                    {/* 6. Rejection Rate */}
                    <div className="kpi">
                      <div className="kpi-header">
                        <div className="kpi-icon-wrap coral"><AlertTriangle size={16} /></div>
                        <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                      </div>
                      <div className="kpi-body">
                        <span className="kpi-label">Rejection Rate</span>
                        <div className="kpi-value">{fmtPct(d.rejectionRate)}</div>
                        <div className="kpi-sub">
                          <span className="kpi-change-tag down">{d.rejectedCount}</span>
                          <span>rejected</span>
                        </div>
                      </div>
                    </div>

                    {/* 7. Avg CIBIL */}
                    <div className="kpi">
                      <div className="kpi-header">
                        <div className="kpi-icon-wrap indigo"><Users size={16} /></div>
                        <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                      </div>
                      <div className="kpi-body">
                        <span className="kpi-label">Avg CIBIL</span>
                        <div className="kpi-value">{Math.round(d.avgCibil)}</div>
                        <div className="kpi-sub"><span>Applicant profile</span></div>
                      </div>
                    </div>

                    {/* 8. Collection Eff. */}
                    <div className="kpi">
                      <div className="kpi-header">
                        <div className="kpi-icon-wrap green"><TrendingUp size={16} /></div>
                        <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                      </div>
                      <div className="kpi-body">
                        <span className="kpi-label">Collection Eff.</span>
                        <div className="kpi-value">{fmtPct(d.collEff)}</div>
                        <div className="kpi-sub"><span>{fmtAmt(d.totalPaid)} collected</span></div>
                      </div>
                    </div>

                    {/* 9. NPA Rate */}
                    <div className="kpi">
                      <div className="kpi-header">
                        <div className="kpi-icon-wrap coral"><AlertTriangle size={16} /></div>
                        <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                      </div>
                      <div className="kpi-body">
                        <span className="kpi-label">NPA Rate</span>
                        <div className="kpi-value">{fmtPct(d.npaRate)}</div>
                        <div className="kpi-sub">
                          <span className="kpi-change-tag down">{d.npaCount}</span>
                          <span>of {d.activeCount} active</span>
                        </div>
                      </div>
                    </div>

                    {/* 10. Outstanding */}
                    <div className="kpi">
                      <div className="kpi-header">
                        <div className="kpi-icon-wrap blue"><Landmark size={16} /></div>
                        <button className="kpi-info-btn" type="button" title="View details"><HelpCircle size={14} /></button>
                      </div>
                      <div className="kpi-body">
                        <span className="kpi-label">Outstanding</span>
                        <div className="kpi-value">{fmtAmt(d.principalOut)}</div>
                        <div className="kpi-sub"><span>Active exposure</span></div>
                      </div>
                    </div>

                  </div>
                </div>
              )}

              {/* See More Metrics Toggle Button */}
              <div style={{ margin: "6px 0 10px 0", display: "flex", justifyContent: "flex-end" }}>
                <button 
                  className="action-pill" 
                  onClick={() => setShowSecondaryMetrics(!showSecondaryMetrics)}
                  type="button"
                >
                  <span>{showSecondaryMetrics ? "Hide Additional Metrics" : "See More Metrics"}</span>
                  <ChevronDown size={14} style={{ transform: showSecondaryMetrics ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
                </button>
              </div>

              {/* Redesigned Charts (Original Overview Selection) */}
              <div className="chart-grid full">
                {/* 1. Application & Disbursement Trend (Full width) */}
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group">
                      <h3>Application & Disbursement Trend</h3>
                      <p style={{ fontSize: "12px", color: "var(--text-tertiary)", marginTop: "2px" }}>Monthly pipeline demand vs completed disbursements</p>
                    </div>

                  </div>
                  
                  <ResponsiveContainer width="100%" height={280}>
                    <AreaChart data={d.monthly} margin={{ top: 8, right: 20, left: -20, bottom: 0 }}>
                      <defs>
                        <linearGradient id="areaGradApps" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#6366f1" stopOpacity={0.2} />
                          <stop offset="100%" stopColor="#6366f1" stopOpacity={0.0} />
                        </linearGradient>
                        <linearGradient id="areaGradDisb" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#0d9488" stopOpacity={0.2} />
                          <stop offset="100%" stopColor="#0d9488" stopOpacity={0.0} />
                        </linearGradient>
                      </defs>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} />
                      <YAxis allowDecimals={false} {...AXIS} />
                      <Tooltip />
                      <Area 
                        type="monotone" 
                        dataKey="applications" 
                        name="Applications" 
                        stroke="#6366f1" 
                        strokeWidth={2} 
                        fill="url(#areaGradApps)" 
                        dot={{ r: 4, strokeWidth: 0, fill: "#6366f1" }}
                        activeDot={{ r: 6 }}
                        isAnimationActive={false} 
                      />
                      <Area 
                        type="monotone" 
                        dataKey="disbursed" 
                        name="Disbursed" 
                        stroke="#0d9488" 
                        strokeWidth={2} 
                        fill="url(#areaGradDisb)" 
                        dot={{ r: 4, strokeWidth: 0, fill: "#0d9488" }}
                        activeDot={{ r: 6 }}
                        isAnimationActive={false} 
                      />
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="chart-grid">
                {/* 2. Branch Volume (Rounded Bar Chart with tracks) */}
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group">
                      <h3>Branch Volume</h3>
                      <p style={{ fontSize: "12px", color: "var(--text-tertiary)", marginTop: "2px" }}>Application count by branch location</p>
                    </div>
                  </div>
                  
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={d.branchApps} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} />
                      <YAxis allowDecimals={false} {...AXIS} />
                      <Tooltip />
                      <Bar 
                        dataKey="value" 
                        name="Applications" 
                        fill="#6366f1" 
                        radius={[8, 8, 0, 0]}
                        background={{ fill: "#f2f4f7", radius: 8 }}
                        isAnimationActive={false} 
                      />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                {/* 3. Product Mix (Pie Chart) */}
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group">
                      <h3>Product Mix</h3>
                      <p style={{ fontSize: "12px", color: "var(--text-tertiary)", marginTop: "2px" }}>Loan type distribution</p>
                    </div>
                  </div>
                  
                  <ResponsiveContainer width="100%" height={260}>
                    <PieChart>
                      <Pie 
                        data={d.loanMix} 
                        dataKey="value" 
                        nameKey="name" 
                        cx="50%" 
                        cy="50%" 
                        innerRadius={45} 
                        outerRadius={85} 
                        strokeWidth={2} 
                        stroke="#fff" 
                        isAnimationActive={false} 
                        label={({ name, value, cx, x, y }) => (
                          <text 
                            x={x} 
                            y={y} 
                            textAnchor={x > cx ? "start" : "end"} 
                            dominantBaseline="central" 
                            fill="var(--text-primary)" 
                            fontSize={11} 
                            fontWeight={500}
                          >
                            {name} ({value})
                          </text>
                        )}
                      >
                        {d.loanMix.map((_, i) => <Cell key={`ov-pie-${i}`} fill={C.pie[i % C.pie.length]} />)}
                      </Pie>
                      <Tooltip />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </>
          )}

          {/* ════════════════════ SALES PAGE ════════════════════ */}
          {page === "sales" && (
            <>
              <div className="kpi-row cols-4">
                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap coral"><AlertTriangle size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Rejection Rate</span>
                    <div className="kpi-value">{fmtPct(d.rejectionRate)}</div>
                    <div className="kpi-sub"><span className="kpi-change-tag down">{d.rejectedCount} rejected</span></div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap amber"><Clock size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Pending Review</span>
                    <div className="kpi-value">{d.pendingCount}</div>
                    <div className="kpi-sub">Cases awaiting credit check</div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap indigo"><Users size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Average CIBIL Score</span>
                    <div className="kpi-value">{Math.round(d.avgCibil)}</div>
                    <div className="kpi-sub">Applicants tier credit mix</div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap green"><Landmark size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Sanctioned Value</span>
                    <div className="kpi-value">{fmtAmt(d.sanctionedAmt)}</div>
                    <div className="kpi-sub">Total credit volume approved</div>
                  </div>
                </div>
              </div>

              <div className="chart-grid">
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Branch Performance</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={d.branchApps} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} />
                      <YAxis allowDecimals={false} {...AXIS} />
                      <Tooltip />
                      <Bar dataKey="value" name="Applications" fill={C.indigo} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Product Mix (Loan Types)</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <PieChart>
                      <Pie data={d.loanMix} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={50} outerRadius={85} strokeWidth={2} stroke="#fff" isAnimationActive={false} label={({ name, value, cx, x, y }) => <text x={x} y={y} textAnchor={x > cx ? "start" : "end"} dominantBaseline="central" fill="#101828" fontSize={11} fontWeight={500}>{name} ({value})</text>}>
                        {d.loanMix.map((_, i) => <Cell key={`pm-${i}`} fill={C.pie[i % C.pie.length]} />)}
                      </Pie>
                      <Tooltip />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="chart-panel" style={{ width: "100%" }}>
                <div className="chart-panel-header">
                  <div className="chart-title-group"><h3>Monthly Sanctioned Trend</h3></div>
                </div>
                <ResponsiveContainer width="100%" height={260}>
                  <LineChart data={d.monthly} margin={{ top: 8, right: 20, left: -20, bottom: 0 }}>
                    <CartesianGrid {...GRID} />
                    <XAxis dataKey="name" {...AXIS} />
                    <YAxis {...AXIS} />
                    <Tooltip formatter={(v: any) => [`₹${Number(v).toFixed(1)} L`, "Sanctioned"]} />
                    <Line type="monotone" dataKey="sanctionedAmt" name="Sanctioned (₹L)" stroke={C.indigo} strokeWidth={2} dot={{ r: 4, fill: C.indigo, strokeWidth: 0 }} isAnimationActive={false} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </>
          )}

          {/* ════════════════════ CREDIT PAGE ════════════════════ */}
          {page === "credit" && (
            <>
              <div className="kpi-row cols-4">
                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap indigo"><Users size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Average CIBIL</span>
                    <div className="kpi-value">{Math.round(d.avgCibil)}</div>
                    <div className="kpi-sub">Applicants risk scoring</div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap green"><ShieldCheck size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Approval Rate</span>
                    <div className="kpi-value">{fmtPct(d.approvalRate)}</div>
                    <div className="kpi-sub"><span className="kpi-change-tag up">{d.approvedCount} Approved</span></div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap coral"><AlertTriangle size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Rejection Rate</span>
                    <div className="kpi-value">{fmtPct(d.rejectionRate)}</div>
                    <div className="kpi-sub"><span className="kpi-change-tag down">{d.rejectedCount} Rejected</span></div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap blue"><Activity size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">NPA Rate</span>
                    <div className="kpi-value">{fmtPct(d.npaRate)}</div>
                    <div className="kpi-sub">{d.npaCount} accounts in NPA list</div>
                  </div>
                </div>
              </div>

              <div className="chart-grid">
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Approval Rate by CIBIL Tier</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={d.cibilBands} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} />
                      <YAxis tickFormatter={(v) => `${v}%`} {...AXIS} />
                      <Tooltip formatter={(v: any) => [`${v}%`, "Approval Rate"]} />
                      <Bar dataKey="approvalRate" name="Approval %" fill={C.indigo} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Segment Risk Analysis</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={d.segs} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} interval={0} tick={{ fontSize: 9 }} />
                      <YAxis tickFormatter={(v) => `${v}%`} {...AXIS} />
                      <Tooltip formatter={(v: any) => [`${v}%`]} />
                      <Bar dataKey="collEff" name="Collection Eff." fill={C.teal} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                      <Bar dataKey="npaRate" name="NPA Rate" fill={C.coral} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="chart-grid">
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Submissions by CIBIL Band</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={240}>
                    <BarChart data={d.cibilBands} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} />
                      <YAxis allowDecimals={false} {...AXIS} />
                      <Tooltip />
                      <Bar dataKey="count" name="Applications" fill={C.sage} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Segment Distribution</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={240}>
                    <BarChart data={d.segs} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} interval={0} tick={{ fontSize: 9 }} />
                      <YAxis allowDecimals={false} {...AXIS} />
                      <Tooltip />
                      <Bar dataKey="count" name="Applications" fill={C.indigo} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </>
          )}

          {/* ════════════════════ OPERATIONS PAGE ════════════════════ */}
          {page === "operations" && (
            <>
              <div className="kpi-row cols-4">
                {d.stageTat.map((s, idx) => {
                  const Icon = [Users, ShieldCheck, Clock, Landmark][idx % 4];
                  const color = ["indigo", "green", "amber", "blue"][idx % 4];
                  return (
                    <div className="kpi" key={s.name}>
                      <div className="kpi-header">
                        <div className={`kpi-icon-wrap ${color}`}>
                          <Icon size={20} />
                        </div>
                      </div>
                      <div className="kpi-body">
                        <span className="kpi-label">{s.name} TAT</span>
                        <div className="kpi-value">{s.hours} Hrs</div>
                        <div className="kpi-sub">
                          <span>{s.pending} cases pending</span>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              <div className="chart-grid">
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Stage Turnaround Time (TAT)</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={d.stageTat} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} interval={0} tick={{ fontSize: 9 }} />
                      <YAxis {...AXIS} />
                      <Tooltip />
                      <Bar dataKey="hours" name="Avg TAT (Hrs)" fill={C.indigo} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Stage Load (Pending backlog)</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={d.stageTat} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} interval={0} tick={{ fontSize: 9 }} />
                      <YAxis allowDecimals={false} {...AXIS} />
                      <Tooltip />
                      <Bar dataKey="pending" name="Pending Cases" fill={C.amber} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </>
          )}

          {/* ════════════════════ PORTFOLIO PAGE ════════════════════ */}
          {page === "portfolio" && (
            <>
              <div className="kpi-row cols-4">
                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap blue"><Landmark size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Principal Outstanding</span>
                    <div className="kpi-value">{fmtAmt(d.principalOut)}</div>
                    <div className="kpi-sub">Total active credit exposure</div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap green"><TrendingUp size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Collection Efficiency</span>
                    <div className="kpi-value">{fmtPct(d.collEff)}</div>
                    <div className="kpi-sub">{fmtAmt(d.totalPaid)} collected out of {fmtAmt(d.totalDue)} due</div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap coral"><AlertTriangle size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">NPA Rate</span>
                    <div className="kpi-value">{fmtPct(d.npaRate)}</div>
                    <div className="kpi-sub">{d.npaCount} active NPA loan accounts</div>
                  </div>
                </div>

                <div className="kpi">
                  <div className="kpi-header"><div className="kpi-icon-wrap indigo"><Users size={20} /></div></div>
                  <div className="kpi-body">
                    <span className="kpi-label">Active Loans</span>
                    <div className="kpi-value">{d.activeCount.toLocaleString("en-IN")}</div>
                    <div className="kpi-sub">With active repayment schedules</div>
                  </div>
                </div>
              </div>

              <div className="chart-grid">
                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Days Past Due (DPD) Distribution</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={d.dpdDist} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} />
                      <YAxis allowDecimals={false} {...AXIS} />
                      <Tooltip />
                      <Bar dataKey="value" name="Loans" isAnimationActive={false} radius={[6, 6, 0, 0]}>
                        {d.dpdDist.map((e, i) => <Cell key={`dpd-${i}`} fill={e.name === "0 DPD" ? C.teal : e.name === "1-30" ? C.sage : e.name === "31-60" ? C.amber : C.coral} />)}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="chart-panel">
                  <div className="chart-panel-header">
                    <div className="chart-title-group"><h3>Sanctioned Amount Trend</h3></div>
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <AreaChart data={d.monthly} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                      <defs>
                        <linearGradient id="indigoAreaGrad" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#6366f1" stopOpacity={0.15} />
                          <stop offset="100%" stopColor="#6366f1" stopOpacity={0.00} />
                        </linearGradient>
                      </defs>
                      <CartesianGrid {...GRID} />
                      <XAxis dataKey="name" {...AXIS} />
                      <YAxis {...AXIS} />
                      <Tooltip formatter={(v: any) => [`₹${Number(v).toFixed(1)} L`, "Sanctioned"]} />
                      <Area type="monotone" dataKey="sanctionedAmt" name="Sanctioned (₹L)" stroke={C.indigo} strokeWidth={2} fill="url(#indigoAreaGrad)" dot={{ r: 4, fill: C.indigo, strokeWidth: 0 }} isAnimationActive={false} />
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </>
          )}

          {/* ════════════════════ PENDING QUEUE PAGE ════════════════════ */}
          {page === "queue" && (() => {
            const pendByStage = groupCount(d.pendingOps, (r) => stageLabel(r.stage));
            const pendByUser = groupCount(d.pendingOps, (r) => r.assignedTo).slice(0, 6);
            return (
              <>
                <div className="kpi-row cols-4">
                  <div className="kpi">
                    <div className="kpi-header"><div className="kpi-icon-wrap amber"><Clock size={20} /></div></div>
                    <div className="kpi-body">
                      <span className="kpi-label">Pending Backlog</span>
                      <div className="kpi-value">{d.pendingOps.length}</div>
                      <div className="kpi-sub">Cases requiring stage reviews</div>
                    </div>
                  </div>

                  <div className="kpi">
                    <div className="kpi-header"><div className="kpi-icon-wrap blue"><Activity size={20} /></div></div>
                    <div className="kpi-body">
                      <span className="kpi-label">Stages Active</span>
                      <div className="kpi-value">{new Set(d.pendingOps.map((r) => r.stage)).size}</div>
                      <div className="kpi-sub">Departments with pending tasks</div>
                    </div>
                  </div>

                  <div className="kpi">
                    <div className="kpi-header"><div className="kpi-icon-wrap indigo"><Users size={20} /></div></div>
                    <div className="kpi-body">
                      <span className="kpi-label">Assigned Credit Officers</span>
                      <div className="kpi-value">{new Set(d.pendingOps.map((r) => r.assignedTo)).size}</div>
                      <div className="kpi-sub">Handling pending reviews</div>
                    </div>
                  </div>

                  <div className="kpi">
                    <div className="kpi-header"><div className="kpi-icon-wrap green"><Landmark size={20} /></div></div>
                    <div className="kpi-body">
                      <span className="kpi-label">Locations (Branches)</span>
                      <div className="kpi-value">{new Set(d.pendingOps.map((r) => r.branch)).size}</div>
                      <div className="kpi-sub">Branches with backlog queue</div>
                    </div>
                  </div>
                </div>

                <div className="chart-grid">
                  <div className="chart-panel">
                    <div className="chart-panel-header">
                      <div className="chart-title-group"><h3>Pending by Stage</h3></div>
                    </div>
                    <ResponsiveContainer width="100%" height={240}>
                      <BarChart data={pendByStage} margin={{ top: 8, right: 16, left: -20, bottom: 0 }}>
                        <CartesianGrid {...GRID} />
                        <XAxis dataKey="name" {...AXIS} interval={0} tick={{ fontSize: 9 }} />
                        <YAxis allowDecimals={false} {...AXIS} />
                        <Tooltip />
                        <Bar dataKey="value" name="Pending Review" fill={C.amber} radius={[6, 6, 0, 0]} isAnimationActive={false} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="chart-panel">
                    <div className="chart-panel-header">
                      <div className="chart-title-group"><h3>Officer Workload Distribution</h3></div>
                    </div>
                    <ResponsiveContainer width="100%" height={240}>
                      <BarChart data={pendByUser} layout="vertical" margin={{ top: 8, right: 16, left: 10, bottom: 0 }}>
                        <CartesianGrid {...GRID} horizontal={false} vertical />
                        <XAxis type="number" allowDecimals={false} {...AXIS} />
                        <YAxis type="category" dataKey="name" width={80} {...AXIS} tick={{ fontSize: 9 }} />
                        <Tooltip />
                        <Bar dataKey="value" name="Pending Cases" fill={C.indigo} radius={[0, 6, 6, 0]} isAnimationActive={false} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="data-table-wrap">
                  <div className="data-table-header">
                    <div className="data-table-header-left">
                      <h3>Operations Queue List</h3>
                      <p>{d.pendingOps.length} credit approvals awaiting action</p>
                    </div>
                    <div className="data-table-header-actions">
                      <button className="action-pill" type="button" onClick={() => alert("Downloading queue logs CSV...")}><Download size={14} /><span>Export</span></button>
                    </div>
                  </div>
                  <div className="data-table-scroll">
                    <table>
                      <thead>
                        <tr>
                          <th className="table-checkbox-col"><input type="checkbox" className="table-checkbox" /></th>
                          <th>Application ID</th>
                          <th>Stage Name</th>
                          <th>Officer Assigned</th>
                          <th>Stage Start Date</th>
                          <th>Branch Location</th>
                          <th>Loan Type</th>
                          <th>Risk Status</th>
                          <th className="table-action-col"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {d.pendingOps.length === 0 ? (
                          <tr><td colSpan={9} className="empty-state">No pending reviews found.</td></tr>
                        ) : (
                          d.pendingOps.map((r) => (
                            <tr key={r.logId}>
                              <td className="table-checkbox-col"><input type="checkbox" className="table-checkbox" /></td>
                              <td className="table-client-name">{r.applicationId}</td>
                              <td>{stageLabel(r.stage)}</td>
                              <td>{r.assignedTo}</td>
                              <td>{r.stageStart.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })}</td>
                              <td>{r.branch}</td>
                              <td>{r.loanType}</td>
                              <td>
                                <span className="badge pending">
                                  <span className="badge-dot" />
                                  <span>Pending Approval</span>
                                </span>
                              </td>
                              <td className="table-action-col">
                                <button className="table-action-btn" type="button"><MoreHorizontal size={14} /></button>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            );
          })()}

        </main>
      </div>
    </div>
  );
}

export default App;
