# Bhutta Marble & Granite Showroom — Marble Ledger Source Code

**Generated:** 23 June 2026  
**Stack:** React + Vite + TypeScript + Clerk Auth  
**Features:** Area · Skirting · Border · Katba calculator · Editable rates · Print/PDF · Excel export · Quote history · Private login

---

## How to Rebuild This App

1. Create a new **React + Vite + TypeScript** project on Replit
2. Install these packages:
   ```
   @clerk/react  @clerk/themes  wouter  xlsx
   ```
3. Replace `src/App.tsx` with the code in **Section 1** below
4. Replace `src/index.css` with the code in **Section 2** below
5. Set these environment variables:
   - `VITE_CLERK_PUBLISHABLE_KEY` — your Clerk publishable key
   - `VITE_ALLOWED_EMAILS` — comma-separated emails allowed to access the app
   - `VITE_CLERK_PROXY_URL` — set by Replit Clerk integration automatically

---

## Section 1: Main App Code — `src/App.tsx`  (724 lines)

```tsx
import { useState, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  ClerkProvider,
  SignIn,
  useClerk,
  useUser,
} from "@clerk/react";
import { publishableKeyFromHost } from "@clerk/react/internal";
import { shadcn } from "@clerk/themes";
import { Switch, Route, Router as WouterRouter, useLocation, Redirect } from "wouter";

// ── Clerk setup (copy verbatim per skill) ────────────────────────────────────
const clerkPubKey = publishableKeyFromHost(
  window.location.hostname,
  import.meta.env.VITE_CLERK_PUBLISHABLE_KEY,
);
const clerkProxyUrl = import.meta.env.VITE_CLERK_PROXY_URL;
const basePath = import.meta.env.BASE_URL.replace(/\/$/, "");

function stripBase(path: string): string {
  return basePath && path.startsWith(basePath)
    ? path.slice(basePath.length) || "/"
    : path;
}

if (!clerkPubKey) {
  throw new Error("Missing VITE_CLERK_PUBLISHABLE_KEY");
}

const ALLOWED_EMAILS = (import.meta.env.VITE_ALLOWED_EMAILS ?? "")
  .split(",")
  .map((e: string) => e.trim().toLowerCase())
  .filter(Boolean);

// ── Marble data ───────────────────────────────────────────────────────────────
type MarbleItem = { n: string; r: number };

const DEFAULT_MARBLE: MarbleItem[] = [
  { n: "Verona Dhabba", r: 100 },
  { n: "White Gray (A)", r: 86.5 },
  { n: "White Gray (B)", r: 82.5 },
  { n: "Strawberry (Red & White)", r: 120 },
  { n: "Ziyārat Gray", r: 78 },
  { n: "Silky Black", r: 85 },
  { n: "Koshainik", r: 95 },
  { n: "Royal Fancy", r: 105 },
  { n: "Ziyārat Gray White", r: 100 },
  { n: "Tavīra", r: 92 },
  { n: "Verona Plain", r: 130 },
  { n: "White Gray (Khānka White)", r: 105 },
  { n: "Bādal Dowezi", r: 75 },
  { n: "Rehmān White", r: 115 },
  { n: "Tippi Būtsīna", r: 95 },
];
const DEFAULT_BORDER: MarbleItem[] = [
  { n: "Mehroon", r: 55 },
  { n: "Black 2''", r: 25 },
  { n: "Black 3''", r: 30 },
  { n: "Strawberry", r: 40 },
];
const DEFAULT_BORDER_DESIGN: MarbleItem[] = [
  { n: "Design 4''", r: 115 },
  { n: "Design 6''", r: 125 },
];
const DEFAULT_KATBA: MarbleItem[] = [
  { n: "Small (12''×15'')", r: 1600 },
  { n: "Medium (A) (12''×18'')", r: 2000 },
  { n: "Medium (I) (12''×18'')", r: 2500 },
  { n: "Large (A) (12''×24'')", r: 3000 },
];

// ── Types ─────────────────────────────────────────────────────────────────────
interface RoomRow {
  id: number; name: string;
  lf: string; li: string; wf: string; wi: string;
  floorMarble: string; floorCustomRate: string;
  borderMarble: string; borderCustomRate: string;
}
interface KatbaRow {
  id: number; location: string;
  type: string; typeCustomRate: string; qty: string;
}
interface RoomResult {
  name: string; L: number; W: number;
  area: number; areaCost: number;
  skirt: number; skirtCost: number;
  border: number; borderCost: number;
  floorName: string; floorRate: number;
  borderName: string; borderRate: number;
}
interface KatbaResult {
  location: string; typeName: string; rate: number; qty: number; cost: number;
}
interface Results {
  rooms: RoomResult[];
  katbaItems: KatbaResult[];
  totals: {
    area: number; areaCost: number;
    skirt: number; skirtCost: number;
    border: number; borderCost: number;
    katbaQty: number; katbaCost: number; grand: number;
  };
}

interface HistoryEntry {
  id: string;
  clientName: string;
  savedAt: string; // ISO string
  results: Results;
}

const HISTORY_KEY = "bhutta_marble_history";

function loadHistory(): HistoryEntry[] {
  try {
    return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
  } catch {
    return [];
  }
}

function saveHistory(entries: HistoryEntry[]) {
  localStorage.setItem(HISTORY_KEY, JSON.stringify(entries));
}

let nextId = 1;
function makeRoom(num: number, f: string, b: string): RoomRow {
  return { id: nextId++, name: `Room ${num}`, lf: "", li: "", wf: "", wi: "", floorMarble: f, floorCustomRate: "", borderMarble: b, borderCustomRate: "" };
}
function makeKatba(t: string): KatbaRow {
  return { id: nextId++, location: "", type: t, typeCustomRate: "", qty: "1" };
}
function ftin(d: number) {
  const t = Math.round(d * 12); return `${Math.floor(t / 12)}'${t % 12}"`;
}
function getRateForName(list: MarbleItem[], name: string) {
  return list.find((m) => m.n === name)?.r ?? 0;
}

// ── Sub-components ────────────────────────────────────────────────────────────
function MarbleSelect({ value, customRate, onChange, onCustomRateChange, list, borderList, borderDesignList, hasGroups }: {
  value: string; customRate: string;
  onChange: (v: string) => void; onCustomRateChange: (v: string) => void;
  list?: MarbleItem[]; borderList?: MarbleItem[]; borderDesignList?: MarbleItem[]; hasGroups?: boolean;
}) {
  return (
    <div>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        {hasGroups ? (
          <>
            <optgroup label="Simple">{(borderList || []).map((m) => <option key={m.n} value={m.n}>{m.n} — Rs {m.r}/-</option>)}</optgroup>
            <optgroup label="Design">{(borderDesignList || []).map((m) => <option key={m.n} value={m.n}>{m.n} — Rs {m.r}/-</option>)}</optgroup>
            <option value="Custom">Custom rate…</option>
          </>
        ) : (
          <>{(list || []).map((m) => <option key={m.n} value={m.n}>{m.n} — Rs {m.r}/-</option>)}<option value="Custom">Custom rate…</option></>
        )}
      </select>
      {value === "Custom" && (
        <input type="number" placeholder="Rs/-" className="custom-rate-input" value={customRate} onChange={(e) => onCustomRateChange(e.target.value)} />
      )}
    </div>
  );
}

function RateGroup({ title, unit, items, onChange }: { title: string; unit: string; items: MarbleItem[]; onChange: (i: number, r: number) => void }) {
  return (
    <div className="rate-group">
      <div className="rate-group-title">{title} <span className="rate-unit">{unit}</span></div>
      <table className="rate-table">
        <thead><tr><th>Name</th><th style={{ width: 110 }}>Rate (Rs)</th></tr></thead>
        <tbody>
          {items.map((m, idx) => (
            <tr key={m.n}>
              <td>{m.n}</td>
              <td><div className="rate-input-wrap"><input type="number" min="0" step="0.5" value={m.r} onChange={(e) => onChange(idx, parseFloat(e.target.value) || 0)} /></div></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ── Ledger (the main app content) ─────────────────────────────────────────────
function Ledger() {
  const { signOut } = useClerk();
  const { user } = useUser();

  const [marble, setMarble] = useState<MarbleItem[]>(DEFAULT_MARBLE.map((m) => ({ ...m })));
  const [border, setBorder] = useState<MarbleItem[]>(DEFAULT_BORDER.map((m) => ({ ...m })));
  const [borderDesign, setBorderDesign] = useState<MarbleItem[]>(DEFAULT_BORDER_DESIGN.map((m) => ({ ...m })));
  const [katbaList, setKatbaList] = useState<MarbleItem[]>(DEFAULT_KATBA.map((m) => ({ ...m })));
  const [rooms, setRooms] = useState<RoomRow[]>(() => [makeRoom(1, DEFAULT_MARBLE[0].n, DEFAULT_BORDER[0].n), makeRoom(2, DEFAULT_MARBLE[0].n, DEFAULT_BORDER[0].n)]);
  const [katbaRows, setKatbaRows] = useState<KatbaRow[]>(() => [makeKatba(DEFAULT_KATBA[0].n)]);
  const [results, setResults] = useState<Results | null>(null);
  const [ratesOpen, setRatesOpen] = useState(false);
  const [history, setHistory] = useState<HistoryEntry[]>(() => loadHistory());
  const [historyOpen, setHistoryOpen] = useState(false);
  const [clientName, setClientName] = useState("");
  const [savedId, setSavedId] = useState<string | null>(null);
  const [expandedId, setExpandedId] = useState<string | null>(null);

  const allBorderList = [...border, ...borderDesign];

  const resetRates = useCallback(() => {
    setMarble(DEFAULT_MARBLE.map((m) => ({ ...m })));
    setBorder(DEFAULT_BORDER.map((m) => ({ ...m })));
    setBorderDesign(DEFAULT_BORDER_DESIGN.map((m) => ({ ...m })));
    setKatbaList(DEFAULT_KATBA.map((m) => ({ ...m })));
  }, []);

  const updateRate = useCallback(
    (setter: React.Dispatch<React.SetStateAction<MarbleItem[]>>) => (idx: number, r: number) => {
      setter((prev) => prev.map((m, i) => (i === idx ? { ...m, r } : m)));
    }, []
  );

  const saveToHistory = useCallback((res: Results) => {
    const name = clientName.trim() || "Unnamed Client";
    const entry: HistoryEntry = {
      id: Date.now().toString(),
      clientName: name,
      savedAt: new Date().toISOString(),
      results: res,
    };
    const updated = [entry, ...history];
    setHistory(updated);
    saveHistory(updated);
    setSavedId(entry.id);
    setHistoryOpen(true);
    setTimeout(() => setSavedId(null), 3000);
  }, [clientName, history]);

  const deleteHistory = useCallback((id: string) => {
    const updated = history.filter((e) => e.id !== id);
    setHistory(updated);
    saveHistory(updated);
    if (expandedId === id) setExpandedId(null);
  }, [history, expandedId]);

  const addRoom = useCallback(() => setRooms((p) => [...p, makeRoom(p.length + 1, marble[0].n, border[0].n)]), [marble, border]);
  const removeRoom = useCallback((id: number) => setRooms((p) => p.filter((r) => r.id !== id)), []);
  const updateRoom = useCallback((id: number, f: keyof RoomRow, v: string) => setRooms((p) => p.map((r) => r.id === id ? { ...r, [f]: v } : r)), []);
  const addKatba = useCallback(() => setKatbaRows((p) => [...p, makeKatba(katbaList[0].n)]), [katbaList]);
  const removeKatba = useCallback((id: number) => setKatbaRows((p) => p.filter((k) => k.id !== id)), []);
  const updateKatba = useCallback((id: number, f: keyof KatbaRow, v: string) => setKatbaRows((p) => p.map((k) => k.id === id ? { ...k, [f]: v } : k)), []);

  const calculate = useCallback(() => {
    const roomResults: RoomResult[] = [];
    const totals = { area: 0, areaCost: 0, skirt: 0, skirtCost: 0, border: 0, borderCost: 0, katbaQty: 0, katbaCost: 0, grand: 0 };
    for (const row of rooms) {
      const lf = parseFloat(row.lf) || 0, li = parseFloat(row.li) || 0, wf = parseFloat(row.wf) || 0, wi = parseFloat(row.wi) || 0;
      if (!lf && !li && !wf && !wi) continue;
      const L = lf + li / 12, W = wf + wi / 12;
      let floorRate: number, floorName: string;
      if (row.floorMarble === "Custom") { floorRate = parseFloat(row.floorCustomRate) || 0; floorName = `Custom (Rs ${floorRate}/-)`;
      } else { floorRate = getRateForName(marble, row.floorMarble); floorName = row.floorMarble; }
      let borderRate: number, borderName: string;
      if (row.borderMarble === "Custom") { borderRate = parseFloat(row.borderCustomRate) || 0; borderName = `Custom (Rs ${borderRate}/-)`;
      } else { borderRate = getRateForName(allBorderList, row.borderMarble); borderName = row.borderMarble; }
      const area = L * W, areaCost = area * floorRate;
      const skirt = 2 * (L + W), skirtCost = skirt * floorRate;
      const innerL = Math.max(L - 2, 0), innerW = Math.max(W - 2, 0);
      const bord = 2 * (innerL + innerW), borderCost = bord * borderRate;
      totals.area += area; totals.areaCost += areaCost;
      totals.skirt += skirt; totals.skirtCost += skirtCost;
      totals.border += bord; totals.borderCost += borderCost;
      roomResults.push({ name: row.name || "Room", L, W, area, areaCost, skirt, skirtCost, border: bord, borderCost, floorName, floorRate, borderName, borderRate });
    }
    const katbaItems: KatbaResult[] = [];
    for (const row of katbaRows) {
      const qty = parseFloat(row.qty) || 0; if (!qty) continue;
      let rate: number, typeName: string;
      if (row.type === "Custom") { rate = parseFloat(row.typeCustomRate) || 0; typeName = `Custom (Rs ${rate}/-)`;
      } else { rate = getRateForName(katbaList, row.type); typeName = row.type; }
      const cost = qty * rate;
      totals.katbaQty += qty; totals.katbaCost += cost;
      katbaItems.push({ location: row.location || "—", typeName, rate, qty, cost });
    }
    totals.grand = totals.areaCost + totals.skirtCost + totals.borderCost + totals.katbaCost;
    setResults({ rooms: roomResults, katbaItems, totals });
    setTimeout(() => document.getElementById("results-section")?.scrollIntoView({ behavior: "smooth" }), 50);
  }, [rooms, katbaRows, marble, allBorderList, katbaList]);

  const exportSpreadsheet = useCallback((res: Results) => {
    const wb = XLSX.utils.book_new();
    const rows: (string | number)[][] = [
      ["Bhutta Marble & Granite Showroom — Marble Ledger"], [],
      ["Room Results"],
      ["Room", "Length (ft)", "Width (ft)", "Floor Marble", "Rate (Rs/sqft)", "Area (sqft)", "Area Cost (Rs)", "Skirting (run ft)", "Skirting Cost (Rs)", "Border Marble", "Border Rate (Rs/run ft)", "Border (run ft)", "Border Cost (Rs)", "Room Total (Rs)"],
    ];
    for (const r of res.rooms) {
      rows.push([r.name, +r.L.toFixed(2), +r.W.toFixed(2), r.floorName, r.floorRate, +r.area.toFixed(2), +r.areaCost.toFixed(0), +r.skirt.toFixed(2), +r.skirtCost.toFixed(0), r.borderName, r.borderRate, +r.border.toFixed(2), +r.borderCost.toFixed(0), +(r.areaCost + r.skirtCost + r.borderCost).toFixed(0)]);
    }
    if (res.katbaItems.length) {
      rows.push([], ["Katba (Decorative Tiles)"], ["Location", "Type", "Rate (Rs/piece)", "Qty", "Cost (Rs)"]);
      for (const k of res.katbaItems) rows.push([k.location, k.typeName, k.rate, k.qty, +k.cost.toFixed(0)]);
    }
    rows.push([], ["Combined Totals"], ["Metric", "Value", "Cost (Rs)"],
      ["Total Floor Area (sqft)", +res.totals.area.toFixed(2), +res.totals.areaCost.toFixed(0)],
      ["Total Skirting (run ft)", +res.totals.skirt.toFixed(2), +res.totals.skirtCost.toFixed(0)],
      ["Total Border (run ft)", +res.totals.border.toFixed(2), +res.totals.borderCost.toFixed(0)],
      ["Total Katba (pcs)", res.totals.katbaQty, +res.totals.katbaCost.toFixed(0)],
      ["Grand Total (Rs)", "", +res.totals.grand.toFixed(0)],
    );
    const ws = XLSX.utils.aoa_to_sheet(rows as unknown as XLSX.CellObject[][]);
    ws["!cols"] = [{ wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 28 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 18 }, { wch: 20 }, { wch: 14 }, { wch: 16 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, ws, "Marble Ledger");
    XLSX.writeFile(wb, `Bhutta-Marble-Ledger-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, []);

  const exportHistory = useCallback(() => {
    if (history.length === 0) return;
    const wb = XLSX.utils.book_new();

    // ── Sheet 1: Summary (one row per saved quote) ──────────────────────────
    const summaryRows: (string | number)[][] = [
      ["Bhutta Marble & Granite Showroom — Quote History"],
      [`Exported: ${new Date().toLocaleString("en-PK")}`],
      [],
      ["#", "Client Name", "Date", "Time", "Rooms", "Floor Area (sqft)", "Area Cost (Rs)", "Skirting Cost (Rs)", "Border Cost (Rs)", "Katba Cost (Rs)", "Grand Total (Rs)"],
    ];
    history.forEach((e, i) => {
      const d = new Date(e.savedAt);
      summaryRows.push([
        i + 1,
        e.clientName,
        d.toLocaleDateString("en-PK", { day: "2-digit", month: "short", year: "numeric" }),
        d.toLocaleTimeString("en-PK", { hour: "2-digit", minute: "2-digit" }),
        e.results.rooms.length,
        +e.results.totals.area.toFixed(2),
        +e.results.totals.areaCost.toFixed(0),
        +e.results.totals.skirtCost.toFixed(0),
        +e.results.totals.borderCost.toFixed(0),
        +e.results.totals.katbaCost.toFixed(0),
        +e.results.totals.grand.toFixed(0),
      ]);
    });
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows as unknown as XLSX.CellObject[][]);
    wsSummary["!cols"] = [{ wch: 4 }, { wch: 22 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 16 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

    // ── Sheet 2+: one sheet per saved quote ─────────────────────────────────
    history.forEach((e) => {
      const d = new Date(e.savedAt);
      const label = `${e.clientName.slice(0, 20)} ${d.toLocaleDateString("en-PK", { day: "2-digit", month: "short" })}`;
      const rows: (string | number)[][] = [
        [`Client: ${e.clientName}`],
        [`Date: ${d.toLocaleDateString("en-PK", { day: "2-digit", month: "long", year: "numeric" })} ${d.toLocaleTimeString("en-PK", { hour: "2-digit", minute: "2-digit" })}`],
        [],
        ["Room", "Length (ft)", "Width (ft)", "Floor Marble", "Rate", "Area (sqft)", "Area Cost", "Skirting (run ft)", "Skirting Cost", "Border Marble", "Border Rate", "Border (run ft)", "Border Cost", "Room Total (Rs)"],
      ];
      for (const r of e.results.rooms) {
        rows.push([r.name, +r.L.toFixed(2), +r.W.toFixed(2), r.floorName, r.floorRate, +r.area.toFixed(2), +r.areaCost.toFixed(0), +r.skirt.toFixed(2), +r.skirtCost.toFixed(0), r.borderName, r.borderRate, +r.border.toFixed(2), +r.borderCost.toFixed(0), +(r.areaCost + r.skirtCost + r.borderCost).toFixed(0)]);
      }
      if (e.results.katbaItems.length) {
        rows.push([], ["Katba (Decorative Tiles)"], ["Location", "Type", "Rate (Rs/piece)", "Qty", "Cost (Rs)"]);
        for (const k of e.results.katbaItems) rows.push([k.location, k.typeName, k.rate, k.qty, +k.cost.toFixed(0)]);
      }
      rows.push([], ["Grand Total (Rs)", "", "", "", "", "", "", "", "", "", "", "", "", +e.results.totals.grand.toFixed(0)]);
      const ws = XLSX.utils.aoa_to_sheet(rows as unknown as XLSX.CellObject[][]);
      ws["!cols"] = [{ wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 26 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 22 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 14 }];
      XLSX.utils.book_append_sheet(wb, ws, label.slice(0, 31));
    });

    XLSX.writeFile(wb, `Bhutta-Quote-History-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [history]);

  return (
    <div className="wrap">
      <header className="title-block">
        <h1>Bhutta Marble &amp; <span>Granite</span> Showroom</h1>
        <div className="sub" style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
          <span>MARBLE LEDGER · AREA · SKIRTING · BORDER</span>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: "0.72rem", color: "var(--ink-soft)" }}>{user?.primaryEmailAddress?.emailAddress}</span>
            <button
              onClick={() => signOut({ redirectUrl: basePath || "/" })}
              style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: "0.7rem", border: "1px solid var(--stone-300)", background: "transparent", color: "var(--ink-soft)", padding: "4px 10px", borderRadius: 3, cursor: "pointer" }}
            >
              Sign out
            </button>
          </div>
        </div>
      </header>

      {/* Quote History Panel */}
      <div className="panel rates-panel">
        <button className="rates-toggle" onClick={() => setHistoryOpen((o) => !o)} aria-expanded={historyOpen}>
          <span className="rates-toggle-icon">{historyOpen ? "▾" : "▸"}</span>
          Quote History
          <span className="rates-badge">{history.length} saved</span>
        </button>
        {historyOpen && (
          <div className="rates-body">
            {history.length === 0 ? (
              <p className="empty" style={{ padding: "12px 0" }}>No saved quotes yet. Calculate a ledger and save it.</p>
            ) : (
              <div className="history-list">
                {history.map((entry) => {
                  const d = new Date(entry.savedAt);
                  const dateStr = d.toLocaleDateString("en-PK", { day: "2-digit", month: "short", year: "numeric" });
                  const timeStr = d.toLocaleTimeString("en-PK", { hour: "2-digit", minute: "2-digit" });
                  const isExpanded = expandedId === entry.id;
                  return (
                    <div key={entry.id} className={`history-entry${savedId === entry.id ? " history-entry--new" : ""}`}>
                      <div className="history-row">
                        <button className="history-expand" onClick={() => setExpandedId(isExpanded ? null : entry.id)}>
                          <span className="rates-toggle-icon">{isExpanded ? "▾" : "▸"}</span>
                          <strong>{entry.clientName}</strong>
                          <span className="history-meta">{dateStr} · {timeStr}</span>
                          <span className="history-total">Rs {entry.results.totals.grand.toFixed(0)}/-</span>
                        </button>
                        <button className="history-del" title="Delete" onClick={() => deleteHistory(entry.id)}>✕</button>
                      </div>
                      {isExpanded && (
                        <div className="history-detail">
                          <table className="totals">
                            <thead><tr><th>Metric</th><th>Value</th><th>Cost (Rs)</th></tr></thead>
                            <tbody>
                              {entry.results.rooms.map((r, i) => (
                                <tr key={i}><td>{r.name}</td><td>{r.area.toFixed(2)} sq ft · {r.skirt.toFixed(2)} run ft skirting</td><td>{(r.areaCost + r.skirtCost + r.borderCost).toFixed(0)}</td></tr>
                              ))}
                              {entry.results.katbaItems.map((k, i) => (
                                <tr key={`k${i}`}><td>{k.location || "Katba"}</td><td>{k.qty} pcs · {k.typeName}</td><td>{k.cost.toFixed(0)}</td></tr>
                              ))}
                              <tr className="grand"><td colSpan={2}>Grand Total</td><td>{entry.results.totals.grand.toFixed(0)}/-</td></tr>
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
            {history.length > 0 && (
              <div className="rates-footer" style={{ marginTop: 14 }}>
                <button className="export-btn" onClick={exportHistory}>
                  ⬇ Export All History (.xlsx)
                </button>
                <span className="print-hint" style={{ marginLeft: 10 }}>{history.length} quote{history.length !== 1 ? "s" : ""} — Summary + one sheet per client</span>
              </div>
            )}
          </div>
        )}
      </div>

      {/* Edit Rates Panel */}
      <div className="panel rates-panel">
        <button className="rates-toggle" onClick={() => setRatesOpen((o) => !o)} aria-expanded={ratesOpen}>
          <span className="rates-toggle-icon">{ratesOpen ? "▾" : "▸"}</span>
          Edit Marble Rates
          <span className="rates-badge">update prices</span>
        </button>
        {ratesOpen && (
          <div className="rates-body">
            <p className="hint" style={{ marginBottom: 20 }}>Adjust any rate below — changes apply immediately when you calculate.</p>
            <div className="rates-grid">
              <RateGroup title="Floor Marbles" unit="Rs / sq ft" items={marble} onChange={updateRate(setMarble)} />
              <RateGroup title="Border — Simple" unit="Rs / running ft" items={border} onChange={updateRate(setBorder)} />
              <RateGroup title="Border — Design" unit="Rs / running ft" items={borderDesign} onChange={updateRate(setBorderDesign)} />
              <RateGroup title="Katba (Decorative Tiles)" unit="Rs / piece" items={katbaList} onChange={updateRate(setKatbaList)} />
            </div>
            <div className="rates-footer">
              <button className="reset-btn" onClick={resetRates}>↺ Reset to defaults</button>
            </div>
          </div>
        )}
      </div>

      {/* Katba Panel */}
      <div className="panel">
        <h2>Katba (Decorative Tiles)</h2>
        <p className="hint">Optional. Add Katba pieces by type and quantity — priced per piece, not by area.</p>
        <table>
          <thead><tr><th style={{ width: "30%" }}>Room / Location</th><th style={{ width: "40%" }}>Katba Type</th><th style={{ width: "20%" }}>Quantity</th><th style={{ width: "10%" }}></th></tr></thead>
          <tbody>
            {katbaRows.map((row) => (
              <tr key={row.id}>
                <td><input type="text" placeholder="e.g. Hall entrance" value={row.location} onChange={(e) => updateKatba(row.id, "location", e.target.value)} /></td>
                <td><MarbleSelect value={row.type} customRate={row.typeCustomRate} onChange={(v) => updateKatba(row.id, "type", v)} onCustomRateChange={(v) => updateKatba(row.id, "typeCustomRate", v)} list={katbaList} /></td>
                <td><input type="number" min="0" placeholder="qty" className="dim-num" value={row.qty} onChange={(e) => updateKatba(row.id, "qty", e.target.value)} /></td>
                <td><button className="rm-btn" onClick={() => removeKatba(row.id)}>✕</button></td>
              </tr>
            ))}
            {!katbaRows.length && <tr><td colSpan={4}><p className="empty">No Katba rows yet.</p></td></tr>}
          </tbody>
        </table>
        <div className="row-actions"><button className="add" onClick={addKatba}>+ Add Katba</button></div>
      </div>

      {/* Room Measurements */}
      <div className="panel">
        <h2>Room Measurements</h2>
        <p className="hint">Enter length and width in feet &amp; inches. Pick the floor marble (used for area + skirting) and the border marble (inset border line). Border is fixed 1 ft in from all walls.</p>
        <table>
          <thead><tr><th style={{ width: "16%" }}>Room</th><th style={{ width: "18%" }}>Length</th><th style={{ width: "18%" }}>Width</th><th style={{ width: "22%" }}>Floor Marble</th><th style={{ width: "20%" }}>Border Marble</th><th style={{ width: "6%" }}></th></tr></thead>
          <tbody>
            {rooms.map((row) => (
              <tr key={row.id}>
                <td><input type="text" placeholder="Room" value={row.name} onChange={(e) => updateRoom(row.id, "name", e.target.value)} /></td>
                <td><div className="dim-pair">
                  <input type="number" min="0" placeholder="ft" className="dim-num" value={row.lf} onChange={(e) => updateRoom(row.id, "lf", e.target.value)} /><span>ft</span>
                  <input type="number" min="0" max="11" placeholder="in" className="dim-num" value={row.li} onChange={(e) => updateRoom(row.id, "li", e.target.value)} /><span>in</span>
                </div></td>
                <td><div className="dim-pair">
                  <input type="number" min="0" placeholder="ft" className="dim-num" value={row.wf} onChange={(e) => updateRoom(row.id, "wf", e.target.value)} /><span>ft</span>
                  <input type="number" min="0" max="11" placeholder="in" className="dim-num" value={row.wi} onChange={(e) => updateRoom(row.id, "wi", e.target.value)} /><span>in</span>
                </div></td>
                <td><MarbleSelect value={row.floorMarble} customRate={row.floorCustomRate} onChange={(v) => updateRoom(row.id, "floorMarble", v)} onCustomRateChange={(v) => updateRoom(row.id, "floorCustomRate", v)} list={marble} /></td>
                <td><MarbleSelect value={row.borderMarble} customRate={row.borderCustomRate} onChange={(v) => updateRoom(row.id, "borderMarble", v)} onCustomRateChange={(v) => updateRoom(row.id, "borderCustomRate", v)} hasGroups borderList={border} borderDesignList={borderDesign} /></td>
                <td><button className="rm-btn" onClick={() => removeRoom(row.id)}>✕</button></td>
              </tr>
            ))}
            {!rooms.length && <tr><td colSpan={6}><p className="empty">No rooms added yet.</p></td></tr>}
          </tbody>
        </table>
        <div className="row-actions">
          <button className="add" onClick={addRoom}>+ Add Room</button>
          <button className="calc" onClick={calculate}>Calculate Ledger →</button>
        </div>
      </div>

      {/* Results */}
      {results && (
        <div id="results-section">
          {(results.rooms.length > 0 || results.katbaItems.length > 0) ? (
            <>
              <div className="print-bar no-print">
                <button className="print-btn" onClick={() => window.print()}>⎙ Print / Save PDF</button>
                <button className="export-btn" onClick={() => exportSpreadsheet(results)}>⬇ Download Spreadsheet (.xlsx)</button>
                <span className="print-hint">Opens in Excel or Google Sheets</span>
              </div>
              <div className="save-bar no-print">
                <input
                  type="text"
                  className="client-name-input"
                  placeholder="Client name (e.g. Mr. Ahmed)"
                  value={clientName}
                  onChange={(e) => setClientName(e.target.value)}
                />
                <button className="save-btn" onClick={() => saveToHistory(results)}>
                  ✦ Save to History
                </button>
                {savedId && <span className="save-confirm">✓ Saved!</span>}
              </div>
              <div className="panel">
                <h2>Results — Per Room</h2>
                <p className="hint">Area = L × W. Skirting = 2×(L+W). Border = inner rect inset 1ft.</p>
                {results.rooms.map((r, i) => (
                  <div className="room-card" key={i}>
                    <h3>{r.name} <span className="dims">{ftin(r.L)} × {ftin(r.W)} ({r.L.toFixed(2)} × {r.W.toFixed(2)} ft)</span></h3>
                    <p className="calc-line">Area = {r.L.toFixed(2)} × {r.W.toFixed(2)} = <strong>{r.area.toFixed(2)} sq ft</strong> — {r.floorName} (Rs {r.floorRate}/sqft)</p>
                    <p className="calc-line">Skirting = 2×({r.L.toFixed(2)}+{r.W.toFixed(2)}) = <strong>{r.skirt.toFixed(2)} run ft</strong> — {r.floorName}</p>
                    <p className="calc-line">Border = ({Math.max(r.L-2,0).toFixed(2)}×{Math.max(r.W-2,0).toFixed(2)}) perimeter = <strong>{r.border.toFixed(2)} run ft</strong> — {r.borderName} (Rs {r.borderRate}/run ft)</p>
                    <div className="metric-grid">
                      <div className="metric"><div className="label">Area</div><div className="value">{r.area.toFixed(2)} <small>sq ft</small></div><div className="cost">Rs {r.areaCost.toFixed(0)}/-</div></div>
                      <div className="metric"><div className="label">Skirting</div><div className="value">{r.skirt.toFixed(2)} <small>run ft</small></div><div className="cost">Rs {r.skirtCost.toFixed(0)}/-</div></div>
                      <div className="metric"><div className="label">Border</div><div className="value">{r.border.toFixed(2)} <small>run ft</small></div><div className="cost">Rs {r.borderCost.toFixed(0)}/-</div></div>
                    </div>
                  </div>
                ))}
                {results.katbaItems.length > 0 && (
                  <div className="room-card">
                    <h3>Katba (Decorative Tiles)</h3>
                    {results.katbaItems.map((k, i) => (
                      <p className="calc-line" key={i}>{k.location}: {k.qty} × {k.typeName} (Rs {k.rate}/piece) = <strong>Rs {k.cost.toFixed(0)}/-</strong></p>
                    ))}
                  </div>
                )}
              </div>
              <div className="panel">
                <h2>Combined Totals</h2>
                <table className="totals">
                  <thead><tr><th>Metric</th><th>Feet-Inches / ft</th><th>Decimal</th><th>Cost (Rs)</th></tr></thead>
                  <tbody>
                    <tr><td>Total Floor Area</td><td>{results.totals.area.toFixed(2)} sq ft</td><td>{results.totals.area.toFixed(2)}</td><td>{results.totals.areaCost.toFixed(0)}</td></tr>
                    <tr><td>Total Skirting</td><td>{results.totals.skirt.toFixed(2)} run ft</td><td>{results.totals.skirt.toFixed(2)}</td><td>{results.totals.skirtCost.toFixed(0)}</td></tr>
                    <tr><td>Total Border / Strip</td><td>{results.totals.border.toFixed(2)} run ft</td><td>{results.totals.border.toFixed(2)}</td><td>{results.totals.borderCost.toFixed(0)}</td></tr>
                    <tr><td>Total Katba</td><td>{results.totals.katbaQty} pcs</td><td>{results.totals.katbaQty}</td><td>{results.totals.katbaCost.toFixed(0)}</td></tr>
                    <tr className="grand"><td colSpan={3}>Grand Total (Rs)</td><td>{results.totals.grand.toFixed(0)}/-</td></tr>
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <div className="panel"><p className="empty">No room dimensions entered yet.</p></div>
          )}
        </div>
      )}
    </div>
  );
}

// ── Access denied screen ───────────────────────────────────────────────────────
function AccessDenied() {
  const { signOut } = useClerk();
  const { user } = useUser();
  return (
    <div style={{ minHeight: "100dvh", display: "flex", alignItems: "center", justifyContent: "center", background: "var(--stone-50)" }}>
      <div className="panel" style={{ maxWidth: 440, textAlign: "center", padding: "40px 32px" }}>
        <div style={{ width: 8, height: 8, background: "var(--veinred)", borderRadius: "50%", margin: "0 auto 16px" }} />
        <h2 style={{ fontFamily: "'Fraunces',serif", fontSize: "1.3rem", marginBottom: 8 }}>Access Restricted</h2>
        <p style={{ fontSize: "0.88rem", color: "var(--ink-soft)", marginBottom: 24 }}>
          <strong>{user?.primaryEmailAddress?.emailAddress}</strong> is not authorised to access this app.
          Please contact the showroom owner.
        </p>
        <button className="calc" onClick={() => signOut({ redirectUrl: basePath || "/" })}>Sign out</button>
      </div>
    </div>
  );
}

// ── Sign-in page ──────────────────────────────────────────────────────────────
function SignInPage() {
  return (
    <div style={{ minHeight: "100dvh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", background: "var(--stone-50)", padding: "24px 16px" }}>
      <div style={{ marginBottom: 28, textAlign: "center" }}>
        <h1 style={{ fontFamily: "'Fraunces',serif", fontSize: "1.9rem", margin: "0 0 4px", letterSpacing: "-0.01em" }}>
          Bhutta Marble &amp; <span style={{ color: "var(--veinred)" }}>Granite</span>
        </h1>
        <p style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: "0.72rem", color: "var(--ink-soft)", margin: 0, letterSpacing: "0.05em" }}>MARBLE LEDGER · PRIVATE ACCESS</p>
      </div>
      <SignIn
        routing="path"
        path={`${basePath}/sign-in`}
        signUpUrl={`${basePath}/sign-up`}
        appearance={{
          theme: shadcn,
          variables: {
            colorPrimary: "#8a3324",
            colorForeground: "#2a2521",
            colorMutedForeground: "#5c544a",
            colorBackground: "#fbf9f6",
            colorInput: "#ffffff",
            colorInputForeground: "#2a2521",
            colorNeutral: "#cfc6b8",
            fontFamily: "'Inter', sans-serif",
            borderRadius: "3px",
          },
          elements: {
            rootBox: "w-full flex justify-center",
            cardBox: "w-[420px] max-w-full overflow-hidden",
            card: "!shadow-none !border !border-[#d8d0c2] !bg-[#fbf9f6] !rounded-none",
            footer: "!shadow-none !border-t !border-[#d8d0c2] !bg-[#f6f4f0] !rounded-none",
            headerTitle: "text-[#2a2521]",
            headerSubtitle: "text-[#5c544a]",
            socialButtonsBlockButtonText: "text-[#2a2521]",
            formFieldLabel: "text-[#2a2521]",
            footerActionLink: "text-[#8a3324]",
            footerActionText: "text-[#5c544a]",
            dividerText: "text-[#5c544a]",
            identityPreviewEditButton: "text-[#8a3324]",
            formFieldSuccessText: "text-[#2a2521]",
            alertText: "text-[#2a2521]",
            logoBox: "hidden",
            logoImage: "hidden",
            socialButtonsBlockButton: "border-[#cfc6b8]",
            formButtonPrimary: "bg-[#8a3324] hover:bg-[#b85c4a]",
            formFieldInput: "border-[#cfc6b8] bg-white text-[#2a2521]",
            footerAction: "bg-[#f6f4f0]",
            dividerLine: "bg-[#d8d0c2]",
            alert: "border-[#cfc6b8]",
            otpCodeFieldInput: "border-[#cfc6b8]",
            formFieldRow: "",
            main: "",
          },
        }}
      />
    </div>
  );
}

// ── Home: redirect signed-in users, show sign-in for signed-out ───────────────
function Home() {
  const { isSignedIn, isLoaded } = useUser();
  if (!isLoaded) return null;
  if (isSignedIn) return <Redirect to="/ledger" />;
  return <Redirect to="/sign-in" />;
}

// ── Protected ledger route with email guard ───────────────────────────────────
function ProtectedLedger() {
  const { isSignedIn, isLoaded, user } = useUser();
  if (!isLoaded) return null;
  if (!isSignedIn) return <Redirect to="/sign-in" />;
  const email = user?.primaryEmailAddress?.emailAddress?.toLowerCase() ?? "";
  if (ALLOWED_EMAILS.length > 0 && !ALLOWED_EMAILS.includes(email)) return <AccessDenied />;
  return <Ledger />;
}

// ── Router ────────────────────────────────────────────────────────────────────
function ClerkProviderWithRoutes() {
  const [, setLocation] = useLocation();
  return (
    <ClerkProvider
      publishableKey={clerkPubKey}
      proxyUrl={clerkProxyUrl}
      signInUrl={`${basePath}/sign-in`}
      signUpUrl={`${basePath}/sign-up`}
      routerPush={(to) => setLocation(stripBase(to))}
      routerReplace={(to) => setLocation(stripBase(to), { replace: true })}
    >
      <Switch>
        <Route path="/" component={Home} />
        <Route path="/ledger" component={ProtectedLedger} />
        <Route path="/sign-in/*?" component={SignInPage} />
        <Route path="/sign-up/*?" component={SignInPage} />
      </Switch>
    </ClerkProvider>
  );
}

export default function App() {
  return (
    <WouterRouter base={basePath}>
      <ClerkProviderWithRoutes />
    </WouterRouter>
  );
}

```

---

## Section 2: Styles — `src/index.css`  (758 lines)

```css
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
@import "tailwindcss";

:root {
  --stone-50: #f6f4f0;
  --stone-100: #ece8e1;
  --stone-300: #cfc6b8;
  --ink: #2a2521;
  --ink-soft: #5c544a;
  --veinred: #8a3324;
  --veinred-soft: #b85c4a;
  --line: #d8d0c2;
  --paper: #fbf9f6;
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  background: var(--stone-50);
  background-image:
    radial-gradient(circle at 18% 25%, rgba(138, 51, 36, 0.05), transparent 40%),
    radial-gradient(circle at 80% 70%, rgba(42, 37, 33, 0.05), transparent 45%);
  color: var(--ink);
  font-family: 'Inter', sans-serif;
  padding: 36px 18px 60px;
}

.wrap {
  max-width: 980px;
  margin: 0 auto;
}

header.title-block {
  border-bottom: 2px solid var(--ink);
  padding-bottom: 18px;
  margin-bottom: 28px;
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  flex-wrap: wrap;
  gap: 10px;
}

header.title-block h1 {
  font-family: 'Fraunces', serif;
  font-weight: 700;
  font-size: 2.5rem;
  margin: 0;
  letter-spacing: -0.01em;
}

header.title-block h1 span {
  color: var(--veinred);
}

header.title-block .sub {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.78rem;
  color: var(--ink-soft);
  text-align: right;
  line-height: 1.5;
}

.panel {
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: 2px;
  padding: 22px 24px;
  margin-bottom: 28px;
  box-shadow: 0 1px 0 rgba(0, 0, 0, 0.02);
}

.panel h2 {
  font-family: 'Fraunces', serif;
  font-size: 1.15rem;
  margin: 0 0 4px;
  display: flex;
  align-items: center;
  gap: 10px;
}

.panel h2::before {
  content: '';
  width: 8px;
  height: 8px;
  background: var(--veinred);
  border-radius: 50%;
  display: inline-block;
  flex-shrink: 0;
}

.panel .hint {
  font-size: 0.83rem;
  color: var(--ink-soft);
  margin: 0 0 16px;
}

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.88rem;
}

thead th {
  text-align: left;
  font-family: 'JetBrains Mono', monospace;
  font-weight: 500;
  font-size: 0.7rem;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--ink-soft);
  border-bottom: 1px solid var(--ink);
  padding: 8px 6px;
}

tbody td {
  border-bottom: 1px solid var(--line);
  padding: 7px 6px;
  vertical-align: middle;
}

tbody tr:hover {
  background: rgba(138, 51, 36, 0.03);
}

input[type=text],
input[type=number] {
  width: 100%;
  border: 1px solid var(--stone-300);
  background: #fff;
  border-radius: 3px;
  padding: 6px 7px;
  font-family: 'Inter', sans-serif;
  font-size: 0.86rem;
  color: var(--ink);
  outline: none;
}

input[type=text]:focus,
input[type=number]:focus {
  border-color: var(--veinred-soft);
}

input.dim-num {
  width: 58px;
  text-align: center;
}

select {
  width: 100%;
  border: 1px solid var(--stone-300);
  background: #fff;
  border-radius: 3px;
  padding: 6px 5px;
  font-family: 'Inter', sans-serif;
  font-size: 0.83rem;
  color: var(--ink);
  outline: none;
}

select:focus {
  border-color: var(--veinred-soft);
}

.dim-pair {
  display: flex;
  align-items: center;
  gap: 4px;
}

.dim-pair span {
  color: var(--ink-soft);
  font-size: 0.78rem;
  white-space: nowrap;
}

.rm-btn {
  background: none;
  border: none;
  color: var(--veinred-soft);
  cursor: pointer;
  font-size: 1.1rem;
  line-height: 1;
  padding: 2px 6px;
}

.rm-btn:hover {
  color: var(--veinred);
}

.row-actions {
  display: flex;
  gap: 10px;
  margin-top: 14px;
  flex-wrap: wrap;
}

button.add,
button.calc {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.78rem;
  letter-spacing: 0.03em;
  border: 1px solid var(--ink);
  background: transparent;
  color: var(--ink);
  padding: 9px 16px;
  border-radius: 3px;
  cursor: pointer;
  transition: all 0.15s ease;
}

button.add:hover {
  background: var(--ink);
  color: var(--paper);
}

button.calc {
  background: var(--veinred);
  color: #fff;
  border-color: var(--veinred);
}

button.calc:hover {
  background: var(--veinred-soft);
  border-color: var(--veinred-soft);
}

.room-card {
  border: 1px solid var(--line);
  border-radius: 2px;
  padding: 16px 18px;
  margin-bottom: 14px;
  background: var(--paper);
}

.room-card h3 {
  font-family: 'Fraunces', serif;
  font-size: 1.05rem;
  margin: 0 0 10px;
  border-bottom: 1px dashed var(--line);
  padding-bottom: 8px;
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  flex-wrap: wrap;
  gap: 6px;
  color: var(--ink);
}

.room-card h3 .dims {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.78rem;
  color: var(--ink-soft);
  font-weight: 400;
}

.calc-line {
  font-size: 0.85rem;
  color: var(--ink-soft);
  margin: 3px 0;
  font-family: 'JetBrains Mono', monospace;
  word-break: break-word;
}

.calc-line strong {
  color: var(--ink);
  font-family: 'Inter', sans-serif;
  font-weight: 600;
}

.metric-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin-top: 12px;
}

.metric {
  background: var(--stone-100);
  border-radius: 3px;
  padding: 10px 12px;
}

.metric .label {
  font-size: 0.68rem;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: var(--ink-soft);
  font-family: 'JetBrains Mono', monospace;
}

.metric .value {
  font-family: 'Fraunces', serif;
  font-size: 1.25rem;
  color: var(--ink);
  margin-top: 2px;
}

.metric .value small {
  font-size: 0.7rem;
  color: var(--ink-soft);
  font-family: 'JetBrains Mono', monospace;
}

.metric .cost {
  font-size: 0.78rem;
  color: var(--veinred);
  margin-top: 3px;
  font-family: 'JetBrains Mono', monospace;
}

table.totals th {
  border-bottom: 1px solid var(--ink);
}

table.totals td {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.85rem;
}

table.totals tr.grand td {
  border-top: 2px solid var(--ink);
  border-bottom: none;
  font-weight: 600;
  font-size: 0.95rem;
  font-family: 'Inter', sans-serif;
  color: var(--veinred);
  padding-top: 12px;
}

.empty {
  color: var(--ink-soft);
  font-size: 0.85rem;
  font-style: italic;
  padding: 10px 0;
}

.custom-rate-input {
  margin-top: 4px;
  width: 100%;
}

/* ── Edit Rates Panel ── */
.rates-panel {
  padding: 0;
  overflow: hidden;
}

.rates-toggle {
  width: 100%;
  display: flex;
  align-items: center;
  gap: 10px;
  background: none;
  border: none;
  padding: 18px 24px;
  font-family: 'Fraunces', serif;
  font-size: 1.05rem;
  color: var(--ink);
  cursor: pointer;
  text-align: left;
}

.rates-toggle:hover {
  background: rgba(138, 51, 36, 0.04);
}

.rates-toggle::before {
  content: '';
  width: 8px;
  height: 8px;
  background: var(--veinred);
  border-radius: 50%;
  display: inline-block;
  flex-shrink: 0;
}

.rates-toggle-icon {
  font-size: 0.9rem;
  color: var(--veinred);
  font-family: monospace;
}

.rates-badge {
  margin-left: auto;
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.68rem;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--ink-soft);
  background: var(--stone-100);
  border: 1px solid var(--line);
  padding: 3px 8px;
  border-radius: 2px;
}

.rates-body {
  padding: 0 24px 22px;
  border-top: 1px solid var(--line);
}

.rates-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 20px;
  margin-bottom: 16px;
}

.rate-group {
  background: var(--stone-50);
  border: 1px solid var(--line);
  border-radius: 2px;
  padding: 12px 14px;
}

.rate-group-title {
  font-family: 'Fraunces', serif;
  font-size: 0.9rem;
  font-weight: 600;
  color: var(--ink);
  margin-bottom: 10px;
}

.rate-unit {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.68rem;
  color: var(--ink-soft);
  font-weight: 400;
  margin-left: 4px;
}

.rate-table {
  width: 100%;
  font-size: 0.83rem;
}

.rate-table thead th {
  font-size: 0.65rem;
  padding: 5px 4px;
}

.rate-table tbody td {
  padding: 5px 4px;
  border-bottom: 1px solid var(--stone-100);
}

.rate-table tbody tr:last-child td {
  border-bottom: none;
}

.rate-input-wrap input[type=number] {
  width: 90px;
  text-align: right;
  padding: 4px 6px;
  font-size: 0.83rem;
}

.rates-footer {
  display: flex;
  justify-content: flex-end;
  padding-top: 4px;
}

button.reset-btn {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.75rem;
  letter-spacing: 0.03em;
  border: 1px solid var(--stone-300);
  background: transparent;
  color: var(--ink-soft);
  padding: 7px 14px;
  border-radius: 3px;
  cursor: pointer;
  transition: all 0.15s ease;
}

button.reset-btn:hover {
  border-color: var(--veinred);
  color: var(--veinred);
}

@media (max-width: 680px) {
  .rates-grid {
    grid-template-columns: 1fr;
  }
}

.print-bar {
  display: flex;
  align-items: center;
  gap: 14px;
  margin-bottom: 18px;
  flex-wrap: wrap;
}

button.print-btn {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.78rem;
  letter-spacing: 0.03em;
  border: 1px solid var(--ink);
  background: var(--ink);
  color: var(--paper);
  padding: 9px 18px;
  border-radius: 3px;
  cursor: pointer;
  transition: all 0.15s ease;
}

button.print-btn:hover {
  background: var(--veinred);
  border-color: var(--veinred);
}

button.export-btn {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.78rem;
  letter-spacing: 0.03em;
  border: 1px solid var(--ink);
  background: transparent;
  color: var(--ink);
  padding: 9px 18px;
  border-radius: 3px;
  cursor: pointer;
  transition: all 0.15s ease;
}

button.export-btn:hover {
  background: var(--ink);
  color: var(--paper);
}

.print-hint {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.72rem;
  color: var(--ink-soft);
}

@media (max-width: 680px) {
  .metric-grid {
    grid-template-columns: 1fr;
  }

  header.title-block .sub {
    text-align: left;
  }

  table {
    font-size: 0.8rem;
  }
}

/* ── Save bar ─────────────────────────────────────────────────────────────── */
.save-bar {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 18px;
  flex-wrap: wrap;
}

input.client-name-input {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.82rem;
  border: 1px solid var(--stone-300);
  background: var(--paper);
  color: var(--ink);
  padding: 8px 12px;
  border-radius: 3px;
  width: 220px;
  outline: none;
  transition: border-color 0.15s ease;
}

input.client-name-input:focus {
  border-color: var(--veinred);
}

button.save-btn {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.78rem;
  letter-spacing: 0.03em;
  border: 1px solid var(--veinred);
  background: var(--veinred);
  color: #fff;
  padding: 9px 18px;
  border-radius: 3px;
  cursor: pointer;
  transition: all 0.15s ease;
}

button.save-btn:hover {
  background: #6d2519;
  border-color: #6d2519;
}

.save-confirm {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.78rem;
  color: #3a7d44;
  font-weight: 600;
  letter-spacing: 0.02em;
}

/* ── History panel ────────────────────────────────────────────────────────── */
.history-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.history-entry {
  border: 1px solid var(--stone-300);
  border-radius: 4px;
  background: var(--paper);
  overflow: hidden;
  transition: border-color 0.2s ease;
}

.history-entry--new {
  border-color: #3a7d44;
  box-shadow: 0 0 0 2px rgba(58, 125, 68, 0.12);
}

.history-row {
  display: flex;
  align-items: center;
}

button.history-expand {
  flex: 1;
  display: flex;
  align-items: center;
  gap: 10px;
  background: transparent;
  border: none;
  padding: 11px 14px;
  cursor: pointer;
  font-family: inherit;
  text-align: left;
  color: var(--ink);
  font-size: 0.85rem;
}

button.history-expand:hover {
  background: var(--stone-100);
}

.history-meta {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.72rem;
  color: var(--ink-soft);
}

.history-total {
  margin-left: auto;
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.82rem;
  font-weight: 700;
  color: var(--veinred);
}

button.history-del {
  background: transparent;
  border: none;
  color: var(--ink-soft);
  font-size: 0.8rem;
  padding: 0 14px;
  cursor: pointer;
  line-height: 1;
  transition: color 0.15s;
}

button.history-del:hover {
  color: var(--veinred);
}

.history-detail {
  padding: 0 14px 12px;
  border-top: 1px solid var(--stone-200);
}

.history-detail table.totals {
  margin-top: 10px;
}

@media print {
  @page {
    margin: 18mm 14mm;
    size: A4;
  }

  body {
    background: #fff !important;
    background-image: none !important;
    padding: 0;
    font-size: 11pt;
    color: #000;
  }

  .no-print,
  .panel:not(#print-results):has(table#inputTable),
  .panel:has(#katbaTable),
  .panel:has(.row-actions) {
    display: none !important;
  }

  .wrap {
    max-width: 100%;
  }

  header.title-block {
    border-bottom: 2pt solid #000;
    padding-bottom: 10pt;
    margin-bottom: 16pt;
  }

  header.title-block h1 {
    font-size: 20pt;
  }

  .panel {
    background: #fff !important;
    border: 1px solid #ccc;
    box-shadow: none;
    margin-bottom: 14pt;
    padding: 12pt 14pt;
    break-inside: avoid;
  }

  .room-card {
    break-inside: avoid;
    border: 1px solid #ccc;
    background: #fff !important;
    padding: 10pt 12pt;
    margin-bottom: 10pt;
  }

  .metric-grid {
    grid-template-columns: repeat(3, 1fr);
  }

  .metric {
    background: #f0ede8 !important;
  }

  table.totals tr.grand td {
    color: #8a3324;
  }

  button,
  .row-actions {
    display: none !important;
  }
}

```

---

*Bhutta Marble & Granite Showroom · Marble Ledger · 2026*