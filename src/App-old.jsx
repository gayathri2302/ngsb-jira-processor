import { useState, useCallback, useMemo } from "react";
import * as XLSX from 'xlsx';

// ─── theme definitions ────────────────────────────────────────────────────────

const THEMES = {
  dark: {
    bg: "#060E18",
    bgSecondary: "#081521",
    bgTertiary: "#0A1520",
    bgQuaternary: "#0D1B2A",
    bgHover: "#112030",
    bgModal: "#0D1B2A",
    border: "#1E3A5A",
    borderLight: "#2A3D5A",
    text: "#C8D8E8",
    textSecondary: "#7A9BB5",
    textTertiary: "#4A7FA5",
    textMuted: "#596B7A",
    accent: "#E8D5B0",
    accentSecondary: "#7AB8E0",
    scrollbarTrack: "#060E18",
    scrollbarThumb: "#1E3A5A",
    buttonBg: "#0F2340",
    buttonBorder: "#2A6A9A",
    inputBg: "#0A1520",
    inputBorder: "#2A3D5A",
    inputFocus: "#4A7FA5",
    tableRowEven: "#0A1520",
    tableRowOdd: "#0D1B2A",
    tableHeader: "#081521",
    totalRow: "#0F2340",
  },
  light: {
    bg: "#F8FAFC",
    bgSecondary: "#FFFFFF",
    bgTertiary: "#F1F5F9",
    bgQuaternary: "#E2E8F0",
    bgHover: "#E8EDF3",
    bgModal: "#FFFFFF",
    border: "#CBD5E1",
    borderLight: "#E2E8F0",
    text: "#1A3A52",
    textSecondary: "#475569",
    textTertiary: "#0066CC",
    textMuted: "#64748B",
    accent: "#0066CC",
    accentSecondary: "#0052A3",
    scrollbarTrack: "#F1F5F9",
    scrollbarThumb: "#CBD5E1",
    buttonBg: "#0066CC",
    buttonBorder: "#0052A3",
    inputBg: "#FFFFFF",
    inputBorder: "#CBD5E1",
    inputFocus: "#0066CC",
    tableRowEven: "#FFFFFF",
    tableRowOdd: "#F8FAFC",
    tableHeader: "#F1F5F9",
    totalRow: "#E2E8F0",
  },
};

// ─── helpers ──────────────────────────────────────────────────────────────────

const INVALID_SHEET_CHARS = /[/\\?\*\[\]:]/g;

const STATUS_ORDER = [
  "Backlog", "To Do", "IN DEVELOPMENT", "Code Review",
  "Ready For Deployment", "Deployed - Ready for Testing",
  "QA In Test", "QA Defect", "QA Complete", "Passed QA",
  "Ready for UAT", "Regression Test", "Dev On Hold",
];

const STATUS_PALETTE = {
  "Backlog": { bg: "#E8E8ED", text: "#555" },
  "To Do": { bg: "#E2F0E9", text: "#276745" },
  "IN DEVELOPMENT": { bg: "#DAEAF6", text: "#1A5276" },
  "Code Review": { bg: "#C5D9EE", text: "#154360" },
  "Ready For Deployment": { bg: "#FEF3CD", text: "#7D5A00" },
  "Deployed - Ready for Testing": { bg: "#FFF8DC", text: "#6B5900" },
  "QA In Test": { bg: "#FDDEDE", text: "#7B2020" },
  "QA Defect": { bg: "#FFBDBD", text: "#6B0000" },
  "QA Complete": { bg: "#D6F5E0", text: "#1D6B35" },
  "Passed QA": { bg: "#B7EAC7", text: "#145A27" },
  "Ready for UAT": { bg: "#C9E8BA", text: "#2D5A1A" },
  "Regression Test": { bg: "#C9D9EF", text: "#1C3860" },
  "Dev On Hold": { bg: "#FFE5B4", text: "#8B5E00" },
};

function parseEpicInfo(el) {
  if (!el) return { epicNum: "", comp: "Unknown", sheetName: "Unknown" };
  const rMatch = el.match(/R[\d]+\.[\d]+[\w.]*/);
  const epicNum = rMatch ? rMatch[0] : "";
  const parts = el.split("|").map(p => p.trim()).filter(p => p && !p.match(/^R[\d]/));
  const comp = parts[parts.length - 1] || el.slice(0, 40);
  const raw = epicNum ? `${comp} - ${epicNum}` : comp.slice(0, 40);
  const sheetName = raw.replace(INVALID_SHEET_CHARS, "").slice(0, 31).trim();
  return { epicNum, comp, sheetName };
}

function statusSortKey(s) {
  const idx = STATUS_ORDER.indexOf(s);
  return idx === -1 ? 99 : idx;
}

// ─── XLSX builder ────────────────────────────────────────────────────────

async function buildXlsx(rows) {
  const epicGroups = {};
  rows.forEach(r => {
    const el = r["Epic Link"] || "";
    if (!epicGroups[el]) epicGroups[el] = [];
    epicGroups[el].push(r);
  });

  function epicSortKey(el) {
    const m = el.match(/R([\d]+)\.([\d]+)([\w.]*)/);
    if (m) return `${m[1].padStart(4, "0")}.${m[2].padStart(4, "0")}.${m[3]}`;
    return "zz" + el;
  }

  const epicLinks = Object.keys(epicGroups).sort((a, b) => epicSortKey(a).localeCompare(epicSortKey(b)));

  const COLS = ["Key", "Summary", "Status", "Start Date", "Dev End date", "Assignee",
    "Estimated effort (hrs)", "Date of UAT deployment", "Updated",
    "QA Start Date", "QA End Date", "Labels", "PM Notes (Offshore)"];

  const wb = XLSX.utils.book_new();

  // Summary sheet
  const summaryData = [["Epic / Component", "Epic #", "Tickets", "Statuses"]];
  epicLinks.forEach(el => {
    const { epicNum, comp } = parseEpicInfo(el);
    const grp = epicGroups[el];
    const statuses = [...new Set(grp.map(r => r.Status))].sort((a, b) => statusSortKey(a) - statusSortKey(b)).join(", ");
    summaryData.push([comp, epicNum || "-", grp.length, statuses]);
  });
  const sumWS = XLSX.utils.aoa_to_sheet(summaryData);
  sumWS["!cols"] = [{ wch: 45 }, { wch: 16 }, { wch: 8 }, { wch: 60 }];
  XLSX.utils.book_append_sheet(wb, sumWS, "Summary");

  // Epic sheets
  epicLinks.forEach(el => {
    const { comp, epicNum, sheetName } = parseEpicInfo(el);
    const grp = epicGroups[el];

    const statuses = [...new Set(grp.map(r => r.Status))].sort((a, b) => statusSortKey(a) - statusSortKey(b));
    const aoa = [];
    const merges = [];

    // Title row
    aoa.push([`${comp}  |  ${epicNum}`]);
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: COLS.length - 1 } });

    statuses.forEach(status => {
      // Status header row
      aoa.push([`◆  ${status.toUpperCase()}`]);
      merges.push({ s: { r: aoa.length - 1, c: 0 }, e: { r: aoa.length - 1, c: COLS.length - 1 } });

      // Column header row
      aoa.push(COLS);

      // Data rows
      grp.filter(r => r.Status === status).forEach(r => {
        aoa.push(COLS.map(c => r[c] || ""));
      });

      // Blank gap
      aoa.push([]);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = merges;
    ws["!cols"] = [
      { wch: 14 }, { wch: 60 }, { wch: 22 }, { wch: 12 }, { wch: 12 }, { wch: 28 },
      { wch: 10 }, { wch: 18 }, { wch: 18 }, { wch: 13 }, { wch: 12 }, { wch: 12 }, { wch: 20 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  // Pivot sheet
  const allStatuses = [...new Set(rows.map(r => r.Status))].sort((a, b) => statusSortKey(a) - statusSortKey(b));
  const pivotHeader = ["Epic", "Epic #", "Component", ...allStatuses, "Total"];
  const pivotRows = [pivotHeader];
  epicLinks.forEach(el => {
    const { epicNum, comp } = parseEpicInfo(el);
    const grp = epicGroups[el];
    const counts = allStatuses.map(s => grp.filter(r => r.Status === s).length);
    pivotRows.push([el.split("|").slice(0, 2).join(" | ").trim(), epicNum, comp, ...counts, grp.length]);
  });
  const totalsRow = ["TOTAL", "", "", ...allStatuses.map(s => rows.filter(r => r.Status === s).length), rows.length];
  pivotRows.push(totalsRow);

  const pivotWS = XLSX.utils.aoa_to_sheet(pivotRows);
  const pivotCols = [{ wch: 50 }, { wch: 16 }, { wch: 40 }, ...allStatuses.map(() => ({ wch: 22 })), { wch: 8 }];
  pivotWS["!cols"] = pivotCols;
  XLSX.utils.book_append_sheet(wb, pivotWS, "Pivot - Status");

  // Owner pivot
  const allOwners = [...new Set(rows.map(r => r.Assignee || "Unassigned"))].sort();
  const ownerHeader = ["Assignee", ...allStatuses, "Total"];
  const ownerRows = [ownerHeader];
  allOwners.forEach(owner => {
    const grp = rows.filter(r => (r.Assignee || "Unassigned") === owner);
    const counts = allStatuses.map(s => grp.filter(r => r.Status === s).length);
    ownerRows.push([owner, ...counts, grp.length]);
  });
  const ownerTotals = ["TOTAL", ...allStatuses.map(s => rows.filter(r => r.Status === s).length), rows.length];
  ownerRows.push(ownerTotals);

  const ownerWS = XLSX.utils.aoa_to_sheet(ownerRows);
  ownerWS["!cols"] = [{ wch: 35 }, ...allStatuses.map(() => ({ wch: 22 })), { wch: 8 }];
  XLSX.utils.book_append_sheet(wb, ownerWS, "Pivot - Owners");

  XLSX.writeFile(wb, "NGSB_Sprint_Epics.xlsx");
}

// ─── pivot table component ────────────────────────────────────────────────────

function PivotTable({ rows, theme }) {
  const [pivotMode, setPivotMode] = useState("status");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterEpic, setFilterEpic] = useState("all");
  const [filterOwner, setFilterOwner] = useState("all");

  const statuses = useMemo(() =>
    [...new Set(rows.map(r => r.Status))].sort((a, b) => statusSortKey(a) - statusSortKey(b)),
    [rows]);

  const epicLinks = useMemo(() => {
    const els = [...new Set(rows.map(r => r["Epic Link"] || ""))];
    return els.sort((a, b) => {
      const ka = a.match(/R([\d]+)\.([\d]+)/);
      const kb = b.match(/R([\d]+)\.([\d]+)/);
      if (ka && kb) return ka[0].localeCompare(kb[0]);
      return a.localeCompare(b);
    });
  }, [rows]);

  const owners = useMemo(() =>
    [...new Set(rows.map(r => r.Assignee || "Unassigned"))].sort(),
    [rows]);

  const filtered = useMemo(() => rows.filter(r => {
    if (filterStatus !== "all" && r.Status !== filterStatus) return false;
    if (filterEpic !== "all" && (r["Epic Link"] || "") !== filterEpic) return false;
    if (filterOwner !== "all" && (r.Assignee || "Unassigned") !== filterOwner) return false;
    return true;
  }), [rows, filterStatus, filterEpic, filterOwner]);

  const [expandedTickets, setExpandedTickets] = useState(null);

  const selectStyleThemed = {
    background: theme.inputBg, border: `1px solid ${theme.inputBorder}`, color: theme.textSecondary,
    padding: "5px 10px", borderRadius: "6px", fontSize: "11px", fontFamily: "'DM Mono',monospace", cursor: "pointer",
    transition: "all 0.3s ease"
  };

  return (
    <div style={{ fontFamily: "'DM Mono', monospace" }}>
      {/* Controls */}
      <div style={{ display: "flex", gap: "12px", marginBottom: "20px", flexWrap: "wrap", alignItems: "center" }}>
        <div style={{ display: "flex", gap: 0, border: `1px solid ${theme.borderLight}`, borderRadius: "6px", overflow: "hidden" }}>
          {["status", "owner"].map(m => (
            <button key={m} onClick={() => setPivotMode(m)} style={{
              padding: "6px 16px", background: pivotMode === m ? theme.buttonBg : "transparent",
              color: pivotMode === m ? theme.accent : theme.textSecondary, border: "none", cursor: "pointer",
              fontFamily: "inherit", fontSize: "11px", letterSpacing: "0.05em", textTransform: "uppercase",
              transition: "all 0.3s ease"
            }}>{m === "status" ? "By Status" : "By Owner"}</button>
          ))}
        </div>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={selectStyleThemed}>
          <option value="all">All Statuses</option>
          {statuses.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <select value={filterEpic} onChange={e => setFilterEpic(e.target.value)} style={selectStyleThemed}>
          <option value="all">All Epics</option>
          {epicLinks.map(el => {
            const { epicNum, comp } = parseEpicInfo(el);
            return <option key={el} value={el}>{epicNum ? `${epicNum} – ${comp}` : comp}</option>;
          })}
        </select>
        <select value={filterOwner} onChange={e => setFilterOwner(e.target.value)} style={selectStyleThemed}>
          <option value="all">All Owners</option>
          {owners.map(owner => <option key={owner} value={owner}>{owner}</option>)}
        </select>
        <span style={{ color: theme.textTertiary, fontSize: "11px", fontFamily: "inherit" }}>
          {filtered.length} tickets
        </span>
      </div>

      {/* Ticket drill-down modal */}
      {expandedTickets && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={() => setExpandedTickets(null)}>
          <div style={{ background: "#0D1B2A", border: "1px solid #1E3A5A", borderRadius: "10px", padding: "24px", maxWidth: "800px", width: "90%", maxHeight: "80vh", overflow: "auto" }}
            onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
              <span style={{ color: "#E8D5B0", fontFamily: "inherit", fontSize: "12px", letterSpacing: "0.08em" }}>{expandedTickets.label} — {expandedTickets.tickets.length} tickets</span>
              <button onClick={() => setExpandedTickets(null)} style={{ background: "none", border: "none", color: "#7A9BB5", cursor: "pointer", fontSize: "18px" }}>✕</button>
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11px", fontFamily: "inherit" }}>
              <thead>
                <tr>{["Key", "Summary", "Status", "Assignee"].map(h => (
                  <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: "#4A7FA5", borderBottom: "1px solid #1E3A5A", letterSpacing: "0.06em", fontSize: "10px", textTransform: "uppercase" }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {expandedTickets.tickets.map((t, i) => {
                  const pal = STATUS_PALETTE[t.Status] || { bg: "#1A2B3C", text: "#aaa" };
                  return (
                    <tr key={i} style={{ borderBottom: "1px solid #0A1520" }}>
                      <td style={{ padding: "6px 10px", color: "#7AB8E0", whiteSpace: "nowrap" }}>{t.Key}</td>
                      <td style={{ padding: "6px 10px", color: "#C8D8E8", maxWidth: "300px" }}>{t.Summary}</td>
                      <td style={{ padding: "6px 10px" }}><span style={{ background: pal.bg, color: pal.text, padding: "2px 8px", borderRadius: "4px", fontSize: "10px", whiteSpace: "nowrap" }}>{t.Status}</span></td>
                      <td style={{ padding: "6px 10px", color: "#8AABB5", whiteSpace: "nowrap" }}>{t.Assignee}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {pivotMode === "status" ? (
        <StatusPivot rows={filtered} statuses={statuses} epicLinks={epicLinks} onCellClick={setExpandedTickets} />
      ) : (
        <OwnerPivot rows={filtered} statuses={statuses} owners={owners} onCellClick={setExpandedTickets} />
      )}
    </div>
  );
}

function StatusPivot({ rows, statuses, epicLinks, onCellClick }) {
  const epicRowData = useMemo(() => epicLinks.map(el => {
    const { epicNum, comp } = parseEpicInfo(el);
    const grp = rows.filter(r => (r["Epic Link"] || "") === el);
    if (grp.length === 0) return null;
    const counts = statuses.map(s => grp.filter(r => r.Status === s));
    return { el, epicNum, comp, grp, counts };
  }).filter(Boolean), [rows, statuses, epicLinks]);

  const totals = statuses.map(s => rows.filter(r => r.Status === s));

  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>
        <thead>
          <tr>
            <th style={thStyle("#081521", true)}>Epic</th>
            <th style={thStyle("#081521", true)}>Component</th>
            {statuses.map(s => {
              const pal = STATUS_PALETTE[s] || { bg: "#1A2B3C", text: "#aaa" };
              return <th key={s} style={{ ...thStyle("#081521"), writingMode: "vertical-rl", transform: "rotate(180deg)", height: "90px", verticalAlign: "bottom", padding: "8px 6px" }}>
                <span style={{ background: pal.bg, color: pal.text, padding: "2px 6px", borderRadius: "3px", fontSize: "9px", whiteSpace: "nowrap" }}>{s}</span>
              </th>;
            })}
            <th style={thStyle("#081521", true)}>Total</th>
          </tr>
        </thead>
        <tbody>
          {epicRowData.map(({ el, epicNum, comp, grp, counts }, ri) => (
            <tr key={el} style={{ background: ri % 2 === 0 ? "#0A1520" : "#0D1B2A" }}>
              <td style={{ ...tdStyle, color: "#7AB8E0", whiteSpace: "nowrap", fontSize: "10px" }}>{epicNum || "—"}</td>
              <td style={{ ...tdStyle, color: "#C8D8E8", maxWidth: "220px", wordBreak: "break-word" }}>{comp}</td>
              {counts.map((tickets, si) => (
                <td key={si} style={{ ...tdStyle, textAlign: "center", cursor: tickets.length > 0 ? "pointer" : "default" }}
                  onClick={() => tickets.length > 0 && onCellClick({
                    label: `${comp} / ${statuses[si]}`, tickets
                  })}>
                  {tickets.length > 0 ? (
                    <span style={{
                      background: STATUS_PALETTE[statuses[si]]?.bg || "#1A2B3C",
                      color: STATUS_PALETTE[statuses[si]]?.text || "#aaa",
                      padding: "2px 8px", borderRadius: "12px", fontWeight: "600", fontSize: "11px",
                      cursor: "pointer", transition: "opacity 0.15s"
                    }}>{tickets.length}</span>
                  ) : <span style={{ color: "#1E3A5A" }}>·</span>}
                </td>
              ))}
              <td style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: "#E8D5B0" }}>{grp.length}</td>
            </tr>
          ))}
          <tr style={{ background: "#0F2340", borderTop: "2px solid #2A3D5A" }}>
            <td colSpan={2} style={{ ...tdStyle, fontWeight: "700", color: "#E8D5B0", letterSpacing: "0.06em", fontSize: "10px" }}>TOTAL</td>
            {totals.map((tickets, si) => (
              <td key={si} style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: "#E8D5B0", cursor: tickets.length > 0 ? "pointer" : "default" }}
                onClick={() => tickets.length > 0 && onCellClick({ label: statuses[si], tickets })}>
                {tickets.length || "·"}
              </td>
            ))}
            <td style={{ ...tdStyle, textAlign: "center", fontWeight: "800", color: "#E8D5B0" }}>{rows.length}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function OwnerPivot({ rows, statuses, owners, onCellClick }) {
  const ownerData = useMemo(() => owners.map(owner => {
    const grp = rows.filter(r => (r.Assignee || "Unassigned") === owner);
    if (grp.length === 0) return null;
    const counts = statuses.map(s => grp.filter(r => r.Status === s));
    return { owner, grp, counts };
  }).filter(Boolean), [rows, statuses, owners]);

  const totals = statuses.map(s => rows.filter(r => r.Status === s));

  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>
        <thead>
          <tr>
            <th style={thStyle("#081521", true)}>Assignee</th>
            {statuses.map(s => {
              const pal = STATUS_PALETTE[s] || { bg: "#1A2B3C", text: "#aaa" };
              return <th key={s} style={{ ...thStyle("#081521"), writingMode: "vertical-rl", transform: "rotate(180deg)", height: "90px", verticalAlign: "bottom", padding: "8px 6px" }}>
                <span style={{ background: pal.bg, color: pal.text, padding: "2px 6px", borderRadius: "3px", fontSize: "9px", whiteSpace: "nowrap" }}>{s}</span>
              </th>;
            })}
            <th style={thStyle("#081521", true)}>Total</th>
          </tr>
        </thead>
        <tbody>
          {ownerData.map(({ owner, grp, counts }, ri) => (
            <tr key={owner} style={{ background: ri % 2 === 0 ? "#0A1520" : "#0D1B2A" }}>
              <td style={{ ...tdStyle, color: "#8AABB5", whiteSpace: "nowrap" }}>{owner}</td>
              {counts.map((tickets, si) => (
                <td key={si} style={{ ...tdStyle, textAlign: "center", cursor: tickets.length > 0 ? "pointer" : "default" }}
                  onClick={() => tickets.length > 0 && onCellClick({ label: `${owner} / ${statuses[si]}`, tickets })}>
                  {tickets.length > 0 ? (
                    <span style={{
                      background: STATUS_PALETTE[statuses[si]]?.bg || "#1A2B3C",
                      color: STATUS_PALETTE[statuses[si]]?.text || "#aaa",
                      padding: "2px 8px", borderRadius: "12px", fontWeight: "600", fontSize: "11px"
                    }}>{tickets.length}</span>
                  ) : <span style={{ color: "#1E3A5A" }}>·</span>}
                </td>
              ))}
              <td style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: "#E8D5B0" }}>{grp.length}</td>
            </tr>
          ))}
          <tr style={{ background: "#0F2340", borderTop: "2px solid #2A3D5A" }}>
            <td style={{ ...tdStyle, fontWeight: "700", color: "#E8D5B0", letterSpacing: "0.06em", fontSize: "10px" }}>TOTAL</td>
            {totals.map((tickets, si) => (
              <td key={si} onClick={() => tickets.length > 0 && onCellClick({ label: statuses[si], tickets })}
                style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: "#E8D5B0", cursor: tickets.length > 0 ? "pointer" : "default" }}>
                {tickets.length || "·"}
              </td>
            ))}
            <td style={{ ...tdStyle, textAlign: "center", fontWeight: "800", color: "#E8D5B0" }}>{rows.length}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

const selectStyle = {
  background: "#0A1520", border: "1px solid #2A3D5A", color: "#7A9BB5",
  padding: "5px 10px", borderRadius: "6px", fontSize: "11px", fontFamily: "'DM Mono',monospace", cursor: "pointer"
};
const thStyle = (bg, bold = false) => ({
  background: bg, color: "#4A7FA5", padding: "8px 12px", textAlign: "left",
  letterSpacing: "0.08em", fontSize: "10px", textTransform: "uppercase",
  borderBottom: "1px solid #1E3A5A", fontWeight: bold ? "700" : "500", whiteSpace: "nowrap"
});
const tdStyle = { padding: "6px 12px", borderBottom: "1px solid #0A1520", verticalAlign: "middle", fontSize: "11px" };

// ─── ticket list view ──────────────────────────────────────────────────────────

function TicketList({ rows }) {
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterEpic, setFilterEpic] = useState("all");
  const [sortKey, setSortKey] = useState("Status");

  const statuses = useMemo(() => [...new Set(rows.map(r => r.Status))].sort((a, b) => statusSortKey(a) - statusSortKey(b)), [rows]);
  const epicLinks = useMemo(() => [...new Set(rows.map(r => r["Epic Link"] || ""))].sort(), [rows]);

  const filtered = useMemo(() => {
    let r = rows;
    if (filterStatus !== "all") r = r.filter(t => t.Status === filterStatus);
    if (filterEpic !== "all") r = r.filter(t => (t["Epic Link"] || "") === filterEpic);
    if (search) {
      const q = search.toLowerCase();
      r = r.filter(t => t.Key?.toLowerCase().includes(q) || t.Summary?.toLowerCase().includes(q) || t.Assignee?.toLowerCase().includes(q));
    }
    return [...r].sort((a, b) => {
      if (sortKey === "Status") return statusSortKey(a.Status) - statusSortKey(b.Status);
      return (a[sortKey] || "").localeCompare(b[sortKey] || "");
    });
  }, [rows, filterStatus, filterEpic, search, sortKey]);

  return (
    <div>
      <div style={{ display: "flex", gap: "10px", marginBottom: "16px", flexWrap: "wrap", alignItems: "center" }}>
        <input placeholder="Search key, summary, assignee…" value={search} onChange={e => setSearch(e.target.value)}
          style={{ ...selectStyle, width: "260px", outline: "none" }} />
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={selectStyle}>
          <option value="all">All Statuses</option>
          {statuses.map(s => <option key={s}>{s}</option>)}
        </select>
        <select value={filterEpic} onChange={e => setFilterEpic(e.target.value)} style={selectStyle}>
          <option value="all">All Epics</option>
          {epicLinks.map(el => { const { epicNum, comp } = parseEpicInfo(el); return <option key={el} value={el}>{epicNum ? `${epicNum} – ${comp}` : comp}</option>; })}
        </select>
        <select value={sortKey} onChange={e => setSortKey(e.target.value)} style={selectStyle}>
          {["Status", "Key", "Assignee"].map(k => <option key={k}>Sort: {k}</option>)}
        </select>
        <span style={{ color: "#4A7FA5", fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>{filtered.length} tickets</span>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>
          <thead>
            <tr>
              {["Key", "Summary", "Status", "Assignee", "Epic", "Start Date", "Dev End date", "QA End Date"].map(h => (
                <th key={h} style={thStyle("#081521", true)}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((t, i) => {
              const { epicNum, comp } = parseEpicInfo(t["Epic Link"] || "");
              const pal = STATUS_PALETTE[t.Status] || { bg: "#1A2B3C", text: "#aaa" };
              return (
                <tr key={i} style={{ background: i % 2 === 0 ? "#0A1520" : "#0D1B2A", transition: "background 0.1s" }}
                  onMouseEnter={e => e.currentTarget.style.background = "#112030"}
                  onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? "#0A1520" : "#0D1B2A"}>
                  <td style={{ ...tdStyle, color: "#7AB8E0", whiteSpace: "nowrap", fontWeight: "600" }}>{t.Key}</td>
                  <td style={{ ...tdStyle, color: "#C8D8E8", maxWidth: "320px" }}>{t.Summary}</td>
                  <td style={{ ...tdStyle, whiteSpace: "nowrap" }}><span style={{ background: pal.bg, color: pal.text, padding: "2px 8px", borderRadius: "4px", fontSize: "10px" }}>{t.Status}</span></td>
                  <td style={{ ...tdStyle, color: "#8AABB5", whiteSpace: "nowrap" }}>{t.Assignee || "—"}</td>
                  <td style={{ ...tdStyle, color: "#4A7FA5", whiteSpace: "nowrap", fontSize: "10px" }}>{epicNum || comp.slice(0, 20)}</td>
                  <td style={{ ...tdStyle, color: "#596B7A", whiteSpace: "nowrap" }}>{t["Start Date"] || "—"}</td>
                  <td style={{ ...tdStyle, color: "#596B7A", whiteSpace: "nowrap" }}>{t["Dev End date"] || "—"}</td>
                  <td style={{ ...tdStyle, color: "#596B7A", whiteSpace: "nowrap" }}>{t["QA End Date"] || "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── main app ─────────────────────────────────────────────────────────────────

export default function App() {
  const [rows, setRows] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState("pivot");
  const [generating, setGenerating] = useState(false);
  const [fileName, setFileName] = useState("");

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setLoading(true);
    setError(null);
    setFileName(file.name);
    try {
      const text = await file.text();
      const parser = new DOMParser();
      const doc = parser.parseFromString(text, "text/html");
      const tables = doc.querySelectorAll("table");
      let dataTable = null;
      for (const t of tables) {
        const rows = t.querySelectorAll("tr");
        if (rows.length > 5) { dataTable = t; break; }
      }
      if (!dataTable) throw new Error("No data table found in file");

      const allRows = dataTable.querySelectorAll("tr");
      let headers = null;
      const data = [];
      for (const tr of allRows) {
        const cells = [...tr.querySelectorAll("td,th")].map(td => td.textContent.trim());
        if (!headers) {
          if (cells.includes("Key") && cells.includes("Summary")) headers = cells;
        } else {
          if (cells.length === headers.length && cells[0]) {
            const obj = {};
            headers.forEach((h, i) => obj[h] = cells[i] || "");
            data.push(obj);
          }
        }
      }
      if (!data.length) throw new Error("No ticket rows parsed");
      setRows(data);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  const onDrop = useCallback(e => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const handleExport = async () => {
    if (!rows) return;
    setGenerating(true);
    try { await buildXlsx(rows); }
    catch (e) { setError("Export failed: " + e.message); }
    finally { setGenerating(false); }
  };

  return (
    <div style={{
      minHeight: "100vh", background: "#060E18",
      fontFamily: "'DM Mono', monospace", color: "#C8D8E8",
      padding: "0"
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:ital,wght@0,300;0,400;0,500;1,400&family=Syne:wght@600;700;800&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width:6px; height:6px; }
        ::-webkit-scrollbar-track { background:#060E18; }
        ::-webkit-scrollbar-thumb { background:#1E3A5A; border-radius:3px; }
        input:focus { border-color:#4A7FA5 !important; }
        select option { background:#0D1B2A; }
      `}</style>

      {/* Header */}
      <div style={{ background: "#081521", borderBottom: "1px solid #1E3A5A", padding: "16px 32px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: "12px" }}>
          <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: "20px", color: "#E8D5B0", letterSpacing: "-0.02em" }}>NGSB</span>
          <span style={{ color: "#2A3D5A", fontSize: "18px" }}>|</span>
          <span style={{ fontSize: "12px", color: "#4A7FA5", letterSpacing: "0.1em", textTransform: "uppercase" }}>Jira Sprint Processor</span>
        </div>
        {rows && (
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <span style={{ fontSize: "11px", color: "#4A7FA5" }}>{rows.length} tickets · {fileName}</span>
            <button onClick={handleExport} disabled={generating} style={{
              background: generating ? "#0A1520" : "#0F2340", border: "1px solid #2A6A9A",
              color: generating ? "#4A7FA5" : "#7AB8E0", padding: "7px 18px", borderRadius: "6px",
              cursor: generating ? "not-allowed" : "pointer", fontSize: "11px", fontFamily: "inherit",
              letterSpacing: "0.06em", transition: "all 0.2s"
            }}>
              {generating ? "⏳ Generating…" : "⬇ Export Excel"}
            </button>
            <button onClick={() => { setRows(null); setFileName(""); }} style={{
              background: "transparent", border: "1px solid #2A3D5A", color: "#4A7FA5",
              padding: "7px 14px", borderRadius: "6px", cursor: "pointer", fontSize: "11px", fontFamily: "inherit"
            }}>✕ Clear</button>
          </div>
        )}
      </div>

      <div style={{ padding: "28px 32px" }}>
        {/* Upload zone */}
        {!rows && (
          <div onDrop={onDrop} onDragOver={e => e.preventDefault()}
            style={{
              border: "1px dashed #2A3D5A", borderRadius: "12px", padding: "64px",
              textAlign: "center", marginBottom: "24px", transition: "border-color 0.2s", cursor: "pointer",
              background: "linear-gradient(135deg,#081521 0%,#060E18 100%)"
            }}
            onDragEnter={e => e.currentTarget.style.borderColor = "#4A7FA5"}
            onDragLeave={e => e.currentTarget.style.borderColor = "#2A3D5A"}
            onClick={() => document.getElementById("file-in").click()}>
            <input id="file-in" type="file" accept=".xls,.xlsx,.html,.htm" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
            <div style={{ fontSize: "36px", marginBottom: "12px" }}>📊</div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: "16px", color: "#C8D8E8", marginBottom: "8px" }}>Drop your IQVIA Jira export here</div>
            <div style={{ fontSize: "12px", color: "#4A7FA5" }}>Supports .xls / .xlsx / HTML exports from Jira</div>
            {loading && <div style={{ marginTop: "16px", color: "#7AB8E0", fontSize: "12px" }}>⏳ Parsing file…</div>}
            {error && <div style={{ marginTop: "16px", color: "#FF8080", fontSize: "12px" }}>⚠ {error}</div>}
          </div>
        )}

        {rows && (
          <>
            {/* Tabs */}
            <div style={{ display: "flex", gap: 0, borderBottom: "1px solid #1E3A5A", marginBottom: "24px" }}>
              {[["pivot", "Pivot Tables"], ["tickets", "All Tickets"]].map(([key, label]) => (
                <button key={key} onClick={() => setTab(key)} style={{
                  padding: "10px 24px", background: "transparent", border: "none",
                  borderBottom: tab === key ? "2px solid #7AB8E0" : "2px solid transparent",
                  color: tab === key ? "#E8D5B0" : "#4A7FA5", cursor: "pointer",
                  fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: "12px", letterSpacing: "0.06em",
                  textTransform: "uppercase", marginBottom: "-1px", transition: "color 0.15s"
                }}>{label}</button>
              ))}
            </div>

            {tab === "pivot" && <PivotTable rows={rows} />}
            {tab === "tickets" && <TicketList rows={rows} />}
          </>
        )}
      </div>
    </div>
  );
}
