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
    emptyCell: "#1E3A5A",
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
    emptyCell: "#CBD5E1",
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

    aoa.push([`${comp}  |  ${epicNum}`]);
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: COLS.length - 1 } });

    statuses.forEach(status => {
      aoa.push([`◆  ${status.toUpperCase()}`]);
      merges.push({ s: { r: aoa.length - 1, c: 0 }, e: { r: aoa.length - 1, c: COLS.length - 1 } });
      aoa.push(COLS);
      grp.filter(r => r.Status === status).forEach(r => {
        aoa.push(COLS.map(c => r[c] || ""));
      });
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

// ─── multi-select component ──────────────────────────────────────────────────

function MultiSelect({ options, selected, onChange, label, theme, renderOption }) {
  const [isOpen, setIsOpen] = useState(false);

  const toggleOption = (value) => {
    if (selected.includes(value)) {
      onChange(selected.filter(v => v !== value));
    } else {
      onChange([...selected, value]);
    }
  };

  const selectAll = () => onChange(options);
  const clearAll = () => onChange([]);

  const displayText = selected.length === 0 ? `All ${label}` :
    selected.length === options.length ? `All ${label}` :
    `${selected.length} ${label}`;

  return (
    <div style={{ position: "relative" }}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        style={{
          background: theme.inputBg, border: `1px solid ${theme.inputBorder}`, color: theme.textSecondary,
          padding: "5px 10px", borderRadius: "6px", fontSize: "11px", fontFamily: "'DM Mono',monospace",
          cursor: "pointer", transition: "all 0.3s ease", display: "flex", alignItems: "center", gap: "6px",
          minWidth: "140px", justifyContent: "space-between"
        }}
      >
        <span>{displayText}</span>
        <span style={{ fontSize: "9px" }}>▼</span>
      </button>
      {isOpen && (
        <>
          <div style={{ position: "fixed", inset: 0, zIndex: 999 }} onClick={() => setIsOpen(false)} />
          <div style={{
            position: "absolute", top: "100%", left: 0, marginTop: "4px", background: theme.bgModal,
            border: `1px solid ${theme.border}`, borderRadius: "6px", padding: "8px", minWidth: "200px",
            maxHeight: "300px", overflowY: "auto", zIndex: 1000, boxShadow: "0 4px 12px rgba(0,0,0,0.3)",
            transition: "all 0.3s ease"
          }}>
            <div style={{ display: "flex", gap: "6px", marginBottom: "8px", paddingBottom: "8px", borderBottom: `1px solid ${theme.borderLight}` }}>
              <button onClick={selectAll} style={{
                flex: 1, padding: "4px 8px", background: theme.bgTertiary, border: `1px solid ${theme.borderLight}`,
                color: theme.textTertiary, borderRadius: "4px", fontSize: "10px", cursor: "pointer",
                transition: "all 0.2s ease"
              }}>All</button>
              <button onClick={clearAll} style={{
                flex: 1, padding: "4px 8px", background: theme.bgTertiary, border: `1px solid ${theme.borderLight}`,
                color: theme.textTertiary, borderRadius: "4px", fontSize: "10px", cursor: "pointer",
                transition: "all 0.2s ease"
              }}>None</button>
            </div>
            {options.map(option => {
              const isSelected = selected.includes(option);
              return (
                <div
                  key={option}
                  onClick={() => toggleOption(option)}
                  style={{
                    padding: "6px 8px", cursor: "pointer", borderRadius: "4px",
                    background: isSelected ? theme.bgHover : "transparent",
                    color: isSelected ? theme.accent : theme.textSecondary,
                    fontSize: "11px", display: "flex", alignItems: "center", gap: "8px",
                    transition: "all 0.15s ease", marginBottom: "2px"
                  }}
                  onMouseEnter={e => !isSelected && (e.currentTarget.style.background = theme.bgTertiary)}
                  onMouseLeave={e => !isSelected && (e.currentTarget.style.background = "transparent")}
                >
                  <span style={{ width: "14px", height: "14px", border: `1px solid ${theme.borderLight}`, borderRadius: "3px",
                    background: isSelected ? theme.accent : "transparent", display: "flex", alignItems: "center",
                    justifyContent: "center", fontSize: "10px", color: isSelected ? theme.bg : "transparent" }}>✓</span>
                  {renderOption ? renderOption(option) : <span>{option}</span>}
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}

function PivotTable({ rows, theme }) {
  const [pivotMode, setPivotMode] = useState("status");
  const [filterStatus, setFilterStatus] = useState([]);
  const [filterEpic, setFilterEpic] = useState([]);
  const [filterOwner, setFilterOwner] = useState([]);

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
    if (filterStatus.length > 0 && !filterStatus.includes(r.Status)) return false;
    if (filterEpic.length > 0 && !filterEpic.includes(r["Epic Link"] || "")) return false;
    if (filterOwner.length > 0 && !filterOwner.includes(r.Assignee || "Unassigned")) return false;
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
        <MultiSelect
          options={statuses}
          selected={filterStatus}
          onChange={setFilterStatus}
          label="Statuses"
          theme={theme}
          renderOption={(s) => {
            const pal = STATUS_PALETTE[s] || { bg: "#1A2B3C", text: "#aaa" };
            return <span style={{ background: pal.bg, color: pal.text, padding: "2px 6px", borderRadius: "3px", fontSize: "10px" }}>{s}</span>;
          }}
        />
        <MultiSelect
          options={epicLinks}
          selected={filterEpic}
          onChange={setFilterEpic}
          label="Epics"
          theme={theme}
          renderOption={(el) => {
            const { epicNum, comp } = parseEpicInfo(el);
            return <span style={{ fontSize: "10px" }}>{epicNum ? `${epicNum} – ${comp.slice(0, 30)}` : comp.slice(0, 30)}</span>;
          }}
        />
        <MultiSelect
          options={owners}
          selected={filterOwner}
          onChange={setFilterOwner}
          label="Owners"
          theme={theme}
        />
        <span style={{ color: theme.textTertiary, fontSize: "11px", fontFamily: "inherit" }}>
          {filtered.length} tickets
        </span>
      </div>

      {expandedTickets && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={() => setExpandedTickets(null)}>
          <div style={{ background: theme.bgModal, border: `1px solid ${theme.border}`, borderRadius: "10px", padding: "24px", maxWidth: "800px", width: "90%", maxHeight: "80vh", overflow: "auto", transition: "all 0.3s ease" }}
            onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
              <span style={{ color: theme.accent, fontFamily: "inherit", fontSize: "12px", letterSpacing: "0.08em" }}>{expandedTickets.label} — {expandedTickets.tickets.length} tickets</span>
              <button onClick={() => setExpandedTickets(null)} style={{ background: "none", border: "none", color: theme.textSecondary, cursor: "pointer", fontSize: "18px" }}>✕</button>
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11px", fontFamily: "inherit" }}>
              <thead>
                <tr>{["Key", "Summary", "Status", "Assignee"].map(h => (
                  <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: theme.textTertiary, borderBottom: `1px solid ${theme.border}`, letterSpacing: "0.06em", fontSize: "10px", textTransform: "uppercase" }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {expandedTickets.tickets.map((t, i) => {
                  const pal = STATUS_PALETTE[t.Status] || { bg: "#1A2B3C", text: "#aaa" };
                  return (
                    <tr key={i} style={{ borderBottom: `1px solid ${theme.borderLight}` }}>
                      <td style={{ padding: "6px 10px", color: theme.accentSecondary, whiteSpace: "nowrap" }}>{t.Key}</td>
                      <td style={{ padding: "6px 10px", color: theme.text, maxWidth: "300px" }}>{t.Summary}</td>
                      <td style={{ padding: "6px 10px" }}><span style={{ background: pal.bg, color: pal.text, padding: "2px 8px", borderRadius: "4px", fontSize: "10px", whiteSpace: "nowrap" }}>{t.Status}</span></td>
                      <td style={{ padding: "6px 10px", color: theme.textSecondary, whiteSpace: "nowrap" }}>{t.Assignee}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {pivotMode === "status" ? (
        <StatusPivot rows={filtered} statuses={statuses} epicLinks={epicLinks} onCellClick={setExpandedTickets} theme={theme} />
      ) : (
        <OwnerPivot rows={filtered} statuses={statuses} owners={owners} onCellClick={setExpandedTickets} theme={theme} />
      )}
    </div>
  );
}

function StatusPivot({ rows, statuses, epicLinks, onCellClick, theme }) {
  const epicRowData = useMemo(() => epicLinks.map(el => {
    const { epicNum, comp } = parseEpicInfo(el);
    const grp = rows.filter(r => (r["Epic Link"] || "") === el);
    if (grp.length === 0) return null;
    const counts = statuses.map(s => grp.filter(r => r.Status === s));
    return { el, epicNum, comp, grp, counts };
  }).filter(Boolean), [rows, statuses, epicLinks]);

  const totals = statuses.map(s => rows.filter(r => r.Status === s));

  const thStyle = (bold = false) => ({
    background: theme.tableHeader, color: theme.textTertiary, padding: "8px 12px", textAlign: "left",
    letterSpacing: "0.08em", fontSize: "10px", textTransform: "uppercase",
    borderBottom: `1px solid ${theme.border}`, fontWeight: bold ? "700" : "500", whiteSpace: "nowrap",
    transition: "all 0.3s ease"
  });

  const tdStyle = {
    padding: "6px 12px", borderBottom: `1px solid ${theme.borderLight}`, verticalAlign: "middle", fontSize: "11px",
    transition: "all 0.3s ease"
  };

  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>
        <thead>
          <tr>
            <th style={thStyle(true)}>Epic</th>
            <th style={thStyle(true)}>Component</th>
            {statuses.map(s => {
              const pal = STATUS_PALETTE[s] || { bg: "#1A2B3C", text: "#aaa" };
              return <th key={s} style={{ ...thStyle(), writingMode: "vertical-rl", transform: "rotate(180deg)", height: "90px", verticalAlign: "bottom", padding: "8px 6px" }}>
                <span style={{ background: pal.bg, color: pal.text, padding: "2px 6px", borderRadius: "3px", fontSize: "9px", whiteSpace: "nowrap" }}>{s}</span>
              </th>;
            })}
            <th style={thStyle(true)}>Total</th>
          </tr>
        </thead>
        <tbody>
          {epicRowData.map(({ el, epicNum, comp, grp, counts }, ri) => (
            <tr key={el} style={{ background: ri % 2 === 0 ? theme.tableRowEven : theme.tableRowOdd, transition: "background 0.3s ease" }}>
              <td style={{ ...tdStyle, color: theme.accentSecondary, whiteSpace: "nowrap", fontSize: "10px" }}>{epicNum || "—"}</td>
              <td style={{ ...tdStyle, color: theme.text, maxWidth: "220px", wordBreak: "break-word" }}>{comp}</td>
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
                  ) : <span style={{ color: theme.emptyCell }}>·</span>}
                </td>
              ))}
              <td style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: theme.accent }}>{grp.length}</td>
            </tr>
          ))}
          <tr style={{ background: theme.totalRow, borderTop: `2px solid ${theme.borderLight}`, transition: "background 0.3s ease" }}>
            <td colSpan={2} style={{ ...tdStyle, fontWeight: "700", color: theme.accent, letterSpacing: "0.06em", fontSize: "10px" }}>TOTAL</td>
            {totals.map((tickets, si) => (
              <td key={si} style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: theme.accent, cursor: tickets.length > 0 ? "pointer" : "default" }}
                onClick={() => tickets.length > 0 && onCellClick({ label: statuses[si], tickets })}>
                {tickets.length || "·"}
              </td>
            ))}
            <td style={{ ...tdStyle, textAlign: "center", fontWeight: "800", color: theme.accent }}>{rows.length}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function OwnerPivot({ rows, statuses, owners, onCellClick, theme }) {
  const ownerData = useMemo(() => owners.map(owner => {
    const grp = rows.filter(r => (r.Assignee || "Unassigned") === owner);
    if (grp.length === 0) return null;
    const counts = statuses.map(s => grp.filter(r => r.Status === s));
    return { owner, grp, counts };
  }).filter(Boolean), [rows, statuses, owners]);

  const totals = statuses.map(s => rows.filter(r => r.Status === s));

  const thStyle = (bold = false) => ({
    background: theme.tableHeader, color: theme.textTertiary, padding: "8px 12px", textAlign: "left",
    letterSpacing: "0.08em", fontSize: "10px", textTransform: "uppercase",
    borderBottom: `1px solid ${theme.border}`, fontWeight: bold ? "700" : "500", whiteSpace: "nowrap",
    transition: "all 0.3s ease"
  });

  const tdStyle = {
    padding: "6px 12px", borderBottom: `1px solid ${theme.borderLight}`, verticalAlign: "middle", fontSize: "11px",
    transition: "all 0.3s ease"
  };

  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>
        <thead>
          <tr>
            <th style={thStyle(true)}>Assignee</th>
            {statuses.map(s => {
              const pal = STATUS_PALETTE[s] || { bg: "#1A2B3C", text: "#aaa" };
              return <th key={s} style={{ ...thStyle(), writingMode: "vertical-rl", transform: "rotate(180deg)", height: "90px", verticalAlign: "bottom", padding: "8px 6px" }}>
                <span style={{ background: pal.bg, color: pal.text, padding: "2px 6px", borderRadius: "3px", fontSize: "9px", whiteSpace: "nowrap" }}>{s}</span>
              </th>;
            })}
            <th style={thStyle(true)}>Total</th>
          </tr>
        </thead>
        <tbody>
          {ownerData.map(({ owner, grp, counts }, ri) => (
            <tr key={owner} style={{ background: ri % 2 === 0 ? theme.tableRowEven : theme.tableRowOdd, transition: "background 0.3s ease" }}>
              <td style={{ ...tdStyle, color: theme.textSecondary, whiteSpace: "nowrap" }}>{owner}</td>
              {counts.map((tickets, si) => (
                <td key={si} style={{ ...tdStyle, textAlign: "center", cursor: tickets.length > 0 ? "pointer" : "default" }}
                  onClick={() => tickets.length > 0 && onCellClick({ label: `${owner} / ${statuses[si]}`, tickets })}>
                  {tickets.length > 0 ? (
                    <span style={{
                      background: STATUS_PALETTE[statuses[si]]?.bg || "#1A2B3C",
                      color: STATUS_PALETTE[statuses[si]]?.text || "#aaa",
                      padding: "2px 8px", borderRadius: "12px", fontWeight: "600", fontSize: "11px"
                    }}>{tickets.length}</span>
                  ) : <span style={{ color: theme.emptyCell }}>·</span>}
                </td>
              ))}
              <td style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: theme.accent }}>{grp.length}</td>
            </tr>
          ))}
          <tr style={{ background: theme.totalRow, borderTop: `2px solid ${theme.borderLight}`, transition: "background 0.3s ease" }}>
            <td style={{ ...tdStyle, fontWeight: "700", color: theme.accent, letterSpacing: "0.06em", fontSize: "10px" }}>TOTAL</td>
            {totals.map((tickets, si) => (
              <td key={si} onClick={() => tickets.length > 0 && onCellClick({ label: statuses[si], tickets })}
                style={{ ...tdStyle, textAlign: "center", fontWeight: "700", color: theme.accent, cursor: tickets.length > 0 ? "pointer" : "default" }}>
                {tickets.length || "·"}
              </td>
            ))}
            <td style={{ ...tdStyle, textAlign: "center", fontWeight: "800", color: theme.accent }}>{rows.length}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ─── ticket list view ──────────────────────────────────────────────────────────

function TicketList({ rows, theme }) {
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState([]);
  const [filterEpic, setFilterEpic] = useState([]);
  const [sortKey, setSortKey] = useState("Status");

  const statuses = useMemo(() => [...new Set(rows.map(r => r.Status))].sort((a, b) => statusSortKey(a) - statusSortKey(b)), [rows]);
  const epicLinks = useMemo(() => [...new Set(rows.map(r => r["Epic Link"] || ""))].sort(), [rows]);

  const filtered = useMemo(() => {
    let r = rows;
    if (filterStatus.length > 0) r = r.filter(t => filterStatus.includes(t.Status));
    if (filterEpic.length > 0) r = r.filter(t => filterEpic.includes(t["Epic Link"] || ""));
    if (search) {
      const q = search.toLowerCase();
      r = r.filter(t => t.Key?.toLowerCase().includes(q) || t.Summary?.toLowerCase().includes(q) || t.Assignee?.toLowerCase().includes(q));
    }
    return [...r].sort((a, b) => {
      if (sortKey === "Status") return statusSortKey(a.Status) - statusSortKey(b.Status);
      return (a[sortKey] || "").localeCompare(b[sortKey] || "");
    });
  }, [rows, filterStatus, filterEpic, search, sortKey]);

  const selectStyleThemed = {
    background: theme.inputBg, border: `1px solid ${theme.inputBorder}`, color: theme.textSecondary,
    padding: "5px 10px", borderRadius: "6px", fontSize: "11px", fontFamily: "'DM Mono',monospace", cursor: "pointer",
    transition: "all 0.3s ease"
  };

  const thStyle = (bold = false) => ({
    background: theme.tableHeader, color: theme.textTertiary, padding: "8px 12px", textAlign: "left",
    letterSpacing: "0.08em", fontSize: "10px", textTransform: "uppercase",
    borderBottom: `1px solid ${theme.border}`, fontWeight: bold ? "700" : "500", whiteSpace: "nowrap",
    transition: "all 0.3s ease"
  });

  const tdStyle = {
    padding: "6px 12px", borderBottom: `1px solid ${theme.borderLight}`, verticalAlign: "middle", fontSize: "11px",
    transition: "all 0.3s ease"
  };

  return (
    <div>
      <div style={{ display: "flex", gap: "10px", marginBottom: "16px", flexWrap: "wrap", alignItems: "center" }}>
        <input placeholder="Search key, summary, assignee…" value={search} onChange={e => setSearch(e.target.value)}
          style={{ ...selectStyleThemed, width: "260px", outline: "none" }} />
        <MultiSelect
          options={statuses}
          selected={filterStatus}
          onChange={setFilterStatus}
          label="Statuses"
          theme={theme}
          renderOption={(s) => {
            const pal = STATUS_PALETTE[s] || { bg: "#1A2B3C", text: "#aaa" };
            return <span style={{ background: pal.bg, color: pal.text, padding: "2px 6px", borderRadius: "3px", fontSize: "10px" }}>{s}</span>;
          }}
        />
        <MultiSelect
          options={epicLinks}
          selected={filterEpic}
          onChange={setFilterEpic}
          label="Epics"
          theme={theme}
          renderOption={(el) => {
            const { epicNum, comp } = parseEpicInfo(el);
            return <span style={{ fontSize: "10px" }}>{epicNum ? `${epicNum} – ${comp.slice(0, 30)}` : comp.slice(0, 30)}</span>;
          }}
        />
        <select value={sortKey} onChange={e => setSortKey(e.target.value)} style={selectStyleThemed}>
          {["Status", "Key", "Assignee"].map(k => <option key={k}>Sort: {k}</option>)}
        </select>
        <span style={{ color: theme.textTertiary, fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>{filtered.length} tickets</span>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "11px", fontFamily: "'DM Mono',monospace" }}>
          <thead>
            <tr>
              {["Key", "Summary", "Status", "Assignee", "Epic", "Start Date", "Dev End date", "QA End Date"].map(h => (
                <th key={h} style={thStyle(true)}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((t, i) => {
              const { epicNum, comp } = parseEpicInfo(t["Epic Link"] || "");
              const pal = STATUS_PALETTE[t.Status] || { bg: "#1A2B3C", text: "#aaa" };
              return (
                <tr key={i} style={{ background: i % 2 === 0 ? theme.tableRowEven : theme.tableRowOdd, transition: "background 0.1s" }}
                  onMouseEnter={e => e.currentTarget.style.background = theme.bgHover}
                  onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? theme.tableRowEven : theme.tableRowOdd}>
                  <td style={{ ...tdStyle, color: theme.accentSecondary, whiteSpace: "nowrap", fontWeight: "600" }}>{t.Key}</td>
                  <td style={{ ...tdStyle, color: theme.text, maxWidth: "320px" }}>{t.Summary}</td>
                  <td style={{ ...tdStyle, whiteSpace: "nowrap" }}><span style={{ background: pal.bg, color: pal.text, padding: "2px 8px", borderRadius: "4px", fontSize: "10px" }}>{t.Status}</span></td>
                  <td style={{ ...tdStyle, color: theme.textSecondary, whiteSpace: "nowrap" }}>{t.Assignee || "—"}</td>
                  <td style={{ ...tdStyle, color: theme.textTertiary, whiteSpace: "nowrap", fontSize: "10px" }}>{epicNum || comp.slice(0, 20)}</td>
                  <td style={{ ...tdStyle, color: theme.textMuted, whiteSpace: "nowrap" }}>{t["Start Date"] || "—"}</td>
                  <td style={{ ...tdStyle, color: theme.textMuted, whiteSpace: "nowrap" }}>{t["Dev End date"] || "—"}</td>
                  <td style={{ ...tdStyle, color: theme.textMuted, whiteSpace: "nowrap" }}>{t["QA End Date"] || "—"}</td>
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
  const [themeMode, setThemeMode] = useState("dark");

  const theme = THEMES[themeMode];

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
      minHeight: "100vh", background: theme.bg,
      fontFamily: "'DM Mono', monospace", color: theme.text,
      padding: "0", transition: "background 0.3s ease, color 0.3s ease"
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:ital,wght@0,300;0,400;0,500;1,400&family=Syne:wght@600;700;800&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width:6px; height:6px; }
        ::-webkit-scrollbar-track { background:${theme.scrollbarTrack}; transition: background 0.3s ease; }
        ::-webkit-scrollbar-thumb { background:${theme.scrollbarThumb}; border-radius:3px; transition: background 0.3s ease; }
        input:focus { border-color:${theme.inputFocus} !important; }
        select option { background:${theme.bgQuaternary}; }
      `}</style>

      {/* Header */}
      <div style={{ background: theme.bgSecondary, borderBottom: `1px solid ${theme.border}`, padding: "16px 32px", display: "flex", alignItems: "center", justifyContent: "space-between", transition: "all 0.3s ease" }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: "12px" }}>
          <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: "20px", color: theme.accent, letterSpacing: "-0.02em" }}>NGSB</span>
          <span style={{ color: theme.borderLight, fontSize: "18px" }}>|</span>
          <span style={{ fontSize: "12px", color: theme.textTertiary, letterSpacing: "0.1em", textTransform: "uppercase" }}>Jira Sprint Processor</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
          <button onClick={() => setThemeMode(themeMode === "dark" ? "light" : "dark")} style={{
            background: theme.inputBg, border: `1px solid ${theme.inputBorder}`, color: theme.textSecondary,
            padding: "7px 14px", borderRadius: "6px", cursor: "pointer", fontSize: "16px", fontFamily: "inherit",
            transition: "all 0.3s ease", display: "flex", alignItems: "center", gap: "6px"
          }} title={`Switch to ${themeMode === "dark" ? "light" : "dark"} mode`}>
            {themeMode === "dark" ? "☀️" : "🌙"}
          </button>
          {rows && (
            <>
              <span style={{ fontSize: "11px", color: theme.textTertiary }}>{rows.length} tickets · {fileName}</span>
              <button onClick={handleExport} disabled={generating} style={{
                background: generating ? theme.bgTertiary : theme.buttonBg, border: `1px solid ${theme.buttonBorder}`,
                color: generating ? theme.textTertiary : theme.accentSecondary, padding: "7px 18px", borderRadius: "6px",
                cursor: generating ? "not-allowed" : "pointer", fontSize: "11px", fontFamily: "inherit",
                letterSpacing: "0.06em", transition: "all 0.3s ease"
              }}>
                {generating ? "⏳ Generating…" : "⬇ Export Excel"}
              </button>
              <button onClick={() => { setRows(null); setFileName(""); }} style={{
                background: "transparent", border: `1px solid ${theme.borderLight}`, color: theme.textTertiary,
                padding: "7px 14px", borderRadius: "6px", cursor: "pointer", fontSize: "11px", fontFamily: "inherit",
                transition: "all 0.3s ease"
              }}>✕ Clear</button>
            </>
          )}
        </div>
      </div>

      <div style={{ padding: "28px 32px" }}>
        {/* Upload zone */}
        {!rows && (
          <div onDrop={onDrop} onDragOver={e => e.preventDefault()}
            style={{
              border: `1px dashed ${theme.borderLight}`, borderRadius: "12px", padding: "64px",
              textAlign: "center", marginBottom: "24px", transition: "border-color 0.2s, background 0.3s ease", cursor: "pointer",
              background: `linear-gradient(135deg,${theme.bgSecondary} 0%,${theme.bg} 100%)`
            }}
            onDragEnter={e => e.currentTarget.style.borderColor = theme.textTertiary}
            onDragLeave={e => e.currentTarget.style.borderColor = theme.borderLight}
            onClick={() => document.getElementById("file-in").click()}>
            <input id="file-in" type="file" accept=".xls,.xlsx,.html,.htm" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
            <div style={{ fontSize: "36px", marginBottom: "12px" }}>📊</div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: "16px", color: theme.text, marginBottom: "8px" }}>Drop your IQVIA Jira export here</div>
            <div style={{ fontSize: "12px", color: theme.textTertiary }}>Supports .xls / .xlsx / HTML exports from Jira</div>
            {loading && <div style={{ marginTop: "16px", color: theme.accentSecondary, fontSize: "12px" }}>⏳ Parsing file…</div>}
            {error && <div style={{ marginTop: "16px", color: "#FF8080", fontSize: "12px" }}>⚠ {error}</div>}
          </div>
        )}

        {rows && (
          <>
            {/* Tabs */}
            <div style={{ display: "flex", gap: 0, borderBottom: `1px solid ${theme.border}`, marginBottom: "24px" }}>
              {[["pivot", "Pivot Tables"], ["tickets", "All Tickets"]].map(([key, label]) => (
                <button key={key} onClick={() => setTab(key)} style={{
                  padding: "10px 24px", background: "transparent", border: "none",
                  borderBottom: tab === key ? `2px solid ${theme.accentSecondary}` : "2px solid transparent",
                  color: tab === key ? theme.accent : theme.textTertiary, cursor: "pointer",
                  fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: "12px", letterSpacing: "0.06em",
                  textTransform: "uppercase", marginBottom: "-1px", transition: "all 0.3s ease"
                }}>{label}</button>
              ))}
            </div>

            {tab === "pivot" && <PivotTable rows={rows} theme={theme} />}
            {tab === "tickets" && <TicketList rows={rows} theme={theme} />}
          </>
        )}
      </div>
    </div>
  );
}
