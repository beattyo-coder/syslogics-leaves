import * as XLSX from 'xlsx';
import { MONTHS_HE } from './data.js';

export function exportToExcel(employees, leaves) {
  const approved = leaves.filter(l => l.status === "approved");
  const activeMonths = [...new Set(approved.map(l => new Date(l.startDate).getMonth()))].sort((a, b) => a - b);
  if (activeMonths.length === 0) { alert("אין נתונים לייצוא"); return; }

  const wb = XLSX.utils.book_new();

  // ═══ Sheet 1: Monthly Summary ═══
  const header1 = ["עובד"];
  const header2 = [""];
  activeMonths.forEach(m => {
    header1.push(MONTHS_HE[m], "");
    header2.push("חופשה", "מחלה");
  });
  header1.push("סה״כ חופשה", "סה״כ מחלה", "סה״כ היעדרויות");
  header2.push("", "", "");

  const rows = [header1, header2];

  employees.forEach(emp => {
    const row = [emp.name];
    let tv = 0, ts = 0;
    activeMonths.forEach(m => {
      const v = approved.filter(l => l.empId === emp.id && l.type === "vacation" && new Date(l.startDate).getMonth() === m).reduce((s, l) => s + l.days, 0);
      const s = approved.filter(l => l.empId === emp.id && l.type === "sick" && new Date(l.startDate).getMonth() === m).reduce((s, l) => s + l.days, 0);
      row.push(v || "", s || "");
      tv += v; ts += s;
    });
    row.push(tv || "", ts || "", (tv + ts) || "");
    rows.push(row);
  });

  // Totals
  const totals = ["סה״כ"];
  let gv = 0, gs = 0;
  activeMonths.forEach(m => {
    const v = approved.filter(l => l.type === "vacation" && new Date(l.startDate).getMonth() === m).reduce((s, l) => s + l.days, 0);
    const s = approved.filter(l => l.type === "sick" && new Date(l.startDate).getMonth() === m).reduce((s, l) => s + l.days, 0);
    totals.push(v || "", s || "");
    gv += v; gs += s;
  });
  totals.push(gv, gs, gv + gs);
  rows.push(totals);

  const ws1 = XLSX.utils.aoa_to_sheet(rows);

  // Merge month headers
  const merges = [];
  let col = 1;
  activeMonths.forEach(() => {
    merges.push({ s: { r: 0, c: col }, e: { r: 0, c: col + 1 } });
    col += 2;
  });
  ws1["!merges"] = merges;

  // Column widths
  ws1["!cols"] = rows[0].map((_, i) => ({ wch: i === 0 ? 14 : 12 }));

  // RTL
  ws1["!sheetViews"] = [{ rightToLeft: true }];
  XLSX.utils.book_append_sheet(wb, ws1, "סיכום חודשי");

  // ═══ Sheet 2: Detailed list ═══
  const detailRows = [["עובד", "סוג", "מתאריך", "עד תאריך", "ימים", "הערה"]];
  approved.sort((a, b) => new Date(a.startDate) - new Date(b.startDate)).forEach(l => {
    const emp = employees.find(e => e.id === l.empId);
    detailRows.push([emp?.name || "", l.type === "vacation" ? "חופשה" : "מחלה", l.startDate, l.endDate, l.days, l.note || ""]);
  });
  const ws2 = XLSX.utils.aoa_to_sheet(detailRows);
  ws2["!cols"] = [{ wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(wb, ws2, "פירוט");

  // ═══ Individual sheets per employee ═══
  employees.forEach(emp => {
    const empLeaves = approved.filter(l => l.empId === emp.id);
    if (empLeaves.length === 0) return;

    const eRows = [["סוג", "מתאריך", "עד תאריך", "ימים", "הערה"]];
    empLeaves.sort((a, b) => new Date(a.startDate) - new Date(b.startDate)).forEach(l => {
      eRows.push([l.type === "vacation" ? "חופשה" : "מחלה", l.startDate, l.endDate, l.days, l.note || ""]);
    });
    const tv = empLeaves.filter(l => l.type === "vacation").reduce((s, l) => s + l.days, 0);
    const ts = empLeaves.filter(l => l.type === "sick").reduce((s, l) => s + l.days, 0);
    eRows.push([]);
    eRows.push(["סה״כ חופשה", "", "", tv, ""]);
    eRows.push(["סה״כ מחלה", "", "", ts, ""]);
    eRows.push(["סה״כ היעדרויות", "", "", tv + ts, ""]);

    const ews = XLSX.utils.aoa_to_sheet(eRows);
    ews["!cols"] = [{ wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(wb, ews, emp.name);
  });

  XLSX.writeFile(wb, `היעדרויות_Syslogics_2026.xlsx`);
}
