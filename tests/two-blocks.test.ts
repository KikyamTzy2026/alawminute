import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { classify, DEFAULT_RULES } from "../lib/classify";
import { buildSchedule } from "../lib/schedule";
import { toMin } from "../lib/time";
import { loadWorkbook, detectLayout, writeWorkbook, dataRowIndices } from "../lib/excel";

(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Billing");
  ws.addRow(["File Number", "Comments", "Time Started", "Time Ended", "Minutes", "Initial Audit", "Time Started", "Time Ended", "Minutes", "Final Audit", "Error"]);
  const comments = ["Duplicate file. Moved to Duplicate folder.", "Completed. FDL has been drafted.", "Not our case. Mail was moved to completed.", "Moved to Completed folder. Filer name was X.", "Duplicate file."];
  comments.forEach((c, i) => ws.addRow([`F${i}`, c]));
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  const loaded = loadWorkbook("two.xlsx", buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
  const rows = loaded.sheets["Billing"]; const lay = detectLayout(rows);
  assert.equal(lay.blocks.length, 2);
  assert.deepEqual(lay.blocks[0], { startCol: 2, endCol: 3, minutesCol: 4 });
  assert.deepEqual(lay.blocks[1], { startCol: 6, endCol: 7, minutesCol: 8 });

  const recs = dataRowIndices(rows, lay).map((i) => {
    const cat = classify(String(rows[i][lay.commentCol]), DEFAULT_RULES);
    return { key: i, cat, minutes: cat === "Completed" ? 15 : 7, block: cat === "Completed" ? 1 : 0 };
  });
  const lunch = { lunchOn: true, lunchStart: toMin("11:30"), lunchEnd: toMin("12:30") };
  const b1 = buildSchedule(recs.filter((r) => r.block === 0), { startMin: toMin("08:00"), endMin: toMin("11:30"), ...lunch });
  const b2 = buildSchedule(recs.filter((r) => r.block === 1), { startMin: toMin("12:30"), endMin: toMin("17:00"), ...lunch });
  assert.equal(b1.slots.size, 3); assert.equal(b2.slots.size, 2);
  const out = await writeWorkbook(loaded, { sheetName: "Billing", layout: lay, groups: [
    { rows: recs.map((r) => ({ key: r.key, slot: b1.slots.get(r.key) })), maxMinutes: 7 },
    { rows: recs.map((r) => ({ key: r.key, slot: b2.slots.get(r.key) })), maxMinutes: 15 },
  ] });
  const chk = new ExcelJS.Workbook(); await chk.xlsx.load(out as any);
  const w = chk.worksheets[0];
  assert.equal(w.columnCount, 11);                       // no extra columns added
  assert.equal(w.getCell("F1").value, "Initial Audit"); assert.equal(w.getCell("K1").value, "Error");
  // row 2 = Duplicate -> block 1 only; row 3 = Completed -> block 2 only
  assert.ok(w.getCell("C2").value instanceof Date); assert.equal(w.getCell("G2").value, null);
  assert.ok(w.getCell("G3").value instanceof Date); assert.equal(w.getCell("C3").value, null);
  assert.equal((w.getCell("E2").value as any).formula, "MINUTE(MOD(D2-C2,1))");
  assert.equal((w.getCell("I3").value as any).formula, "MINUTE(MOD(H3-G3,1))");
  const hh = (d: any) => d.getUTCHours() * 60 + d.getUTCMinutes();
  assert.equal(hh(w.getCell("D6").value), toMin("11:30"));  // last billing record ends at 11:30
  assert.equal(hh(w.getCell("H5").value), toMin("17:00"));  // last posting record ends at 5:00 PM
  console.log("ok - two blocks: layout detection, per-block schedule, columns untouched");

  // sheet with only ONE set -> second set appended after the last column
  const wb2 = new ExcelJS.Workbook(); const s2 = wb2.addWorksheet("S");
  s2.addRow(["Comments", "Time Started", "Time Ended", "Minutes"]); s2.addRow(["Completed."]); s2.addRow(["Duplicate file."]);
  const b2buf = Buffer.from(await wb2.xlsx.writeBuffer());
  const l2 = loadWorkbook("one.xlsx", b2buf.buffer.slice(b2buf.byteOffset, b2buf.byteOffset + b2buf.byteLength) as ArrayBuffer);
  const lay2 = detectLayout(l2.sheets["S"]); assert.equal(lay2.blocks.length, 1);
  const o2 = await writeWorkbook(l2, { sheetName: "S", layout: lay2, groups: [
    { rows: [{ key: 1, slot: { start: 480, end: 495 } }, { key: 2 }], maxMinutes: 15 },
    { rows: [{ key: 1 }, { key: 2, slot: { start: 780, end: 787 } }], maxMinutes: 7 }] });
  const c2 = new ExcelJS.Workbook(); await c2.xlsx.load(o2 as any); const x = c2.worksheets[0];
  assert.deepEqual([5, 6, 7].map((c) => x.getCell(1, c).value), ["Time Started", "Time Ended", "Minutes"]);
  console.log("ok - missing second set is appended once");
})().catch((e) => { console.error(e); process.exit(1); });
