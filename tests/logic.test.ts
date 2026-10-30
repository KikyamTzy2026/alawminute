import assert from "node:assert/strict";
import fs from "node:fs";
import { classify, DEFAULT_RULES } from "../lib/classify";
import { buildSchedule, checkSchedule, ScheduleSettings } from "../lib/schedule";
import { toMin, fmt12 } from "../lib/time";
import { loadWorkbook, detectLayout, writeWorkbook, dataRowIndices } from "../lib/excel";
import ExcelJS from "exceljs";

const S: ScheduleSettings = { startMin: toMin("08:00"), endMin: toMin("13:30"), lunchOn: true, lunchStart: toMin("11:30"), lunchEnd: toMin("12:30") };
let n = 0; const ok = (name: string, fn: () => void | Promise<void>) => Promise.resolve(fn()).then(() => { n++; console.log("ok -", name); });

(async () => {
  await ok("classification examples", () => {
    const c = (t: string) => classify(t, DEFAULT_RULES);
    for (const t of ["Completed.", "COmpleted.", "Completed. FDL has been drafted.", "FDL was drafted.", "Mail was forwarded.", "Moved to Completed folder.", "Filer name was X.", "Email and downloaded documents were saved to the file."]) assert.equal(c(t), "Completed", t);
    for (const t of ["Did not proceed.", "Duplicate file. Moved to Duplicate folder.", "Not our case. Mail was moved to completed.", "Duplicate case.", "Completed. Did not proceed."]) assert.equal(c(t), "Did not proceed", t);
    assert.equal(c(""), "Needs Review"); assert.equal(c("   "), "Needs Review"); assert.equal(c("hello there"), "Needs Review");
  });
  await ok("7-min DNP and 15-min Completed schedule, backward from end, original order", () => {
    const r = buildSchedule([{ key: 0, minutes: 7 }, { key: 1, minutes: 15 }, { key: 2, minutes: 7 }], S);
    assert.deepEqual(r.slots.get(2), { start: toMin("13:23"), end: toMin("13:30") });
    assert.deepEqual(r.slots.get(1), { start: toMin("13:08"), end: toMin("13:23") });
    assert.deepEqual(r.slots.get(0), { start: toMin("13:01"), end: toMin("13:08") });
  });
  await ok("no record overlaps lunch; whole record moves before the break", () => {
    const items = Array.from({ length: 20 }, (_, i) => ({ key: i, minutes: i % 3 === 0 ? 15 : 7 }));
    const r = buildSchedule(items, S);
    const rows = items.map((i) => ({ key: i.key, minutes: i.minutes, slot: r.slots.get(i.key) }));
    assert.equal(checkSchedule(rows, S).size, 0);
    for (const s of r.slots.values()) assert.ok(s.end <= S.lunchStart || s.start >= S.lunchEnd, `${fmt12(s.start)}-${fmt12(s.end)}`);
    const lastBefore = [...r.slots.values()].find((s) => s.end === S.lunchStart);
    assert.ok(lastBefore, "a record ends exactly at 11:30");
    // order preserved: starts strictly ascending by key
    const starts = items.map((i) => r.slots.get(i.key)?.start).filter((x): x is number => x !== undefined);
    assert.deepEqual(starts, [...starts].sort((a, b) => a - b));
  });
  await ok("record that would straddle 12:30 end is pushed before 11:30 (never 11:25-12:32)", () => {
    const s2 = { ...S, endMin: toMin("12:32") };
    const r = buildSchedule([{ key: 0, minutes: 7 }], s2);
    assert.deepEqual(r.slots.get(0), { start: toMin("11:23"), end: toMin("11:30") });
  });
  await ok("schedule that exceeds workday reports unscheduled, drops nothing silently", () => {
    const items = Array.from({ length: 40 }, (_, i) => ({ key: i, minutes: 15 }));
    const r = buildSchedule(items, S);
    assert.equal(r.slots.size + r.unscheduled.length, 40);
    assert.ok(r.unscheduled.length > 0);
    assert.deepEqual(r.unscheduled, r.unscheduled.slice().sort((a, b) => a - b));
    assert.ok(r.unscheduled.every((k) => k < r.slots.size + r.unscheduled.length && !r.slots.has(k)));
    assert.ok([...r.slots.keys()].every((k) => k >= r.unscheduled.length)); // the EARLIEST rows are the ones left out
  });
  await ok("zero-minute (unprocessed) rows are left without times", () => {
    const r = buildSchedule([{ key: 0, minutes: 0 }, { key: 1, minutes: 7 }], S);
    assert.ok(!r.slots.has(0)); assert.deepEqual(r.unscheduled, [0]);
  });

  const file = process.argv[2] ?? "/mnt/user-data/uploads/MARK-_Daily_Internal_Inbound_E-Serve_Processing_10_15_2026-25771c54f1dcf4bfb55def8c.xls";
  if (fs.existsSync(file)) {
    await ok("real .xls -> Excel output has real time values and correct formulas", async () => {
      const buf = fs.readFileSync(file);
      const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
      const wb = loadWorkbook("x.xls", ab);
      const rows = wb.sheets[wb.sheetNames[0]];
      const lay = detectLayout(rows);
      assert.equal(lay.headers[lay.commentCol], "Comments");
      assert.equal(lay.startCol, 10); assert.equal(lay.endCol, 11); assert.equal(lay.minutesCol, 12);
      const recs: { key: number; minutes: number }[] = [];
      for (const i of dataRowIndices(rows, lay)) {
        const cat = classify(String(rows[i][lay.commentCol] ?? ""), DEFAULT_RULES);
        recs.push({ key: i, minutes: cat === "Completed" ? 15 : cat === "Did not proceed" ? 7 : 0 });
      }
      const wide = { ...S, endMin: toMin("17:00") };
      const sch = buildSchedule(recs, wide);
      assert.equal(sch.unscheduled.length, 0);
      const out = await writeWorkbook(wb, { sheetName: wb.sheetNames[0], layout: lay, groups: [{ rows: recs.map((r) => ({ key: r.key, slot: sch.slots.get(r.key) })), maxMinutes: 15 }] });
      const chk = new ExcelJS.Workbook(); await chk.xlsx.load(out as any);
      const ws = chk.worksheets[0];
      assert.equal(ws.getCell("K2").numFmt, "h:mm AM/PM");
      const kv: any = ws.getCell("K2").value; // ExcelJS returns time-formatted numbers as Date => a real time value, not text
      assert.ok(kv instanceof Date, "Time Started is a real Excel time");
      assert.equal(kv.getUTCHours() * 60 + kv.getUTCMinutes(), sch.slots.get(recs[0].key)!.start);
      const m: any = ws.getCell("M2").value;
      assert.equal(m.formula, "MINUTE(MOD(L2-K2,1))");
      assert.equal(ws.getCell("A2").value, "25-013734");
      assert.equal(ws.columnCount, 16);
      console.log("   sample row 2:", ws.getCell("K2").value, ws.getCell("L2").value, m.formula, "| last row ends", fmt12((sch.slots.get(recs[recs.length - 1].key) as any).end));
    });
    await ok("sample file does not fit 8:00-1:30 with lunch (warning scenario)", () => {
      const buf = fs.readFileSync(file);
      const wb = loadWorkbook("x.xls", buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
      const rows = wb.sheets[wb.sheetNames[0]]; const lay = detectLayout(rows);
      const recs = dataRowIndices(rows, lay).map((i) => ({ key: i, minutes: classify(String(rows[i][lay.commentCol]), DEFAULT_RULES) === "Completed" ? 15 : 7 }));
      assert.equal(recs.length, 39);
      const sch = buildSchedule(recs, S);
      console.log(`   scheduled ${sch.slots.size}, unscheduled ${sch.unscheduled.length}`);
      assert.ok(sch.unscheduled.length > 0);
    });
  }
  console.log(`\nAll ${n} checks passed.`);
})().catch((e) => { console.error(e); process.exit(1); });
