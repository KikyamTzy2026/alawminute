import assert from "node:assert/strict";
import fs from "node:fs";
import { buildSchedule } from "../lib/schedule";
import { toMin } from "../lib/time";
import { loadWorkbook, detectLayout, writeWorkbook, dataRowIndices, distinctValues } from "../lib/excel";

const file = process.argv[2] ?? "/mnt/user-data/uploads/Daily_Internal_Billing_Time_Cost_to_Post_10_28__October_29__2026__1-9e3ba68645bcbf30973a013d.xls";
const out = process.argv[3];
(async () => {
  if (!fs.existsSync(file)) { console.log("billing sample not present - skipped"); return; }
  const buf = fs.readFileSync(file);
  const wb = loadWorkbook("b.xls", buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
  const rows = wb.sheets[wb.sheetNames[0]];
  const lay = detectLayout(rows, { billing: true });
  assert.equal(lay.blocks.length, 2);
  assert.ok(lay.totalCol >= 0);
  assert.deepEqual(lay.blocks.map((b) => [b.startCol, b.endCol, b.minutesCol]), [[21, 22, 23], [25, 26, 27]]);
  assert.equal(lay.totalCol, 28);
  // my rows only: Processor Name == MARK
  const pn = lay.headers.findIndex((h) => /^processor name$/i.test(h));
  assert.equal(distinctValues(rows, lay, pn)[0].value, "MARK");
  const mine = dataRowIndices(rows, lay).filter((i) => String(rows[i][pn]).trim() === "MARK");
  assert.equal(mine.length, 38);
  // first 20 = Completed (billing 15 + posting 15), rest = Did not proceed (billing 7 only)
  const lunch = { lunchOn: true, lunchStart: toMin("11:30"), lunchEnd: toMin("12:30") };
  const it1 = mine.map((k, n) => ({ key: k, minutes: n < 20 ? 15 : 7 }));
  const it2 = mine.map((k, n) => ({ key: k, minutes: n < 20 ? 15 : 0 })).filter((x) => x.minutes > 0);
  const s1 = buildSchedule(it1, { startMin: toMin("08:00"), endMin: toMin("11:30"), ...lunch });
  const s2 = buildSchedule(it2, { startMin: toMin("12:30"), endMin: toMin("17:00"), ...lunch });
  console.log(`billing set: ${s1.slots.size} scheduled / ${s1.unscheduled.length} not; posting set: ${s2.slots.size} / ${s2.unscheduled.length}`);
  const bytes = await writeWorkbook(wb, { sheetName: wb.sheetNames[0], layout: lay, billing: true, groups: [
    { rows: mine.map((k) => ({ key: k, slot: s1.slots.get(k) })), maxMinutes: 15 },
    { rows: mine.map((k) => ({ key: k, slot: s2.slots.get(k) })), maxMinutes: 15 },
  ] });
  if (out) fs.writeFileSync(out, bytes);
  console.log("ok - billing file written", bytes.length, "bytes");
})().catch((e) => { console.error(e); process.exit(1); });
