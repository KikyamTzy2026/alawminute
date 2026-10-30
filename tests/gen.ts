import fs from "node:fs";
import { classify, DEFAULT_RULES } from "../lib/classify";
import { buildSchedule } from "../lib/schedule";
import { toMin } from "../lib/time";
import { loadWorkbook, detectLayout, writeWorkbook, dataRowIndices } from "../lib/excel";
const [inp, out, end] = process.argv.slice(2);
const buf = fs.readFileSync(inp);
const wb = loadWorkbook(inp, buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
const rows = wb.sheets[wb.sheetNames[0]]; const lay = detectLayout(rows);
const recs = dataRowIndices(rows, lay).map((i) => ({ key: i, minutes: classify(String(rows[i][lay.commentCol]), DEFAULT_RULES) === "Completed" ? 15 : 7 }));
const sch = buildSchedule(recs, { startMin: toMin("08:00"), endMin: toMin(end ?? "13:30"), lunchOn: true, lunchStart: toMin("11:30"), lunchEnd: toMin("12:30") });
writeWorkbook(wb, { sheetName: wb.sheetNames[0], layout: lay, groups: [{ rows: recs.map((r) => ({ key: r.key, slot: sch.slots.get(r.key) })), maxMinutes: 15 }] })
  .then((b) => { fs.writeFileSync(out, b); console.log("rows", recs.length, "scheduled", sch.slots.size); });
