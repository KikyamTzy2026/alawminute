import * as XLSX from "xlsx";
import { Slot } from "./types";
import { toExcelTime } from "./time";
import { readXlsStyles, XlsCellStyle } from "./xlsStyles";

export interface LoadedWorkbook {
  fileName: string;
  /** The original file bytes, exactly as uploaded (never modified). */
  original: ArrayBuffer;
  isXlsx: boolean;
  sheetNames: string[];
  sheets: Record<string, unknown[][]>;
}

export interface BlockCols { startCol: number; endCol: number; minutesCol: number }

export interface Layout {
  /** every Time Started / Time Ended / Minutes triplet found in the header row, left to right */
  blocks: BlockCols[];
  headerRowIdx: number; // 0-based index into the sheet rows
  headers: string[];
  commentCol: number;   // 0-based, -1 if unknown
  totalCol: number;     // "TOTAL MINUTES" column or -1 (only used to recognise the Billing format)
  startCol: number;     // existing "Time Started" column or -1
  endCol: number;
  minutesCol: number;
}

const cellText = (v: unknown) => (v === null || v === undefined ? "" : String(v));

export function loadWorkbook(fileName: string, data: ArrayBuffer): LoadedWorkbook {
  const wb = XLSX.read(data, { type: "array", cellDates: false });
  const isXlsx = /\.xlsx$/i.test(fileName);
  const sheets: Record<string, unknown[][]> = {};
  for (const name of wb.SheetNames) {
    sheets[name] = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name], {
      header: 1, defval: "", blankrows: true, raw: true,
    });
  }
  return { fileName, original: data, isXlsx, sheetNames: wb.SheetNames, sheets };
}

/**
 * `billing` only changes how the comment column is guessed (Notes / Resolution style columns, never a free-text
 * guess). Without it the detection is exactly what every other project has always used.
 */
export function detectLayout(rows: unknown[][], opts: { billing?: boolean } = {}): Layout {
  const lim = Math.min(rows.length, 15);
  let headerRowIdx = 0;
  let best = -1;
  for (let i = 0; i < lim; i++) {
    const filled = (rows[i] ?? []).filter((c) => cellText(c).trim()).length;
    const hasKey = (rows[i] ?? []).some((c) => /comment/i.test(cellText(c)));
    const score = (hasKey ? 100 : 0) + filled;
    if (score > best) { best = score; headerRowIdx = i; }
  }
  const headers = (rows[headerRowIdx] ?? []).map((c) => cellText(c).replace(/\s+/g, " ").trim());
  const find = (re: RegExp) => headers.findIndex((h) => re.test(h));
  let commentCol = find(/comment/i);
  if (commentCol < 0 && opts.billing) {
    commentCol = find(/^notes?$/i);
    if (commentCol < 0) { const idx = headers.map((h, i) => (/resolution/i.test(h) ? i : -1)).filter((i) => i >= 0); commentCol = idx.length ? idx[idx.length - 1] : -1; }
  }
  if (commentCol < 0 && !opts.billing) {
    // fall back to the column with the longest average text
    let bestAvg = 0;
    headers.forEach((_, c) => {
      const vals = rows.slice(headerRowIdx + 1, headerRowIdx + 60).map((r) => cellText(r?.[c]));
      const textVals = vals.filter((v) => v && Number.isNaN(Number(v)));
      const avg = textVals.reduce((a, v) => a + v.length, 0) / Math.max(vals.length, 1);
      if (avg > bestAvg && avg > 12) { bestAvg = avg; commentCol = c; }
    });
  }
  const blocks: BlockCols[] = [];
  for (let i = 0; i + 2 < headers.length; i++) {
    if (/^time\s*started$/i.test(headers[i]) && /^time\s*ended$/i.test(headers[i + 1]) && /^minutes$/i.test(headers[i + 2])) {
      blocks.push({ startCol: i, endCol: i + 1, minutesCol: i + 2 });
      i += 2;
    }
  }
  return {
    blocks, totalCol: find(/^total\s*minutes$/i), headerRowIdx, headers, commentCol,
    startCol: find(/^time\s*started$/i),
    endCol: find(/^time\s*ended$/i),
    minutesCol: find(/^minutes$/i),
  };
}

export interface OutRow { key: number; slot?: Slot }

export interface WriteGroup {
  rows: OutRow[];       // every data row (slot undefined => cells of this block are left empty)
  maxMinutes: number;   // longest scheduled duration in this block, decides which Minutes formula is safe
}

export interface WriteOptions {
  sheetName: string;
  layout: Layout;
  /** one group per Time Started/Ended/Minutes block, in left-to-right order */
  groups: WriteGroup[];
  /**
   * Billing project only. true: rows without a time in a set only have their two time cells emptied (their Minutes
   * formula stays), the sheet's own Minutes formulas are kept, and TOTAL MINUTES is refreshed.
   * false/undefined (every other project): the original behaviour.
   */
  billing?: boolean;
}

const colLetter = (n: number) => { // 1-based
  let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

/** Rebuild a legacy .xls as an ExcelJS workbook keeping values, number formats, fonts, fills, borders, widths, heights, merges. */
function buildFromXls(ExcelJS: any, original: ArrayBuffer) {
  const sj = XLSX.read(original, { type: "array", cellStyles: true, cellNF: true, cellDates: false });
  const styles = readXlsStyles(original);
  const wb = new ExcelJS.Workbook();
  for (const name of sj.SheetNames) {
    const src = sj.Sheets[name];
    const ws = wb.addWorksheet(name);
    let maxR = -1, maxC = -1;
    for (const addr of Object.keys(src)) {
      if (addr[0] === "!") continue;
      const c = src[addr] as XLSX.CellObject;
      if (c.t === "z" || c.v === undefined || c.v === "") continue;
      const { r, c: cc } = XLSX.utils.decode_cell(addr);
      if (r > maxR) maxR = r;
      if (cc > maxC) maxC = cc;
    }
    const cols = (src["!cols"] ?? []) as { width?: number; hidden?: boolean }[];
    for (let c = 0; c <= maxC; c++) if (cols[c]?.width) { const col = ws.getColumn(c + 1); col.width = cols[c].width; if (cols[c].hidden) col.hidden = true; }
    const rowsInfo = (src["!rows"] ?? []) as { hpt?: number; hidden?: boolean }[];
    for (let r = 0; r <= maxR; r++) if (rowsInfo[r]?.hpt) ws.getRow(r + 1).height = rowsInfo[r].hpt;
    for (let r = 0; r <= maxR; r++) {
      for (let c = 0; c <= maxC; c++) {
        const cell = src[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
        const st: XlsCellStyle | undefined = styles?.styleOf(name, r, c);
        const target = ws.getCell(r + 1, c + 1);
        if (cell && cell.f && typeof cell.f === "string" && cell.f.trim()) {
          // keep the original formulas (e.g. =MINUTE(W2-V2), =SUM(X2,AB2)) with their cached result
          target.value = { formula: cell.f.replace(/^=/, ""), result: cell.v as any };
        } else if (cell && (cell.t === "n" || cell.t === "s" || cell.t === "b") && cell.v !== undefined && cell.v !== "") target.value = cell.v as any;
        if (st) {
          target.font = { name: st.font.name, size: st.font.size, bold: st.font.bold, italic: st.font.italic, strike: st.font.strike, underline: st.font.underline || undefined, ...(st.font.color ? { color: { argb: st.font.color } } : {}) };
          if (st.fill) target.fill = { type: "pattern", pattern: "solid", fgColor: { argb: st.fill } };
          target.alignment = { horizontal: st.align.h as any, vertical: st.align.v as any, wrapText: st.align.wrap || undefined, indent: st.align.indent || undefined };
          const b = (x?: [string, string?]) => (x ? { style: x[0], ...(x[1] ? { color: { argb: x[1] } } : {}) } : undefined);
          if (st.border.l || st.border.r || st.border.t || st.border.b) target.border = { left: b(st.border.l), right: b(st.border.r), top: b(st.border.t), bottom: b(st.border.b) } as any;
          if (st.numFmt && st.numFmt !== "General") target.numFmt = st.numFmt;
        } else if (cell?.z && cell.z !== "General") target.numFmt = String(cell.z);
      }
    }
    for (const m of (src["!merges"] ?? []) as XLSX.Range[]) ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1);
    // keep the header filter buttons, limited to the real data area so the range is always valid
    const af = (src as any)["!autofilter"]?.ref as string | undefined;
    if (af && maxR >= 0 && maxC >= 0) {
      const r = XLSX.utils.decode_range(af);
      ws.autoFilter = { from: { row: r.s.r + 1, column: r.s.c + 1 }, to: { row: Math.max(r.s.r + 1, Math.min(r.e.r, maxR) + 1), column: Math.min(r.e.c, maxC) + 1 } };
    }
    // ExcelJS' default DPI values are not needed and are a known source of Excel repair prompts
    ws.pageSetup = { orientation: "portrait" } as any;
  }
  return wb;
}

/**
 * Fills ONLY the Time Started / Time Ended / Minutes cells. Every other cell, header and format of the
 * original workbook is kept as it was (a legacy .xls is rebuilt with its formatting, then saved as .xlsx).
 */
export async function writeWorkbook(src: LoadedWorkbook, o: WriteOptions): Promise<Uint8Array> {
  const mod: any = await import("exceljs");
  const ExcelJS = mod.default ?? mod;
  let wb: any;
  if (src.isXlsx) { wb = new ExcelJS.Workbook(); await wb.xlsx.load(src.original); }
  else wb = buildFromXls(ExcelJS, src.original);
  const ws = wb.getWorksheet(o.sheetName);
  if (!ws) throw new Error(`Worksheet "${o.sheetName}" not found.`);

  const hdrRow = o.layout.headerRowIdx + 1;
  let lastCol = Math.max(ws.columnCount, o.layout.headers.length);
  const headerCell = (c: number) => ws.getCell(hdrRow, c);
  const ensure = (idx: number | undefined, title: string): number => {
    if (idx !== undefined && idx >= 0) return idx + 1;
    lastCol += 1; // missing column: appended once, never duplicated
    const cell = headerCell(lastCol);
    cell.value = title;
    const prev = headerCell(lastCol - 1);
    if (prev?.style) cell.style = JSON.parse(JSON.stringify(prev.style));
    ws.getColumn(lastCol).width = 14;
    return lastCol;
  };

  const minCols: number[] = [];
  o.groups.forEach((g, gi) => {
    const l = o.layout;
    const known = l.blocks[gi] ?? (gi === 0 ? { startCol: l.startCol, endCol: l.endCol, minutesCol: l.minutesCol } : undefined);
    const cStart = ensure(known?.startCol, "Time Started");
    const cEnd = ensure(known?.endCol, "Time Ended");
    const cMin = ensure(known?.minutesCol, "Minutes");
    const L = { s: colLetter(cStart), e: colLetter(cEnd) };
    minCols.push(cMin);
    for (const r of g.rows) {
      const rowNo = r.key + 1;
      const sCell = ws.getCell(rowNo, cStart);
      const eCell = ws.getCell(rowNo, cEnd);
      const mCell = ws.getCell(rowNo, cMin);
      if (!r.slot) {
        sCell.value = null; eCell.value = null;
        if (!o.billing) mCell.value = null; // Billing keeps the sheet's own Minutes formula
        continue;
      }
      const dur = r.slot.end - r.slot.start;
      sCell.value = toExcelTime(r.slot.start);
      eCell.value = toExcelTime(r.slot.end);
      sCell.numFmt = "h:mm AM/PM";
      eCell.numFmt = "h:mm AM/PM";
      // MINUTE() only returns 0-59, so it is used when every duration is < 60 min;
      // MOD(...,1) keeps it correct across midnight. Longer durations fall back to total minutes.
      const formula = g.maxMinutes < 60
        ? `MINUTE(MOD(${L.e}${rowNo}-${L.s}${rowNo},1))`
        : `ROUND(MOD(${L.e}${rowNo}-${L.s}${rowNo},1)*1440,0)`;
      const existing = mCell.value as any;
      if (o.billing && existing && typeof existing === "object" && existing.formula) mCell.value = { formula: existing.formula, result: dur };
      else { mCell.value = { formula, result: dur }; mCell.numFmt = "0"; }
    }
  });
  // TOTAL MINUTES: keep the user's own formula (e.g. =SUM(X2,AB2)) and only refresh its cached result;
  // write =SUM(...) of the Minutes cells if the cell has no formula yet.
  if (o.billing && o.layout.totalCol >= 0 && minCols.length) {
    const totals = new Map<number, number>();
    for (const g of o.groups) for (const r of g.rows) if (r.slot) totals.set(r.key, (totals.get(r.key) ?? 0) + (r.slot.end - r.slot.start));
    totals.forEach((total, key) => {
      const rowNo = key + 1;
      const cell = ws.getCell(rowNo, o.layout.totalCol + 1);
      const ex: any = cell.value;
      const f = ex && typeof ex === "object" && typeof ex.formula === "string" ? ex.formula : `SUM(${minCols.map((c) => colLetter(c) + rowNo).join(",")})`;
      cell.value = { formula: f, result: total };
    });
  }
  wb.calcProperties = { ...(wb.calcProperties ?? {}), fullCalcOnLoad: true };
  const out = await wb.xlsx.writeBuffer();
  return new Uint8Array(out as ArrayBuffer);
}

/** 0-based sheet row indices below the header that contain at least one value (blank rows are ignored). */
export function dataRowIndices(rows: unknown[][], layout: Layout): number[] {
  const out: number[] = [];
  for (let i = layout.headerRowIdx + 1; i < rows.length; i++) {
    if ((rows[i] ?? []).some((c) => cellText(c).trim() !== "")) out.push(i);
  }
  return out;
}

/** Distinct non-empty values of a column below the header, most frequent first (max 60). */
export function distinctValues(rows: unknown[][], layout: Layout, col: number): { value: string; count: number }[] {
  const m = new Map<string, number>();
  for (const i of dataRowIndices(rows, layout)) {
    const v = cellText(rows[i]?.[col]).trim();
    if (v) m.set(v, (m.get(v) ?? 0) + 1);
  }
  return [...m.entries()].sort((a, b) => b[1] - a[1]).slice(0, 60).map(([value, count]) => ({ value, count }));
}
