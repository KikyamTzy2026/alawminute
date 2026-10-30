import * as XLSX from "xlsx";

/** Minimal BIFF8 (.xls) style reader: fonts, fills, borders, alignment, number formats per cell. */

export interface XlsCellStyle {
  font: { name: string; size: number; bold: boolean; italic: boolean; underline: boolean; strike: boolean; color?: string };
  fill?: string; // ARGB
  align: { h?: string; v?: string; wrap: boolean; indent: number };
  border: { l?: [string, string?]; r?: [string, string?]; t?: [string, string?]; b?: [string, string?] };
  numFmt: string;
}

const DEFAULT_PALETTE = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 " +
  "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF " +
  "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 " +
  "003366 339966 003300 333300 993300 993366 333399 333333").split(" ");

const BORDER = ["", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"];
const HALIGN = [undefined, "left", "center", "right", "fill", "justify", "centerContinuous"];
const VALIGN = ["top", "middle", "bottom", "justify", "distributed"];

export interface XlsStyleReader {
  /** style of a cell, or undefined if the sheet or cell has no explicit formatting record */
  styleOf(sheetName: string, row: number, col: number): XlsCellStyle | undefined;
}

export function readXlsStyles(data: ArrayBuffer): XlsStyleReader | null {
  let content: Uint8Array;
  try {
    const cfb = XLSX.CFB.read(new Uint8Array(data), { type: "array" });
    const entry = XLSX.CFB.find(cfb, "/Workbook") || XLSX.CFB.find(cfb, "/Book");
    if (!entry || !entry.content) return null;
    content = Uint8Array.from(entry.content as ArrayLike<number>);
  } catch { return null; }

  const dv = new DataView(content.buffer, content.byteOffset, content.byteLength);
  const u16 = (o: number) => dv.getUint16(o, true);
  const u32 = (o: number) => dv.getUint32(o, true);

  const palette = DEFAULT_PALETTE.slice();
  const fonts: any[] = [];
  const formats: Record<number, string> = {};
  const xfs: any[] = [];
  const sheetNames: string[] = [];   // worksheet BOUNDSHEET names in stream order
  const cellXf: Record<string, Map<number, number>> = {}; // sheet -> (row*256+col) -> xf
  const table = (XLSX.SSF as any).get_table?.() ?? {};

  let pos = 0, bofDepth = 0, sheetIdx = -1, inGlobals = true;
  while (pos + 4 <= content.length) {
    const rt = u16(pos), len = u16(pos + 2), d = pos + 4;
    if (d + len > content.length) break;
    if (rt === 0x0809) { // BOF
      const type = u16(d + 2);
      if (bofDepth === 0 && type === 0x0005) inGlobals = true;
      else if (type === 0x0010) { inGlobals = false; sheetIdx++; }
      bofDepth++;
    } else if (rt === 0x000a) { // EOF
      bofDepth--;
      if (bofDepth <= 0) bofDepth = 0;
    } else if (inGlobals) {
      if (rt === 0x0031) { // FONT
        const cch = content[d + 14];
        const flag = content[d + 15];
        let name = "";
        for (let i = 0; i < cch; i++) name += String.fromCharCode(flag & 1 ? u16(d + 16 + i * 2) : content[d + 16 + i]);
        fonts.push({ size: u16(d) / 20, italic: !!(u16(d + 2) & 2), strike: !!(u16(d + 2) & 8), icv: u16(d + 4), bold: u16(d + 6) >= 700, underline: content[d + 10] !== 0, name });
      } else if (rt === 0x041e) { // FORMAT
        const cch = u16(d + 2), flag = content[d + 4];
        let s = "";
        for (let i = 0; i < cch; i++) s += String.fromCharCode(flag & 1 ? u16(d + 5 + i * 2) : content[d + 5 + i]);
        formats[u16(d)] = s;
      } else if (rt === 0x0092) { // PALETTE
        const n = u16(d);
        for (let i = 0; i < n; i++) {
          const h = (v: number) => v.toString(16).padStart(2, "0");
          palette[i] = (h(content[d + 2 + i * 4]) + h(content[d + 3 + i * 4]) + h(content[d + 4 + i * 4])).toUpperCase();
        }
      } else if (rt === 0x00e0) { // XF
        const b1 = u32(d + 10), b2 = u32(d + 14), fc = u16(d + 18);
        xfs.push({
          font: u16(d), fmt: u16(d + 2),
          h: content[d + 6] & 7, wrap: !!(content[d + 6] & 8), v: (content[d + 6] >> 4) & 7, indent: content[d + 8] & 15,
          dl: b1 & 15, dr: (b1 >> 4) & 15, dt: (b1 >> 8) & 15, db: (b1 >> 12) & 15,
          cl: (b1 >> 16) & 127, cr: (b1 >> 23) & 127, ct: b2 & 127, cb: (b2 >> 7) & 127,
          pat: (b2 >>> 26) & 63, fg: fc & 127,
        });
      } else if (rt === 0x0085) { // BOUNDSHEET
        const dt = content[d + 5];
        const cch = content[d + 6], flag = content[d + 7];
        let name = "";
        for (let i = 0; i < cch; i++) name += String.fromCharCode(flag & 1 ? u16(d + 8 + i * 2) : content[d + 8 + i]);
        if (dt === 0) sheetNames.push(name);
      }
    } else if (sheetIdx >= 0 && sheetNames[sheetIdx] !== undefined) {
      const sn = sheetNames[sheetIdx];
      const m = (cellXf[sn] ||= new Map());
      const put = (r: number, c: number, xf: number) => m.set(r * 256 + c, xf);
      if ([0x00fd, 0x0203, 0x027e, 0x0201, 0x0204, 0x0205, 0x0006].includes(rt)) put(u16(d), u16(d + 2), u16(d + 4));
      else if (rt === 0x00be) { // MULBLANK
        const r = u16(d), c0 = u16(d + 2), c1 = u16(d + len - 2);
        for (let c = c0; c <= c1; c++) put(r, c, u16(d + 4 + (c - c0) * 2));
      } else if (rt === 0x00bd) { // MULRK
        const r = u16(d), c0 = u16(d + 2), c1 = u16(d + len - 2);
        for (let c = c0; c <= c1; c++) put(r, c, u16(d + 4 + (c - c0) * 6));
      }
    }
    pos = d + len;
  }
  if (!xfs.length) return null;

  const colorOf = (icv: number): string | undefined => {
    if (icv >= 8 && icv <= 63) return "FF" + palette[icv - 8];
    if (icv === 0) return "FF000000";
    return undefined; // 64/65/32767 = automatic
  };
  const cache = new Map<number, XlsCellStyle>();
  const build = (i: number): XlsCellStyle | undefined => {
    if (cache.has(i)) return cache.get(i);
    const x = xfs[i]; if (!x) return undefined;
    const f = fonts[x.font >= 4 ? x.font - 1 : x.font] ?? { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strike: false, icv: 32767 };
    const bd = (line: number, col: number): [string, string?] | undefined => (line ? [BORDER[line] || "thin", colorOf(col)] : undefined);
    const st: XlsCellStyle = {
      font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: f.underline, strike: f.strike, color: colorOf(f.icv) },
      fill: x.pat === 1 ? colorOf(x.fg) : undefined,
      align: { h: HALIGN[x.h] as string | undefined, v: VALIGN[x.v], wrap: x.wrap, indent: x.indent },
      border: { l: bd(x.dl, x.cl), r: bd(x.dr, x.cr), t: bd(x.dt, x.ct), b: bd(x.db, x.cb) },
      numFmt: formats[x.fmt] ?? table[x.fmt] ?? "General",
    };
    cache.set(i, st);
    return st;
  };

  return {
    styleOf(sheetName, row, col) {
      const xf = cellXf[sheetName]?.get(row * 256 + col);
      return xf === undefined ? undefined : build(xf);
    },
  };
}
