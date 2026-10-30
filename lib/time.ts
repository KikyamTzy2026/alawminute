export function toMin(hhmm: string): number {
  const m = /^(\d{1,2}):(\d{2})$/.exec((hhmm ?? "").trim());
  if (!m) return NaN;
  const h = Number(m[1]), mi = Number(m[2]);
  if (h > 23 || mi > 59) return NaN;
  return h * 60 + mi;
}

export function fmt12(min: number): string {
  const m = ((Math.round(min) % 1440) + 1440) % 1440;
  const h24 = Math.floor(m / 60), mi = m % 60;
  const ap = h24 >= 12 ? "PM" : "AM";
  const h = h24 % 12 === 0 ? 12 : h24 % 12;
  return `${h}:${String(mi).padStart(2, "0")} ${ap}`;
}

/** Excel stores a time of day as a fraction of 24h. */
export function toExcelTime(min: number): number {
  return min / 1440;
}
