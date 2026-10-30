"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Category, RecordRow, Rules, Settings, Slot } from "@/lib/types";
import { classify, DEFAULT_RULES } from "@/lib/classify";
import { buildSchedule, checkSchedule, crossOverlaps, ScheduleSettings, validateSettings } from "@/lib/schedule";
import { fmt12, toMin } from "@/lib/time";
import { dataRowIndices, detectLayout, distinctValues, Layout, loadWorkbook, LoadedWorkbook, writeWorkbook } from "@/lib/excel";

const DEFAULT_SETTINGS: Settings = {
  mode: "single",
  b1Name: "Billing", b1Start: "08:00", b1End: "11:30",
  b2Name: "Posting", b2Start: "12:30", b2End: "17:00",
  startTime: "08:00", endTime: "13:30",
  lunchOn: true, lunchStart: "11:30", lunchEnd: "12:30",
  minCompleted: 15, minDnp: 7, minReview: 0,
  minCompleted2: 15, minDnp2: 0, minReview2: 0,
  blankAs: "Needs Review",
};
const CATS: Category[] = ["Completed", "Did not proceed", "Needs Review", "Exclude"];
const badge: Record<Category, string> = {
  "Completed": "bg-emerald-100 text-emerald-800",
  "Did not proceed": "bg-amber-100 text-amber-800",
  "Needs Review": "bg-rose-100 text-rose-800",
  "Exclude": "bg-slate-200 text-slate-600",
};

const cellText = (v: unknown) => (v === null || v === undefined ? "" : String(v));
const lines = (s: string) => s.split("\n").map((x) => x.trim()).filter(Boolean);
type BSlot = Slot & { block: number };

export default function Page() {
  const [wb, setWb] = useState<LoadedWorkbook | null>(null);
  const [sheet, setSheet] = useState("");
  const [commentCol, setCommentCol] = useState(-1);
  const [filterCol, setFilterCol] = useState(-1);
  const [filterVal, setFilterVal] = useState("");
  const [settings, setSettings] = useState<Settings>(DEFAULT_SETTINGS);
  const [dnpText, setDnpText] = useState(DEFAULT_RULES.dnp.join("\n"));
  const [doneText, setDoneText] = useState(DEFAULT_RULES.completed.join("\n"));
  const [overrides, setOverrides] = useState<Record<number, Category>>({});
  const [signature, setSignature] = useState<string | null>(null);
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [drag, setDrag] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  // Settings and keyword rules are remembered in this browser only.
  useEffect(() => {
    try {
      const s2 = localStorage.getItem("ttg:settings:v2");
      const s1 = localStorage.getItem("ttg:settings");
      if (s2) setSettings({ ...DEFAULT_SETTINGS, ...JSON.parse(s2) });
      else if (s1) setSettings({ ...DEFAULT_SETTINGS, ...JSON.parse(s1), mode: "single" }); // older versions switched projects automatically
      const r = localStorage.getItem("ttg:rules");
      if (r) { const p = JSON.parse(r); setDnpText(p.dnp); setDoneText(p.completed); }
    } catch { /* storage may be unavailable */ }
  }, []);
  useEffect(() => {
    try {
      localStorage.setItem("ttg:settings:v2", JSON.stringify(settings));
      localStorage.setItem("ttg:rules", JSON.stringify({ dnp: dnpText, completed: doneText }));
    } catch { /* ignore */ }
  }, [settings, dnpText, doneText]);

  const rows = wb && sheet ? wb.sheets[sheet] : [];
  const two = settings.mode === "two"; // "two" = the Billing project; everything Billing-only is gated by this flag
  const layout: Layout | null = useMemo(() => (rows.length ? detectLayout(rows, { billing: two }) : null), [rows, two]);
  const rules: Rules = useMemo(() => ({ dnp: lines(dnpText), completed: lines(doneText) }), [dnpText, doneText]);

  const allDataRows = useMemo(() => (layout ? dataRowIndices(rows, layout) : []), [rows, layout]);
  const filterValues = useMemo(() => (layout && filterCol >= 0 ? distinctValues(rows, layout, filterCol) : []), [rows, layout, filterCol]);
  const inScope = useMemo(() => {
    const fv = filterVal.trim().toLowerCase();
    if (!two || filterCol < 0 || !fv) return allDataRows; // the row filter belongs to Billing only
    return allDataRows.filter((i) => cellText(rows[i]?.[filterCol]).trim().toLowerCase() === fv);
  }, [allDataRows, rows, filterCol, filterVal, two]);

  const records: RecordRow[] = useMemo(() => {
    if (!layout || (!two && commentCol < 0)) return [];
    return inScope.map((i) => {
      const comment = commentCol >= 0 ? cellText(rows[i]?.[commentCol]) : "";
      // Standard: an empty comment is always "Needs Review". Billing may choose how empty comments count.
      const detected: Category = two && !comment.trim() ? settings.blankAs : classify(comment, rules);
      return { key: i, rowNumber: i + 1, comment, detected, category: overrides[i] ?? detected };
    });
  }, [rows, layout, inScope, commentCol, rules, overrides, settings.blankAs, two]);

  const nBlocks = two ? 2 : 1;
  const minutesFor = useCallback((c: Category, b: number) => {
    if (c === "Exclude") return 0;
    const m = b === 0
      ? (c === "Completed" ? settings.minCompleted : c === "Did not proceed" ? settings.minDnp : settings.minReview)
      : (c === "Completed" ? settings.minCompleted2 : c === "Did not proceed" ? settings.minDnp2 : settings.minReview2);
    return Number.isFinite(m) ? m : 0;
  }, [settings]);

  const lunchOn = settings.lunchOn, lunchStart = toMin(settings.lunchStart), lunchEnd = toMin(settings.lunchEnd);
  const scheds: ScheduleSettings[] = useMemo(() => (two
    ? [{ startMin: toMin(settings.b1Start), endMin: toMin(settings.b1End), lunchOn, lunchStart, lunchEnd },
       { startMin: toMin(settings.b2Start), endMin: toMin(settings.b2End), lunchOn, lunchStart, lunchEnd }]
    : [{ startMin: toMin(settings.startTime), endMin: toMin(settings.endTime), lunchOn, lunchStart, lunchEnd }]),
  [two, settings.b1Start, settings.b1End, settings.b2Start, settings.b2End, settings.startTime, settings.endTime, lunchOn, lunchStart, lunchEnd]);
  const blockNames = [settings.b1Name || "Set 1", settings.b2Name || "Set 2"];

  const settingErrors = useMemo(() => {
    const e: string[] = [];
    scheds.forEach((sc, i) => validateSettings(sc).forEach((m) => e.push(two ? `${blockNames[i]}: ${m}` : m)));
    const mins: [string, number][] = [["Did not proceed", settings.minDnp], ["Completed", settings.minCompleted], ["Needs Review", settings.minReview]];
    if (two) mins.push(["Did not proceed (2nd set)", settings.minDnp2], ["Completed (2nd set)", settings.minCompleted2], ["Needs Review (2nd set)", settings.minReview2]);
    for (const [k, v] of mins) if (!Number.isFinite(v) || v < 0 || !Number.isInteger(v)) e.push(`Minutes for "${k}" must be a whole number of 0 or more.`);
    return e;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scheds, settings, two]);

  const currentSig = JSON.stringify({ sheet, commentCol, filterCol, filterVal, settings, cats: records.map((r) => [r.key, r.category]) });
  const generated = signature !== null && signature === currentSig && settingErrors.length === 0;
  const stale = signature !== null && signature !== currentSig;

  // One schedule per set of time columns. A row can appear in both sets (e.g. Billing AND Posting).
  const result = useMemo(() => {
    if (!generated) return null;
    const slots: Map<number, Slot>[] = [];
    const unscheduled: Set<number>[] = [];
    const checks = new Map<number, string[]>();
    let idleMinutes = 0;
    scheds.forEach((sc, b) => {
      const items = records.filter((r) => r.category !== "Exclude").map((r) => ({ key: r.key, minutes: minutesFor(r.category, b) })).filter((i) => i.minutes > 0);
      const r = buildSchedule(items, sc);
      slots.push(r.slots); unscheduled.push(new Set(r.unscheduled)); idleMinutes += r.idleMinutes;
      checkSchedule(items.map((i) => ({ key: i.key, minutes: i.minutes, slot: r.slots.get(i.key) })), sc)
        .forEach((m, k) => checks.set(k, [...(checks.get(k) ?? []), ...(two ? m.map((x) => `${blockNames[b]}: ${x}`) : m)]));
    });
    if (scheds.length > 1) {
      const all = slots.flatMap((m, b) => [...m.entries()].map(([key, slot]) => ({ key: key * 2 + b, slot })));
      crossOverlaps(all).forEach((ck) => { const k = Math.floor(ck / 2); checks.set(k, [...(checks.get(k) ?? []), "Overlaps the other set"]); });
    }
    return { slots, unscheduled, idleMinutes, checks };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [generated, records, minutesFor, scheds, two]);

  function suggestFilter(w: LoadedWorkbook, name: string) {
    const rws = w.sheets[name]; const lay = detectLayout(rws);
    const idx = lay.headers.findIndex((h) => /^processor\s*name$/i.test(h));
    if (idx >= 0) {
      const top = distinctValues(rws, lay, idx)[0];
      if (top) { setFilterCol(idx); setFilterVal(top.value); return; }
    }
    setFilterCol(-1); setFilterVal("");
  }

  async function onFile(f: File | undefined | null) {
    if (!f) return;
    setError(""); setBusy(true); setNotice("");
    try {
      if (!/\.(xlsx|xls)$/i.test(f.name)) throw new Error("Please upload an .xlsx or .xls file.");
      const loaded = loadWorkbook(f.name, await f.arrayBuffer());
      if (!loaded.sheetNames.length) throw new Error("This workbook has no worksheets.");
      setWb(loaded); setSignature(null);
      pickSheet(loaded, loaded.sheetNames[0], true);
    } catch (e) { setError(e instanceof Error ? e.message : "Could not read that file."); }
    finally { setBusy(false); }
  }
  function pickSheet(w: LoadedWorkbook, name: string, announce = false, billing = two) {
    setSheet(name); setOverrides({}); setSignature(null);
    const lay = detectLayout(w.sheets[name], { billing });
    setCommentCol(lay.commentCol);
    if (billing) suggestFilter(w, name); else { setFilterCol(-1); setFilterVal(""); }
    // Never switches the project by itself; only points out that this sheet looks like the Billing format.
    if (announce && !billing && lay.blocks.length >= 2 && lay.totalCol >= 0)
      setNotice("This sheet has two sets of Time Started / Time Ended / Minutes and a TOTAL MINUTES column, which is the Billing format. Choose Project → Billing to use the Billing options.");
  }

  function changeProject(next: "single" | "two") {
    setSettings((s) => ({ ...s, mode: next })); setOverrides({}); setSignature(null); setNotice("");
    if (wb && sheet) pickSheet(wb, sheet, false, next === "two");
  }

  async function download() {
    if (!wb || !layout || !result) return;
    setBusy(true); setError("");
    try {
      // Standard: every listed row is written (excluded rows are emptied), as always.
      // Billing: excluded and filtered-out rows are never modified.
      const touched = two ? records.filter((r) => r.category !== "Exclude") : records;
      const groups = scheds.map((_, b) => ({
        rows: touched.map((r) => ({ key: r.key, slot: result.slots[b].get(r.key) })),
        maxMinutes: Math.max(0, ...[...result.slots[b].values()].map((v) => v.end - v.start)),
      }));
      const bytes = await writeWorkbook(wb, { sheetName: sheet, layout, groups, billing: two });
      const blob = new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = wb.fileName.replace(/\.(xlsx|xls)$/i, "") + "_scheduled.xlsx";
      a.click(); URL.revokeObjectURL(a.href);
    } catch (e) { setError(e instanceof Error ? e.message : "Could not create the Excel file."); }
    finally { setBusy(false); }
  }

  // ---- summary ----
  const count = (c: Category) => records.filter((r) => r.category === c).length;
  const active = records.filter((r) => r.category !== "Exclude");
  const neededByBlock = scheds.map((_, b) => active.reduce((a, r) => a + minutesFor(r.category, b), 0));
  const totalMinutes = neededByBlock.reduce((a, x) => a + x, 0);
  const capacity = (sc: ScheduleSettings) => (Number.isNaN(sc.startMin) || Number.isNaN(sc.endMin) || sc.endMin <= sc.startMin ? 0 :
    (sc.endMin - sc.startMin) - (sc.lunchOn && !Number.isNaN(sc.lunchStart) && !Number.isNaN(sc.lunchEnd) ? Math.max(0, Math.min(sc.endMin, sc.lunchEnd) - Math.max(sc.startMin, sc.lunchStart)) : 0));
  const availByBlock = scheds.map(capacity);
  const scheduledCount = result ? result.slots.reduce((a, m) => a + m.size, 0) : 0;
  const unscheduledCount = result ? result.unscheduled.reduce((a, s) => a + s.size, 0) : 0;
  const scheduledMinutes = result ? result.slots.reduce((a, m) => a + [...m.values()].reduce((x, v) => x + v.end - v.start, 0), 0) : 0;
  const problems = result ? [...result.checks.values()].flat().length : 0;

  const set = <K extends keyof Settings>(k: K, v: Settings[K]) => setSettings((s) => ({ ...s, [k]: v }));
  const num = (v: string) => (v.trim() === "" ? NaN : Number(v));
  const nv = (n: number) => (Number.isNaN(n) ? "" : n);
  const slotText = (b: number, key: number) => { const sl = result?.slots[b].get(key); return sl ? `${fmt12(sl.start)} – ${fmt12(sl.end)} (${sl.end - sl.start})` : "—"; };

  return (
    <main className="mx-auto max-w-7xl space-y-6 p-4 md:p-8">
      <header>
        <h1 className="text-2xl font-bold">Excel Time Tracker Generator</h1>
        <p className="text-sm text-slate-600">Upload your work file, review the detected categories, and generate Time Started, Time Ended and Minutes. Everything runs in your browser — your file is never uploaded to a server.</p>
      </header>

      <section
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => { e.preventDefault(); setDrag(false); onFile(e.dataTransfer.files?.[0]); }}
        onClick={() => fileRef.current?.click()}
        className={`cursor-pointer rounded-xl border-2 border-dashed p-8 text-center transition ${drag ? "border-blue-500 bg-blue-50" : "border-slate-300 bg-white hover:border-blue-400"}`}
      >
        <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { onFile(e.target.files?.[0]); e.target.value = ""; }} />
        <p className="font-medium">{wb ? `Loaded: ${wb.fileName}` : "Drag & drop an Excel file here, or click to browse"}</p>
        <p className="text-xs text-slate-500">.xlsx or .xls — your original file is never modified; a new file is downloaded.</p>
      </section>
      {error && <div role="alert" className="rounded-lg border border-rose-300 bg-rose-50 p-3 text-sm text-rose-800">{error}</div>}
      {notice && <div className="rounded-lg border border-sky-200 bg-sky-50 p-3 text-sm text-sky-900">{notice}</div>}

      {wb && layout && (
        <>
          <section className="grid gap-4 rounded-xl bg-white p-4 shadow-sm md:grid-cols-2">
            <label className="text-sm font-medium md:col-span-2">Project
              <select className="mt-1 block w-full rounded border p-2 md:w-1/2" value={settings.mode} onChange={(e) => changeProject(e.target.value as "single" | "two")}>
                <option value="single">Standard — one set of Time Started / Time Ended / Minutes</option>
                <option value="two">Billing — two sets + TOTAL MINUTES</option>
              </select>
            </label>
            <label className="text-sm font-medium">Worksheet
              <select className="mt-1 block w-full rounded border p-2" value={sheet} onChange={(e) => pickSheet(wb, e.target.value)}>
                {wb.sheetNames.map((n) => <option key={n}>{n}</option>)}
              </select>
            </label>
            <label className="text-sm font-medium">{two ? "Comment / resolution column" : "Comment column"} {commentCol < 0 && (two
              ? <span className="font-normal text-amber-700">(none — empty comments are treated as “{settings.blankAs}”)</span>
              : <span className="text-rose-600">(not detected — please choose)</span>)}
              <select className="mt-1 block w-full rounded border p-2" value={commentCol} onChange={(e) => { setCommentCol(Number(e.target.value)); setOverrides({}); }}>
                <option value={-1}>{two ? "— none —" : "— select —"}</option>
                {layout.headers.map((h, i) => h ? <option key={i} value={i}>{h}</option> : null)}
              </select>
            </label>
            {two && (
            <div className="md:col-span-2">
              <div className="grid gap-3 md:grid-cols-2">
                <label className="text-sm font-medium">Only process rows where column…
                  <select className="mt-1 block w-full rounded border p-2" value={filterCol} onChange={(e) => { setFilterCol(Number(e.target.value)); setFilterVal(""); setOverrides({}); }}>
                    <option value={-1}>— all rows —</option>
                    {layout.headers.map((h, i) => h ? <option key={i} value={i}>{h}</option> : null)}
                  </select>
                </label>
                <label className="text-sm font-medium">…equals
                  <input list="filter-values" disabled={filterCol < 0} className="mt-1 block w-full rounded border p-2 disabled:bg-slate-100" value={filterVal} onChange={(e) => { setFilterVal(e.target.value); setOverrides({}); }} placeholder="e.g. MARK" />
                  <datalist id="filter-values">{filterValues.map((v) => <option key={v.value} value={v.value}>{v.count} rows</option>)}</datalist>
                </label>
              </div>
              <p className="mt-1 text-xs text-slate-500">{inScope.length} of {allDataRows.length} rows will be processed. All other rows are left exactly as they are in your file.</p>
            </div>
            )}
            {two ? (
              <p className="text-xs text-slate-500 md:col-span-2">
                Time columns found: {layout.blocks.length === 0 ? "none yet — they will be added at the end of the sheet" : layout.blocks.map((b, i) => `Set ${i + 1} (${layout.headers[b.startCol]} / ${layout.headers[b.endCol]} / ${layout.headers[b.minutesCol]})`).join(" · ")}
                {layout.totalCol >= 0 ? ` · ${layout.headers[layout.totalCol]} (kept as your formula, result refreshed)` : ""}.
                {layout.blocks.length < 2 ? " Billing uses two sets, so the missing set will be added at the end." : ""}
              </p>
            ) : (
              <p className="text-xs text-slate-500 md:col-span-2">
                Output columns: {layout.startCol >= 0 ? "Time Started (existing, will be updated)" : "Time Started (will be added)"} · {layout.endCol >= 0 ? "Time Ended (existing)" : "Time Ended (will be added)"} · {layout.minutesCol >= 0 ? "Minutes (existing)" : "Minutes (will be added)"}
              </p>
            )}
          </section>

          <section className="grid gap-6 rounded-xl bg-white p-4 shadow-sm lg:grid-cols-3">
            <div className="space-y-3">
              <h2 className="font-semibold">Work minutes per record</h2>
              <p className="text-xs text-slate-500">{two ? "0 means that set stays empty for that category." : "0 means unprocessed."}</p>
              <table className="w-full text-sm">
                <thead><tr className="text-left text-xs text-slate-500"><th></th><th>{two ? blockNames[0] : "Minutes"}</th>{two && <th>{blockNames[1]}</th>}</tr></thead>
                <tbody>
                  {([["Did not proceed", "minDnp", "minDnp2"], ["Completed", "minCompleted", "minCompleted2"], ["Needs Review", "minReview", "minReview2"]] as const).map(([label, k1, k2]) => (
                    <tr key={k1}>
                      <td className="py-1 pr-2 font-medium">{label}</td>
                      <td className="pr-2"><input aria-label={`${label} minutes`} type="number" min={0} step={1} className="w-20 rounded border p-1" value={nv(settings[k1])} onChange={(e) => set(k1, num(e.target.value))} /></td>
                      {two && <td><input aria-label={`${label} minutes, ${blockNames[1]}`} type="number" min={0} step={1} className="w-20 rounded border p-1" value={nv(settings[k2])} onChange={(e) => set(k2, num(e.target.value))} /></td>}
                    </tr>
                  ))}
                </tbody>
              </table>
              {two && (              <Field label="A row with an empty comment counts as">
                <select className="w-full rounded border p-2" value={settings.blankAs} onChange={(e) => set("blankAs", e.target.value as Category)}>
                  {(["Needs Review", "Completed", "Did not proceed"] as Category[]).map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>)}
            </div>
            <div className="space-y-3">
              <h2 className="font-semibold">Workday</h2>
              {!two ? (
                <>
                  <Field label="Workday start time"><input type="time" className="w-full rounded border p-2" value={settings.startTime} onChange={(e) => set("startTime", e.target.value)} /></Field>
                  <Field label="Work end time (last record ends here)"><input type="time" className="w-full rounded border p-2" value={settings.endTime} onChange={(e) => set("endTime", e.target.value)} /></Field>
                </>
              ) : (
                <>
                  {([["b1Name", "b1Start", "b1End", "1st set (left)"], ["b2Name", "b2Start", "b2End", "2nd set (right)"]] as const).map(([nk, sk, ek, title]) => (
                    <div key={nk} className="rounded border p-2">
                      <Field label={`${title} name`}><input className="w-full rounded border p-1" value={settings[nk]} onChange={(e) => set(nk, e.target.value)} /></Field>
                      <div className="mt-2 grid grid-cols-2 gap-2">
                        <Field label="Start"><input type="time" className="w-full rounded border p-1" value={settings[sk]} onChange={(e) => set(sk, e.target.value)} /></Field>
                        <Field label="End (last record ends here)"><input type="time" className="w-full rounded border p-1" value={settings[ek]} onChange={(e) => set(ek, e.target.value)} /></Field>
                      </div>
                    </div>
                  ))}
                </>
              )}
              <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={settings.lunchOn} onChange={(e) => set("lunchOn", e.target.checked)} /> Exclude lunch break</label>
              <div className="grid grid-cols-2 gap-2">
                <Field label="Lunch start"><input type="time" disabled={!settings.lunchOn} className="w-full rounded border p-2 disabled:bg-slate-100" value={settings.lunchStart} onChange={(e) => set("lunchStart", e.target.value)} /></Field>
                <Field label="Lunch end"><input type="time" disabled={!settings.lunchOn} className="w-full rounded border p-2 disabled:bg-slate-100" value={settings.lunchEnd} onChange={(e) => set("lunchEnd", e.target.value)} /></Field>
              </div>
            </div>
            <div className="space-y-3">
              <h2 className="font-semibold">Comment detection (one keyword per line)</h2>
              <Field label="Did not proceed keywords (checked first)"><textarea rows={4} className="w-full rounded border p-2 font-mono text-xs" value={dnpText} onChange={(e) => setDnpText(e.target.value)} /></Field>
              <Field label="Completed keywords"><textarea rows={4} className="w-full rounded border p-2 font-mono text-xs" value={doneText} onChange={(e) => setDoneText(e.target.value)} /></Field>
              <button className="text-xs text-blue-700 underline" onClick={() => { setDnpText(DEFAULT_RULES.dnp.join("\n")); setDoneText(DEFAULT_RULES.completed.join("\n")); }}>Reset keywords</button>
            </div>
            {settingErrors.length > 0 && <ul className="lg:col-span-3 list-disc rounded-lg bg-rose-50 p-3 pl-6 text-sm text-rose-800">{settingErrors.map((e) => <li key={e}>{e}</li>)}</ul>}
          </section>

          <section className="grid grid-cols-2 gap-3 md:grid-cols-4 lg:grid-cols-8">
            <Stat label="Total records" value={records.length} />
            <Stat label="Completed" value={count("Completed")} />
            <Stat label="Did not proceed" value={count("Did not proceed")} />
            <Stat label="Needs review" value={count("Needs Review")} tone={count("Needs Review") ? "warn" : undefined} />
            <Stat label="Total work minutes" value={totalMinutes} />
            <Stat label="Scheduled work time" value={result ? `${scheduledMinutes} min` : "—"} />
            <Stat label={two ? "Scheduled entries" : "Scheduled records"} value={result ? scheduledCount : "—"} />
            <Stat label="Could not schedule" value={result ? unscheduledCount : "—"} tone={result && unscheduledCount ? "bad" : undefined} />
          </section>

          <section className="space-y-3">
            <div className="flex flex-wrap items-center gap-3">
              <button disabled={settingErrors.length > 0 || (!two && commentCol < 0) || !records.length}
                onClick={() => setSignature(currentSig)}
                className="rounded-lg bg-blue-600 px-5 py-2 font-medium text-white shadow hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-40">Generate Schedule</button>
              <button disabled={!result || busy || problems > 0}
                onClick={download}
                className="rounded-lg bg-emerald-600 px-5 py-2 font-medium text-white shadow hover:bg-emerald-700 disabled:cursor-not-allowed disabled:opacity-40">{busy ? "Working…" : "Download Excel"}</button>
              {stale && <span className="text-sm text-amber-700">Settings or categories changed — click Generate Schedule again.</span>}
              <span className="text-sm text-slate-500">Needed / available: {scheds.map((_, b) => `${two ? blockNames[b] + " " : ""}${neededByBlock[b]}/${availByBlock[b]} min`).join(" · ")}</span>
            </div>
            {result && unscheduledCount > 0 && (
              <div role="alert" className="rounded-lg border border-amber-300 bg-amber-50 p-3 text-sm text-amber-900">
                <b>{two ? `${unscheduledCount} entr${unscheduledCount === 1 ? "y" : "ies"}` : `${unscheduledCount} record(s)`} could not be scheduled.</b>{" "}
                {scheds.map((sc, b) => neededByBlock[b] > availByBlock[b] ? (
                  <span key={b}>{two ? `${blockNames[b]}: ` : ""}needs {neededByBlock[b]} min but only {availByBlock[b]} min are available between {fmt12(sc.startMin)} and {fmt12(sc.endMin)}{settings.lunchOn ? " (lunch excluded)" : ""}. </span>
                ) : null)}
                Nothing was deleted or shortened: the earliest rows are left without times. Move the start/end time, shorten the minutes, or exclude rows, then generate again.
              </div>
            )}
            {result && result.idleMinutes > 0 && (
              <div className="rounded-lg border border-sky-200 bg-sky-50 p-3 text-sm text-sky-900">
                Records are never split across lunch, so {result.idleMinutes} idle minute(s) remain right after lunch.
              </div>
            )}
            {result && problems > 0 && <div role="alert" className="rounded-lg border border-rose-300 bg-rose-50 p-3 text-sm text-rose-800">Validation found {problems} problem(s) — see the status column. Download is disabled until they are fixed.</div>}
            {count("Needs Review") > 0 && <div className="rounded-lg border border-rose-200 bg-rose-50 p-3 text-sm text-rose-800">{count("Needs Review")} record(s) need review. Pick a category for them in the table, or leave them (they get {two ? (settings.minReview > 0 || settings.minReview2 > 0 ? "the Needs Review minutes" : "no time entry") : settings.minReview > 0 ? `${settings.minReview} min each` : "no time entry"}).</div>}
          </section>

          <section className="overflow-x-auto rounded-xl bg-white shadow-sm">
            <table className="min-w-full text-sm">
              <thead className="bg-slate-100 text-left text-xs uppercase text-slate-600">
                <tr>{["Row", "Original comment", "Category", ...(two ? [blockNames[0], blockNames[1]] : ["Minutes", "Time Started", "Time Ended"]), "Status"].map((h) => <th key={h} className="px-3 py-2">{h}</th>)}</tr>
              </thead>
              <tbody>
                {records.map((r) => {
                  const msgs: string[] = [];
                  if (commentCol >= 0 && !r.comment.trim()) msgs.push("Missing comment");
                  if (r.category === "Needs Review") msgs.push("Unrecognized category");
                  let anySlot = false;
                  if (result && r.category !== "Exclude") {
                    scheds.forEach((_, b) => {
                      if (result.slots[b].has(r.key)) anySlot = true;
                      if (result.unscheduled[b].has(r.key)) msgs.push(`${two ? blockNames[b] + ": " : ""}does not fit in workday`);
                    });
                    if (!anySlot && !msgs.length) msgs.push("No time assigned");
                    msgs.push(...(result.checks.get(r.key) ?? []));
                  }
                  const ok = !!result && anySlot && msgs.length === 0;
                  const sl0 = result?.slots[0].get(r.key) as BSlot | undefined;
                  return (
                    <tr key={r.key} className={`border-t ${r.category === "Exclude" ? "opacity-50" : ""}`}>
                      <td className="px-3 py-2 text-slate-500">{r.rowNumber}</td>
                      <td className="max-w-md px-3 py-2">{r.comment || <i className="text-slate-400">(empty)</i>}</td>
                      <td className="px-3 py-2">
                        <select value={r.category} onChange={(e) => setOverrides((o) => ({ ...o, [r.key]: e.target.value as Category }))}
                          className={`rounded px-2 py-1 text-xs font-medium ${badge[r.category]}`}>
                          {CATS.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                        {r.category !== r.detected && <span className="ml-1 text-xs text-slate-400" title={`Detected: ${r.detected}`}>edited</span>}
                      </td>
                      {two ? (
                        <>
                          <td className="px-3 py-2 text-xs">{r.category === "Exclude" ? "—" : minutesFor(r.category, 0) > 0 ? (result ? slotText(0, r.key) : `${minutesFor(r.category, 0)} min`) : "—"}</td>
                          <td className="px-3 py-2 text-xs">{r.category === "Exclude" ? "—" : minutesFor(r.category, 1) > 0 ? (result ? slotText(1, r.key) : `${minutesFor(r.category, 1)} min`) : "—"}</td>
                        </>
                      ) : (
                        <>
                          <td className="px-3 py-2">{r.category === "Exclude" ? "—" : minutesFor(r.category, 0)}</td>
                          <td className="px-3 py-2">{sl0 ? fmt12(sl0.start) : "—"}</td>
                          <td className="px-3 py-2">{sl0 ? fmt12(sl0.end) : "—"}</td>
                        </>
                      )}
                      <td className={`px-3 py-2 text-xs ${ok ? "text-emerald-700" : msgs.length ? "text-rose-700" : "text-slate-400"}`}>
                        {ok ? "✓ OK" : msgs.length ? msgs.join("; ") : result ? (r.category === "Exclude" ? "Excluded" : "") : "Not generated"}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {!records.length && <p className="p-6 text-sm text-slate-500">{two ? "No records to process. Check the row filter above." : "No records found. Choose the comment column above."}</p>}
          </section>
        </>
      )}
    </main>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label className="block text-sm font-medium">{label}<div className="mt-1 font-normal">{children}</div></label>;
}
function Stat({ label, value, tone }: { label: string; value: number | string; tone?: "warn" | "bad" }) {
  const c = tone === "bad" ? "border-rose-300 bg-rose-50" : tone === "warn" ? "border-amber-300 bg-amber-50" : "border-slate-200 bg-white";
  return <div className={`rounded-lg border p-3 ${c}`}><div className="text-xl font-bold">{value}</div><div className="text-xs text-slate-600">{label}</div></div>;
}
