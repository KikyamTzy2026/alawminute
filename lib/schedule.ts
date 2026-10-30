import { Slot } from "./types";

export interface ScheduleSettings {
  startMin: number;
  endMin: number;
  lunchOn: boolean;
  lunchStart: number;
  lunchEnd: number;
}

export function validateSettings(s: ScheduleSettings): string[] {
  const errs: string[] = [];
  if (Number.isNaN(s.startMin) || Number.isNaN(s.endMin)) errs.push("Workday start/end time is invalid.");
  else if (s.endMin <= s.startMin) errs.push("Work end time must be after the workday start time.");
  if (s.lunchOn) {
    if (Number.isNaN(s.lunchStart) || Number.isNaN(s.lunchEnd)) errs.push("Lunch break time is invalid.");
    else if (s.lunchEnd <= s.lunchStart) errs.push("Lunch break end must be after lunch break start.");
  }
  return errs;
}

export interface ScheduleItem { key: number; minutes: number }

export interface ScheduleResult {
  slots: Map<number, Slot>;
  unscheduled: number[]; // keys that did not fit (never silently dropped)
  idleMinutes: number;   // gaps created by moving records in front of lunch
}

/**
 * Schedules items BACKWARD from endMin, preserving the given order.
 * A record is never split: if it would touch the lunch break, the whole record is moved so that it
 * ends exactly when lunch starts. Records that no longer fit after workday start are reported.
 */
export function buildSchedule(items: ScheduleItem[], s: ScheduleSettings): ScheduleResult {
  const slots = new Map<number, Slot>();
  const unscheduled: number[] = [];
  let cursor = s.endMin;
  let idle = 0;
  let full = false;

  for (let i = items.length - 1; i >= 0; i--) {
    const it = items[i];
    if (full || !(it.minutes > 0)) { unscheduled.push(it.key); continue; }
    let end = cursor;
    if (s.lunchOn && end > s.lunchStart && end - it.minutes < s.lunchEnd) {
      idle += end - s.lunchStart;
      end = s.lunchStart;
    }
    const start = end - it.minutes;
    if (start < s.startMin) { full = true; unscheduled.push(it.key); continue; }
    slots.set(it.key, { start, end });
    cursor = start;
  }
  unscheduled.reverse();
  return { slots, unscheduled, idleMinutes: idle };
}

export interface CheckRow { key: number; minutes: number; slot?: Slot }

/** Independent re-check of a finished schedule; returns messages per key. */
export function checkSchedule(rows: CheckRow[], s: ScheduleSettings): Map<number, string[]> {
  const out = new Map<number, string[]>();
  const add = (k: number, m: string) => out.set(k, [...(out.get(k) ?? []), m]);
  let prevEnd = -Infinity;
  for (const r of rows) {
    if (!r.slot) continue;
    const { start, end } = r.slot;
    if (end - start !== r.minutes) add(r.key, "Duration mismatch");
    if (end <= start) add(r.key, "Non-positive duration");
    if (s.lunchOn && start < s.lunchEnd && end > s.lunchStart) add(r.key, "Lunch-break conflict");
    if (start < prevEnd) add(r.key, "Overlaps previous record");
    if (start < s.startMin || end > s.endMin) add(r.key, "Outside workday");
    prevEnd = end;
  }
  return out;
}

/** Records from different blocks must not overlap each other either (one person, one task at a time). */
export function crossOverlaps(slots: { key: number; slot: Slot }[]): Set<number> {
  const bad = new Set<number>();
  const sorted = [...slots].sort((a, b) => a.slot.start - b.slot.start || a.slot.end - b.slot.end);
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i].slot.start < sorted[i - 1].slot.end) { bad.add(sorted[i].key); bad.add(sorted[i - 1].key); }
  }
  return bad;
}
