import { Category, Rules } from "./types";

export const DEFAULT_RULES: Rules = {
  dnp: [
    "did not proceed",
    "didn't proceed",
    "not proceed",
    "not our case",
    "duplicate",
  ],
  completed: [
    "completed",
    "complete",
    "fdl",
    "drafted",
    "forwarded",
    "filer name",
    "saved to the file",
    "documents were saved",
  ],
};

const norm = (s: string) =>
  s.toLowerCase().replace(/[\u2018\u2019]/g, "'").replace(/\s+/g, " ").trim();

export function classify(comment: string, rules: Rules): Category {
  const c = norm(comment ?? "");
  if (!c) return "Needs Review";
  // Priority 1: explicit "do not proceed" outcomes always win
  // (e.g. "Not our case. Mail was moved to completed.")
  if (rules.dnp.some((k) => k.trim() && c.includes(norm(k)))) return "Did not proceed";
  // Priority 2: completed work
  if (rules.completed.some((k) => k.trim() && c.includes(norm(k)))) return "Completed";
  // Priority 3: ambiguous -> never guess
  return "Needs Review";
}
