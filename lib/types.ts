export type Category = "Completed" | "Did not proceed" | "Needs Review" | "Exclude";

export interface Rules {
  dnp: string[];
  completed: string[];
}

export interface Settings {
  /** "single" = one Time Started/Ended/Minutes block; "two" = e.g. Billing (morning) + Posting (afternoon) */
  mode: "single" | "two";
  b1Name: string; b1Start: string; b1End: string;
  b2Name: string; b2Start: string; b2End: string;
  /** minutes per record for the SECOND set in "two" mode (0 = leave the second set empty for that category) */
  minDnp2: number; minCompleted2: number; minReview2: number;
  /** how a row with an empty comment is classified */
  blankAs: Category;
  startTime: string; // "HH:MM" 24h
  endTime: string;
  lunchOn: boolean;
  lunchStart: string;
  lunchEnd: string;
  minCompleted: number;
  minDnp: number;  // first set (or the only set)
  minReview: number; // 0 = leave Needs Review rows unprocessed
}

export interface RecordRow {
  key: number;        // 0-based index of the row inside the sheet
  rowNumber: number;  // 1-based Excel row number
  comment: string;
  detected: Category; // automatic result
  category: Category; // current (possibly user-corrected)
}

export interface Slot { start: number; end: number } // minutes from midnight
