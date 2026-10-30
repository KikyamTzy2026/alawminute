# Excel Time Tracker Generator

Upload an Excel work file (.xlsx / .xls), auto-classify each row's comment, and fill **Time Started**, **Time Ended** and **Minutes**
scheduled backward from your work end time, skipping the lunch break. 100% client-side: the file never leaves your browser.

## Run locally
```bash
npm install
npm run dev      # http://localhost:3000
npm test         # classification + scheduling + Excel output checks
npm run build && npm start
```
Requires Node 18.18+ (Node 20/22 recommended).

## Deploy to Vercel
1. Push this folder to a GitHub repo (`git init && git add . && git commit -m init`, then push).
2. In Vercel: **Add New → Project → Import** the repo. Framework is auto-detected as Next.js. No environment variables needed.
3. Click **Deploy**. (CLI alternative: `npx vercel --prod`.)

## Structure
```
app/page.tsx        dashboard UI (upload, settings, summary, preview, download)
app/layout.tsx, app/globals.css
lib/classify.ts     keyword classifier (Did not proceed is checked before Completed)
lib/schedule.ts     backward scheduler + independent validator
lib/excel.ts        read .xls/.xlsx (SheetJS), write workbook (ExcelJS)
lib/xlsStyles.ts    reads cell formatting out of legacy .xls files
lib/time.ts         time helpers
tests/logic.test.ts automated checks
```

## Billing project (exclusive option)
Choose **Project → Billing** for the Billing sheet (two sets of Time Started / Time ended / Minutes, Initial Auditor, TOTAL MINUTES).
Nothing below applies to **Standard**, which behaves exactly as before.
* Row filter: only rows where a column (default *Processor Name*) equals your name are processed; every other row is left exactly as it is.
* Two sets, each with its own start/end time (defaults Billing 8:00–11:30 AM, Posting 12:30–5:00 PM) and its own minutes per category
  (0 = that set stays empty for that category). A row can have time in both sets.
* The sheet's own formulas are kept: Minutes (`=MINUTE(W2-V2)`), second Minutes, and `TOTAL MINUTES` (`=SUM(X2,AB2)`); only their results refresh.
* If the sheet has no comment text, choose how an empty comment counts (Needs Review / Completed / Did not proceed) and fix single rows in the table.
* The app never switches projects by itself; for a Billing-looking sheet it only shows a hint.

## How scheduling works
* Records keep their original order and are placed backward from **Work end time**.
* A record is never split. If it would touch the lunch break (default 11:30 AM–12:30 PM), the whole record is moved to end at lunch start.
  This can leave idle minutes right after lunch; the app tells you how many.
* If records do not fit between workday start and end, the earliest rows are left without times and clearly flagged. Nothing is deleted or shortened.
* Time cells are real Excel time values formatted `h:mm AM/PM`. Existing Time Started / Time Ended / Minutes columns are updated in place; missing ones are appended.
* Minutes formula: `=MINUTE(MOD(L2-K2,1))` with references taken from the real column positions (MOD handles midnight).
  `MINUTE()` only returns 0–59, so if any duration is 60+ minutes the app uses `=ROUND(MOD(L2-K2,1)*1440,0)` instead.
* Rows that are excluded or unscheduled have their Time/Minutes cells cleared in the output; all other original data is untouched.
* Only the Time Started / Time Ended / Minutes cells are changed. Legacy .xls files are rebuilt as .xlsx with their original fonts, header fill, alignment, number formats (e.g. the Date (EST) column), column widths and row heights.
