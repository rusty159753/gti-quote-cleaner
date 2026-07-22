# GTI Quote Cleaner

A single web page that turns a raw GTI **Quotes List** export (`.xls`) into a
clean, sortable spreadsheet (`.xlsx`) — one row per quote, real dates, formatted
numbers, filterable table.

**Everything runs in your browser.** Your file is never uploaded, stored, or sent
anywhere. No login, no install, no admin rights needed.

## One-line usage

Open the page, drop in the raw `.xls` export, click **Clean file**, and download
the cleaned `.xlsx`.

## What it does

The cleaner applies a fixed, deterministic set of rules (the same input always
produces the same output):

- Removes the repeating page furniture (the `QUOTES LIST` / `FOR DATE RANGE`
  banners and the repeated column-header rows).
- Drops the empty spacer columns the export uses for PDF layout.
- Renames and reorders columns to a clean 12-column layout.
- **Consolidates each quote into a single row**, even when a quote was split
  across several physical rows or across a page break.
- Turns Est. Date into a real date (`mm/dd/yyyy`), Amount and SQFT into formatted
  numbers, and keeps Account No. / Quote / Invoice as text (so long ID codes never
  turn into scientific notation).
- Normalizes the **Created By** names so the same person is counted once.
- Splits each product spec onto its own line inside the Product cell.
- Produces a **run summary**: rows read in, quotes out, quotes ordered (those with
  an invoice), and any items that need a human's eye.

Output columns, in order:

```
Account No. | Account Name | Quote | Est. Date | Job Name | Amount |
Invoice | GTI Comments | Product | SQFT | Created By | Internal Note
```

## Files

| File | Purpose |
|---|---|
| `index.html` | The page and its styling. |
| `app.js` | Loads the in-browser engine and wires up the buttons. |
| `cleaner.py` | The cleaning rules. This is where all the logic lives. |
| `README.md` | This file. |

The page loads [Pyodide](https://pyodide.org) (Python compiled to run in the
browser) from a public CDN, plus two small pure-Python libraries — `xlrd` (to read
the legacy `.xls`) and `openpyxl` (to write the `.xlsx`). The first visit downloads
the engine once (~7 MB) and the browser caches it; later visits are fast. **None of
this transmits your file** — only program code is fetched, and your spreadsheet is
processed entirely inside the tab.

## Maintaining the "Created By" name mapping

The cleaner can normalize the **Created By** column so the same person is counted
once (for example two spellings folded into one name). It **flags anything it does
not recognize** in the run summary and never guesses a name's spelling.

**The names are not stored in this code.** No operator names live in the
repository. Each person enters the mapping in their own browser, where it is saved
locally and never uploaded. This keeps the public site free to host while keeping
real names off the public repo.

To set it up, open the tool and expand **"Manage 'Created By' names (optional)"**.
Enter one name per line as `WHAT THE EXPORT TYPES = Clean Name`, for example:

```
RAWNAME = Clean Name
OTHERSPELLING = Clean Name
CleanAlready = CleanAlready
```

- Map a raw value to the clean name you want (e.g. `SOMENAME = Some Name`).
- If two different spellings should count as one person, map both to the same
  clean name.
- To keep a name that already comes through correctly and stop it being flagged,
  map it to itself (`Some Name = Some Name`).

Click **Save names** to store them in that browser. Use **Export file** to save a
private `gti-created-by-names.json` you can keep somewhere safe, and **Import file**
to load it into another browser or computer — so you set the list up once and reuse
it. The mapping applies the next time you clean a file.

> Because the names live only in each browser, a brand-new browser starts with an
> empty list and will flag every name for review until you Save or Import a mapping.
> That is expected — nothing is broken.

## If GTI changes its export format

The cleaner finds each column by **matching the export's own header text**, not by
fixed column position, so small layout shifts are handled automatically. If GTI
renames a required column or drops one, the tool **stops and tells you plainly**
rather than producing a wrong file — it will name the missing column. If that
happens, update the header labels in the `FINAL_TO_SOURCE` table in `cleaner.py` to
match the new export.

## Two deliberate design decisions

These are the only two places where the implementation makes an explicit judgment
beyond the literal wording of the spec. Both are safe and deterministic; they are
recorded here so a future maintainer understands why.

1. **Est. Date accepts real Excel dates as well as `yyyy-mm-dd` text.** The
   original spec assumed dates arrive as ISO text. The real exports store Est. Date
   as genuine Excel date cells (especially once the file has been opened in Excel),
   so the cleaner handles both and only flags a date it genuinely cannot read. This
   is lossless — no date is ever guessed.

2. **Page-split rows collapse repeated fragments.** When a quote is split across a
   page break, the export repeats its single-value fields (Job Name, Created By,
   SQFT, etc.) on the continuation row while the Account Name wraps. When joining
   those fields the cleaner drops a fragment that exactly repeats the one before it,
   so a company name whose halves land on different pages (e.g. `ACME GLASS` +
   `PORTLAND, LLC.`) joins to `ACME GLASS PORTLAND, LLC.`, while a job name repeated
   on both rows (e.g. `JOBREF` + `JOBREF`) stays `JOBREF`. Without this, page-split
   quotes would show doubled names and unreadable SQFT values.

## Deployment (GitHub Pages)

This is a static site — no build step. Commit `index.html`, `app.js`, `cleaner.py`,
and `README.md` to the repository, then enable **GitHub Pages** on the `main`
branch (root). The page is served at the Pages URL; share that link with users.

> Do not commit real quote exports. `.gitignore` already excludes `*.xls` /
> `*.xlsx` so sample data never lands in this public repo.
