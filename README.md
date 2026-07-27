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
- Renames and reorders columns to a clean 13-column layout.
- **Consolidates each quote into a single row**, even when a quote was split
  across several physical rows or across a page break.
- Turns the quote date into a real date (`mm/dd/yyyy`), Amount and SQFT into
  formatted numbers, and keeps Account / Quote / Invoice as text (so long ID codes
  never turn into scientific notation).
- Normalizes the **Created By** names so the same person is counted once.
- Splits each product spec onto its own line inside the Product cell.
- Distills the free-text **GTI Comments** into two filterable columns —
  **Quote Category** and **Ordered** (the invoice date) — and no longer emits the
  raw comment column itself (see below).
- Produces a **run summary**: rows read in, quotes out, quotes ordered (those with
  an invoice), a count per quote category, and any items that need a human's eye.

Output columns, in order (displayed header in **bold**, internal field in
parentheses where it differs):

```
Account (Account No.) | Customer (Account Name) | Quote | Date (Est. Date) |
Quote Category | Job Name | Amount | Product | SQFT | Created By |
Internal Note | Invoice | Ordered (Invoice Date)
```

- The free-text **GTI Comments** column is no longer emitted. Its content is
  distilled into **Quote Category** and **Ordered** (both derived from it), so the
  wide unstructured column no longer clutters the sheet.
- Displayed headers are shortened for readability: Account No. → **Account**,
  Account Name → **Customer**, Est. Date → **Date**, Invoice Date → **Ordered**.
- **Internal Note now wraps** and is set to the same width as the **Product**
  column; the **Account** column is narrowed to fit its ~4-digit account numbers.

## Ordered (invoice date) and Quote Category

**Quote Category is a 3-value field: `Duplicated Quote`, `Web Quote`, or
`Phone/Email`.** It used to be derived purely from the shape of the GTI Comments
text, but that broke down for invoiced quotes: GTI Comments holds only the
**latest** lifecycle event, so once a quote is invoiced its comment becomes
`Created Invoice Number: N on <date>` and any earlier web-quote or duplicate
marker is gone for good — comment shape alone can no longer say "was this
originally a web quote" once it's been invoiced.

**Created By** does not have that problem — it's a separate export column that
invoicing never touches. In the reference export, every quote whose comment
says `Estimate created from Web Quote N` has Created By exactly `WEBSITE` (63
of 63), and that holds whether or not the quote was later invoiced (67 more
invoiced quotes still show `Created By == WEBSITE`). `WEBSITE` never appears
on a duplicated or blank-comment quote. So the rule is now:

| Priority | Quote Category | Rule |
|---|---|---|
| 1 | `Duplicated Quote` | Comment matches `Quote based on Est No: N` |
| 2 | `Web Quote` | else, raw (pre-name-mapping) Created By == `WEBSITE` |
| 3 | `Phone/Email` | else — every blank-comment quote (a CSR took a call or email) and every invoiced quote whose Created By is a person, not WEBSITE |

Reference-export counts: `Duplicated Quote` = 48, `Web Quote` = 130,
`Phone/Email` = 739 (48 + 130 + 739 = 917, the full quote count).

**Ordered (Invoice Date) is now fully independent of Quote Category.** It is
derived purely from a comment matching `Created Invoice Number: N on <date>`,
regardless of what category the row ends up in:

- **Ordered is a real date** (`mm/dd/yyyy`), so it sorts and filters like the quote
  **Date**. It is stored **date-only** — the time of day in the comment is
  dropped. A comment the cleaner cannot read a date from leaves the cell blank and
  is reported in the run summary; nothing is ever guessed or back-filled.
- The invoice **number** is not duplicated into a new column: it is already in the
  existing `Invoice` column, and across the reference export the two agreed on all
  350 invoiced quotes.
- **A comment that matches none of the three known shapes** (invoice, web-quote,
  duplicate) is still flagged in the run summary (`flag_unrecognized_comments`),
  even though it no longer changes Quote Category — this should normally be zero,
  and a non-zero count means GTI started writing a comment shape nobody has seen
  before. Pass the examples along so the pattern can be added.

One wrinkle worth knowing about: when a quote lands on a page break, the export can
push the tail of its invoice comment (`... on Jun 30 2026 7:32AM`) onto the next
page's header line, leaving the visible comment ending at `on`. Those lines are
otherwise discarded as page furniture. The cleaner picks that date back up — it is
the only text on such a line and it is only ever used for a comment that has no
date of its own — and reports how many it rebuilt.

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

Two folders exist in the local working copy but are deliberately **not published**,
because both quote real figures and real names:

| Local-only folder | Purpose |
|---|---|
| `_local_validation/` | Regression harness. Runs the pipeline against a private sample export and asserts a fixed baseline. |
| `Build files/` | Planning and session handoff notes. |

Both are listed in `.gitignore`. Keep them there.

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

`FINAL_TO_SOURCE` covers only the columns copied from the export. Invoice Date and
Quote Category are listed separately in `DERIVED_COLUMNS`, and the comment shapes
they recognize live in the regular expressions just below it — that is the place to
add a new comment pattern if `Other` ever shows a non-zero count.

## Four deliberate design decisions

These are the only places where the implementation makes an explicit judgment
beyond the literal wording of the spec. All are safe and deterministic; they are
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

3. **An invoice date stranded on a page-header line is recovered.** A page break
   can leave the date fragment of an invoice comment on the following page's
   sub-header row, which furniture removal correctly discards — the visible comment
   then ends at `on`. Such a row holds a bare date-time in the Comments column and
   nothing else, so the value is identifiable rather than lost. The cleaner reads it
   back, but only ever applies it to a comment that carries no date of its own, so a
   stray match can never overwrite a real value. In the reference export this
   affected 7 of 350 invoiced quotes, and the run summary reports the count.

4. **Quote Category uses Created By, not just Comments, to detect web origin —
   checked against the RAW, pre-name-mapping value.** GTI Comments only ever
   holds the latest lifecycle event, so invoicing permanently overwrites a
   quote's original "Estimate created from Web Quote N" comment; Created By is
   never touched by invoicing, so it stays a reliable signal. This is grounded
   in the real export (100% correlation, 0 exceptions: every Web-Quote-shaped
   comment has Created By == WEBSITE, whether or not later invoiced) and
   confirmed by the business owner — duplicated quotes are always CSR-created
   and can never happen through the website, so there is no ambiguity between
   the Duplicated Quote and Web Quote checks. The RAW Created By value (before
   the optional per-browser name-normalization mapping runs) is used
   specifically so that mapping — an unrelated feature for tidying operator
   names — can never accidentally relabel the literal string `WEBSITE` and
   misclassify a web-origin quote.

## Deployment (GitHub Pages)

This is a static site — no build step. Commit `index.html`, `app.js`, `cleaner.py`,
and `README.md` to the repository, then enable **GitHub Pages** on the `main`
branch (root). The page is served at the Pages URL; share that link with users.

> Do not commit real quote exports. `.gitignore` already excludes `*.xls` /
> `*.xlsx`, plus the `_local_validation/` and `Build files/` folders, so neither
> sample data nor internal notes land in this public repo. Stage files by name
> rather than with `git add -A`.
