# xls2spread

*[中文說明](README.md) · English*

Turns a spreadsheet like Taiwan's "Fixed Asset Construction, Improvement and Expansion Plan
and Cost-Benefit Analysis" into a **two-page spread with locked top and bottom rules**.
Adjust the row heights, print straight to PDF, or write the result back into the xlsx and
keep tweaking in Excel.

## The problem

A government fixed-asset cost-benefit table is a **table that spans a two-page spread**:
odd pages carry the fund and project names and the funding sources, even pages carry the
project details and the budget figures. The same row has to line up across both pages, and
the top and bottom rules — the two lines that bound the table — have to land in the same
place on every page, or the whole document jumps up and down as you flip through it.

Laying this out in Excel runs into four walls:

1. **Row heights are chained to the cell grid.** Both halves of a row have to accommodate
   whichever side is taller, so the short side just sits there empty. On the real file,
   30% of the total row height was blank space created this way.
2. **The bottom rule follows the last row.** Excel draws table borders as far as the
   content goes, so every page ends somewhere different.
3. **Everything is coupled.** Change one row's height and the page breaks downstream all
   shift, so you go back and fix the breaks by hand, over and over.
4. **Alignment only works by hand.** To get a hanging indent after a line wrap, you insert
   a manual line break inside the cell and then color the padding characters white.

## The approach

**Excel is only the data source; the layout belongs to HTML/CSS.**

At the core it is an allocation problem: distribute every data row inside a **fixed
rule-to-rule height** so the final page count is as sensible as possible — usually the
smallest, but never at the cost of how the data reads. So the layout is recomputed in full
on every redraw, not stamped out once:

- The page setup is read and used as-is: the column break decides where the spread splits,
  the print titles decide the repeating header, the paper margins decide the text block,
  and cell styles (font, borders, alignment, indent, number format) are reproduced exactly —
  including the white-text indent trick. Distributed alignment (the "T o t a l" look) wraps
  each character and pushes them apart with flex rather than relying on CSS
  `text-align-last`, so every browser agrees.
- **The rules are locked.** Each page's table height is fixed; any shortfall is taken up by
  a single blank row, and the column rules run all the way down to the bottom rule. On the
  last page the totals and notes are pinned to the bottom, with the foot of the notes flush
  with the other pages' bottom rule.
- **Row heights are not copied from Excel.** On load every row is squeezed to the height the
  text on both pages actually needs, and both halves share one value — which is exactly why
  the odd and even pages always line up.
- **Page breaks are computed.** Raise a row height until it no longer fits and the trailing
  rows are pushed to the next page; lower it and rows flow back up. The bottom rule never
  moves. Pin a break somewhere specific when you want one.
- **Rows invisible in Excel are dropped.** Rows carrying the hidden flag, plus rows whose
  height is squeezed below their own font size (this file has 18 of them, 4.5–4.9pt, holding
  leftover values nobody cleaned out). You cannot see them in Excel; copying them over and
  then giving them a real height would surface content that should not be there.
  209 rows → 191 rows on the real file.
- **The widow rule is not forced.** A group heading (no project content of its own, and the
  next row is indented further) or a blank spacer row landing at the foot of a page makes the
  content look cut off. The tool tries to push it to the next page, but **only when that does
  not cost an extra page**; if it would, the row stays put and gets an orange underline so you
  can decide. (Measured on the real file: 3 stranded headings before, 0 after, with the group
  count unchanged at 11.)
- When you are done, Cmd+P is the deliverable — or export an xlsx carrying the new row
  heights and page breaks back to Excel.

## Two ways to use it

**`index.html` (the main one)** — a single file. Open it, drop an .xlsx on it, start
adjusting. Nothing to install, no backend, and the file never leaves your machine. Unzipping
uses the browser's built-in `DecompressionStream` and the XML is read with `DOMParser` — no
external libraries at all.

**`xls2spread.py` (only for .xls)** — browsers cannot read the old BIFF .xls format. This
script does the conversion and nothing else (via LibreOffice, so it needs `soffice`); the
layout always happens back in index.html:

```bash
python3 xls2spread.py cost-benefit-analysis.xls
```

If you have Excel, Save As → Excel Workbook (.xlsx) does the same thing and you will not
need this script.

## Putting it online

Everything is static, so pushing it is enough:

| Platform | How |
|---|---|
| GitHub Pages | Push the repo → Settings → Pages → pick a branch. With `index.html` at the root it is the home page |
| Hugging Face Space | Create a Space, choose **Static**, put `index.html` at the root (a Static Space's entry point must be named index.html) |

Because all the parsing happens in the browser, **hosting it publicly still does not give
anyone access to the budget files you open in it**. Browser requirements: Chrome 80+ /
Safari 16.4+ / Firefox 113+ (for `DecompressionStream`).

`.gitignore` already blocks `*.xls` and `*.xlsx` so real budget files used for testing do
not get pushed.

## Where the layout comes from

All of it is read from Excel; none of it is hard-coded:

| Layout element | Source |
|---|---|
| Left / right page columns | The column page break (between F and G in this file) |
| Rows per page | The horizontal page breaks in the source |
| Repeating header | The print title rows (rows 1–6) |
| Paper size, margins | Page setup |
| Row height, column width, font, alignment, indent, borders, number format | Cell styles |
| Totals + notes | Everything after the row with a thick bottom border, pinned to the rules |

Where the output differs from the source, it is deliberate:

- **The rules are fixed.** Every page's table height is locked to one value, a blank row
  takes up any shortfall, and the vertical rules run down to the bottom rule. Excel and
  LibreOffice let the bottom rule follow the last row, so every page ends at a different
  height.
- **Row rules are dropped.** Data rows keep only the vertical column rules; horizontal lines
  appear on the header and on the top and bottom rules.
- **The first spread's title gets no "(cont.)".** The source's title rows repeat on every
  page, which would print "(cont.)" on every even page; only page 2 should read
  "… Summary", with "… Summary (cont.)" starting on page 4.
- **Automatic breaks by row height** (on by default): raise a height until it no longer fits
  and the trailing rows move to the next page; lower it and rows flow back up. The bottom
  rule never moves. Turn it off and the bottom rule follows the content instead — rows past
  the edge of the paper will not print.

  **The source's own manual breaks are hard boundaries and rows never flow across them** —
  they are semantic (one fund per page), not height-driven. To let a spot flow, click the
  "⛔ Manual break" marker in that page's top-left corner to remove it, or press
  "Clear breaks" for pure height-driven flow (18 groups → 11 on this file, once the row
  heights come down by 30%).

## It reflows the HTML way by default, instead of copying Excel

Excel's row heights are forced by the cell grid (both halves of a row have to match the
taller side), and its breaks are cut to fit that grid. HTML has neither constraint, so **the
first load reflows everything from the content**:

- Every row is squeezed to the height the text on both pages actually needs (total row
  height on this file: 10890pt → 7594pt, 30% less)
- Breaks are decided purely by height, ignoring the source's break points (18 groups → 11 on
  this file, with each page 90–99% full)

Press "Reset" for Excel's original layout, or "Reflow" to do it again.

The default for `Scale` is derived from Excel's own fitToPage algorithm (shrink the wider
half-page to the width of the text block). Column widths are stored in the file as a count
of characters, and converting that to a real size means multiplying by the digit width of
the default font — a baseline that gets distorted by file conversion, so the derived value
is off by a few percent (80.7% derived here, against 77.0% measured from LibreOffice's
actual output). If it matters, drag "Scale" to the value you want; it is saved
automatically.

## Working with it

| What you want | How |
|---|---|
| Resize one row | Drag its bottom edge / select it and press ↑↓ (Shift = ×5) / type a number under "Selected row" |
| Fit one row automatically | Double-click it |
| Break after a specific row | Click the running total on the left ruler |
| Remove a break that is blocking the flow | Click "⛔ Manual break" in that page's top-left corner |
| Reflow purely by height | "Clear breaks" |
| A page looks crowded on top, empty at the bottom | "This page → Even heights" spreads the leftover space across that page's rows ("All pages → Even heights" does the lot) |
| More air between rows | The "Row spacing" slider, applied to every page |
| Document-wide | Scale, row spacing, line height, font size, rule-to-rule height, last-page bottom rule |
| On-screen zoom | Under View → Zoom on the left: **Fit width** (the whole spread, gutter and ruler included, fills the window), **Fit height** (one page fills it), **Custom** (drag the slider). It re-fits when the window changes and centres itself when the spread is narrower than the window. Screen only — printing is always full size |
| Back to Excel | "Export xlsx" writes the row heights and breaks into the original file |
| Forgot how something works | "How to use" on the toolbar opens a cheat sheet you can copy |
| Want to take a change back | ⌘Z / Ctrl+Z to undo, ⇧⌘Z to redo; deleting a row and dragging to reorder are not undoable |
| Saving | Automatic, in the browser; "Export settings" saves a json you can carry around |

The same row is always the same height on both pages (one value applied to both halves),
which is what keeps the odd and even pages aligned.

The totals and notes on the last page are pinned to the bottom: with "Last page ends on the
bottom rule" checked, **the foot of the notes lands on the same line as the other pages'
bottom rule**, and the last page's own bottom rule (under the totals) is pushed up by however
tall the notes are. Uncheck it to set the last page's height by hand.

There are two ways to measure the notes; switch in "Totals & notes", where the panel also tells you
how far the last page's rule moved, in pt and mm:

- **Shrink to content (default)**: note rows keep only the height their text needs, dropping the blank
  space the source left, so the rule sits lower and the last page can take a few more rows.
- **Keep source note heights**: each note row keeps the height the author gave it in Excel and only grows
  when the text no longer fits, which puts the rule exactly where the source had it.

Either way the foot of the notes stays flush with the other pages' rule, and a note row is always the
same height on both halves of the spread.

**Re-wrapping across the spread** (Totals & notes → Re-wrap across the spread): a spread's notes are
one sentence running across both sheets, but in the source workbook the left and right cells wrap on
their own — once the browser's fonts are used the two halves often end up with different line counts
(2 on the left, 1 on the right) and the sentence reads broken. This button joins the two halves into
one paragraph and lays it out across "left width → right width → left width…", writing each line's
left part back into the left cell and its right part into the right cell:

- both halves get the same number of lines, each line on the same baseline (aligned across the gutter)
- Chinese punctuation never starts a line (，。、) and never ends one (（「); Latin words are not split
- notes you have not edited are re-wrapped too; edited ones keep your text, and every press starts
  from the source text, so the result is stable
- when the right half is a note of its own (starting with a number or 註), the halves are wrapped
  separately instead of being joined
- notes split over several cells (e.g. a lone 註： cell) cannot be joined and are skipped, with a note
  in the panel

The re-wrap writes the on-screen text through the same path as manual editing, so ⌘Z undoes it and
you can always reset back to the workbook's own content.

**Running totals** (in mm on paper, header already deducted): each page shows
`Used 195.1 / 219.2mm · 24.1mm left` in its top-right corner, and the left ruler labels the
running total row by row, marking the last one `▏full`. A red underline means that cell's
text is being clipped — that row needs to be taller.

"Row spacing" and row height are stored separately: the row height records only what the
text needs, and the spacing is added at layout time. So you can drag the spacing at any
point, and "All pages → Even heights" will not wipe it out.

**Export xlsx**: writes the adjusted row heights (spacing included) and page breaks back
into the worksheet and copies every other byte across untouched — formulas, styles, column
widths, macros and shapes all survive, so you can keep tweaking in Excel. Note that Excel
prints with its own fitToPage ratio (about 5% smaller than this tool's default), so pages
look slightly looser there than in the preview; the break points are unaffected.

Print with A4 portrait, margins set to "None", headers and footers off, and scale at 100%
(the layout is already at real size). Measured in Chrome: 36 pages, 209.9×297mm, no blank
pages.

**When something shows up in the PDF that should not be there**: a color block, icon, border
or watermark in the same corner of every page is almost always a Chrome extension (PDF
toolbar, screenshot, translate, ad blocker) painting over the page and getting printed with
it. It is not part of this layout. Print once from an incognito window, which loads no
extensions, and it disappears.

## Checks and flags

- **Cross-foot** (the "Checks" section): re-adds the hierarchy using the numbers currently
  on screen (grand total = Σ agencies → agency = Σ funds → fund = Σ major categories →
  category = Σ projects) plus the cross-checks (the agency rows, the project-type totals),
  listing every mismatch. **Source check** reads the "variance" block Excel computed when the
  file was saved and verifies it is all zeros. Results open in a scrollable, copyable panel
  (✗ red, ✓ green, with "Copy all").
- **Cells with no alignment set**: where the source leaves alignment unset (Excel's
  "General"), left / right is guessed from the content, and those cells get a red background
  so you can confirm them by hand. Red does print, and it is in the exported xlsx too;
  uncheck "Red" under "View" to hide it.
- **Empty cells**: cells inside the table rules with no value (fully blank spacer rows do not
  count) get a **light orange** background, and the count plus the first few addresses are
  reported on load. Orange is display only and never reaches the exported xlsx. If any are
  still on screen when you press "Print / Save PDF", you get one confirmation first (cells
  you fill in while editing disappear from the count). Uncheck "Orange" under "View" to hide
  it.
- Two more flags work the same way: **purple** for a cell aligned differently from the rest
  of its column, **green** for a cell mixing fonts against the document's convention (hover
  to see which font). Both are on screen only.
- The settings panel on the left is grouped by **what a control affects** — Selected row,
  Row heights & breaks, Selected column, Totals & notes, Structure, Text, Layout, Checks,
  View, Editing & settings. Sections collapse, and which ones you left open is remembered.
  The toolbar across the top holds the document-wide actions: undo/redo, Print / Save PDF,
  Export xlsx, New file, and the language, theme and panel toggles.

## Interface language

The UI ships in Traditional Chinese and English. The **EN / 中文** button — top-right of the
drop screen, and in the toolbar once a file is open — switches between them, and the choice
is remembered in the browser. Only the interface is translated; your spreadsheet's own
content is never touched.
