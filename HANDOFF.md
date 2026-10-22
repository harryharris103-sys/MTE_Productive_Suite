# MTE Productive Suite – handoff note (v0.1.2)

Single self-contained file: `MTE_Productive_Suite.html` (~145 KB). Only external dependency: JSZip 3.10.1 from cdnjs (needed for .xlsx/.pptx/.ods/.docx import/export). Everything else is inline.

## Architecture
- Landing page of app cards; each module is an IIFE registered with `MTE.registerApp(...)`; exporters with `MTE.addExporter(appId,{ext,label,run})`; `MTE.save(blob,name)` downloads. Helpers: `$`, `tb(text,mime)`, `esc`, `NL`, `TAB`, `fillSel('#sx','sheets')`.
- Modules: Word (docx import/export, pagination, footnotes, comments, ruler margins), Spreadsheets, Presenter, Notes, Paint (layers). Persistence is localStorage (`mte-deck-v1`, `mte-notes-v1`, spreadsheet workbook key).

## Spreadsheets
- Data model: `WB={sheets:[{name,cells:{A1:{v,b,bg,co,al,nf,n}},w:{colIndex:px},objs:[]}],a:activeIndex,names:{}}`. `v` is raw text (formulas start with `=`). `nf` is one of n2, comma, cur, pct, int.
- Grid: 52 columns (A–AZ) x 300 rows (`NC`,`NR`); `L(i)` index→letters, `ci(s)` letters→index, `pk(key)`→[col,row], `kk(c,r)`→key.
- Calc: `ev()` rewrites the formula (defined names via `nm_`, `Sheet!A1` / `'My Sheet'!A1:B3` → `SC()/SR()`, ranges → `R()`, cells → `C()`) and runs it with `Function`, whitelisting identifiers from `FN`. Cross-sheet reads go through `xs()` (swaps `cells`, depth limit 40 → `#CYCLE`).
- Ribbon: config object `T` in the block just before `fillSel('#sx','sheets')`; tab order is fixed in `rib()`. Add a feature = add `B('Label',fn)` to a tab.
- xlsx export `xl()` writes values, formulas, styles.xml (bold, fill, colour, alignment, number format), column widths. `importX()` reads the same back. ODS/XML2003/SYLK/DIF/PRN/XLTX/Unicode/DOS exporters sit before the `pdf` entry.
- Floating objects (pictures, shapes, icons) live in `sheet.objs` and are drawn by `rO()` over the grid (not scrolled with cells, not in .xlsx). Ink is a canvas; strokes are kept in `ST` for Ink Replay / Ink to Shape (not saved).

## Presenter
- Deck: `D={slides:[{bg,notes,tr,items:[{id,t:'text'|'img'|'vid'|'rect'|'ell',x,y,w,h,...}]}]}` on a 960x540 stage.
- `pptx()` (async) exports text, shapes, images, embedded video (blob fetched; grey poster image), notes (notesMaster + notesSlides), transitions. `openPptx()` imports the same set (scales by slide size). Videos imported from a .pptx are blob URLs and do not survive a page reload.

## Testing recipe (worked well)
Playwright + `/opt/pw-browsers/chromium-1194/chrome-linux/chrome`, route the jszip URL to a local copy (no network), drive via DOM (`td[data-k="A1"]`, `#rib button`, `#sx` export select + `waitForEvent('download')`). Validate outputs with python-pptx, openpyxl and `soffice --headless --convert-to`.

## Pitfalls hit
- A global string replace once edited two modules (a spreadsheet export leaked into Word, and a styles override into the .pptx content types). Always assert match counts (`count==1`) when patching.
- The file uses `.v` for raw values; `rv(k)` is the computed value, `raw(k)` the raw text.

## Known gaps
- Not possible in a browser-only app: .xls binary, .xlsm, .xlsb, .xlam, .xltm, .xlt, .xps, .mht, Strict Open XML (the `.xls` offered is an HTML table Excel can open).
- Sparklines are text formulas (SPARKLINE/SPARKCOL/WINLOSS), shown as errors in Excel/LibreOffice.
- Translate/Thesaurus open external sites; Spelling uses browser spell-check in the formula bar only.
- Row/column insert/delete and sort do not rewrite formulas. Dates import as serial numbers. Formulas cannot use `1/0`-style errors beyond JS semantics.
- Excel ribbon items not built: Themes/Colors/Fonts/Effects, Background, Breaks, Print Titles, Arrange (Bring Forward, Align, Group, Rotate), Selection Pane, Watch Window, Calculation Options, Custom Views, Split, New Window, Switch Windows, Macros, Add-ins, Data Validation, Check-box/What-If tools, Subtotal.
- Word/Paint/Notes were not changed in this round beyond the fix above.
