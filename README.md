<h1>
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="public/logo-dark.svg" />
    <img src="public/logo.svg" alt="Intact" height="56" />
  </picture>
</h1>

**Intact — private .docx editor.** Change what you need. Keep the rest intact.

Fix a Word document in seconds, privately: no install, no account, no upload — and nothing you
didn't touch changes. Open a `.docx`, edit its text, and download it again. Or start a new one, from
a blank page, and write it here: it is saved as Word writes it, so it opens in Word as it looks
here. The file is unzipped, parsed and rebuilt entirely in your browser.

Where this is heading, and what it deliberately won't do, is in [VISION.md](VISION.md).

Built with [Vue 3](https://vuejs.org/), [Vue Router](https://router.vuejs.org/) (hash mode, so it
works on GitHub Pages without rewrites), [Vite](https://vite.dev/) and [JSZip](https://stuk.github.io/jszip/).

### Password-protected files, older `.doc` files and rights management

Intact opens `.docx` files (and `.dotx`, `.docm`, `.dotm`), including **files encrypted with a
password** (Word's File > Info > Protect Document > Encrypt with Password, or LibreOffice's Save with
password): it asks for the password, decrypts the file in the browser and downloads it encrypted with
the same password again (see Features). The password is checked on this device and never sent or
stored. It doesn't open these, and says so when you try, instead of failing:

- **Word 97–2003 `.doc` files**, the older binary format many archives still hold. Open the file in
  Microsoft Word (or LibreOffice, or Word for the web) and save it as a Word Document (`.docx`):
  File > Save As > Word Document. Then open the `.docx` here.
- **Files protected by rights management** (a sensitivity label that encrypts, or IRM). Only the
  label's owner or the organisation's tools can remove the protection.
- **Files encrypted the oldest ways** (RC4, from Word 97–2003's encryption kept in a `.docx`, or for a
  certificate rather than a password). Open them in Word, which re-encrypts as Word 2010 and later do.

A password that only restricts editing (Restrict Editing) is fine too: such files open here,
read-only or for filling in, as Word opens them. Compare documents (Review > Compare) says the same of
a version it can't read, and asks for the password of one encrypted with a password.

## Features

- Open `.docx` (file picker, by mouse or keyboard, or by dropping it anywhere on the start page; a document dropped on the editor's page opens in a new tab, and a file dropped where the editor can't take it is refused, never opened by the browser in the app's place) or start a blank document; Word templates (`.dotx`) and
  files with macros (`.docm`, `.dotm`) open too and are saved as what they are (the header shows the
  extension). Macros are never run: a bar says the file contains them and that they are kept, as
  they are, and "Before you release" lists them
- Start a new document from a template, as Word does when a template is opened to write with: a
  template opens with a bar that says downloading saves the template itself, and "Start a new
  document" makes it download as a `.docx` instead (a `.dotm` as a `.docm`, macros kept), only the
  main part's content type changed, so the quote, letter or invoice written from it never replaces
  the template. "Edit the template instead" goes back. The template file opened is never changed
- The interface in English or Indonesian (Bahasa Indonesia): the browser's language chooses it, and
  the start page's Language menu changes it (remembered in this browser). The start page, the
  header, the ribbon and its tools, the status bar, Page setup, Page number and its Format page
  numbers, Fill in and the Paragraph dialog's labels are translated, with Word's own Indonesian
  names (Beranda, Sisipkan, Tata Letak, Referensi, Tinjau; Pengaturan halaman, Lacak perubahan);
  the rest is in English for now (see Limitations). Only the interface changes: a document's text,
  styles and fields, and the file saved, are the same in both. The strings are looked up by their
  English text (`src/i18n.js`, `src/locales/id.js`), loaded only when Indonesian is chosen. A
  blank document started in Indonesian is in Indonesian (`w:lang` id-ID), named "Tanpa judul", so
  the browser checks its spelling in Indonesian where it can
- Indonesian government documents (naskah dinas), first on the start page when the interface is in
  Indonesian: a surat dinas (an office's letter), a nota dinas (an internal memo between officials),
  a surat undangan (an invitation to a meeting or event: day and date, time, place, agenda) and a
  berita acara (an official record of an event or a handover, signed by both parties), in
  Indonesian, laid out as the rules for official documents (tata naskah dinas) lay them out: on A4,
  in Arial 12 pt, single-spaced, with 3 cm on the left and 2 cm elsewhere; the office's letterhead
  (kop: its name in capitals, its unit, address and contacts, over a thick-and-thin line, as Word's
  border writes it), "Nomor", "Sifat", "Lampiran" and "Hal" lined up at a tab stop with their colons,
  the signatory's position, name and NIP at the right, "Tembusan", and in the berita acara both
  parties' signatures side by side in a table without borders. Then the others an office writes
  most: a keputusan (a decree: its title, "Menimbang" and "Mengingat" with their points, "MEMUTUSKAN",
  "Menetapkan" and its articles, KESATU, KEDUA, KETIGA, in a table without borders, as decrees lay
  them out; where and when it was set; to whom copies go; and its lampiran, the attachment, in a
  section of its own on a new page, to turn landscape in Page setup for a wide table), a surat tugas
  (an assignment: its basis, the employee's name, NIP, rank and position lined up, and what to do), a
  surat edaran (a circular: background, purpose, scope, basis, what is asked, closing), a notula
  rapat (minutes: when, where, who led and who took notes, what was discussed and decided, and who
  does what next, signed by both), a laporan kegiatan (an activity report: introduction, what was
  done, results, conclusions and recommendations), a telaahan staf (a staff study for a superior:
  the problem, assumptions, facts, analysis, conclusion and recommendation) and a surat kuasa (a
  power of attorney, signed by whoever gives it and whoever receives it). Their points are Word's
  numbered lists (a., b.; 1., 2.; A., B., with 1., 2. under them), so a point added with Enter is the
  next letter or number; their pages are numbered from the second, at the top in the middle between
  dashes (- 2 -), as the rules for official documents number them, and the decree's attachment
  from 1 again. The deciding official and the decree's title are typed once and shown in capitals
  where a decree writes them so (Word's All caps). Their placeholders are in Indonesian
  ("[Nomor surat]", "[Nama jabatan]"), for Fill in. F4 and Folio are in Page setup, for offices
  that print on them
- Start from a template, on the start page: a letter, a memo, meeting minutes, a quote, an invoice,
  a CV, a report and a novel's manuscript, each a new Word document made in the browser (nothing is fetched), in Word's
  own styles (Title, Heading 1 and 2, Table Grid with a shaded heading row that repeats on each page,
  List Paragraph with Word's bullets, Footer with a page number), with lines such as "To:" and
  "Date:" lined up at a tab stop, and placeholders that say what to type ("[Recipient name]",
  "[Quote number]"). Fill in opens with them, so each is typed once for every place it appears. On
  Letter paper where the browser's language is of a country that uses it (the United States,
  Canada, Mexico, the Philippines and others in the Americas), else A4, with 2.54 cm margins. The CV
  keeps contact details in the text, where screening software reads them, and the report numbers
  its pages. The manuscript is in standard manuscript format, as agents and publishers ask for it:
  Times New Roman 12 pt, double-spaced, each paragraph's first line indented half an inch and no
  space between paragraphs (all Normal's, so text typed or pasted follows); a title page with the
  author's contact details, "about [Word count, rounded] words" and the title halfway down; a header
  on the other pages with "[Surname] / [Short title] /" and the page number (a different first
  page); each chapter a Heading 1 that starts a new page a third of the way down (and is listed,
  and can be moved, in the Headings pane); and a scene break as a centred #. Edits to a starter
  aren't tracked, as in a blank document
- Templates for teachers, as a group of their own on the start page: in English, a test (Part A's
  multiple choice in two columns between continuous section breaks, its questions and answer
  choices numbered 1. and A. by Word's own multilevel list, so a choice added with Enter is the next
  letter; Part B's questions with ruled lines to answer on), a worksheet (matching, fill in the
  blanks, a question to write about and a self-check), a rubric (criteria against four levels and a
  score, on a landscape page, its rows tall enough to write in), a lesson plan (objectives,
  materials, the lesson's stages, differentiation, assessment and reflection) and a letter to
  parents whose [Parent name], [Student name] and [Class] are a class list's columns for Mail
  merge, with a reply slip. When the interface is in Indonesian, in Indonesian (untuk guru), after
  the naskah dinas: a lembar soal (the school's letterhead, the student's details, the general
  instructions, pilihan ganda A–D in two columns, isian and uraian, each part numbered from 1), a
  kisi-kisi soal (a landscape table from capaian pembelajaran to nomor soal, signed by the head of
  school and the teacher), a modul ajar in Kurikulum Merdeka's parts (informasi umum, komponen
  inti, lampiran), a rubrik penilaian and a surat untuk orang tua for Mail merge, on A4 in Times New
  Roman 12 pt with 2 cm margins (the letter as an official letter, in Arial). Answer lines are a
  table whose cells have a line under them, so Word and Intact draw them alike
- Mail merge from a list, as Word's Mailings tab does it (Download ▾ > Mail merge…, or From a list…
  in Fill in): choose a CSV saved from Excel or Google Sheets (its first row names the columns;
  commas, semicolons or tabs between fields, quoted fields with commas and line breaks, UTF-8 or
  Excel's Windows-1252), and each row fills the document's placeholders. Each placeholder is matched
  to the column of its name (whatever the case, spaces, underscores or accents) and can be given
  another, or left as it is; a row at a time is previewed. It makes either one document, a copy for
  each row, each starting on a new page (`w:pageBreakBefore`), or a .zip of a document for each
  row, named by a column. The document as it would be downloaded is merged, in its XML: a value goes
  into the run where its placeholder starts (so it takes its formatting, also across runs), with
  Word's line breaks for an address's lines; in one document, the ids that must be unique
  (bookmarks, tracked changes, drawings) are moved up for each copy, bookmarks' names made its own,
  and comments stay with the first copy. Up to 1,000 rows. The list is read on this device, and
  the document itself isn't changed
- Right-to-left text, as Word shows it: paragraphs (`w:bidi`, directly or from their style) are laid
  out and edited right to left; Word's left and right alignment in them are logical (start and end),
  as Word and LibreOffice read them, and the alignment buttons align as shown; runs marked
  right-to-left (`w:rtl`) and right-to-left tables (`w:bidiVisual`, first column at the right) too.
  Direction embeddings and overrides (`w:dir`, `w:bdo`), smart tags and inline custom XML are kept
  around their text, also when it is edited, with their properties
- Line numbers in the margin, as Word prints them for sections with line numbering
  (`w:lnNumType`): every line of the text but tables, notes, headers, footers, text boxes and
  paragraphs that suppress them, every n-th number (count by), from the starting number, restarting
  each section or never, or each page, where Word last began one (its `w:lastRenderedPageBreak`
  marks, page breaks and "page break before"), so the numbers match Word's printed pages while the
  edits don't move its page breaks. Drawn over the page, never in it, and kept up to date as you type
- Edit text in paragraphs, headings, lists and table cells
- Bold, italic, underline, strikethrough, superscript and subscript (each turns the other off, as in
  Word), and highlight in Word's 15 highlight colours (the last colour chosen is remembered; existing
  highlight shows and can be removed); paragraph styles (Normal, Title, Heading 1–3, and the document's own: headings 4–9 it defines, the
  styles in Word's style gallery that aren't hidden, such as Quote and Caption, and any style its text
  uses; a paragraph in another style shows that style's name); alignment
- The Numbering library (Numbering ▾), as Word's Numbering Library and Multilevel List: Indonesia's
  legal numbering as its regulations write it (UU No. 12/2011, Lampiran II): **Pasal 1 → (1) → a. →
  1. → a) → 1)**, a Pasal centred on its own line with its ayat restarting in each one; **BAB I →
  Pasal 1 → (1) → a.**, chapters in Roman numerals with the Pasal numbered on through them, as Word
  does with `w:lvlRestart w:val="0"`; **1. → 1.1. → 1.1.1.** for a report's numbered headings and
  paragraphs; and Word's own 1. → a. → i. The selected paragraphs become a new list of that kind, each
  keeping its level, and Tab and Shift+Tab (or Increase and Decrease Indent) move a paragraph between
  Pasal, ayat and huruf. Written as Word writes a multilevel list (`w:abstractNum` with nine levels,
  `w:multiLevelType` multilevel, each level's text, format, indent and suffix; a centred level's
  `w:jc` in its own properties), so Word and LibreOffice number it the same; lists in files that
  restart a level only after a given one (`w:lvlRestart`) are numbered so here too
- Bulleted and numbered lists that behave like Word's: the Bullets and Numbering buttons start and
  end lists; Enter on an empty item (or Backspace at its start) ends the list; Tab and Shift+Tab at
  the start of an item change its level
- Hidden text, for a test's answer key, as Word's Font > Hidden: the Hidden text button or
  Ctrl+Shift+H hides the selected text (`w:vanish`), or shows it again when all of it is hidden. It is
  shown faintly here with Word's dotted line under it, and left out when printing (and in the PDF),
  as Word leaves it out, unless Print hidden text is chosen in the PDF dialog: the test prints
  without its answers, its key with them
- Read Aloud, as Word's Review > Read Aloud (Ctrl+Alt+Space): from the caret's paragraph, or the
  selected text, paragraph by paragraph, in the text's language, the paragraph and the word being
  read highlighted (with the CSS Custom Highlight API, so the page is never changed), with Pause,
  Resume, and Esc or Stop. Only voices on the device are used (the Web Speech API's local voices):
  online voices, which send the text to a server to be spoken, never are, and where there is no
  voice on the device for the text's language, it says so. Hidden text (a test's answers) isn't read
- Tabs where Word puts them: a tab alone in its run (as Word writes a table of contents' tabs and
  Insert answer line writes one) or alone in its paragraph ends at its stop as Word lines it up,
  left, right, centred or at the decimal point of the text after it, measured on the page and laid
  out again as the text changes; its leader (dots, hyphens, a line) or its underline is drawn across
  it, and prints (`src/editor/tabLayout.js`). Only the page's styles change, never the text saved
- Insert > Answer line: a line from the caret to the paragraph's right edge, to write a name or an
  answer on ("Name: ______"), as teachers make one in Word: an underlined tab (a run of its own) to a
  right tab stop at the margin, added to the paragraph if it hasn't one there; what is typed after
  it isn't underlined
- Insert > Line, Arrow and Text box, as Word's Insert > Shapes writes them: a line or an arrow (a
  DrawingML shape, `wps:wsp`, `prst` line or straightConnector1 with an arrowhead) in front of the
  text where the caret is, anchored to its paragraph, selected to be dragged into place, resized
  and turned with the picture tools (Layout Options without a picture's crop, effects and styles),
  to label a diagram or a picture; a text box (5 × 1.5 cm, with a thin border) in line with the
  text, to type in. Lines, arrows, rectangles and ellipses already in a document are drawn as Word
  draws them (their outline, fill and arrowheads) and laid out as pictures are, instead of a
  placeholder
- Lists started by typing, as Word's AutoFormat: "1. ", "1) " or "(1) ", "a. ", "A) ", "(i) " and the
  like, "* " or "- " typed at the start of a paragraph make it a numbered list in that style (starting
  from the number typed: "3. " starts at 3) or a bulleted list (Word's bullet, or a dash), written as
  Word writes such a list. A number that follows on from a numbered list just above, in its style,
  continues it. Ctrl+Z straight after gives back the text as typed; the AutoFormat button in the
  status bar turns this off, with smart quotes. Tracked while tracking
- Line and paragraph spacing and indents, as Word's Line and Paragraph Spacing menu and Paragraph
  dialog: line spacing (1.0, 1.15, 1.5, 2.0, 2.5, 3.0; Ctrl+1, Ctrl+5, Ctrl+2), add or remove space
  before (Ctrl+0) and after the paragraph, and in the Paragraph dialog left and right indents, a
  first-line or hanging indent (as Word's dialog shows one, Left is where the first line starts:
  "Left 0, Hanging 1.27 cm" is the usual hanging indent, `w:left="720" w:hanging="720"`, and a list
  item's Left is where its number is), space before and after, and line spacing single, 1.5 lines, double,
  at least, exactly or multiple. The dialog shows what the paragraph has (from its style when not its
  own) and writes only what is changed in it, into the paragraph's own `w:spacing` and `w:ind`, with
  everything else in its properties kept (and Word's "lines" values, which would win, taken off what
  is changed). Increase and Decrease Indent (Ctrl+M, Ctrl+Shift+M) move a paragraph's left indent to
  the next half inch, or a list item's level. Tracked while tracking (`w:pPrChange`)
- Lengths in centimetres or inches, as Word shows them: inches for a document on North American
  paper (Letter, Legal, Executive, Tabloid), centimetres for the others, in Page setup, the
  Paragraph dialog (indents and tab stops) and a picture's Layout Options. The Units choice in Page
  setup and the Paragraph dialog switches them, converting what the dialog shows (a length left as
  it was still counts as unchanged, so it isn't rewritten), and is remembered in this browser for
  every document (`src/editor/units.js`)
- Tab stops, as Word's Tabs dialog, in the Paragraph dialog: the paragraph's stops (its styles' and
  list level's, and its own) are listed with their alignment (left, centre, right, decimal, bar) and
  leader (dots, dashes, line); Set adds one at a position in centimetres or inches (replacing one there),
  × removes one and Clear all takes them all off. Only the paragraph's own `w:tabs` is written: a
  stop for each one it doesn't inherit, and a clear (`w:val="clear"`) for a style's stop taken off,
  as Word writes them; tracked while tracking. On the page, tabs go to Word's default stops
  (`w:defaultTabStop`, half an inch unless the file says otherwise), or, where a paragraph's first
  stop is a left one, to that stop, so a letter's "Nomor:" and "Perihal:" lines line up as in Word
- Print the document (Ctrl+P, or the browser's Print): only the page is printed, not the editor
  around it, on the document's paper size and in its margins, with a new page at each page break,
  "page break before" and section break; tracked changes and highlights print as they show.
  Headers and footers print on every page, in the margins as Word places them (the text starts
  below a header taller than its margin): one that is the same on every page repeats as it shows,
  pictures and tables too; page numbers (`PAGE`, `NUMPAGES`, also as simple fields) are numbered on
  each page, "Page 2 of 3", and a different first page or different even pages get their own, in
  CSS page-margin boxes (Chrome, Edge), which hold a header's or footer's text, laid out by its
  alignment and Word's centre and right tab stops, in its font and colour. A document of several
  sections prints as Word prints it: each section on its own paper (a portrait decree, then its
  landscape attachment), with its own margins, header and footer, and page numbers in its own format
  and from its own start (i, ii for a report's front matter, then 1, 2), or numbered on from the
  section before; a section that starts on the same page (continuous) shares that page. Each section's
  different first page is followed (nothing on the attachment's first page, its others numbered)
  where the section starts its page numbers at a number, as an attachment does. A page number
  written without a result (as some software writes one, a whole field in one run) shows # as the
  others do, prints numbered, and is saved as it was
- Save as PDF from Download's menu (▾): a short dialog says that the browser makes the PDF from
  its print window, on this device (Save as PDF, or Microsoft Print to PDF, as the destination), what
  it will be called (the file's name: the page's title is the file's name while printing), about
  how many pages it has, and what it will show that you may not want sent: tracked changes and
  highlights, which print as they show (accept or reject the changes first for a clean copy).
  Comments and the marks drawn over the page don't print
- About how many pages, as Word would print it: the status bar says "≈ 2 pages", and a dashed line
  across the page, numbered in the margin, shows about where each page begins. Worked out from the
  paper size and margins (and a header or footer taller than its margin), line by line as Word fills
  a page, with Word's widow and orphan control, headings kept with the text after them, table rows
  kept whole, and a new page at each page break, "page break before" and section break that starts
  one; kept up to date as you type. Zoomed in or out, where the browser lays the text out at that
  size, the pages are worked out on a hidden copy at 100%, as printed, so the count and the breaks
  don't change with the zoom. Drawn over the page, never in it; clicking the count shows or
  hides the lines (remembered in this browser). The count matches the printout; Word may break a
  line or two differently with its own fonts
- Tab and Shift+Tab move between table cells; Tab in the last cell adds a row, as in Word (a copy of
  the last row's cells, empty: widths, merged columns, shading and paragraph formatting, but not a
  header row's repeat or vertical merges), tracked while tracking changes
- Insert tables (toolbar grid, or type any size up to Word's 63 columns) at the caret, as Word does:
  the paragraph is split there, the table spans the page in Word's Table Grid style (added to
  `styles.xml` if the file lacks it), and while tracking its rows are tracked insertions
- Change a table's layout from the Table layout menu, as on Word's Layout tab: insert rows above or
  below (as many as are selected; copies of the row, carrying vertical merges and header repeat on
  where they fall inside them), insert columns left or right (cells spanning the place are widened,
  the grid and widths follow, and a table that fitted the page narrows to fit), delete rows, columns
  or the whole table. Rows and columns are tracked while tracking, as Word tracks them: a row
  inserted or deleted is marked on the row (`trPr`), a column on each of its cells (`cellIns`,
  `cellDel`), deleted text struck through, and each row or column is accepted or rejected as one
  change (rejecting an inserted column, or accepting a deleted one, takes its grid column out and
  narrows cells spanning it). Deleting a row or column you inserted removes it; rows and columns
  holding someone's tracked change aren't deleted untracked. Rows and cells that are content
  controls change too: a row added in a repeating section is a new item of it (with its own ids,
  as Word's + adds one), a row added in a plain field stays in it, a field left empty goes, and new
  cells next to a field cell are ordinary cells; locked fields and fields linked to data stay as
  they are. Only the rows and cells a change touches count, so a locked row (a form's total, say)
  or a cell linked to data doesn't stop changes elsewhere: rows are still inserted and deleted
  around it, and columns beside a locked cell. A change that would alter such a field (a column
  adding a cell to a locked row, deleting the column of a linked cell) is refused, and the status
  bar names the field
- Sort a table, as Word's Table Layout > Sort (Table layout menu, Sort…): the rows below the heading
  row (a first row it guesses is one, as Word does, and can be told it isn't) by one column, chosen by
  its heading, as text (in the order of the document's language, numbers in it as numbers: "Item 2"
  before "Item 10"), numbers (as the document writes them: "Rp 1.250.000", "(2,5)", "12%") or dates
  ("17/08/2026" day first, or month first in American English; "17 Agustus 2026"; "2026-08-17"),
  ascending or descending, empty cells last, rows that sort the same in the order they were. Rows are
  moved whole, with their cells' properties, as one undoable step. A column of numbers is sorted as
  numbers without asking. Not while tracking changes (Word records a sort as moved rows; here it
  says so instead), and not tables with cells merged down or with tracked changes to their rows or
  cells
- Table formulas, as Word's Table Layout > Formula (Table layout menu, Formula…): Word's `=` field in
  the cell at the caret, such as `=SUM(ABOVE)` (the dialog offers it, or `=SUM(LEFT)` when there are
  numbers to the left and none above, as Word's does), `=AVERAGE(LEFT)`, `=B2*C2`, `=SUM(B2:B5)*11%`,
  `=ROUND(A1/3;2)`, with SUM, AVERAGE, MIN, MAX, COUNT, PRODUCT, ABS, ROUND and INT (Paste function
  adds one), + − × ÷ ^ and brackets, cells as A1, B2 (Word's columns and rows of the table's grid) and
  ABOVE, LEFT, BELOW and RIGHT, which take the numbers beside the cell up to the first cell that
  isn't one (a heading, or an empty cell), as Word's do. The result shows as it is worked out, and a
  number format (`\# "#,##0.00"`, "Rp #.##0", "0%") writes it with the document's own separators:
  1.234.567,89 in an Indonesian document, 1,234,567.89 in an English one (and arguments separated by
  ";" where "," is the decimal sign, as Word's list separator is). Written as Word writes the field
  (` =SUM(ABOVE) \# "Rp #.##0" ` with its result), so Word updates it too; Update fields (F9) works
  every formula out again here, in the order of the text, after a number in the table changes. A
  formula Word wouldn't understand is refused in the dialog; one in a file shows Word's
  "!Syntax Error" when updated (`src/docx/formulas.js`)
- Shading and borders for table cells, from the Table layout menu (Shading and borders…), as Word's
  Shading and Borders buttons: a fill colour for the selected cells (the theme's colours, lightened
  80% and 40% as Word offers them, the colours the text uses, your colours, any colour by its code,
  or No colour) and All borders, Outside borders (round the selected cells as a block), Bottom
  border (a heading row or a total) or No border (which hides the table style's too). Written in
  each cell's properties (`w:shd`, `w:tcBorders`), one undoable step, the selection kept; tracked
  while tracking as a change of the cells' formatting (`w:tcPrChange`). A cell's own borders in the
  file are drawn as Word draws them
- Merge and split cells from the same menu, as Word does: Merge cells joins the selected cells
  (grown to whole cells where a merged cell sticks out) into one spanning them across (`gridSpan`)
  and down (`vMerge`), their content one after another in it; Split cells… divides a cell into
  columns of equal width (adding grid columns that the other rows span where needed) and rows (the
  row's other cells merged down), and a merged cell splits back into the cells it spans. Cells that
  are text fields merge too (the merged cell is the first one's field; a placeholder isn't moved);
  locked fields, fields linked to data and other kinds of field don't, and the status bar says
  why. While tracking, merging and splitting are tracked as a Word document records them: cells
  merged across a row with `w:hMerge`, each cell's old properties kept in a tracked property change
  (`tcPrChange`); cells merged down a column (or split back) with `w:cellMerge`; the text moved
  shown inserted in the merged cell and deleted where it was; cells split into columns as inserted
  cells, the old grid kept (`tblGridChange`) so rejecting gives the widths back; into rows as
  inserted rows. Each is one change to accept or reject; accepting a merge across a row leaves one
  cell spanning the columns (`gridSpan`), as Word writes a merge. Cells an older Word file merged
  with `w:hMerge` show merged
- Insert pictures (PNG, JPEG, GIF or BMP) from the toolbar, by pasting a picture on its own (such as
  a screenshot) or by dropping picture files on the page: in line with the text at the caret, at
  their size or narrowed to fit between the margins, tracked while tracking. The picture's file,
  relationship and content type are added when saving, only if it is still in the document
- A logo in the header or footer (a letterhead): with the caret in a header or footer, Insert
  picture (or pasting or dropping one) puts it there, in line at the caret, with the picture file's
  relationship in that part's own relationships (`header1.xml.rels`), as Word writes it; the text's
  part is not changed. Undo takes it out, and its file and relationship with it
- Lay out pictures, as Word does: a click selects a picture (a frame with handles); corners resize
  it in proportion, sides stretch it; a floating picture is dragged to a new place (measured from
  the margin and its paragraph, so it moves with the text). Layout Options (the button beside a
  selected picture) set its text wrapping (in line, square, tight, through, top and bottom, behind
  or in front of text), its position (left, centred or right of the margins or at a distance;
  below its paragraph or at the top, middle or bottom of the margins), its exact size in
  centimetres (keeping proportions, or back to its own size) and its alt text. The picture's
  DrawingML is changed where it must be (`wp:inline` ↔ `wp:anchor`, extents, position, wrap,
  `docPr/@descr`) and the rest kept; the page shows the layout as near as a web page can. Resizing
  and moving aren't tracked; the review lists them. Charts, shapes and other drawings stay as they are
- Crop, rotate and add effects to pictures, as Word's Picture Format tab does: crop by numbers or
  with the handles (`a:srcRect`; what shows keeps its scale, the frame follows), rotate with the
  round handle or by 90° and flip (`a:xfrm` `rot`, `flipH`, `flipV`), and set a border, shadow,
  glow, soft edges, reflection, greyscale, brightness and contrast (`a:ln`, `a:effectLst`,
  `a:grayscl`, `a:lum`); effects Intact doesn't set are kept. Tight and Through wrapping follow the
  picture's outline, as in Word: the outline of what isn't transparent is written as the wrap
  polygon, and the page wraps the text around it (CSS `shape-outside`), as it does for a polygon
  Word drew
- Picture styles, shapes, 3-D and artistic effects, as on Word's Picture Format tab: a gallery of
  Word's picture styles (each written out as Word writes it: shape, border, shadow, glow, soft
  edges, reflection and 3-D format; Word keeps no reference to a style), a shape to cut the picture
  to (`a:prstGeom`), Word's 3-D Format (a bevel with its width and height, a material, a lighting
  rig and its angle) and 3-D Rotation (a preset, or X, Y and Z rotation and a perspective;
  `a:sp3d`, `a:scene3d`, `a:camera/a:rot`, `a:lightRig`), and all 22 of Word's artistic effects
  (Marker, Pencil Sketch, Paint Strokes, Watercolor Sponge, Mosaic Bubbles, Glass, Plastic Wrap and
  the rest), each with an amount, drawn into the picture in the browser and saved as a new picture,
  as Word saves them, with the original kept beside it in an extension of the picture (`a:extLst`),
  so the effect can be changed or taken off here (Word shows the drawn picture; Word's own record of
  an effect needs the original in Microsoft's HD Photo format, which a browser can't write). A bevel
  is lit per pixel, as a 3-D renderer lights it: its profile gives the surface's slope near the
  edge, and the light rig (from its direction and angle) and the material give light and shade by
  Lambert's and Phong's laws (`src/docx/lighting.js`); a 3-D rotation is shown with the file's
  angles and field of view. The page shows every effect with the file's own values: a shadow's distance,
  direction, blur, colour and transparency, a glow's size and colour, soft edges' width, a
  reflection's distance and fade, a bevel's width, Word's brightness and contrast (tones moved and
  stretched about the middle, not CSS's multiply) and greyscale, the picture's shape; a turned
  picture in line with the text takes the room of its turned outline, on the page and in the file
  (`wp:effectExtent`), as in Word
- Add comments (toolbar button or Ctrl+Alt+M) on the selected text, or the word at the caret, with
  your name (asked once, shared with Track Changes): they appear in the comments panel and are
  written to `comments.xml` (created if the document has none) next to the comments already there.
  Every comment can be replied to, edited, resolved (or reopened) and deleted in the panel, as in
  Word: a reply is written as a real reply (linked to its comment in `commentsExtended.xml`, which is
  created, with paragraph ids, if the file has none), resolving sets Word's "done" flag, and Undo
  covers all of it. Deleting a comment's text deletes it too. Comments anchored in a header, footer
  or note work the same way (the panel says where they are): a reply is written into that part next
  to its comment, inside its range, and a deleted comment's marks go from it. Previous and Next
  comment (the Review tab) go through the comments in the text from the caret, or from the comment
  shown, round to the other end: each selects the text its comment is on and shows its thread in the
  panel ("Comment 2 of 7, by Editor"); those in headers, footers and notes are in the panel only
- Fill in templates and forms: text fields (content controls) are editable, with placeholder text
  replaced as you type; check boxes tick on click; drop-down lists and date pickers offer their
  choices. From the keyboard, Alt+↓ (or Space on a selected field) ticks a box or opens a list or
  date picker for the field at the caret. Fields around whole table cells or rows work too. Locked
  fields and fields linked to the document's data stay read-only and say why
- Fill in legacy forms, as Word's older form fields (Developer → Legacy Tools) that school, medical
  and club forms still use: check boxes (`FORMCHECKBOX`) are shown as Word draws them and tick on a
  click, drop-down lists (`FORMDROPDOWN`) show their chosen entry and offer the others, and text
  fields (`FORMTEXT`) are typed in, Word's blank (five spaces) selected on a click so typing replaces
  it. Ticking or choosing changes only the field's settings (`w:ffData`: `w:checked`, `w:result`),
  as Word does; the box or entry is drawn by the page, never added to the text. A form protected for
  filling in (Restrict Editing) is filled in as in Word: only its fields take typing, Tab and
  Shift+Tab (and Enter in a text field) go from field to field, and Space, Enter or Alt+↓ ticks a box
  or opens a list. Fields turned off in Word stay as they are and say so
- Sign a form (the toolbar's Draw signature): draw a signature with a finger, pen or mouse, in black
  or blue ink, or type your name in a handwriting font, and it goes in at the caret, or in place of
  the selected "______" line, as a picture in line with the text: cropped to the ink, with a
  transparent background, about 6 cm across, with "Signature" as its alt text. Drawn in this page
  only: it isn't stored or sent anywhere. Tracked while tracking, undone with Ctrl+Z
- Insert a check box or today's date (the toolbar's Insert button), as Word's Developer tab and
  Insert Date and Time do: the check box is Word's check box content control (`w14:checkbox`, ☐ and
  ☒ in MS Gothic), ticked with a click here and in Word; the date goes at the caret in one of a few
  forms in the text's language ("19 November 2026", "Thursday, 19 November 2026", "19/11/2026",
  "2026-11-19"): as text, so it stays the day the form was filled in, or, with Update automatically
  ticked (as in Word's Date and Time dialog, and remembered in this browser), as Word's DATE field
  in the same format (`DATE \@ "d MMMM yyyy"`, its picture worked out from how the date is written
  in that language), showing today's date, which Word changes to the day it opens or prints the
  file: a quote's or letter's date that is always current. Both are tracked while tracking
- Edit footnotes and endnotes in place: they are listed below the text (as in Word's draft view),
  numbered in the order the text refers to them, and edited like the text: typing, formatting,
  Undo, Find and Replace and Track Changes all work in them. Clicking a note's number in the text
  goes to the note, and its number in the note goes back. A note can't be deleted with the text
  around it (nor its number); deleting its reference in the text hides it and renumbers the others,
  as in Word. Only notes that changed are rewritten in `footnotes.xml` / `endnotes.xml`; untouched,
  the parts come back byte for byte. Comments in notes stay with the comments panel
- Insert footnotes and endnotes (the toolbar, or Ctrl+Alt+F and Ctrl+Alt+D, as in Word): a
  reference at the caret (after the selection) and its note below the text with the caret in it,
  numbered in the text's order among the document's own notes (endnotes i, ii, iii), in Word's
  Footnote (Endnote) Text and Reference styles (added to `styles.xml` if the file lacks them). A
  document without such notes gets `footnotes.xml` or `endnotes.xml` (with Word's separators), its
  relationship and content type. A note is saved only while its reference is in the text: Undo, or
  deleting the reference, takes a new one away. The reference is tracked while tracking, and so is
  the note's text
- References, from the Insert menu, as on Word's References tab:
  - Table of contents: "Contents" in Word's TOC Heading style and Word's `TOC` field over headings
    1–3 (\o "1-3" \h \z \u), built at once from the headings as Update fields builds one, linked to
    them; Word fills in the page numbers when it opens the file (`w:updateFields`)
  - Caption…: "Figure 1: text" (or Table, Equation) in Word's Caption style, numbered by a `SEQ`
    field among the others, below a picture or paragraph or above a table, as Word places them
  - Cross-reference…: to a heading (its text) or a caption (its label and number, or all of it), or
    "above"/"below", as a `REF` field to a `_Ref` bookmark added on it, as Word adds one
  - Equation…: typed as in LaTeX (fractions, scripts, roots, sums and integrals with limits,
    brackets that grow, Greek letters, the usual symbols, functions upright; matrices in their
    brackets, `\begin{pmatrix} a & b \\ c & d \end{pmatrix}` and bmatrix, Bmatrix, vmatrix, Vmatrix
    and matrix; cases; working lined up at its = signs, `\begin{aligned} x &= 1 \\ &= 2
    \end{aligned}`; accents, `\hat`, `\bar`, `\vec`, `\dot`, `\tilde`, and `\overline`,
    `\underline`), previewed, and written as Word's own equations (Office Math, `m:oMath`; on its own
    line, `m:oMathPara`, in an empty paragraph) that Word edits as its own: a matrix as `m:m` in a
    delimiter (`m:d`), lines of working as an equation array (`m:eqArr`) whose alignment points are
    the `&`s in its text, as Word writes them (`src/editor/mathInput.js`). An equation array, typed
    here or written by Word, is shown lined up at those points
  - Citation…: a source the document has, or a new one (a book, a journal article or a website), in
    APA (7th edition) or MLA (9th edition) style, with the page cited if wanted, as Word's `CITATION`
    field. Sources are kept as Word keeps them, in its Source Manager's part (`customXml/itemN.xml`,
    `b:Sources`, with its properties part and relationships), so Word's Manage Sources lists them;
    the part is written only when a source or the style is changed, else it goes back byte for byte
  - Bibliography: "References" (APA) or "Works Cited" (MLA) in Heading 1, and Word's `BIBLIOGRAPHY`
    field listing the document's sources in that style, one paragraph each in Word's Bibliography
    style with a half-inch hanging indent (`src/docx/references.js`). A new source or another
    style brings the citations and bibliographies up to date at once
- Headings (the status bar's button), as Word's Navigation pane: the document's headings by level,
  the one the caret is under marked; a click goes to a heading. Kept up to date as you type, and
  shown or hidden as last time in this browser. A heading moves with everything under it (a chapter
  with its scenes), as in Word's pane: dragged onto another heading it goes before that one, and ▲ ▼
  beside it (or Alt+Shift+↑ ↓ with it focused) move it past the heading of its level before or
  after it, under the same heading (a scene stays in its chapter). The paragraphs, tables and the
  like are moved as they are, so the file keeps them exactly; list and note numbers and comments
  follow, and Ctrl+Z puts it back
- Update fields (the toolbar, or F9, as Word's for the whole text): caption numbers (`SEQ`, with
  Word's number formats, \c repeat, \r reset and \s per chapter) are renumbered in the text's
  order; cross-references (`REF`) show the bookmarked text again, or its paragraph's number
  (\r, \n, \w) and "above"/"below" (\p); note numbers in cross-references (`NOTEREF`) follow the
  notes; citations (`CITATION`) and bibliographies (`BIBLIOGRAPHY`, in the text) from the
  document's sources, in its citation style; and tables of contents (`TOC`, in Word's content control or in the text) are rebuilt from
  the headings (\o outline levels, \u paragraphs' own levels, \t styles), as Word builds one:
  each entry in the table's own formatting for its level (Word's TOC 1–9 styles where it has none),
  linked to its heading with \h, through Word's `_Toc` bookmark (added to a heading without one).
  Page numbers are Word's: an entry keeps the page number it had (a new one takes the number above
  it), and the file asks Word to update fields when it opens it (`w:updateFields`). One undoable
  step; the status bar says what was updated, and the review says the fields were updated here
- Page numbers, as Word's Insert Page Number: Word's page number field (`PAGE`) at the caret in a
  header or footer. From the text: into the footer if it is empty (centred), else the caret goes to
  the footer to say where, so a footer's design stays the document's; a document without a footer
  gets one for all pages with the number centred, in Word's Footer style: `footer1.xml`, its
  relationship, content type and the first section's reference, written only while it is in the
  page (Undo takes it away). Tracked while tracking. Page number's menu (▾) puts it where Word's
  Top of Page and Bottom of Page do: at the left, centre or right of the header or footer (in a
  new header for all pages, `header1.xml`, if there is none: a student's "Surname 1" at the top
  right, with the name typed before the number); a header or footer with text keeps its design, and
  the caret goes there to say where. Edit header and Edit footer, in the same menu, go to the end of
  the header or footer shown, or give a document without one an empty one for all pages to type in
  (Undo takes it away)
- Format page numbers, as Word's (Page number ▾ > Format page numbers…): the number format (1, 2, 3;
  - 1 -; a, b, c; A, B, C; i, ii, iii; I, II, III) and where numbering starts (continue from the
  section before, or start at a number), for the section the caret is in, the whole document (the
  format for every section, the start for the first), or from the caret on: a new section (a
  section break, next page, before the caret's paragraph) numbered as chosen while the pages before
  keep theirs, for a report's front matter in i, ii, iii and its chapters from 1. Written in the
  section's `w:pgNumType` (`w:fmt`, `w:start`; Word's default, Arabic numerals that continue, as no
  format and no start), in the order Word's schema requires, with its other settings (chapter
  numbers) kept; a file's numbering is shown and goes back as it was when untouched. Each section
  break says how its section is numbered ("pages i, ii, iii from 1"), printing numbers the pages so
  (`counter(page, lower-roman)`, the first page set to the start), and Ctrl+Z puts it back
- Page setup for the whole document, as Word's: paper size (A4, Letter, Legal, A5, A3, Executive,
  and F4, 21.5 × 33 cm, and Folio, 8½ × 13 in, the office paper of Indonesia and much of South-East
  Asia), or a custom size (its width and height typed in, from 0.1 in to 22 in, as Word allows, and
  turned to the orientation chosen; a document on paper of no listed size shows its own),
  orientation and margins (in centimetres or inches, with Word's Normal, Narrow, Moderate and Wide), set in
  every section's `w:pgSz` and `w:pgMar`. Only what is changed in the dialog is written; the rest of
  each section (headers and footers, columns, gutters) is kept. The page and the printout follow,
  and Ctrl+Z puts it back
- Page setup for part of a document, as Word's "Apply to": the dialog shows the paper and margins
  of the section the caret is in, and applies them to the whole document, this section (where there
  are several; the default then, as in Word), this point forward or the selected text. "This point
  forward" puts a section break (next page) before the caret's paragraph; "Selected text" puts one
  before and one after the selected paragraphs, for a landscape page holding a wide table in a
  portrait document. A break is written as Word writes one: in the paragraph before, a copy of the
  section's properties (`w:sectPr` with `w:type="nextPage"`), so the text before it keeps its paper,
  margins, headers and footers (a paragraph is added beside a table to hold it). A section's own
  paper and margins go into its `w:sectPr`; the rest of it is kept. Where the document has several
  sections, each break is named on the page with its section's paper ("Section 2: A4 landscape ·
  section break (next page)"). One undoable step
- Columns, as Word's Layout > Columns (in Page setup): one, two or three equal columns, the space
  between them and a line between them, for the whole document, this section, this point forward or
  the selected text, written as Word writes them (`w:cols` with `w:num`, `w:space` and `w:sep`).
  Changing only the columns of part of a document puts continuous section breaks around it, as
  Word's Columns dialog does (a test paper's multiple choice in two columns under its heading, on
  the same page): the new section starts `w:type="continuous"`, and the section before keeps how
  it started. The page shows the columns at once, and a break shows as the kind of start the section
  after it has, as Word reads `w:type` ("section break (continuous)"). Columns of a document's own
  widths are kept unless changed here, and then become equal; a column narrower than half an inch is
  refused. One undoable step
- Headers and footers, as Word's Header & Footer options (in Page setup): a different first page
  (`w:titlePg` in every section) and different odd and even pages (`w:evenAndOddHeaders` in
  settings.xml). Switched on, the first-page (or even-page) header and footer the document has are
  put in the page, and those it hasn't are made, empty, as Word starts them (a new part each, in
  Word's Header or Footer style, with the first section's reference to it); the tabs above and below
  the page choose which one to edit ("First page", "Other pages", "Even pages"). A first-page or
  even-page one in the file that its settings don't use isn't in the page (nor found by Find) until
  switched on, and goes back as it was. Switched off, Word's settings say so and new ones aren't
  saved. One undoable step
- Modify Style, as Word's (the Paragraph style menu's last item, Modify style…, or the right-click
  menu): a paragraph style's font, size, bold and italic, colour (or Automatic), alignment, space
  before and after, line spacing and indents (left, right, and a first-line or hanging indent, as
  the Paragraph dialog shows them, in centimetres or inches: Normal with a 1.27 cm first-line
  indent, as a manuscript has it, or a hanging indent for a bibliography), shown as the style has them (from its base styles and the
  document's defaults), with a preview. Only what is changed is written, into the style's own
  `w:rPr` and `w:pPr` in styles.xml (off as off, `w:b w:val="0"`, so a base style's bold doesn't
  come through; a first-line indent written takes the place of a hanging one, and the other way
  round, with Word's character-unit indents that would win over it); every paragraph in the style, and in styles based on it, follows on the page and
  in Word. Paragraphs keep their own formatting. Ctrl+Z puts the style back
- A watermark, as Word's Design > Watermark, in Page setup: DRAFT, CONFIDENTIAL, DO NOT COPY,
  SAMPLE, URGENT, ASAP or your own text, diagonal or horizontal, in Word's semi-transparent silver.
  Written as Word writes one, as WordArt in every header behind the text (a VML text path named
  `PowerPlusWaterMarkObject`, centred between the margins), at the start of each header's first
  paragraph; the rest of the header is kept. A document without a header gets one for all pages,
  in Word's Header style (`header1.xml`, its relationship, content type and the first section's
  reference), saved only while it is in the page. Word's own text watermark shows, and can be
  changed or taken off (a header made only for it goes with it). It is drawn behind the text in the
  middle of each page the page count works out, and printed on every page; Ctrl+Z takes it back
- Edit headers and footers in place, in the page's margins, as in Word: the first page's, other
  pages', even pages' and each section's, chosen with tabs above and below the page (a click in one
  edits it; it is dimmed until then). Typing, formatting, Undo, Find and Replace and Track Changes
  work in them as in the text, and lists in them are numbered on their own. Only the parts that
  changed are rewritten (`header1.xml`, `footer2.xml`…); untouched, they come back byte for byte.
  Comments in them stay with the comments panel, and go with their text when it is deleted
- Edit the text in text boxes and shapes, shown where Word puts them: in line with the text or
  floating left, right or centred with the text wrapping around them, at their size, with their fill
  and outline (theme colours too), inner margins and vertical alignment. Both kinds Word writes are
  edited: DrawingML shapes (`wps:wsp`, Word 2010 and later) and legacy VML text boxes. The text is
  edited like the body (formatting, Undo, Track Changes); on save only the text changes, written into
  the shape and into its copy for older readers (`mc:Fallback`), whose tracked changes get ids of
  their own. A text box left alone is written back exactly as it was. Text boxes and pictures that
  share a run with other content, as LibreOffice writes them, are shown and edited the same way
  (written back as a run each, which means the same)
- See a designed document as Word lays it out: a section set in columns (`w:cols`, with a line
  between them when it has one) is shown in columns; paragraph borders (`w:pBdr`, with their style,
  width, colour and distance from the text) and shading (`w:shd`, theme colours too); drop caps
  (`w:framePr w:dropCap`) beside the lines after them; character spacing (`w:spacing`) and raised or
  lowered text (`w:position`); and theme colours on text with their tint or shade
  (`w:themeColor`, `w:themeShade`), worked out from the document's theme as Word does. None of this
  changes the file
- Fonts: the fonts embedded in a document (`word/fonts/*.odttf`, listed in `fontTable.xml`) are
  read from it, unscrambled as the standard says (ECMA-376 Part 1, 17.8.1) and shown with it, as Word
  does; nothing is fetched. A font the text is set in that is neither embedded nor on this device is
  named in the status bar ("Font shown with a substitute: Playfair Display"), so you know the page
  shows another in its place; the document keeps it
- See equations: Word's equations (Office Math) are drawn as MathML by the browser (fractions,
  roots, scripts, sums and integrals, brackets, matrices, functions, accents, and Word's script,
  double-struck and bold letters), on their own line or in the line of text. They stay read-only
  and are written back exactly as they were (new ones can be typed: see References)
- Fill in placeholders (the toolbar's Fill in button, which shows how many there are, or the
  template bar): text a template leaves to fill in, in brackets — `[Client name]`, `{{date}}`,
  `«Amount»` or `<<Amount>>` — is listed once each, with how many places it appears; focusing one
  highlights them all, and what is typed goes in everywhere it appears, as one undoable edit
  (tracked while tracking), taking each placeholder's formatting as Replace does. Citations and
  references (`[1]`, `[Smith 2020]`), check boxes (`[x]`) and editorial notes (`[sic]`) aren't
  placeholders; those in read-only content are listed and left as they are
- Find and replace (Ctrl/Cmd+F, Ctrl/Cmd+H), with match case and whole words. Matches are found
  across formatting but not across paragraphs; a replacement takes the formatting of the match's
  first character, as in Word. Footnotes, endnotes, headers, footers and text boxes are searched and
  replaced too (a match in another header or footer shows its tab). Comments and locked fields are
  searched, but their matches are read-only and never replaced
- See what you can't edit, read-only and in plain language: comments with their replies and resolved
  state in a side panel (click one to highlight its text); fields labelled by what they
  are ("Page number – Word updates it"); citations and bibliographies by the reference manager they
  come from ("Zotero citation – Zotero updates it in Word; edits made here are lost when it does";
  also EndNote, Mendeley, Citavi and Word's own), and typing in one says so in the status bar. Clicking read-only content (a tracked change, chart,
  equation, locked field or cell) says in the status bar what it is and where to edit it, and puts
  the caret beside it, as Word does; typing in it, or in a header, footer or comment, says why
  nothing happens instead of doing nothing (or typing where the caret was before)
- Show formatting marks (¶ button, Ctrl+Shift+8), as in Word: · for spaces, ° for non-breaking
  spaces, → for tabs, ↵ for line breaks and ¶ at each paragraph end. They are drawn over the page,
  never in it, so they can't change the document; the choice is remembered in this browser
- Paste with basic formatting, as Word's "merge formatting": bold, italic, underline,
  strikethrough, superscript, subscript, paragraphs, line breaks, bulleted/numbered lists and links
  to web and e-mail addresses come through from web pages, Google Docs and Word; fonts, colours,
  styles, other links and pictures in copied text don't (the status bar says when pictures were
  skipped; a picture copied on its own is inserted)
- Paste tables from Excel, Google Sheets, LibreOffice Calc, Word or a web page as tables, as Word
  pastes them: the rows and columns, cells merged across and down (`gridSpan`, `vMerge`), the
  columns' widths in proportion to the copied ones (`<col width>`, or the cells'), and each cell's
  text with the basic formatting above (Excel's bold heading cells, written by class in its
  clipboard HTML, too). The table goes where Insert table puts one, in Word's Table Grid style
  across the page; text copied around it pastes before and after it. One cell copied on its own
  pastes as its text. Tracked while tracking (its rows and their text are insertions); one step for
  Undo. Where a table can't go (a header, footer, note, text box, field or another table's cell),
  and for one wider than Word's 63 columns, its rows are pasted as lines with tabs between the
  cells, and the status bar says so
- Links (toolbar button or Ctrl+K), as in Word: the selected text, or new text at the caret,
  becomes a link to a web address, an e-mail address or a DOI (`10.1000/xyz` becomes
  `https://doi.org/10.1000/xyz`); in a link, the same dialog changes its address or removes it. Links
  are written as Word writes them (`w:hyperlink` with an external relationship in the part the link
  is in, so links in notes work too, and Word's Hyperlink style, added if missing); a link inserted
  while tracking is a tracked insertion. Relationships of links that are deleted again are taken out.
  A web or e-mail address typed in the text becomes a link when a space or Enter follows it, as
  Word's AutoFormat makes it ("Internet and network paths with hyperlinks"): `https://…`,
  `http://…` and `www.…` addresses, `mailto:` links and e-mail addresses, without the punctuation
  around them ("(see www.example.com)." links www.example.com; a bracket the address opened stays in
  it). Ctrl+Z straight after gives the address back as plain text; the AutoFormat button turns it
  off; typed while tracking, the link goes with the tracked insertion. Not in a link already, a
  plain-text field or a field
- Fonts, sizes, colours and character styles from the document itself, as a brand template needs:
  the Font menu offers the theme's heading and body fonts (written as Word writes theme fonts,
  `w:asciiTheme`) and the fonts the document uses; Font size the sizes it uses, then Word's usual
  ones; Font colour (a split button that applies the last colour) Automatic, the theme's ten colours
  (`w:themeColor`) and the colours its text uses, then More colours: a colour picker and a box for a
  colour's code (`#0E6467`, or short as `#0E6`), for a brand colour; colours chosen that way are
  remembered in this browser and offered again as Your colours, in any document (cell shading
  offers them too). Character style its own character styles (such as Strong or Emphasis). The
  document's fonts come first, so text stays on brand, then, under Other fonts, Word's usual ones
  that Windows and macOS have (Arial, Calibri, Cambria, Times New Roman and others), written by name,
  for a blank document or one with few fonts of its own. The selected text's run
  properties are changed and the text is rebuilt as runs with them, so the page shows what is saved.
  Clear formatting (button or Ctrl+Space) takes direct character formatting and the character style
  off, as Word does; highlight, language and a link's style stay. Grow Font and Shrink Font (the
  buttons beside the size, Ctrl+Shift+> and Ctrl+Shift+<) step the selected text through Word's
  sizes (… 10, 11, 12, 14, 16 … 72, then by 10), each run from its own size, and Ctrl+] and Ctrl+[
  change it by 1 pt
- Format Painter (the brush button, or Ctrl+Shift+C and Ctrl+Shift+V), as in Word: a click copies
  the character formatting at the caret (or of the selection's first character) and paints it on
  the text selected next, or the word clicked; a double-click keeps it on for several places, until
  Esc or another click on it. The text's run properties become the copied ones (a tracked
  formatting change while tracking); painting whole paragraphs gives them the copied paragraph
  style too. Ctrl+Shift+V is the browser's paste as plain text until formatting has been copied
- Change Case (the Aa menu, or Shift+F3), as in Word: Sentence case., lowercase, UPPERCASE,
  Capitalize Each Word and tOGGLE cASE for the selection, or the word at the caret; Shift+F3 goes
  round UPPERCASE, lowercase and Capitalize Each Word, the text staying selected. Each run keeps its
  formatting; the text's language decides the letters (Turkish dotted i, for example). While
  tracking, a changed run is a tracked deletion and insertion, as Word records it; text already
  deleted isn't changed
- A right-click menu, as Word's (also the Menu key or Shift+F10): Cut, Copy, Paste and Paste as text
  (the browser may ask to let the page read the clipboard; Ctrl+V never needs it), Accept and Reject
  on a tracked change, Link… or, in a link, Open, Edit and Remove link, New comment and Paragraph…,
  and in a table Insert rows above or below and columns left or right, Merge and Split cells,
  Shading and borders and Delete row, column or table. Dialogs opened from it open where it was.
  Cut deletes as a key press would, so it is tracked while tracking. Locked text offers Copy (and a
  comment where comments are allowed). Shift+right-click, and a long press on a touch screen, give
  the browser's own menu, with its spelling suggestions
- Insert Symbol (the toolbar's Ω Symbol), as Word's Symbol dialog: ©, €, ✓, →, ½, accented letters,
  Greek, arrows, maths and the spaces Word offers (non-breaking, thin, en and em), by group or by
  name ("euro", "tick"), or any character by its Unicode code (`20AC`, `U+00BD`, as Word's
  Character code). It goes in at the caret as typed (tracked while tracking); the symbols used last
  are offered first, remembered in this browser
- Formatting changed while tracking is tracked, as Word records it: bold, italic, underline,
  strikethrough, superscript, subscript, highlight, font, size, colour, character style and clear
  formatting record the text's old run properties (`w:rPrChange`); paragraph style, alignment and
  lists the paragraph's old properties (`w:pPrChange`). They show as tracked formatting changes and
  are accepted or rejected here or in Word; formatting changed twice keeps the first, and formatting
  changed back records nothing. The review lists them as tracked
- Smart quotes and dashes as you type, as Word's AutoFormat (lists too: see above): " and ' become
  curly quotes in the text's language (“ ” ‘ ’ in English, „ “ in German, « » with no-break spaces
  in French, and others), "--" between words an em dash, " - " an en dash and "..." an ellipsis;
  and Word's AutoCorrect symbols: (c) ©, (r) ®, (tm) ™, (e) €, --> →, <-- ←, ==> ⇒, <== ⇐ and
  <=> ⇔, replaced as the last character is typed (Ctrl+Z straight after gives back what was typed;
  "(c)" stays as it is where it is an item's letter: at the start of a paragraph, or after "(b)").
  Typed as a key press would be, so they are tracked while tracking. The AutoFormat button in the
  status bar turns this off (remembered in this browser)
- Zoom, as in Word's status bar: − and + step the page through 50% to 200%, the percentage goes
  back to 100%, and Ctrl+mouse wheel (or a trackpad's pinch) on the page zooms it. Only how the page
  shows changes, not the document, and it prints at its size; remembered in this browser. Formatting
  marks, line numbers and a picture's frame and handles follow the zoom
- AutoCorrect entries of your own, as Word's AutoCorrect Options (the status bar's AutoCorrect, beside AutoFormat): a word
  and what replaces it as you type, such as a long name from its initials ("cllr" → "Councillor"),
  once a space, punctuation or Enter ends the word, in the case it was typed in ("Cllr" →
  "Councillor", "CLLR" → "COUNCILLOR"; an entry with capitals of its own, such as "NYT", is typed as
  it is). Ctrl+Z straight after gives back what was typed. Opened with text selected, the dialog
  offers it as what to replace with. Common English typos are corrected too, as Word corrects them
  ("teh" → "the", "adn" → "and", "dont" → "don’t"; the list is in the dialog), in English text only
  (Indonesian's "teh" is tea), and as typed or with a capital, never in capitals. While tracking,
  the word just typed is your own insertion, so the entry takes its place in it, as in Word. "Replace
  text as you type" switches it all off. The entries are kept in this browser, as your colours are.
  Export… saves them as a CSV file ("autocorrect.csv": a header row "Replace,With", then an entry a
  row, in UTF-8 as Excel's "CSV UTF-8" saves it), and Import… brings a list in, from another browser
  or device or made in a spreadsheet (CSV or tab-separated, with commas, semicolons or tabs, with or
  without its header row): an entry for a word already there replaces it, and the dialog says how
  many were added, changed and skipped; nothing is kept until OK is chosen
- Thesaurus, as Word's (Review › Thesaurus, Shift+F7, or the right-click menu): the word at the caret
  (or selected) looked up, its meanings with the words that say each, the most used first, and
  similar and opposite meanings; a click puts a word in its place, in its case ("Big" → "Large"),
  tracked while tracking. In English, a word in another form is looked up by its base and the
  words are offered in its form, as they would fit where it is: "approves" offers "sanctions",
  "took" "chose" (and "taken" "chosen"), "bigger" "larger" and "more prominent", "levies" (a
  verb) "imposes", with English spelling's rules (stopped, carried, permitted) and the common
  irregular verbs, plurals and comparatives. An -ed form is a past or a past participle as the words
  before it say: "returned" offers "gave", "had returned" "given". A base is looked up only as the
  part of speech its form is ("news" isn't "new"s), and a word that can't be put in that form (a name,
  "O.K.") is left out of it. Any word can be typed in. English from Princeton WordNet 3.1; Indonesian, for text in
  Indonesian (or chosen in the dialog), from Wordnet Bahasa 1.2's words checked by hand. Both are
  part of the app (`src/data/`, built by `scripts/thesaurus.mjs`), loaded the first time they are
  used, and work offline: nothing is sent anywhere
- Copy for the web (Download's menu, or the right-click menu for the selection): the document as a
  website's editor (a CMS such as WordPress), an e-mail or a newsletter takes it: clean HTML with
  headings, paragraphs, bold, italic, strikethrough, superscript and subscript, links, bulleted and
  numbered lists (nested by level) and tables, and none of Word's fonts, colours or styles; plain
  text beside it for fields that take only text (list items with their markers, cells apart by
  tabs). The text is as it will read: deleted tracked text and hidden text are left out. Copy as
  Markdown gives the same as Markdown (with GitHub's tables), for Markdown editors, GitHub and
  Ghost. The status bar says what was copied, and what was left out (pictures, text boxes, notes and
  equations, to add in the site's editor). Nothing is saved as a file: it goes to the clipboard
- Word count, as Word's (a click on the status bar's count): pages (about how many), words,
  characters with and without spaces, paragraphs and sentences, and Word's Readability Statistics:
  sentences per paragraph, words per sentence and characters per word, and for English text the
  share of passive sentences, the Flesch Reading Ease and the Flesch-Kincaid Grade Level, worked out
  in the browser. Sentences of more than 30 words are counted, and Show selects each in turn, to cut
  or split it. A length to write to, as a newsroom or a publisher sets one: a target in words, shown
  in the status bar as "612 / 650 words" (in amber once over it), and a limit for the headline in
  characters, shown as "Headline 52 / 70": the headline is the text's first paragraph when it is in
  a heading style (Title, Heading 1–6). Both stay with the document while it is open and in its
  unsaved copy, never in the file
- The status bar counts the words and characters in the document, and those selected ("12 of 340
  words · 64 of 2104 characters"; with and without spaces in its tooltip); headers and footers aren't
  counted, as in Word; the
  page takes the document's language (and a run in another language its own), so browsers that
  choose a spelling dictionary by language (Firefox) check the text in it
- Track Changes for your own edits (toolbar button or Ctrl+Shift+E), as in Word: typing, deleting,
  Enter, Shift+Enter, paste, Replace and page breaks are recorded as tracked changes with your name
  and the date, shown in colour as you work (deleted text struck through and kept, a deleted
  paragraph break shown as a struck ¶ without merging). Deleting your own insertion removes it.
  Word's setting is saved too, so the recipient keeps tracking; documents saved with Track Changes
  on open with it on. Your name is remembered in this browser only
- Compare documents, as Word's Review > Compare does for the text (Review tab, Compare): another
  version of the document, a `.docx` chosen on this device (read in the browser, never uploaded;
  one encrypted with a password asks for it), against the one open. Each paragraph changed, added
  or taken out since that version is listed, in the order of the text, with the words taken out
  struck through and those put in underlined ("hari ~~Senin~~ <u>Selasa pukul 09.00</u>"), and a
  count of each; a click goes to the paragraph (or to where one was taken out). Both versions are
  read from their XML the same way, as their body's text: paragraphs and table cells in order, with
  tracked deletions left out, as they read once their changes are accepted (this one as Download
  would write it), so a document compared with itself is the same. Nothing in either document is
  changed (`src/editor/compare.js`)
- A redline from Compare, as Word's Compare makes one ("Open redline" or "Download redline"): a new
  document, this version as Download would write it, with the changes since the other version
  written into it as tracked changes in your name (Track Changes' name; "Author" until one is
  given), dated now, to send back to whoever wrote the other version or to accept and reject in
  Word (`src/docx/redline.js`). Words put in are marked inserted where they are, keeping their
  formatting (their runs are split at the change); words taken out are put back, marked deleted,
  in the formatting of the text beside them, with tabs as tabs; a paragraph added is marked
  inserted with a paragraph break (the one before it, as Word records Enter, where the two are
  formatted alike; else its own); a paragraph taken out is put back where it was, deleted with its
  break, in its style where this document has it. Tracked changes already in this version are
  accepted in the redline first, as Word's Compare compares the documents as if they were (the
  status bar says how many). The new changes' ids come after every id the package uses. Accepting
  all its changes gives this version's text, and rejecting them all the other's (checked by the
  tests here, and in LibreOffice with real documents). It is named "… (redline).docx", opens in a
  new tab beside this document (made while the result is read, so the browser lets the tab open), or
  downloads; with a password on this document, it is encrypted with it. This document is unchanged
- Accept and reject tracked changes, anyone's and your own, as in Word's Review tab: when the
  document has tracked changes, the toolbar shows Previous, Next, Accept and Reject. Accept and
  Reject act on the change at the cursor (just click it) or every change in a selection, then go
  to the next; their menus also accept or reject without moving on, or all changes at once. An
  insertion accepted (or a deletion rejected) becomes ordinary text; a paragraph break removed
  joins its paragraph with the next; a formatting change rejected restores the old formatting
  (run or paragraph: style, alignment, list); an inserted row rejected (or a deleted one accepted)
  goes. Tracked paragraph breaks show as a coloured ¶, formatting changes with a dashed underline
  or a bar in the margin, tracked rows tinted. Tracked changes to a table's, row's or cell's
  properties (`tblPrChange`, `tblGridChange`, `trPrChange`, `tcPrChange`) are changes like the
  others: accepted, the record goes; rejected, the old properties come back. Tracked changes inside
  content kept as it is (a locked field, a field linked to data, a locked cell, a field such as
  PAGE) are one change per piece: Accept or Reject resolves all of them in its XML at once, as
  Word's Accept All in it would (insertions, deletions, moves, paragraph breaks, formatting, rows,
  cells and merges), and the piece stays locked or linked
- Show markup, as Word's Display for Review (the Review tab): All markup (as above), Simple markup
  (the text as it will be once the changes are accepted, with a red line in the margin beside each
  paragraph that has a change), No markup (that text alone, comments' highlights too), to read a
  heavily edited draft clean, or Original (the text before the changes: insertions hidden, deleted
  text as ordinary text), read-only, as typing there would edit text that is no longer there. Only
  how the page shows: the changes stay in the page and the file, edits go on being tracked in
  simple and no markup (and show as the view has them), print and PDF follow the view, and every
  document opens with all markup. Previous and Next change, Accept and Reject show all markup
  again, so the change they go to can be seen
- Review your changes before downloading: for a document started here (a blank one, or a new one
  from a template) it doesn't ask for the edits to be tracked, and the CV screening check shows only
  for a CV (or from Help). Download (or Ctrl/Cmd+S) first lists every paragraph that
  changed, with the words deleted and inserted, which changes are tracked, a warning when some
  aren't, and anything else that changes (comments removed with their text, Track Changes switched,
  styles or lists added). An untouched document says "No changes". When the text changed and the
  document has fields Word works out (a table of contents, cross-references, figure and table
  numbers), it says they may be out of date and how to update them in Word (Ctrl+A, F9), and offers
  to have Word update them when the file is opened (`w:updateFields` in `settings.xml`); when
  citations from Word or a reference manager (Zotero, EndNote, Mendeley) were added or removed and
  a bibliography lists them, it says the bibliography needs refreshing (`src/docx/fields.js`). The
  review can be skipped, and opened any time from the toolbar
- Download the result as `.docx` (Ctrl/Cmd+S), or save a PDF (see above)
- Save back to the file opened, as Word's Save, where the browser can (Chrome and Edge on a
  computer, with the File System Access API): a file opened from the start page's Open tile (the
  browser's own file picker) or dropped on it is saved into itself by Save (Ctrl+S) in the header,
  through the review as Download is; the browser asks once whether Intact may save changes to it.
  Download's menu (▾) still has Download a copy, and Save as… picks a file on this device to save to
  (a new manuscript, say), which later Saves write to. A file renamed in the header is saved as a
  new one. The file is written where it is, by the browser: nothing is uploaded. Elsewhere (Firefox,
  Safari, phones), Open and Download work as before
- Share… from Download's menu (▾), where the browser can share files (Safari on iPhone and iPad,
  among others): after "Review your changes", the saved file goes to the device's share sheet, to
  send it back by mail or a messaging app without looking for it in Downloads. It goes from the page
  to the app chosen, on the device; Intact uploads nothing. Where the browser asks for one more tap
  (saving took a moment), a small dialog offers it. Where files can't be shared, the item isn't shown
- Respect the restrictions Word puts on a document, as Word does: with Restrict Editing on
  (`w:documentProtection`, enforced) it opens read-only, comments only (comments can be added and
  answered), filling in forms only (its fields can be filled in) or with every edit tracked (Track
  Changes stays on, and tracked changes can't be accepted or rejected); a document marked "read-only
  recommended" or with a password to modify (`w:writeProtection`) opens read-only. A bar above the
  toolbar says what applies, and typing or clicking in locked text says why nothing changes. "Edit
  anyway" lifts it for the session after saying what it means (going around Restrict Editing, which
  only someone allowed to should do); the restriction stays in the file, so Word still applies it,
  and the review says so. Protection set up but not enforced restricts nothing, as in Word
- Documents encrypted with a password, as Word's Encrypt with Password (File > Info > Protect
  Document) and LibreOffice's Save with password make them: opening one asks for its password, on
  the start page ("Protected with a password"), which is checked and used on this device only, to
  decrypt it in the browser (`src/docx/encryption.js`, with the browser's Web Crypto: no library,
  nothing fetched); a wrong one says so, and the field is there to try again. Word 2010 and later's
  Agile Encryption (AES with SHA-512 and 100,000 rounds, or the other AES and SHA-1/2 settings it
  allows) and Office 2007's Standard Encryption, which LibreOffice still writes (AES, SHA-1), are
  read, from the OLE compound file Word keeps them in (`src/docx/cfb.js`, which reads and writes
  that container, large files' extra FAT sectors too); a file whose integrity check (its HMAC)
  doesn't match is refused as changed or damaged. A bar says the document is protected. Downloaded
  (or saved back, or shared), it is encrypted again with the same password, as Word saves it: Agile
  Encryption, with a new salt and HMAC key each time, in a compound file with Word's data spaces;
  a document opened from Standard Encryption is saved with Agile, as Word re-saves one. Unchanged,
  the file downloads exactly as it was opened. Download's menu has "Protect with password…" for any
  document (typed twice; the dialog says a lost password can't be recovered) and, once there is
  one, "Change or remove password…" (also on the bar); the review says whether the file downloads
  encrypted, with a new password, or without one. The password isn't kept: the key made from it is,
  while the document is open. The copy of unsaved changes kept on this device is encrypted with the
  document's password too, and restoring it asks for the password; the start page says it is
  encrypted. What is made from the document on this device is not encrypted unless said: a mail
  merge's documents (the status bar says so) and a PDF; a redline is (see above). Checked against
  other implementations: files encrypted by msoffcrypto-tool and LibreOffice open here, and files
  encrypted here open in LibreOffice and decrypt with msoffcrypto-tool, integrity check included
- Digitally signed documents (`_xmlsignatures`) open read-only too, as Word opens them. Downloaded
  unchanged, the file is the one opened, byte for byte, so its signatures stay valid; edited (after
  "Edit anyway"), the signatures are removed from the download, as Word removes them when a signed
  document is edited, rather than sent along invalid; the review says which will happen
- See the document's sensitivity label beside its name, as Office shows it: the Microsoft Purview
  label (`MSIP_Label_*` custom properties, or `docMetadata/LabelInfo.xml`, which doesn't store the
  name), with who or what applied it, when, and the marking it adds, or another tool's classification
  property (such as a protective marking). It is never changed: the file keeps it, and the review
  says so
- Check a document before you release it (for example under Freedom of Information): the review lists
  what the file carries beyond the page, as Word's Document Inspector does, and removes what can be
  removed here. Comments (Remove all comments), tracked changes, whose deleted text anyone can still
  read (Accept all changes), hidden text (`w:vanish`; Remove hidden text, untracked as in Word) and
  personal information: the author, last saved by, company and manager properties and the names on
  comments and tracked changes. "Remove it from the downloaded file" takes those out (names become
  "Author", initials "A", `people.xml` loses the accounts behind them) and sets Word's
  `w:removePersonalInformation`, so Word keeps doing so; unticked, the parts go back as they were.
  A CV or letter made from someone else's file or template carries their name as its author: where
  your name is known (given for comments or Track Changes), the review says when the author isn't
  you. "Do this for every file I change that has no comments or tracked changes" makes removing it
  the default, remembered in this browser: then it is chosen for each file you have changed whose
  properties name someone (not a placeholder such as "Author"), with no comments or tracked changes
  (whose names it would change), no restriction, and the edit recorded in its properties. The
  review opens there and says what goes and why; unticked, the file keeps it. Without that choice,
  nothing is removed for you, so the properties of a file you didn't ask to change stay as they were
  Other document and custom properties, custom XML data, embedded files and macros are listed, for
  you to check in Word. Removing is one undoable step each, and not offered while the document is
  restricted
- Saving an edit records it in the document's properties, as Word does: modified now, last modified
  by you (the name given for comments and Track Changes; without one, the last name is kept and the
  review says so) and the revision number one up, worked out from the file as opened so saving twice
  counts once. Nothing changed, nothing is recorded: the properties come back byte for byte. The
  review shows what is recorded and can switch it off (remembered in this browser); with personal
  information removed, no name is recorded
- Redact, for a source's name or an address in a document to be released (the Review tab, or the
  right-click menu): the selected text becomes █, one for each character, as a page is blacked out
  by hand. It is never tracked, even while tracking, because a tracked deletion keeps the text in the
  file for anyone to read in Word; Undo brings it back while the document is open. A link in it goes
  too, and its address goes from the package's relationships when the file is saved (Undo puts it
  back). A selection holding a tracked change, or read-only content such as a field, is refused,
  saying why. The status bar says when the same words are elsewhere in the document, and "Before you
  release" lists the redacted passages and where what was redacted is still in the file: in the text
  (Redact them all blacks out every place there, headers, footers and notes included), in tracked
  changes, in read-only content, in comments or in the document's properties. What was redacted is
  remembered for this only while the document is open, never in its unsaved copy or the file
- Check accessibility before publishing, as Word's Accessibility Checker does (Section 508,
  EN 301 549, WCAG): the review lists errors (pictures without alt text, unless marked decorative;
  tables without a header row), warnings (skipped heading levels, no document language, text whose
  colour is below WCAG's contrast minimum) and tips (unclear link text such as "click here", blank
  paragraphs used for spacing, pictures not in line with text, no title, and text hard to read:
  sentences of more than 30 words, and in English a Flesch-Kincaid grade above 12 once there are
  100 words, as WCAG's reading level asks). "Show" goes to each place
  in turn; for a picture it opens its alt text, and a missing title can be added there and is saved
  in the document's properties
- Check a CV for employers' screening software (applicant tracking systems), which read a CV's text
  before anyone does and miss or mix up what isn't in plain paragraphs: the review warns of contact
  details only in a header or footer, no e-mail address, text in text boxes or columns, and charts
  or diagrams, and gives tips on text laid out side by side in tables, pictures (no text is read in
  them) and a file called just "CV". "Show" goes to each place. A document reads as a CV by its name
  (CV, résumé, curriculum vitae, Lebenslauf) or two section headings such as Experience, Education
  and Skills, in English, Indonesian, German, French, Spanish or Portuguese; for any other the check
  is there, folded and not counted
- Files that aren't `.docx` packages say what they are, instead of failing to unzip: encrypted with a
  password, protected by rights management (a sensitivity label that encrypts, or IRM), a Word
  97–2003 `.doc` with a `.docx` name, or something else. Intact doesn't decrypt files
- Word's keyboard shortcuts: Ctrl+B/I/U, Ctrl+Shift+= / Ctrl+= (superscript / subscript),
  Ctrl+Alt+H (highlight), Shift+F3 (change case), Ctrl+Shift+C / Ctrl+Shift+V (copy and paint formatting), Ctrl+Space (clear character formatting; highlight stays, as in Word), Ctrl+Shift+8 (formatting marks), Ctrl+L/E/R/J (alignment), Ctrl+Shift+N (Normal), Ctrl+Alt+1–3 (headings),
  Ctrl+Shift+L (bullets), Ctrl+M / Ctrl+Shift+M (list level), Ctrl+Enter (page break), Shift+Enter
  (line break), Ctrl+Shift+Space (non-breaking space), Ctrl+Shift+E (Track Changes), Ctrl+K (link), Ctrl+F/H,
  Ctrl+Z/Y, Ctrl+S; Cmd on a Mac. Some browsers keep a few keys for themselves
- The ribbon, as Word's: the toolbar's tools in tabs, one tab's at a time on one row of a laptop's
  screen: File (Download, or Save to the file opened, and Download a copy; Save as; Print; Save as
  PDF; Share; Copy for the web; Mail merge; Protect with password, or Change or remove password;
  Inspect document, the review at Before you release; Check accessibility), Home (styles, fonts,
  formatting, paragraphs, lists, Find, formatting marks), Insert (link, check box, date and
  references, symbol, picture, signature, table, Fill in), Layout (Page setup, page numbers, table
  layout), References (footnotes, endnotes, updating fields), Review (Track Changes, accepting and
  rejecting, the review, comments) and View (the Navigation pane, zoom with Word's levels and back
  to 100%, formatting marks and Show markup), where Word users look for them. Home is the tab chosen
  first, as in Word. A tool can be in more than one tab (formatting marks in Home and View, Show
  markup in Review and View): one control, shown in each. The File and View tabs' own tools are the
  ribbon's: "All tools" and a phone leave them to Download's menu (▾) and the status bar, which have
  them. Undo, Redo and Help are in every tab. The
  Review tab shows how many tracked changes are waiting, and the Layout tab is marked while the caret
  is in a table, as Word shows its table tabs. ← and → (Home, End) choose a tab, as WAI-ARIA's tabs
  have it; Alt+F10 goes to the tab's tools. "All tools" shows every tool at once in one toolbar
  instead, as before ("Tabs" goes back); the tab and the choice are remembered in this browser. A
  phone keeps its one row of every tool, which scrolls
- The toolbar works from the keyboard, as WAI-ARIA's toolbar pattern and Word have it: Alt+F10 goes to
  it, it is one Tab stop, the arrow keys (and Home and End) move along it and Esc goes back to the text
  where the caret was. Its buttons are named and say whether they are on (bold, alignment, lists…) for
  screen readers. Its controls stay in place: Accept, Reject and the count of tracked changes are always
  there, turned off while there is nothing to accept, and keep their labels on wide and laptop screens.
  The icon buttons get short text labels (Insert, Symbol, Picture, Sign, Table, Layout, Footnote, Endnote,
  Update, Page no., Page setup, Find, Fill in) where the toolbar has room for them: each in turn, as
  long as it fits without the toolbar taking another row, worked out again when the window's width
  changes, so the toolbar never grows; their names for screen readers stay as they were
  On a phone it is one row that scrolls sideways, faded at the edge where more controls are, with
  touch-sized buttons. The tools a phone is used for come first (Undo, Redo, Bold, Italic,
  Underline, Highlight, the lists, Insert, Draw signature, Insert picture and New comment), the rest a
  swipe away, and the arrow keys follow that order. A document with tracked changes or comments to
  go through, as one to approve between meetings, puts the review's tools first instead: Undo and
  Redo, then Previous and Next change, Accept and Reject (named, not just a tick and a cross) with
  their menus and the count, then New comment and the comments; the status bar says so once when it
  opens ("2 tracked changes and 1 comment to review: Accept and Reject are first in the toolbar").
  Once nothing is left to review, the editing tools come first again. The header leaves the file's
  name room: Open and Close are their icons (named for screen readers), with less space between
  them. The status bar is one line, and what just happened
  ("Check box inserted", "Shared letter.docx") comes first in it, rather than out of sight at its
  end; zoom is left to the phone's pinch. Nothing makes the document scroll sideways
  there: a picture wider than the screen is shown smaller (its size in the file unchanged), and what
  lies past the page's edge is cut off at it, as on paper
- Help (Ctrl+/, or the toolbar's Help button) lists the keyboard shortcuts and opens the review's
  checks without downloading: Review your changes, Before you release, Check accessibility and
  Check a CV for screening software, each opening the review at its part
- The review starts with what needs a look: a line saying how many things there are to check, then
  warnings and choices (such as recording the edit in the properties). Before you release,
  Accessibility, the changed paragraphs (more than five) and other notes are folded behind their
  counts, and open by themselves when they hold something to check
- Several documents at once, each in a tab of its own (a CV in one, its cover letter in the next):
  the Open button in the header, or Ctrl+O (Word's Open, which in the browser would leave the app),
  opens the start page in a new tab, and this document stays open. A document dropped on the page
  is offered to open in a new tab, and handed to it in memory (a `BroadcastChannel` between the two
  tabs: nothing is stored or sent). Each tab holds a Web Lock for its document's draft, so the start
  page in another tab shows that draft as "Open in another tab" instead of offering to restore a
  second copy whose edits would overwrite each other's; closed there, it can be restored
- Unsaved changes survive a reload, a crash or a closed tab: while a document has changes not
  downloaded, a copy is kept in this browser (IndexedDB), on this device only, a moment after each
  edit ("Changes kept on this device" in the status bar), and the start page offers to restore or
  discard it. The copy holds the edits only, so the download made from it records the edit in the
  document's properties once; a signed document's copy has no signatures, as its changed download
  wouldn't. It goes when the document is downloaded or closed, when it is discarded, and after 7
  days; the five most recent are kept. Where the browser refuses storage (some private windows), editing works
  as before, without it
- Faithful round-trip: everything the editor doesn't understand is written back untouched —
  fonts, colours, sizes, spacing, numbering, images, table layout, headers/footers, footnotes,
  fields, equations, bookmarks, section settings and tracked changes (shown with their author until
  accepted or rejected, here or in Word). The test suite
  checks this on every pull request (see [Testing](#testing))
- Undo/redo for every edit, including paragraph style changes; list numbers update as you edit
- Deleting never quietly takes what you can't edit with it. An edit that would remove a tracked
  change (someone's insertion, deletion or move, a tracked paragraph break, a table with tracked
  rows) is refused, and the status bar says whose change it is and what to do. Deleting the text a
  comment is attached to deletes the comment, as in Word: the comments panel and status bar say so,
  Undo brings it back, and on save it is removed completely (with its replies, from every comment
  part). A comment whose mark survives keeps its range, shrunk to the text that is left. Backspace
  and Delete next to invisible markers (comment ranges, bookmarks, field codes) delete the text
  beyond them, not the marker

## How it works

`src/docx/DocxDocument.js` is the core:

1. **Load** – JSZip opens the package; `_rels/.rels` locates the main part (`word/document.xml`),
   and `styles.xml`, `theme1.xml` and `numbering.xml` are read for display.
2. **Render** – each `<w:p>` becomes an HTML block (`<p>` or `<h1>`–`<h6>`). Its `<w:pPr>` and each
   run's `<w:rPr>` are kept as the original XML in a side table, referenced by `data-ppr` /
   `data-rpr`. Bold/italic/underline/strike become HTML tags so the browser can edit them.
   Unsupported content becomes a non-editable placeholder (`data-raw`) holding the original XML.
   Tables keep their XML structure; cell contents are editable, and rows and columns change through
   `src/docx/tables.js`, which edits a copy of the table's XML by grid column while the rendered
   rows and cells (with any edits in them) are kept. A section in columns is wrapped in a
   `div.docx-columns` for the page only; saving reads the blocks inside it.
3. **Save** – the editor DOM is walked back into WordprocessingML: stored properties are cloned,
   toggles are merged in (respecting the schema order of `CT_RPr`/`CT_PPr`), adjacent runs with the
   same formatting are merged, and placeholders are written back verbatim. Only
   `word/document.xml` changes (and only if the text did: saved as it was when rendered, the
   part goes back byte for byte, so an edit to a header alone, or one undone, leaves it alone), plus `styles.xml` if a heading or the Table Grid style had to be
   added, `numbering.xml` if a new list was started, `footnotes.xml` / `endnotes.xml` if a note was
   edited, a header or footer part if it was edited, and for inserted pictures their files in
   `word/media`, relationships and content types.
4. **Content controls** – a `<w:sdt>` is shown as a frame around its rendered content. Text and
   combo box fields hold an editable `.sdt-content` island; check boxes, drop-downs and dates
   change through clicks. On save its `<w:sdtContent>` is rebuilt and its properties kept, updating
   only what the user changed (`showingPlcHdr`, `w14:checked`, `lastValue`, `fullDate`).
   Placeholder text left alone or cleared is written back verbatim. Locked fields, fields bound to
   custom XML (Word would overwrite edits from the data on open), tables of contents, citations,
   pictures and groups stay read-only and are written back verbatim.
5. **Footnotes and endnotes** – the notes the body refers to are rendered after it, inside the
   editable page, as a `section.docx-notes` (not editable itself) holding one editing host per note
   (`.docx-note`), so typing, formatting, Undo, Find and Track Changes work there as in the body.
   `serializeBlocks` skips the section; `serializeNotes` saves each note that no longer saves the
   same way as when it was rendered into its part, and `writeNotes` writes the part only then. The
   notes, and each note's number, are `data-protected`: `src/editor/protect.js` refuses an edit
   that would remove one with the text around it.
6. **Equations** – `src/docx/omml.js` turns Office Math (`m:oMath`, `m:oMathPara`) into MathML for
   display. The equation itself stays a read-only placeholder holding its original XML.
7. **Find and replace** – `src/editor/find.js` searches runs of text nodes in the same paragraph
   and editing host, and highlights matches with the CSS Custom Highlight API, which doesn't
   change the DOM. Replacing edits the text nodes directly as one undoable step.
8. **Companions** – `src/docx/annotations.js` loads comments (replies and resolved state from
   `commentsExtended.xml`), footnote/endnote text (for references' tooltips) and finds the headers
   and footers each section uses (edited in the page: see 10). Comment ranges are highlighted with the CSS Custom Highlight API between
   the hidden `commentRangeStart` / `commentRangeEnd` markers.
9. **Lists** – while editing, a paragraph's list and level are `data-num-id` / `data-ilvl`
   attributes (set from its own `<w:numPr>`; otherwise its style's numbering applies). A paragraph's
   `<w:numPr>` is only rewritten when the user changed its list or level; its direct indentation is
   then dropped so the list level's applies, as in Word, unless it was set since in the Paragraph
   dialog (`data-ind-set`). Spacing and indents set here are written into the paragraph's stored
   `<w:pPr>` straight away, so undo, tracking (`w:pPrChange`) and saving see them as they are. New
   lists get Word's default nine-level bullet or numbering definitions.
10. **Headers and footers** – rendered into the page, before the text and after the notes, as
   `section.docx-hf-band` bands (not editable themselves, like the notes' section) holding one
   editing host per header or footer part (`.docx-hf-part`), each rendered with its part's
   relationships (`doc.inPart`). The view shows the variant chosen in its tabs with a stylesheet, so
   the choice survives Undo; the header band is in the flow at the top of the page, the footer band
   is placed at its foot, with room kept for it below the text. `serializeBlocks` skips the bands;
   `serializeParts` saves each part whose content no longer saves the same way as when it was
   rendered, and `writeParts` writes it only then.
11. **Text boxes** – `src/docx/textboxes.js` finds a run's text box (DrawingML in `w:drawing` or
   `mc:AlternateContent`, or VML in `w:pict`) and its size, place, fill, outline and insets. It is
   rendered as a `span.textbox` frame (not editable, placed as a picture is) around a
   `.textbox-content` editing host. On save the run is cloned from the raw store; if the text
   changed, the new blocks replace its `w:txbxContent`, and in `update` (once revision ids are known)
   the copies in `mc:Fallback` get the same blocks, their revisions their own ids. The page nests
   paragraphs inside the paragraph the box is anchored in, which the HTML parser doesn't allow, so
   Undo (`src/editor/history.js`) restores copies of the page's nodes rather than parsing its HTML.
12. **Layout and fonts** – `StyleSheet.pPrCss` / `rPrCss` (`src/docx/styles.js`) give borders,
   shading, character spacing, raised text and theme colours as CSS; drop caps float. Embedded fonts
   are loaded with the `FontFace` API (`src/docx/fonts.js`), and fonts that are missing are found by
   measuring text on a canvas, all in the browser.
13. **Protected deletions** – everything that is or holds a tracked change is labelled when rendered
   (`data-tracked`, see `src/docx/revisions.js`); `src/editor/protect.js` checks what each edit would
   remove (`beforeinput` target ranges, or the selection) and refuses edits that would take a
   labelled element away. Comments follow their mark: `src/docx/commentRemoval.js` removes, on save,
   comments whose mark was deleted from `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`
   and `commentsExtensible.xml`, and those parts are only rewritten when a comment was removed.
14. **Encryption** – a password-encrypted `.docx` is an OLE compound file (`src/docx/cfb.js` reads
   its streams and writes new ones) holding `EncryptionInfo` and `EncryptedPackage`
   (`src/docx/encryption.js`, MS-OFFCRYPTO). The password's key is the salted hash hashed again
   (100,000 rounds for Word) and with a block key; it decrypts the random key the package is
   encrypted with, in 4096-byte AES-CBC segments, each with its own IV, checked by an HMAC. Web Crypto
   always pads AES-CBC, so decrypting adds the block that decrypts to padding, and encrypting drops
   it. `DocxDocument.open` decrypts before anything else and keeps the key; `save` and `draft`
   encrypt what they would have written (`sealed`), or give back the file opened when nothing changed.

## Development

```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # production build in dist/
npm test         # round-trip and editing tests in headless Chromium and Firefox
```

The thesaurus's data (`src/data/thesaurus-*.txt`) is built from the open wordnets, with their
licences at the top of each file:

```bash
npm pack wordnet-db@3.1.14 && tar xzf wordnet-db-3.1.14.tgz   # Princeton WordNet 3.1
git clone --depth 1 https://github.com/open-language/id-wordnet   # Wordnet Bahasa 1.2
node scripts/thesaurus.mjs package/dict id-wordnet/database/1.2/wn-msa-all.tab
```

## Testing

`tests/` drives the production build in real browsers (Chromium and Firefox) with
[Playwright](https://playwright.dev/), because editing depends on the browser's `contenteditable`,
which differs between engines. Fixtures are built from XML in the
tests themselves (`tests/helpers.js`). The core checks are:

- opening and downloading a document **without edits** gives back the same document;
- saving a saved file again changes nothing;
- editing one paragraph changes **only** that paragraph.

"The same" is judged by meaning, not bytes (`canonicalize` in `tests/helpers.js` lists the
normalisations, such as merging adjacent runs with identical formatting). A new feature that touches
the document should add its case to the kitchen-sink fixture in `tests/roundtrip.spec.js`.

Real documents written by Word, LibreOffice, pandoc and python-docx live in `tests/fixtures/`
(anonymized; see the README there, and `sources/` for how the generated ones are made). Each goes
through every promise in `tests/fixtures.spec.js`: all its text (body, headers, footers, text
boxes, notes) and pictures are shown; saved without edits, or with an edit undone, every part comes
back byte for byte; typing in a paragraph, heading, list item, table cell, header, footer, text box
or note, pressing Enter or making a word bold changes only that block of that part (as a tracked
change where the document has Track Changes on), the typed text in the formatting around it; the
saved package stays sound (well-formed XML, content types, relationship targets); and the page logs
no errors. A `.docx` added there is tested automatically.
First run: `npx playwright install chromium firefox`. To run one browser:
`npm test -- --project=chromium` (or `firefox`).

Checked by hand as well: documents edited in Intact open correctly in Microsoft Word, and the
editor works in Firefox (November 2026).

## Deploying to GitHub Pages

The workflow in `.github/workflows/deploy.yml` builds and deploys on every push to `main`.
Enable it once under **Settings → Pages → Build and deployment → Source: GitHub Actions**.
The build uses relative asset paths (`base: './'`), so it works under any repository sub-path.

## Self-hosting

Organisations that can't use a public site can serve Intact from their own web server, or on a
network with no internet at all: it is static files with no server code and no external requests.

1. Download `intact-<version>.zip`, its `.sbom.cdx.json` and `SHA256SUMS` from a
   [release](../../releases) (made by `.github/workflows/release.yml` when a `v*` tag is pushed,
   after the tests pass).
2. Check them:

   ```bash
   sha256sum -c SHA256SUMS
   ```

   ```bash
   gh attestation verify intact-<version>.zip --repo bowo-prasetyo/intact
   ```

   The attestation proves the zip was built by this repository's release workflow from the tagged
   commit. The SBOM (CycloneDX) lists what the build contains (Vue, Vue Router, JSZip).
3. Unzip it into any directory your server serves, at any path. Serve `.webmanifest` as
   `application/manifest+json`.
4. Recommended response headers for its path (the page already carries the Content Security
   Policy in a `<meta>` tag; a header adds `frame-ancestors`, which a `<meta>` can't set):

   ```nginx
   location /intact/ {
     add_header Content-Security-Policy "default-src 'none'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' blob: data:; font-src 'self' blob: data:; media-src 'self' blob: data:; connect-src 'none'; worker-src 'self'; manifest-src 'self'; base-uri 'none'; form-action 'none'; object-src 'none'; frame-src 'none'; frame-ancestors 'none'" always;
     add_header Referrer-Policy "no-referrer" always;
     add_header X-Content-Type-Options "nosniff" always;
     add_header Permissions-Policy "camera=(), microphone=(), geolocation=()" always;
   }
   ```

**What the browser enforces.** The Content Security Policy (`vite.config.js`, put into the built
`index.html`) lets the page load only Intact's own files and connect nowhere (`connect-src 'none'`):
even a script that got into the page couldn't send a document anywhere. Links opened from a
document send no referrer. The tests check that a whole editing session breaks no rule of the policy
and that a request to another site is refused (`tests/deployment.spec.js`).

**Offline and installable.** Served over HTTPS (or from `localhost`), the built app registers a
service worker (`src/sw-template.js`, written out as `sw.js` with the list of the build's files)
that keeps a copy of the app's own files, so once opened it also works offline, and the browser
offers to install it as an app (`manifest.webmanifest`), which can open `.docx` files from the
desktop where the system supports it. The service worker caches nothing else: documents are opened
from the device's files and never fetched, cached or stored. A new version replaces the old copy.

The About page shows the version.

## Limitations

- After restoring a draft, the text you had changed while tracking shows as tracked changes kept from
  the file (accept or reject them; they can't be edited inside), and "Review your changes" lists only
  what changed since restoring. A draft is kept per browser: another browser or device doesn't see it
- Simple and no markup hide deleted text, but a deleted paragraph break still shows its two
  paragraphs apart, and a paragraph whose text was all deleted still takes its line; insertions in
  content kept as it is show as plain text. The original likewise keeps an inserted paragraph break
  (and so the paragraphs it made) and shows text with its new formatting where the formatting was
  changed while tracking (the old is in the file, and Reject brings it back)
- Moving a heading with its text (the Headings pane) isn't tracked, so it isn't offered while
  tracking (cut and paste is tracked); it moves only headings of the text itself (not in a table or a
  content control), and not a part that holds a section break, or that a bookmark, comment, field
  (a table of contents, an index) or editing permission starts in and ends outside of, as moving it
  would cut that in two. Word renumbers its own fields (captions, cross-references) when it updates
  them
- Saving back to the file opened needs the File System Access API (Chrome and Edge on a computer);
  a document restored from a draft, brought over from another tab or opened from a phone is
  downloaded instead, and Save as… picks a file again
- Tracked changes inside locked content are accepted or rejected together, not one by one; edits
  inside form fields and text moved by dragging aren't tracked (cut and paste is), nor is ending a list
  with Enter on an empty item, or a picture's layout
- Fonts are chosen from those the document uses and Word's usual ones, and sizes from the document's
  and Word's usual sizes, not from every font installed; a font not on this device is shown with a
  substitute (the status bar names it); a colour by its code is written as that colour, not as a theme colour, and a theme
  colour's tints and shades aren't offered for text. A font, size or colour applies to selected
  text, not to the word at the caret. "Your colours" are kept per browser (another browser or device
  doesn't have them)

- 3-D lighting is worked out for a bevel's surface, not for a picture's depth (extrusion) or contour;
  Word's preset 3-D rotations are shown with angles as near as their look (a rotation set by angles
  is exact); artistic effects are drawn with Intact's own versions of Word's filters, so they look
  alike rather than identical, and in Word they are part of the picture; a wrap outline isn't
  turned with its picture. Field cells that are locked, linked to data or not text
  fields can't be merged
- Printing prints headers and footers once, at the start and the end, as the page shows them, where
  they can't be put on the right pages (the Save as PDF dialog says why): a page number beside a
  picture or table; in a document of several sections, a header or footer with a picture or table
  (a logo in a letterhead: each section still prints on its own paper), a section after the first
  with a different first page that numbers its pages on from the section before (which page is its
  first is Word's to work out: start its numbers at a number, in Format page numbers, to print it),
  the number of pages in a section (`SECTIONPAGES`), or a section that begins on the same page as the
  one before with a header of its own; and in a browser without CSS page-margin boxes (Firefox,
  Safari), where a page number would otherwise print the same on every page. Even pages are the
  printout's even pages, as the browser counts them. Line numbers and comments in the margin aren't printed. Browsers that switch tabs with Ctrl+1 to Ctrl+9 may
  keep Ctrl+1, Ctrl+2 and Ctrl+5 for themselves (the installed app, in its own window, gets them):
  the Line and Paragraph Spacing menu does the same. Ctrl+T and Ctrl+Shift+T (hanging indent) are
  the browser's: use the Paragraph dialog. Indents are set in centimetres or inches and spacing in points
- Lists that define only one level (such as Word's built-in "List Bullet" style) can't be indented
- Tab stops on the page: a tab sharing its run with other tabs or text (in a paragraph of several
  tabs) is put at the paragraph's first stop, if it is a left one, and evenly spaced stops after it,
  else at the default stops, and its leader isn't drawn (they are kept, and Word lays them out)
- A watermark is text only (Word's picture watermarks are kept, not made or shown here), in Word's
  silver, the same in every header; it isn't tracked. It is drawn in the middle of each page as
  the page count works them out, so it may sit a little higher or lower than in Word
- The right-click menu's Paste reads the clipboard through the browser, which may ask first or
  refuse (then use Ctrl+V)
- The Format Painter copies character formatting and a paragraph's style, not its alignment,
  spacing or indents; Ctrl+Shift+C opens the developer tools in some browsers instead (use the
  button)
- Pasting keeps only bold, italic, underline, strikethrough, superscript, subscript, paragraphs,
  lists, tables and links to web and e-mail addresses (not highlight); pictures within copied text and other
  links (their text is kept) are not added. A pasted table's cell shading, borders, number alignment
  and row heights aren't kept (it takes Table Grid's), and a table pasted in another table's cell is
  pasted as text, not as a nested table or into the cells there. Making existing text a link isn't tracked
- Chrome and Edge check spelling in the languages chosen in their settings, whatever the document's
  language
- Text that holds someone's tracked change can't be deleted together with it; accept or reject the
  change first, or select around it
- Deleting a footnote or endnote reference leaves its note in the notes part (Word ignores it), but
  for a note added here; pictures, tables, page breaks, new comments and notes can't go in a note (comments already in notes can be replied to, edited,
  resolved and deleted in the panel)
- A page number (`PAGE`) shows # here, as each page has its own (it prints numbered: see Print;
  Word numbers the pages when it opens the file), and the number of pages (`NUMPAGES`,
  `SECTIONPAGES`) about how many there are, as the status bar says (on a phone, the number in the
  file): drawn over the numbers the file holds, which are kept as they are and aren't typed into
  (End and Shift+End go past one at the end of a line, as past a footnote's or endnote's
  reference, to type or select after it; Home and Shift+Home go before one at the start of a line,
  so text typed there goes before the number, not into it). A
  footer made here is for all pages of the first section and those after it without their own; a
  first page set to have a different footer, and even pages where they differ, don't show it, as in
  Word. Page setup isn't tracked; gutters stay as they are, and columns made here are of equal width (Word's columns of their own widths are set in Word). The page shows every section
  at the width of the document's last one (a landscape section's wide table may be cut off at the
  page's edge here, not in Word, nor printed, where each section is on its own paper); each break is named with its section's paper. "Different first
  page" and "different odd and even pages" are set for the whole document, not section by section.
  Modify Style changes paragraph styles' font, size, bold, italic, colour, alignment, spacing and
  indents: not their borders, tab stops, numbering or what they are based on (a list's indents come
  from its numbering, as in Word, so a list paragraph keeps them); it isn't tracked
  (Word doesn't track style changes either)
- Equations already in a document are shown, not edited (new ones are typed as LaTeX, a part of it:
  environments with their own column settings, such as `array`, aren't); they need a math font to look as in Word (Windows and macOS have one: Cambria Math, STIX
  Two Math; on Linux install one such as Latin Modern Math)
- Fields are preserved; their instructions can't be edited (their results can, until they are
  updated). Update fields doesn't work out page numbers (`PAGE`, `PAGEREF`, `NUMPAGES`: Word does
  when it opens the file), nor dates, document properties or mail-merge fields, nor citations whose
  source the document doesn't have or citations and bibliographies in Word's content controls (Word
  updates them), nor
  fields in notes, headers and footers or written as simple fields (`w:fldSimple`); it isn't
  tracked. Editing a comment's text replaces any formatting within it
- Citations and bibliographies are formatted here in APA or MLA style as their guides set them out,
  close to but not always the same as Word's own styles, which Word applies again when it updates
  them; a citation cites one source, books, journal articles and websites are the kinds of source
  offered, and Word's other styles (Chicago, Harvard, IEEE…) aren't. A table of contents,
  bibliography or caption isn't tracked (it can't be inserted while tracking), and its paragraphs go in
  the text itself, not in a table, note, header, footer or text box. Undo takes a citation out of
  the text, not its source out of the document's sources (as in Word, where a source stays in the list)
- Tables, page breaks and new comments can't go in a header, footer or text box, nor pictures in a
  text box or note (the status bar says so); a note reference in a header stays as it is
- Cell borders are a single ½ pt line in the automatic colour, as Word's Borders button draws them;
  their style, width and colour are set in Word. Borders and shading go on cells, not on the whole
  table or on paragraphs
- A text box's size, position and look are kept, not changed here, and it grows to show all its text
  where Word would cut it off; text wrapping around it follows its rectangle. A text box made here is
  in line with the text. Lines, arrows, rectangles and ellipses are drawn and laid out; other shapes
  without text (stars, callouts, freeforms), groups of shapes, canvases, charts and SmartArt stay as
  they are. A shape's colour, line style and arrowheads are set in Word. A text box laid out on the
  page (rather than beside its paragraph) is shown near its paragraph
- Read Aloud reads with the voices the device has (the system's speech settings add more); a
  browser that lists only online voices reads nothing, by design. Words are highlighted where the
  voice says where it is (not every voice does). Hidden text is printed or left out for the whole
  document
- With no pages, columns are balanced (as Word balances them before a continuous section break), each
  header and footer is shown once, and adjacent paragraphs with the same borders get a border each
  (Word draws one box around them); a drop cap in the margin is shown like one in the text
- Legacy form fields: a text field's maximum length, type (number, date) and format, and a field's
  macros and calculations aren't applied; Word applies them when the form is next filled in there.
  Filling in fields isn't tracked. A signature is a picture of a signature, not a digital signature:
  it doesn't prove who signed (Word's signature lines and digital signatures are kept, not made here)
- Mail merge: in one document, headers and footers are shared by every copy, so a placeholder in
  them is filled in only when every row has the same value for it (the status bar says when one is
  left); a document with footnotes or endnotes is merged into separate documents only. Rows don't
  skip or filter (merge the rows wanted), and values are text as the CSV has it (dates and numbers
  as Excel saved them)
- The starter templates are in English (the manuscript's chapters are "Chapter One", "Chapter Two"),
  but for the naskah dinas, in Indonesian; their text is placeholders to type over (or Fill in), not
  a design to choose from. The naskah dinas follow the common layout of the rules for official
  documents; an office's own rules (a logo beside the letterhead, another font, the date's place)
  are set by typing over them, Insert picture and Page setup
- In Indonesian, the interface is translated where it is used most (see Features); dialogs such as
  Review your changes, Help, comments, mail merge, pictures' Layout Options and messages in the
  status bar are still in English, as is the About page
- Page numbers print in their section's format and from its start where the header or footer goes
  in the page's margin boxes (Chrome and Edge); elsewhere (see Printing above) the header and footer
  print once, and Word numbers the pages when it opens the file
- Placeholders are found in the page (its headers, text, notes and footers), not in comments, and
  only on one line and within one paragraph; a placeholder written as underscores (`________`)
  or as "XXX" isn't recognised. A new document made from a template doesn't record which template
  it came from (Word's attached template)
- Content controls: locked and data-bound ones, and anything but text fields around table cells
  or rows, are read-only
- The tests cover Chromium and Firefox; Safari is untested (find and comment highlighting need
  Safari 17.2 or later)
- Line numbers restart where Word last broke the pages, not where the edits move the breaks to; a
  section in columns has one column of numbers, at the left; right-to-left sections (`w:bidi` in a
  section) keep their numbers at the left and their columns in order
- Legacy binary `.doc` files are not supported, nor files encrypted by rights management
  (sensitivity labels that encrypt, IRM), with RC4 or for a certificate: they say so when opened,
  and Compare says so of a version it can't read. See "Password-protected files, older `.doc` files
  and rights management" above for what to do instead
- Password encryption: a mail merge's documents and a PDF made from a protected document aren't
  encrypted (protect each merged document before sending it, if it needs one); the redline is. A
  password to modify (Word's General Options, `w:writeProtection`) is a different thing, kept as it
  is (see Restrict Editing). Opening a protected document takes a moment (Word's 100,000 rounds of
  hashing, about half a second); the password isn't remembered, so restoring its unsaved changes or
  reopening it asks again. A lost password can't be recovered
- Compare documents compares the text, not formatting, pictures, comments, headers, footers or notes;
  a paragraph moved elsewhere shows as taken out in one place and added in another, and as many
  paragraphs taken out and added in one place are paired as changed. Its redline marks the text's
  changes only: formatting changes, pictures, tables' structure, and text in headers, footers,
  notes and text boxes aren't marked (they are this version's); a paragraph taken out comes back as
  plain text in its style (not its own direct formatting, numbering or pictures). Very long
  documents (tens of thousands of paragraphs that differ) are refused
- Sorting sorts by one column (Word's Sort has up to three), and only a whole table's rows below its
  heading row; formulas are those Word's Formula dialog writes (functions, cells, ranges, ABOVE and
  the like, arithmetic), not bookmarks' values, IF, AND, OR, NOT, DEFINED or TRUE/FALSE, and their
  results update with F9 rather than as you type, as in Word
- The Numbering library's lists are new lists: Word's Define New Multilevel List (other formats,
  linking levels to heading styles) isn't offered
- Restrict Editing applies to the whole document: the parts Word lets everyone (or named people)
  edit in a read-only document (`w:permStart`), and sections left unprotected in a form
  (`w:formProt`), are locked here too; "Edit anyway" opens everything. Formatting restrictions
  (allowed styles only) aren't enforced. Under "tracked changes only", the edits Intact doesn't track
  (in form fields, a picture's layout) aren't prevented in fields, and pictures can't be laid out
- "Before you release" doesn't remove dates from comments and tracked changes, document properties
  other than personal ones, custom XML data, embedded files or macros (it lists them), nor hidden
  text set by a character style or inside content Intact keeps as it is (it says so). A sensitivity
  label can't be applied or changed here
- The accessibility check reads the page as shown: it doesn't check reading order, colour alone
  conveying meaning, or charts, equations and objects kept as they are; a table's header row, the
  document's language and a picture's "decorative" mark are set in Word. Without a core properties
  part, a title can't be added here
- Signatures are recognised, not verified: Intact doesn't check that a signature is valid or who
  signed
- The page count is worked out, not Word's: the browser's fonts and line breaking, the last
  section's paper and margins for every section, footnotes not taking room at the bottom of a page,
  and rows kept whole where Word may split them. A paragraph's "keep with next" and "keep lines
  together" other than headings' aren't followed. On a phone, where the page takes the screen's
  width, pages aren't counted
- The PDF is the browser's: its fonts (a font the device lacks is substituted, as on screen), and
  links and a document outline as the browser makes them. Saving the PDF is the print window's
- Typed addresses become links as Word makes them: an address without `www.` or a scheme (such as
  linkedin.com/in/…) stays text, as in Word; use Ctrl+K for it. Network paths (`\\server\share`)
  aren't made links
- A document dropped on the page is handed to the new tab only while both are open, and a browser
  may block the new tab (allow pop-ups for the site); without `BroadcastChannel` the new tab opens
  on the start page, to open the file there. Without Web Locks (older browsers), a draft open in
  another tab isn't recognised as such
- Share… needs a browser that can share `.docx` files (the Web Share API with files): Safari on
  iPhone, iPad and Mac can; Chrome shares only some kinds of file, so where it can't share a `.docx`,
  Share… isn't offered and Download keeps the file on the device. A date that updates shows the day
  it was inserted here (Word updates it when it opens the file); dates already in a file as fields
  aren't updated here
- AutoCorrect replaces a word typed in one piece of text (not across formatting changes inside the
  word), with plain text (an entry can't carry formatting or a picture, as Word's formatted entries
  can), and a word to replace is one word, without spaces. Your entries are kept per browser: another
  browser or device doesn't have them until they are taken there (Export…, Import…). Word's own AutoCorrect
  list isn't read from the computer: the typos corrected are Intact's own list, in English
- The thesaurus is WordNet's: its meanings and words are a dictionary's, not always those a writer
  reaches for. Words are put in another form by rule and by a list of common irregular words: a
  rarer irregular word may come out regular ("slayed" for "slew"), a plural of several words takes it on its
  last ("levy en masses"), and a longer verb stressed at its end doubles its last letter only if it
  is a common one (permitted, preferred). Typed into the dialog, an -ed form is a past ("gave"),
  having no words before it to say otherwise. Indonesian words are offered as they are (its prefixes
  and suffixes aren't worked out). The Indonesian thesaurus is small (about 2,600
  meanings), as only Wordnet Bahasa's words checked by hand are in it: the rest, made automatically,
  often mix up meanings. Other languages are looked up in English. The English data makes the app's
  offline copy about 0.8 MB bigger (it is loaded only when the thesaurus is first used)
- Copy for the web leaves out pictures, text boxes, footnotes and endnotes, equations and comments
  (it says how many), and underline (on the web it reads as a link); a heading's own bold isn't
  copied as bold. Formatting is read from the page as shown, so a style's bold or italic comes
  through as bold or italic. The browser decides what a site's editor receives on paste: some
  editors keep only the plain text
- Readability scores are worked out as the Flesch formulas define them, with syllables counted by
  rule rather than from a dictionary, and the passive voice found by its form ("was approved"), so
  they are near Word's figures rather than the same. They are for English only; other languages get
  the counts and averages. Sentences end at ., ! and ? (not after common abbreviations such as
  "Mr." or "Nov."), and a paragraph without one is a sentence, as a headline is
- Redact blacks out text, not pictures (delete them), and the length of what was redacted stays
  visible, as on a page blacked out by hand. "Before you release" finds what was redacted only in
  this session, and only as it was selected (the same words in another form, such as "Ms Ruiz" for
  "Ana Ruiz", are not found: use Find). Text kept out of view (such as a field's instructions or a
  picture's alt text) isn't searched; check it in Word's Document Inspector
- The screening check knows what such software commonly misses, not what a given employer's does;
  it doesn't rate how well a CV matches a job
