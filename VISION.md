# Vision

> **Fix a Word document in seconds, privately: no install, no account, no upload — and nothing you
> didn't touch changes. Or write a new one, just as privately, that opens in Word as it looks here.**

This document says who the product is for, what it promises, and what it deliberately won't do.
Use it to decide feature requests: a change that doesn't serve the job below, or that weakens a
promise, needs a very good reason.

## The problem

People regularly get a `.docx` that needs a small change: a typo in a contract, a name in a
template, a date in a letter, a sentence in a CV. The usual options all have a cost:

- **Microsoft Word** needs a licence and an install, which many people don't have on the device in
  front of them (Chromebooks, Linux, borrowed or locked-down machines).
- **Google Docs and Word for the web** need an account and an upload. That's a real barrier for
  confidential documents such as contracts, HR files, medical or legal letters.
- **LibreOffice and other converters** often change things you didn't touch: fonts, spacing,
  numbering, tracked changes. That's a problem when the file goes back to someone who uses Word.

The same costs stand in the way of people who need to *write* a document for someone who uses
Word: a letter, a memo, minutes, a quote or invoice, a CV. Without Word on their device, they either
upload their words to a cloud editor or send a file that looks different when it is opened. And
people who have fixed a few documents here come to want to start the next one here too.

## Who it's for

1. **Privacy-conscious professionals** (legal, HR, finance, healthcare, public sector) who
   mustn't upload a document but need to correct it.
2. **People without Word** on the device they're using.
3. **Anyone sending a file back** to a Word user who will notice if anything else changed.
4. **People who write everyday documents without Word** (letters, memos, minutes, quotes,
   invoices, CVs, reports) and send them to people who use it.

## The job

Two jobs, with one test of success: the recipient can't tell another tool was used.

1. **Fix.** Open a Word file, change some text and basic formatting, and get back a file that is
   the original with **only those edits**.
2. **Write.** Write an everyday document, from a template or a blank page, and get a file that
   **opens in Word exactly as it looks here**.

Fixing came first and still sets the bar: what writing adds must not weaken what fixing promises.

## Promises

These are product requirements, not aspirations. A release that breaks one is a bug.

1. **Private.** The document never leaves the device. No server, no account, no analytics, and no
   third-party requests while a document is open.
2. **Faithful.** Anything the user did not edit is written back with the same meaning: formatting,
   images, tables, fields, content controls, comments, footnotes and tracked changes. Content the
   editor can't edit is shown read-only and kept, never silently changed or dropped. The round-trip
   tests (`tests/roundtrip.spec.js`) enforce this; every new capability adds cases to them. What is
   made here is written as Word writes it, so it shows in Word as it shows here.
3. **Instant.** Opening the page and a document takes seconds, with nothing to install or set up.
4. **Honest.** When something can't be edited or preserved, the user is told. The README and About
   page list known limitations, and the headline claims stay true.

## Non-goals

Saying no to these keeps the product small enough to keep its promises:

- **Complex authoring and desktop publishing.** Writing everyday documents is in (see The job);
  what goes beyond them is out: master and sub-documents, building a style system from nothing,
  linked text boxes and magazine layouts, mail merge from databases or address books (a CSV, which
  everyday mail merge needs, is in), macros and add-ins, and anything that only a print shop would
  need. A request for one of these needs a very good reason, as any request that risks the promises
  does.
- **Collaboration and cloud storage.** No accounts, sharing, comment threads or sync: those need
  a server, and they conflict with "private".
- **Other formats.** No spreadsheets, slides, PDF or legacy `.doc`. Do one format well.
- **Pixel-perfect pagination.** The editor shows content faithfully, not exact Word page breaks. It
  says about how many pages there are and about where they break, worked out and marked as such
  (≈), because whether a CV or a letter still fits on a page is part of fixing it.

## Principles for deciding what's next

- **Fidelity over features.** A feature that risks the round-trip waits until it can be tested.
- **A document made here is a Word document.** Everything new (a table, a style, a header, a
  section) is written as Word writes it, and each kind of thing made here is opened in Microsoft
  Word before release: tests that save a new document and check it as Word reads it, besides the
  round-trip tests, which check what was kept.
- **Writing grows from what fixing needs.** A writing feature comes after the fixing feature it is
  built on (styles applied before styles defined, page setup for the whole document before a
  section's), so the two jobs share one way of writing each thing.
- **Unlock what's locked before adding what's new.** The most valuable work is making preserved
  content editable where it matters for the job (for example the text inside content controls in
  templates), not adding new kinds of content.
- **Real documents decide.** Grow the test fixtures with documents written by Word, Google Docs and
  LibreOffice, and check output in Microsoft Word before each release.

## Direction

**Now — make the promises true and provable**
- Round-trip test suite in CI, gating deploys ✔
- Tracked changes preserved (read-only) ✔
- Undo/redo for all edits; live list numbering ✔
- Close the remaining known losses (inline `customXml`, `dir`/`bdo` wrappers) ✔
- Deleting text never silently drops tracked changes or leaves orphaned comments ✔
- Read-only content says what it is when clicked or typed into, instead of ignoring the keys ✔
- Test in Firefox ✔ (automated and by hand); output checked in Microsoft Word ✔ (by hand)
- Add real-world fixture documents (a letter written in Word ✔; reports, memos and minutes written
  by LibreOffice, pandoc and python-docx ✔; Google Docs welcome), each edited in every kind of place
  by the tests ✔; test Safari

**Next — do the core job better**
- Edit text inside content controls (templates and forms); tick boxes, pick from lists and dates ✔
- Start and end bulleted and numbered lists; Word's Enter, Backspace and Tab behaviour in lists and tables ✔
- Paste with basic formatting (bold, italic, lists) instead of plain text only ✔
- Show list and table spacing as Word does ✔
- Find and replace ✔
- Show comments, headers, footers and notes read-only, in plain language ✔

**Later — only if the job demands it**
- Track your own edits as tracked changes ✔; review your changes before downloading ✔; Word's
  keyboard shortcuts ✔
- Superscript, subscript and highlight ✔; show formatting marks (¶) ✔
- Accept or reject tracked changes ✔, including formatting changes and table rows ✔, and table, row
  and cell property changes ✔
- Add a comment ✔; reply to, edit, resolve and delete comments ✔; add a table row with Tab in the
  last cell ✔
- Comments in headers, footers and notes: reply, edit, resolve, delete ✔; tracked changes in locked
  content: accept and reject ✔
- Insert tables and pictures ✔; insert and delete table rows and columns ✔, tracked ✔; merge and
  split cells ✔, tracked ✔, field cells too ✔; tables whose rows or cells are
  fields ✔; resize, wrap and position pictures ✔; crop, rotate and add effects ✔; picture styles, 3-D and artistic effects ✔; 3-D lighting and all of Word's artistic effects ✔; wrap text to a
  picture's outline ✔
- Edit footnotes and endnotes ✔; show equations ✔
- Apply the document's own paragraph styles ✔; say when fields Word updates may be out of date ✔
- Name reference managers' citations ✔; keep pasted links and insert links (DOIs too) ✔; count the
  words selected ✔; check spelling in the document's language ✔
- Edit headers and footers ✔; edit the text in text boxes and shapes ✔; show designed documents as
  Word lays them out (columns, paragraph borders and shading, drop caps, character spacing, theme
  colours, embedded fonts) and name fonts shown with a substitute ✔
- Font, size, colour and character style from the document's own ✔; clear formatting ✔; track
  formatting changes ✔; smart quotes and dashes ✔; count characters ✔
- Respect Word's restrictions: Restrict Editing (read only, comments, forms, tracked changes),
  read-only recommended, digital signatures (kept valid when unchanged, removed when edited, as in
  Word), with an explicit "Edit anyway" ✔; say what rights-managed and `.doc` files are ✔; open
  password-encrypted files with their password, on the device ✔
- Show the sensitivity label or classification ✔; check for hidden information before release
  (comments, tracked changes, hidden text, personal information) and remove it ✔
- Record edits in the document's properties ✔; check accessibility before publishing ✔
- Installable offline app (PWA) ✔; the privacy promise enforced by a Content Security Policy ✔;
  versioned releases for self-hosting, with checksums, SBOM and build attestation ✔
- Right-to-left text ✔; templates and files with macros ✔; line numbers ✔
- Print just the document ✔; lists started by typing ("1. ", "* ") ✔; line and paragraph spacing
  and indents, with Word's Paragraph dialog ✔; Grow and Shrink Font ✔; AutoCorrect symbols ✔;
  zoom ✔
- Page numbers (and a footer for them) ✔; insert footnotes ✔ and endnotes ✔; page setup for the
  whole document ✔
- Update tables of contents, cross-references and caption numbers (F9) ✔; page numbers stay Word's
- About how many pages, and where they break, drawn over the page ✔; headers, footers and page
  numbers on every printed page ✔; save as PDF from the print window ✔; removing personal
  information as a default you choose ✔; check a CV for employers' screening software ✔
- Typed web and e-mail addresses made links, as Word's AutoFormat ✔; several documents at once, each
  in a tab of its own, with drafts that know where they are open ✔
- Fill in legacy forms (check boxes, drop-down lists and text fields), also when protected for
  filling in, with Tab between fields ✔; draw a signature on a form ✔
- Phones: the most used tools first, messages in view, and Share… to send the file back ✔; insert a
  check box or today's date ✔
- Templates for quotes, letters and invoices: fill in placeholders such as [Client name] once for
  every place they appear ✔; start a new document from a template, leaving the template as it is ✔
- A business's own look: a logo in the header or footer ✔; a brand colour by its code, remembered
  for the next document ✔; shading and borders for table cells (a quote's heading row and total) ✔
- Students' papers: End goes past a footnote's reference ✔; page numbers at the top or bottom, left,
  centre or right, and a header or footer to type in ("Surname 1") ✔; lengths in inches on Letter
  paper, or as chosen ✔
- References: a table of contents, captions and cross-references, equations typed as LaTeX,
  citations and a bibliography in APA or MLA (Word's own sources) ✔; the Navigation pane's
  headings ✔; a hanging indent as Word's dialog shows it ✔; the review's tracking and CV notes only
  where they apply ✔
- Today's date as a field Word updates ✔; page numbers shown as # and the page count as about how
  many pages ✔; labels on the toolbar's icon buttons where there is room ✔
- Tables pasted from Excel, Google Sheets and Word as tables ✔; Format Painter ✔; Change Case
  (Shift+F3) ✔; Insert Symbol ✔
- Tab stops in the Paragraph dialog, and tabs on the page at Word's stops ✔; a right-click menu ✔;
  Word's usual fonts after the document's own ✔; a DRAFT or CONFIDENTIAL watermark ✔

**Later — writing everyday documents**
- Starter templates on the start page: a letter, a memo, meeting minutes, a quote, an invoice, a CV
  and a report, each a Word document with placeholders to fill in ✔; each checked to be written in
  the order Word's schema requires, its parts related and typed ✔
- Mail merge from a CSV: one letter (or one document of them) for each row, through the
  placeholders ✔
- Page setup for one section: a landscape page in a portrait document, its own margins ✔
- Change a style's definition (Heading 1 in another font or colour, Normal's spacing), as Word's
  Modify Style ✔
- Design headers and footers: a different first page, different odd and even pages ✔
- Tests that save a new document of each kind and check it as Word reads it ✔ (starters)
- The tools in tabs, as Word's ribbon (Home, Insert, Layout, References, Review), with every tool in
  one toolbar still a choice ✔; the File and View tabs, where Word users look for saving, printing,
  protecting and checking the file, and for the Navigation pane and zoom ✔
- Long manuscripts and their editing rounds: indents in Modify Style (Normal with a first-line
  indent) ✔; Show markup, simple or none, to read an edited draft clean ✔; Previous and Next
  comment ✔; a manuscript starter in standard manuscript format ✔; the original (before the
  changes) ✔; move a chapter with its text from the Headings pane ✔; save back to the file opened,
  where the browser can ✔
- Public servants who write official documents (Indonesia first): F4 and Folio paper and a custom
  size ✔; page numbers in i, ii, iii or from any number, per section (Format page numbers) ✔; the
  interface in Indonesian ✔; starters for the naskah dinas (surat dinas, nota dinas, surat undangan,
  berita acara) ✔; Indonesian legal numbering (Pasal → ayat → huruf, BAB → Pasal) ✔;
  compare two versions of a document ✔; sort a table and table formulas (=SUM(ABOVE)) ✔; say what
  to do with older `.doc` and password-protected files ✔; print and save as PDF each section on its
  own paper, with its own header, footer and page numbers (a decree and its landscape attachment) ✔;
  cells merged down without a line between them ✔; starters for the other naskah an office writes
  most (keputusan with its lampiran, surat tugas, surat edaran, notula, laporan kegiatan, telaahan
  staf, surat kuasa), their points numbered as Word numbers them and their pages from the second ✔
- Teachers who write tests, worksheets and letters home (Indonesia first, and in English): columns,
  as Word's Columns, with continuous section breaks for part of a page ✔; templates for a test or
  lembar soal (multiple choice in two columns, numbered 1. and A. by Word, ruled answer lines), a
  worksheet, a rubric, a lesson plan or modul ajar, a kisi-kisi soal and a letter to parents ready
  for Mail merge from a class list ✔; table rows shown as tall as Word makes them ✔; equations with
  matrices, cases and working lined up at its = signs ✔; hidden text for answer keys, printed only
  when asked ✔; Read Aloud with voices on the device only ✔; tabs at Word's right, centre and
  decimal stops with their leaders, and Insert answer line ✔; lines, arrows and text boxes ✔
- Reporters who write to length and file stories that name confidential sources: a target in words
  and a headline's limit in characters, in the status bar ✔; Word count with Word's readability
  statistics and the long sentences to show ✔; Redact, never tracked, with "Before you release"
  saying where a redacted name is still in the file ✔. Not dictation or transcription: the browsers'
  speech recognition sends the voice to a server, which the privacy promise rules out
- Writing tools a reporter reaches for in Word: AutoCorrect entries of your own and common typos ✔;
  the thesaurus (Shift+F7), in English and Indonesian, on the device ✔; copy for the web (a CMS, an
  e-mail) or as Markdown, to the clipboard rather than as another file format ✔; synonyms in the form
  of the word they replace ✔; AutoCorrect entries exported and imported as a CSV file ✔
- Managers and executives who review and approve confidential papers (board papers, contracts,
  HR files): open documents encrypted with a password, decrypted on the device, and download them
  encrypted again, or protect any document with one ✔ (the password never leaves the device, so
  "Private" holds; it is the files that most need it); a redline from Compare, the changes since
  the other version as tracked changes, to send back as Word's Compare makes one ✔; on a phone, a
  document to review puts Accept, Reject and the comments first in the toolbar ✔

## How we'll know it's working

- **Fidelity:** zero known round-trip regressions; the fixture set keeps growing.
- **Trust:** users report that recipients didn't notice anything unusual about the file.
- **Speed:** a user goes from opening the page to downloading an edited file in under a minute.
- **Writing:** a document started here (from a template or a blank page) opens in Word looking as
  it did here, and its recipient can't tell another tool made it.

## Name

The product is called **Intact**, with the descriptor **"private .docx editor"** wherever there is
room for it ("Intact — private .docx editor"). The name states the "Faithful" promise: the document
comes back intact, with only your edits. It replaces "Office", which suggested a suite of
applications and invited trademark questions. Use this one name everywhere: the header, the page
title, the package name and the `Application` written into new documents.

Call it an **editor**, never a "word processor". Writing everyday documents is part of the job now,
but fixing documents is still most of what it is used for, and "editor" says so; "word processor"
would promise everything Word does, which the non-goals rule out. "Editor" also keeps the
descriptor clear of Word.

## Brand colour

The brand colour is **Intact Teal**, `#0e6467` (`--brand` and `--accent` in `src/style.css`, and
the icon). Teal says calm, safe and checked, and it sits apart from Word's blue and from the colours
that mean something in a document here: green and red (others' insertions and deletions), purple
(your tracked changes), coral (comments) and yellow (find). Don't use teal for any of those, and
don't bring Word's blue back into the interface. Document content keeps its own colours, such as
Word's hyperlink blue and highlight colours.

| Token | Light | Dark | Use |
|---|---|---|---|
| `--brand` | `#0e6467` | `#0b3f42` | Header bar, with white text |
| `--accent` | `#0e6467` | `#5ec4bd` | Links, selected and focused controls |
| `--active` | `#e0efee` | `#1d3b3c` | Selected item backgrounds |
| `--hover` | `#eef3f3` | `#263033` | Hover backgrounds |
| `--sheet-accent` / `--sheet-tint` / `--sheet-line` | `#0e6467` / `#e3f1f0` / `#9ccfcb` | same | Marks on the always-white page |
| `--own-change` | `#6f2dbd` | `#c4a3ff` | Your tracked changes in the interface (the status bar); 7.5:1 on the surface in both |

Every text pairing meets WCAG AA (white on `--brand` is 6.9:1 in light and 11.6:1 in dark).

## Logo

The mark is a page with a text cursor: the cursor is where you change what you need, the page is the
rest, kept whole. The cursor is also the I of Intact. The wordmark is Inter Display SemiBold,
outlined, so the logo needs no font.

| File | Use |
|---|---|
| `public/logo.svg` / `public/logo-dark.svg` | Mark and wordmark, on light / dark backgrounds |
| `public/favicon.svg` | Browser tab; follows the browser's light or dark theme |
| `public/favicon.ico` | 16, 32 and 48 px, for browsers and tools without SVG icons |
| `public/apple-touch-icon.png` | Home-screen icon, 180 px, white page on teal |
| `.github/social-preview.png` | Link previews (Settings → Social preview), 1280 × 640 |

In the app header the mark is drawn inline, reversed (white page, teal cursor). Keep clear space of
at least the cursor's width around the logo; don't recolour it outside the brand colours, add
effects, or put Word's "W" or blue near it.

## Talking about Word

Lead with what Intact does, not with Word. The start page headline is "Change what you need. Keep
the rest intact.", and its lead line names both jobs: fixing a `.docx`, or starting a new one ready
for Word. Mention Microsoft Word only to describe compatibility, in plain text and never in
a heading or the name: "`.docx` files, the format Microsoft Word uses", not "Edit Word documents".
