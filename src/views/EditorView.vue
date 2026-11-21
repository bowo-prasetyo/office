<template>
  <div class="editor-shell">
    <AppHeader>
      <div class="file">
        <input v-model="baseName" class="filename" :aria-label="t('File name')" spellcheck="false" :style="{ width: `${Math.max(baseName.length, 4) + 2}ch` }" />
        <span class="ext" :title="extTitle">.{{ fileExt }}</span>
        <span v-if="dirty" class="dirty" :title="t('Unsaved changes')">●</span>
        <span v-if="labelChip" class="label-chip" tabindex="0" role="note" :title="labelChip.title" :aria-label="labelChip.title">
          <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path d="M2 2.5h6l6 6-5.5 5.5-6-6z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /><circle cx="5.3" cy="5.8" r="1.1" fill="currentColor" /></svg>
          <span class="text">{{ labelChip.text }}</span>
        </span>
      </div>
      <span class="save-split">
        <button class="primary" type="button" :disabled="saving" :title="fileHandle ? `Save: write the changes into ${fileHandle.name}, the file opened, on this device (Ctrl+S)` : t('Download (Ctrl+S)')" @click="requestDownload()">
          {{ saving ? t('Saving…') : fileHandle ? t('Save') : t('Download') }}
        </button>
        <button ref="saveMenuButton" class="primary arrow" type="button" :title="canShareFiles ? t('More ways to save: a PDF, or share the file') : t('More ways to save: a PDF')" :aria-label="t('More ways to save')" aria-haspopup="menu" :aria-expanded="saveMenu.open" :disabled="saving" @click="toggleSaveMenu">▾</button>
      </span>
      <button class="ghost open-tab" type="button" :title="t('Open another document in a new tab, beside this one (Ctrl+O)')" :aria-label="t('Open another document in a new tab')" @click="openAnother">
        <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path d="M9 2.5h4.5V7M13.5 2.5 8 8M6.5 3.5H3.5a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1v-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" fill="none" /></svg>
        <span class="label">{{ t('Open') }}</span>
      </button>
      <button class="ghost close-doc" type="button" :aria-label="t('Close')" :title="t('Close the document')" @click="close">
        <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" /></svg>
        <span class="label">{{ t('Close') }}</span>
      </button>
    </AppHeader>

    <div v-if="restrictions.length" class="restriction-bar" :class="{ unlocked: unlocked || !lockMode }" role="region" :aria-label="t('Restrictions on this document')">
      <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="3" y="7" width="10" height="7" rx="1" stroke="currentColor" stroke-width="1.4" fill="none" /><path v-if="unlocked" d="M5.5 7V5a2.5 2.5 0 0 1 5-.6" stroke="currentColor" stroke-width="1.4" fill="none" /><path v-else d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" stroke="currentColor" stroke-width="1.4" fill="none" /></svg>
      <p>
        <template v-for="(r, i) in restrictions" :key="i"><strong>{{ r.title }}</strong>{{ ` ${r.text} ` }}</template>
        <span>{{ unlocked ? unlockedText : lockText }}</span>
      </p>
      <button v-if="lockMode" type="button" @click="editAnyway">{{ t('Edit anyway') }}</button>
    </div>

    <!-- Encrypted with a password (Word's Encrypt with Password): it downloads encrypted with it. -->
    <div v-if="passwordProtected" class="restriction-bar unlocked password-bar" role="region" :aria-label="t('Password protection')">
      <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="3" y="7" width="10" height="7" rx="1" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" stroke="currentColor" stroke-width="1.4" fill="none" /><circle cx="8" cy="10.5" r="1" fill="currentColor" /></svg>
      <p><strong>{{ t('Protected with a password.') }}</strong> {{ t('It downloads encrypted with it, as Word saves it, and opens only with it. Unsaved changes kept on this device are encrypted too.') }}</p>
      <button type="button" aria-haspopup="dialog" @click="openPasswordDialog()">{{ t('Change or remove password') }}</button>
    </div>

    <!-- A template: saved as itself, or made into a new document as Word does when one is opened
         to write with. Either way the template file opened is never changed. -->
    <div v-if="doc.openedFormat.template" class="restriction-bar unlocked template-bar" role="region" :aria-label="t('Template')">
      <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3.5 1.8h6.2l3 3v9.4h-9.2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M5.8 7.5h4.4M5.8 9.7h4.4M5.8 11.9h2.6" stroke="currentColor" stroke-width="1.1" stroke-linecap="round" stroke-dasharray="1.4 1.1" /></svg>
      <p v-if="fromTemplate"><strong>{{ t('New document from this template.') }}</strong> It downloads as a .{{ fileExt }}; the template ({{ baseName }}.{{ doc.openedFormat.ext }}) stays as it is, ready for the next one.</p>
      <p v-else><strong>{{ t('Word template.') }}</strong> Downloading saves the template itself. To write a quote, letter or invoice with it, start a new document: your template stays as it is.</p>
      <button v-if="placeholderCount && !textLocked" type="button" @click="openFillIn()">Fill in {{ placeholderCount }} {{ placeholderCount === 1 ? 'placeholder' : 'placeholders' }}…</button>
      <button type="button" @click="setFromTemplate(!fromTemplate)">{{ fromTemplate ? 'Edit the template instead' : 'Start a new document' }}</button>
    </div>

    <!-- One toolbar, one Tab stop: arrow keys move along it, Esc goes back to the text, Alt+F10 comes
         back to it (see editor/toolbar.js). Its controls stay where they are: none appears, goes or
         loses its label when the document gets its first tracked change. -->
    <!-- The ribbon, as Word's: the tools in tabs (Home, Insert, Layout, References, Review), one tab's
         at a time, below the tabs; Undo, Redo and Help in every tab. "All tools" shows them all in
         one toolbar instead (remembered). A phone keeps its one row of tools that scrolls. -->
    <div v-if="ribbon" class="ribbon-tabs" role="tablist" :aria-label="t('Toolbar tabs')" @keydown="onRibbonKeys" @mousedown="keepFocus">
      <button
        v-for="tab in RIBBON_TABS"
        :id="`ribbon-tab-${tab.id}`"
        :key="tab.id"
        type="button"
        role="tab"
        :class="{ context: tab.id === 'layout' && state.table && ribbonTab !== 'layout' }"
        :aria-selected="ribbonTab === tab.id"
        :tabindex="ribbonTab === tab.id ? 0 : -1"
        aria-controls="editor-toolbar"
        :title="t(tab.title)"
        @click="setRibbonTab(tab.id)"
      >{{ t(tab.label) }}<span v-if="tab.id === 'review' && changes.total && ribbonTab !== 'review'" class="tab-badge" :aria-label="t(changes.total === 1 ? ', {n} tracked change' : ', {n} tracked changes', { n: changes.total })">{{ changes.total }}</span></button>
      <span class="ribbon-spacer" />
      <button type="button" class="ribbon-mode" :title="t('Show every tool at once, in one toolbar without tabs')" @click="setRibbon(false)">{{ t('All tools') }}</button>
    </div>
    <div id="editor-toolbar" ref="toolbarEl" class="toolbar" :class="{ ribbon, 'review-first': reviewFirst }" :data-active="ribbon ? ribbonTab : null" role="toolbar" :aria-label="ribbon ? t('{tab} tools', { tab: t(RIBBON_TABS.find((x) => x.id === ribbonTab).label) }) : t('Editing and review')" :aria-labelledby="ribbon ? `ribbon-tab-${ribbonTab}` : null" @mousedown="keepFocus">
      <div class="tb-group">
      <fieldset class="group" :disabled="textLocked" :title="textLocked ? lockMessage : null">
      <button type="button" class="ph ur" :title="t('Undo (Ctrl+Z)')" :aria-label="t('Undo')" @click="undo">↶</button>
      <button type="button" class="ph ur" :title="t('Redo (Ctrl+Y)')" :aria-label="t('Redo')" @click="redo">↷</button>
      <span class="sep" />
      <select data-tab="home" class="style-choice" :value="state.block" :aria-label="t('Paragraph style')" :title="t('Paragraph style (Ctrl+Shift+N, Ctrl+Alt+1–3)')" @change="onStyleMenu">
        <option value="normal">{{ t('Normal') }}</option>
        <option value="title">{{ t('Title') }}</option>
        <option value="h1">{{ t('Heading 1') }}</option>
        <option value="h2">{{ t('Heading 2') }}</option>
        <option value="h3">{{ t('Heading 3') }}</option>
        <optgroup v-if="styleChoices.length" :label="t('In this document')">
          <option v-for="c in styleChoices" :key="c.id" :value="`style:${c.id}`">{{ c.label }}</option>
        </optgroup>
        <option v-if="!knownBlock" :value="state.block" disabled>{{ doc.styleLabel(state.block.slice(6)) }}</option>
        <option value="__modify">{{ t('Modify style…') }}</option>
      </select>
      <select data-tab="home" class="font-choice" :value="fontValue" :aria-label="t('Font')" :title="t('Font: the document’s own fonts first, then Word’s usual ones')" @change="applyFont($event.target.value)">
        <option v-if="!fontValue" value="" disabled hidden>{{ state.font }}</option>
        <optgroup label="Theme fonts">
          <option v-for="f in fontChoices.filter((x) => x.group === 'theme')" :key="f.value" :value="f.value">{{ f.label }}</option>
        </optgroup>
        <optgroup v-if="fontChoices.some((x) => x.group === 'document')" :label="t('In this document')">
          <option v-for="f in fontChoices.filter((x) => x.group === 'document')" :key="f.value" :value="f.value">{{ f.label }}</option>
        </optgroup>
        <optgroup label="Other fonts">
          <option v-for="f in fontChoices.filter((x) => x.group === 'standard')" :key="f.value" :value="f.value" :style="{ fontFamily: `'${f.name}'` }">{{ f.label }}</option>
        </optgroup>
      </select>
      <select data-tab="home" class="size-choice" :value="sizeValue" :aria-label="t('Font size')" :title="t('Font size')" @change="applyFontSize(Number($event.target.value))">
        <option v-if="!sizeValue" value="" disabled hidden>{{ state.size }}</option>
        <optgroup v-if="sizeChoices.used.length" :label="t('In this document')">
          <option v-for="z in sizeChoices.used" :key="`u${z}`" :value="String(z)">{{ z }}</option>
        </optgroup>
        <optgroup label="Other sizes">
          <option v-for="z in sizeChoices.others" :key="`o${z}`" :value="String(z)">{{ z }}</option>
        </optgroup>
      </select>
      <button data-tab="home" type="button" :title="t('Grow font (Ctrl+Shift+>)')" :aria-label="t('Grow font')" @click="growFont(1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M1.5 13.5l3.5-10 3.5 10M2.7 10h4.6M11 6.5l2-2.5 2 2.5" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /></svg>
      </button>
      <button data-tab="home" type="button" :title="t('Shrink font (Ctrl+Shift+<)')" :aria-label="t('Shrink font')" @click="growFont(-1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 13.5l3-8.5 3 8.5M3.1 10.5h3.8M11 4l2 2.5 2-2.5" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /></svg>
      </button>
      <select data-tab="home" v-if="characterStyles.length" class="char-style" :value="state.charStyle" :aria-label="t('Character style')" :title="t('Character style: the document’s own')" @change="applyCharacterStyle($event.target.value)">
        <option value="">{{ t('No character style') }}</option>
        <option v-for="c in characterStyles" :key="c.id" :value="c.id">{{ c.label }}</option>
        <option v-if="state.charStyle && !characterStyles.some((c) => c.id === state.charStyle)" :value="state.charStyle" disabled>{{ doc.styleLabel(state.charStyle) }}</option>
      </select>
      <span class="sep" data-tab="home" />
      <button data-tab="home" type="button" :title="t('Bold (Ctrl+B)')" :aria-label="t('Bold')" :class="{ on: state.bold, ph: true }" :aria-pressed="state.bold" @click="exec('bold')"><b>B</b></button>
      <button data-tab="home" type="button" :title="t('Italic (Ctrl+I)')" :aria-label="t('Italic')" :class="{ on: state.italic, ph: true }" :aria-pressed="state.italic" @click="exec('italic')"><i>I</i></button>
      <button data-tab="home" type="button" :title="t('Underline (Ctrl+U)')" :aria-label="t('Underline')" :class="{ on: state.underline, ph: true }" :aria-pressed="state.underline" @click="exec('underline')"><u>U</u></button>
      <button data-tab="home" type="button" :title="t('Strikethrough')" :aria-label="t('Strikethrough')" :class="{ on: state.strike }" :aria-pressed="state.strike" @click="exec('strikeThrough')"><s>S</s></button>
      <button data-tab="home" type="button" :title="t('Superscript (Ctrl+Shift+=)')" :aria-label="t('Superscript')" :class="{ on: state.sup }" :aria-pressed="state.sup" @click="toggleVertAlign('superscript')"><span class="script">x<sup>2</sup></span></button>
      <button data-tab="home" type="button" :title="t('Subscript (Ctrl+=)')" :aria-label="t('Subscript')" :class="{ on: state.sub }" :aria-pressed="state.sub" @click="toggleVertAlign('subscript')"><span class="script">x<sub>2</sub></span></button>
      <span data-tab="home" class="split ph">
        <button type="button" :title="`Highlight: ${colorLabel(highlightColor)} (Ctrl+Alt+H)`" :class="{ on: !!state.highlight }" @click="applyHighlight(highlightColor, true)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4 9.5l5.5-6 3 3-6 5.5h-2.5zM4 9.5l-1.5 2.5h3" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><rect x="1.5" y="13" width="13" height="2.2" :fill="HIGHLIGHT[highlightColor]" stroke="currentColor" stroke-width=".4" /></svg>
        </button>
        <button ref="paletteButton" type="button" class="arrow" :title="t('Highlight colour')" :aria-label="t('Highlight colour options')" aria-haspopup="dialog" :aria-expanded="palette.open" @click="togglePalette">▾</button>
      </span>
      <span data-tab="home" class="split">
        <button type="button" :title="`Font colour: ${fontColor ? fontColor.label : 'Automatic'}`" @click="applyFontColor(fontColor)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4.5 11.5l3.5-9 3.5 9M5.7 8.5h4.6" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /><rect x="1.5" y="13" width="13" height="2.2" :fill="fontColor ? `#${fontColor.hex}` : '#000'" stroke="currentColor" stroke-width=".4" /></svg>
        </button>
        <button ref="colorButton" type="button" class="arrow" :title="t('Font colour: the document’s theme and its own colours')" :aria-label="t('Font colour options')" aria-haspopup="dialog" :aria-expanded="colorPalette.open" @click="toggleColorPalette">▾</button>
      </span>
      <button data-tab="home" type="button" :title="t('Clear formatting (Ctrl+Space): takes the text’s direct formatting and character style off; highlight stays')" @click="clearCharacterFormatting">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3 3h8M7 3l-2.5 9M9.5 9.5l4 4M13.5 9.5l-4 4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" fill="none" /></svg>
      </button>
      <button data-tab="home" type="button" class="painter-toggle" :title="painter.active ? 'Format Painter is on: select (or click) the text to paint the formatting on – Esc or a click here stops it' : 'Format Painter (Ctrl+Shift+C, Ctrl+Shift+V): copies the formatting at the caret to the text you select next; double-click to paint it on several places'" :aria-label="t('Format Painter')" :class="{ on: painter.active }" :aria-pressed="painter.active" @click="togglePainter" @dblclick="startPainter(true)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="2" y="1.8" width="10" height="4.4" rx=".8" stroke="currentColor" stroke-width="1.3" fill="none" /><path d="M12 4h1.8v3.6H7.6v2.2" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><rect x="6.4" y="9.8" width="2.4" height="4.6" rx=".6" fill="currentColor" /></svg>
      </button>
      <button data-tab="home" ref="caseButton" type="button" :title="t('Change case (Shift+F3): Sentence case, lowercase, UPPERCASE, Capitalize Each Word')" :aria-label="t('Change case')" aria-haspopup="menu" :aria-expanded="caseMenu.open" @click="toggleCaseMenu">
        <span class="case-icon" aria-hidden="true">{{ t('Aa') }}</span>
      </button>
      <span class="sep" data-tab="home" />
      <button data-tab="home" v-for="a in aligns" :key="a.value" type="button" :title="a.title" :aria-label="a.label" :class="{ on: state.align === a.value }" :aria-pressed="state.align === a.value" @click="alignBlocks(a.value)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path :d="a.icon" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <button data-tab="home" ref="spacingButton" type="button" :title="t('Line and paragraph spacing (Ctrl+1, Ctrl+5, Ctrl+2)')" :aria-label="t('Line and paragraph spacing')" aria-haspopup="menu" :aria-expanded="spacingMenu.open" @click="toggleSpacingMenu">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M7.5 3.5h6.5M7.5 8h6.5M7.5 12.5h6.5M3.5 2.5v11M1.8 4.2l1.7-1.7 1.7 1.7M1.8 11.8l1.7 1.7 1.7-1.7" stroke="currentColor" stroke-width="1.4" fill="none" /></svg>
      </button>
      <span class="sep" data-tab="home" />
      <button data-tab="home" type="button" :title="t('Bullets (Ctrl+Shift+L)')" :class="{ on: state.list === 'bullet', ph: true }" :aria-pressed="state.list === 'bullet'" @click="toggleList('bullet')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6.5 4h8M6.5 8h8M6.5 12h8M2.5 4h.01M2.5 8h.01M2.5 12h.01" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" fill="none" /></svg>
      </button>
      <span data-tab="home" class="split ph">
        <button type="button" :title="t('Numbering')" :class="{ on: state.list === 'number', ph: true }" :aria-pressed="state.list === 'number'" @click="toggleList('number')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6.5 4h8M6.5 8h8M6.5 12h8M2 2.5h1.2v3M1.6 5.5h2M1.6 9.6q0-1.6 1.6-1 .9.6-1.6 3h2" stroke="currentColor" stroke-width="1.1" fill="none" /></svg>
      </button>
        <button ref="listLibraryButton" type="button" class="arrow" :title="t('Multilevel lists: Pasal, ayat and huruf; BAB and Pasal; 1.1.1')" :aria-label="t('Numbering library')" aria-haspopup="menu" :aria-expanded="listLibrary.open" @click="toggleListLibrary">▾</button>
      </span>
      <button data-tab="home" type="button" :title="state.list ? 'Decrease list level (Shift+Tab, Ctrl+Shift+M)' : 'Decrease indent (Ctrl+Shift+M)'" :aria-label="t('Decrease indent')" @click="indentBlocks(-1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M7 4h7M7 8h7M7 12h7M4.5 5.5L2 8l2.5 2.5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <button data-tab="home" type="button" :title="state.list ? 'Increase list level (Tab, Ctrl+M)' : 'Increase indent (Ctrl+M)'" :aria-label="t('Increase indent')" @click="indentBlocks(1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M7 4h7M7 8h7M7 12h7M2 5.5L4.5 8 2 10.5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <span class="sep" data-tab="home" />
      <button data-tab="insert" ref="linkButton" type="button" :title="t('Link (Ctrl+K): insert a link to a web address, e-mail address or DOI, or change the one here')" @click="openLinkDialog">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6.6 9.4l2.8-2.8M7.2 4.6l1.1-1.1a2.6 2.6 0 0 1 3.7 3.7l-1.1 1.1M8.8 11.4l-1.1 1.1a2.6 2.6 0 0 1-3.7-3.7l1.1-1.1" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" fill="none" /></svg>
      </button>
      <button data-tab="insert" ref="insertButton" type="button" class="ph" :title="t('Insert a check box or today’s date, or a reference: a table of contents, caption, cross-reference, equation, citation or bibliography')" :aria-label="t('Insert a check box or date, or a reference')" aria-haspopup="menu" :aria-expanded="insertMenu.open" @click="toggleInsertMenu">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="1.8" y="2.8" width="7.4" height="7.4" rx="1" stroke="currentColor" stroke-width="1.3" fill="none" /><path d="M3.4 6.4l1.5 1.5 2.8-3" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" fill="none" /><rect x="6.8" y="7.6" width="7.4" height="6.6" rx="1" stroke="currentColor" stroke-width="1.3" fill="var(--surface)" /><path d="M6.8 9.9h7.4M9 6.6v2M12 6.6v2" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" /></svg><span class="short-label">{{ t('Insert') }}</span>
      </button>
      <button data-tab="insert" type="button" :title="t('Insert a symbol: ©, €, ✓, →, ½ and others, by group or name')" :aria-label="t('Insert symbol')" aria-haspopup="dialog" :aria-expanded="symbolOpen" @click="openSymbols">
        <span class="symbol-icon" aria-hidden="true">Ω</span><span class="short-label">{{ t('Symbol') }}</span>
      </button>
      <button data-tab="insert" type="button" class="ph" :title="t('Insert picture')" :aria-label="t('Insert picture')" @click="pickPicture">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="1.8" y="2.8" width="12.4" height="10.4" rx="1" stroke="currentColor" stroke-width="1.4" fill="none" /><circle cx="5.6" cy="6.3" r="1.3" fill="currentColor" /><path d="M2.5 12.3l3.8-3.8 2.8 2.8 2-2 2.6 2.6" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ t('Picture') }}</span>
      </button>
      <button data-tab="insert" type="button" class="ph" :title="t('Draw signature: sign with your finger, a pen or the mouse, and put it at the caret')" :aria-label="t('Draw signature')" aria-haspopup="dialog" :aria-expanded="signatureOpen" @click="openSignature">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M1.5 11.5c1.6-3.6 3-5.6 3.9-5.2 1.2.5-1.6 5 .1 5.1 1.2.1 2.1-2.6 3.1-2.5.8.1.3 1.9 1.2 2 .6 0 1.1-.6 1.6-1.1" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" fill="none" /><path d="M1.5 14h13" stroke="currentColor" stroke-width="1.1" stroke-linecap="round" /><path d="M11.2 6.6l3-3-1.3-1.3-3 3-.4 1.7z" stroke="currentColor" stroke-width="1.1" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ t('Sign') }}</span>
      </button>
      <input ref="pictureInput" type="file" accept="image/png,image/jpeg,image/gif,image/bmp" multiple hidden @change="onPicturesPicked" />
      <button data-tab="insert" ref="tableButton" type="button" :title="t('Insert table')" :aria-label="t('Insert table')" aria-haspopup="dialog" :aria-expanded="tablePicker.open" @click="toggleTablePicker">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="1.8" y="2.3" width="12.4" height="11.4" rx="1" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M1.8 6h12.4M1.8 9.8h12.4M6 2.3v11.4M10 2.3v11.4" stroke="currentColor" stroke-width="1.1" fill="none" /></svg><span class="short-label">{{ t('Table') }}</span>
      </button>
      <button data-tab="layout" ref="layoutButton" type="button" :title="t('Table layout: insert or delete rows and columns')" :aria-label="t('Table layout')" aria-haspopup="menu" :aria-expanded="tableMenu.open" :disabled="!state.table" @click="toggleTableMenu">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M9 2.3H2.8a1 1 0 00-1 1v9.4a1 1 0 001 1h9.4a1 1 0 001-1V8" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M1.8 8h7M6 2.3v11.4" stroke="currentColor" stroke-width="1.1" fill="none" /><path d="M12.5 1.5v5M10 4h5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" /></svg><span class="short-label">{{ t('Layout') }}</span>
      </button>
      <button data-tab="references" type="button" :title="t('Insert footnote (Ctrl+Alt+F)')" :aria-label="t('Insert footnote')" @click="insertFootnote">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M1.5 9.5h7M1.5 13h9M11 2.5h1.5v4.5M10.8 7h3.2" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" fill="none" /></svg><span class="short-label">{{ t('Footnote') }}</span>
      </button>
      <button data-tab="references" type="button" :title="t('Insert endnote (Ctrl+Alt+D)')" :aria-label="t('Insert endnote')" @click="insertEndnote">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M1.5 4h7M1.5 7.5h5M1.5 13.5h13M11.5 2.5v4.5M13.5 2.5v4.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" fill="none" /></svg><span class="short-label">{{ t('Endnote') }}</span>
      </button>
      <button data-tab="references" type="button" :title="t('Update fields (F9): the table of contents, cross-references, caption and note numbers – page numbers are Word’s')" :aria-label="t('Update fields')" @click="updateAllFields">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M13 6.5A5.2 5.2 0 0 0 3.6 4.6M3 9.5a5.2 5.2 0 0 0 9.4 1.9M3.4 2v2.8h2.8M12.6 14v-2.8H9.8" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ t('Update') }}</span>
      </button>
      <span data-tab="layout" class="split">
        <button type="button" :title="t('Page number: in the header or footer at the caret, or a footer with the page number')" :aria-label="t('Page number')" @click="insertPageNumber()">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3.5 1.8h6.2l3 3v9.4h-9.2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M7 9.3h2.4M6.6 11.2l.6-3.8M8.8 11.2l.6-3.8M6.4 7.8h3.2" stroke="currentColor" stroke-width="1" stroke-linecap="round" fill="none" /></svg><span class="short-label">{{ t('Page no.') }}</span>
        </button>
        <button ref="pageNumberButton" type="button" class="arrow" :title="t('Page number at the top or bottom of the page; edit the header or footer')" :aria-label="t('Page number and header options')" aria-haspopup="menu" :aria-expanded="pageNumberMenu.open" @click="togglePageNumberMenu">▾</button>
      </span>
      <button data-tab="layout" ref="pageSetupButton" type="button" :title="t('Page setup: paper size, orientation and margins')" :aria-label="t('Page setup')" aria-haspopup="dialog" :aria-expanded="pageDialog.open" @click="openPageSetup">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="3" y="1.8" width="10" height="12.4" rx=".8" stroke="currentColor" stroke-width="1.3" fill="none" /><rect x="5.2" y="4" width="5.6" height="8" stroke="currentColor" stroke-width=".9" stroke-dasharray="1.2 1" fill="none" /></svg><span class="short-label">{{ t('Page setup') }}</span>
      </button>
      </fieldset>
      </div>
      <div class="tb-group">
      <!-- The File tab, as Word's File menu (its Backstage): saving, printing, sharing and protecting the
           file, and checking it before it goes. The same as Download's menu (▾), so they show in the
           ribbon's File tab only, not in "All tools" or on a phone, where that menu has them. -->
      <button data-tab="file" type="button" :title="fileHandle ? `Save to ${fileHandle.name} (Ctrl+S)` : t('Download the Word document (Ctrl+S)')" :aria-label="fileHandle ? t('Save') : t('Download')" @click="saveMenuAction('docx')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M8 2v8M4.8 7 8 10.2 11.2 7M2.5 12.5v1.2h11v-1.2" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ fileHandle ? t('Save') : t('Download') }}</span>
      </button>
      <button v-if="fileHandle" data-tab="file" type="button" :title="t('Download a copy, to Downloads')" :aria-label="t('Download a copy')" @click="saveMenuAction('copy')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M8 2v7M5 6.2 8 9.2l3-3M2.5 12.5v1.2h11v-1.2" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" fill="none" /><circle cx="13" cy="3" r="1.5" fill="currentColor" /></svg><span class="short-label">{{ t('Copy') }}</span>
      </button>
      <button v-if="canSaveAs" data-tab="file" type="button" :title="t('Save as: a file on this device, which Save writes to after')" :aria-label="t('Save as')" @click="saveMenuAction('saveAs')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 2.5h9l2 2v9h-11z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M5 2.5v3.5h5V2.5M5 13.5V9.5h6v4" stroke="currentColor" stroke-width="1.2" fill="none" /></svg><span class="short-label">{{ t('Save as') }}</span>
      </button>
      <span class="sep" data-tab="file" />
      <button data-tab="file" type="button" :title="t('Print (Ctrl+P): the document on its own paper, without the editor')" :aria-label="t('Print')" @click="printDocument">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4.5 6V2.5h7V6M4.5 11.5h-2V6h11v5.5h-2" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M4.5 9.5h7v4h-7z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ t('Print') }}</span>
      </button>
      <button data-tab="file" type="button" :title="t('Save as PDF, from the print window')" :aria-label="t('Save as PDF')" aria-haspopup="dialog" @click="saveMenuAction('pdf')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3.5 1.8h6.2l3 3v9.4h-9.2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M5.5 11.5V8h1.2a1 1 0 0 1 0 2H5.5" stroke="currentColor" stroke-width="1" fill="none" /></svg><span class="short-label">{{ t('PDF') }}</span>
      </button>
      <button v-if="canShareFiles" data-tab="file" type="button" :title="t('Share the file: mail, messages…')" :aria-label="t('Share')" @click="saveMenuAction('share')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="12" cy="3.5" r="1.8" stroke="currentColor" stroke-width="1.2" fill="none" /><circle cx="4" cy="8" r="1.8" stroke="currentColor" stroke-width="1.2" fill="none" /><circle cx="12" cy="12.5" r="1.8" stroke="currentColor" stroke-width="1.2" fill="none" /><path d="M5.6 7.1 10.4 4.4M5.6 8.9l4.8 2.7" stroke="currentColor" stroke-width="1.2" /></svg><span class="short-label">{{ t('Share') }}</span>
      </button>
      <button data-tab="file" type="button" :title="t('Copy for the web: the document, for a website’s editor (CMS) or an e-mail')" :aria-label="t('Copy for the web')" @click="saveMenuAction('web')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="8" cy="8" r="5.8" stroke="currentColor" stroke-width="1.2" fill="none" /><path d="M2.2 8h11.6M8 2.2c-2 2-2 9.6 0 11.6M8 2.2c2 2 2 9.6 0 11.6" stroke="currentColor" stroke-width="1.1" fill="none" /></svg><span class="short-label">{{ t('Web') }}</span>
      </button>
      <button data-tab="file" type="button" :title="t('Mail merge: a letter for each row of a list (a CSV from Excel or Google Sheets)')" :aria-label="t('Mail merge')" aria-haspopup="dialog" @click="saveMenuAction('merge')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M1.8 4h9.4v7H1.8z" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round" fill="none" /><path d="M1.8 4l4.7 3.6L11.2 4" stroke="currentColor" stroke-width="1.2" fill="none" /><path d="M5 13.2h9.2V6.5" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" fill="none" /></svg><span class="short-label">{{ t('Mail merge') }}</span>
      </button>
      <span class="sep" data-tab="file" />
      <button data-tab="file" type="button" :title="passwordProtected ? t('Change or remove the password the file is encrypted with') : t('Protect with password: encrypt the file, as Word’s Encrypt with Password')" :aria-label="passwordProtected ? t('Change or remove password') : t('Protect with password')" aria-haspopup="dialog" :class="{ on: passwordProtected }" :aria-pressed="passwordProtected" @click="saveMenuAction('password')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="3" y="7" width="10" height="7" rx="1" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" stroke="currentColor" stroke-width="1.4" fill="none" /></svg><span class="short-label">{{ t('Protect') }}</span>
      </button>
      <button data-tab="file" type="button" :title="t('Inspect document: what the file carries beyond the page (comments, tracked changes, hidden text, personal information), before you release it')" :aria-label="t('Inspect document')" aria-haspopup="dialog" @click="openReview('release')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="7" cy="7" r="4.2" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M10.1 10.1 14 14" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" /><path d="M5.2 7h3.6" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" /></svg><span class="short-label">{{ t('Inspect') }}</span>
      </button>
      <button data-tab="file" type="button" :title="t('Check accessibility: alt text, headings, tables, contrast and the like, before publishing')" :aria-label="t('Check accessibility')" aria-haspopup="dialog" @click="openReview('accessibility')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="8" cy="2.8" r="1.4" fill="currentColor" /><path d="M2.8 5.6 8 6.8l5.2-1.2M8 6.8v3.4M8 10.2l-2.4 4M8 10.2l2.4 4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ t('Accessibility') }}</span>
      </button>
      <span class="sep" data-tab="classic" />
      <button data-tab="home" type="button" :title="t('Find and replace (Ctrl+F, Ctrl+H)')" :aria-label="t('Find and replace')" :class="{ on: findOpen }" :aria-pressed="findOpen" @click="findOpen ? closeFind() : openFind(false)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="6.5" cy="6.5" r="4.2" stroke="currentColor" stroke-width="1.6" fill="none" /><path d="M9.6 9.6L14 14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" /></svg><span class="short-label">{{ t('Find') }}</span>
      </button>
      <button data-tab="insert" ref="fillInButton" type="button" class="fill-in-toggle" :title="textLocked ? `Fill in placeholders – ${lockMessage}` : placeholderCount ? `Fill in placeholders (${placeholderCount}): [Client name], {{date}} and the like, each once for everywhere it appears` : 'Fill in placeholders: [Client name], {{date}} and the like, each once for everywhere it appears'" :aria-label="t('Fill in placeholders')" aria-haspopup="dialog" :aria-expanded="fillIn.open" :class="{ on: fillIn.open }" :disabled="textLocked" @click="fillIn.open ? closeFillIn() : openFillIn()">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4 3H2.5v10H4M12 3h1.5v10H12" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" fill="none" /><path d="M5.5 10.5l4.8-4.8 1.2 1.2-4.8 4.8H5.5z" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ t('Fill in') }}</span>
        <span v-if="placeholderCount" class="badge">{{ placeholderCount }}</span>
      </button>
      <button data-tab="home view" type="button" class="marks-toggle" :title="t('Show formatting marks (Ctrl+Shift+8)')" :aria-label="t('Show formatting marks')" :class="{ on: showMarks }" :aria-pressed="showMarks" @click="toggleMarks">¶</button>
      <span class="sep" data-tab="review" />
      <button data-tab="review" ref="trackButton" type="button" class="track-toggle" :title="lockMode ? `Track Changes: ${lockMode === 'trackedChanges' ? 'always on' : tracking ? 'on' : 'off'} – ${lockMessage}` : 'Track Changes (Ctrl+Shift+E)'" :class="{ on: tracking }" :aria-pressed="tracking" :disabled="!!lockMode" @click="toggleTracking">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 13.5h5M2 10.5h3M10.5 2.5l3 3-6 6H4.5v-3z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" fill="none" /></svg>
        <span class="label">{{ t('Track changes') }}</span>
      </button>
      <span data-tab="review" class="changes-group rv" role="group" :aria-label="t('Tracked changes')">
        <button type="button" :title="t('Previous change')" :disabled="!changes.total" @click="gotoChange(-1)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M10 3L5 8l5 5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
        </button>
        <button type="button" :title="t('Next change')" :disabled="!changes.total" @click="gotoChange(1)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6 3l5 5-5 5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
        </button>
        <fieldset class="group" :disabled="!!lockMode" :title="lockMode ? 'Tracked changes can’t be accepted or rejected: ' + lockMessage : null">
        <span class="split">
          <button type="button" class="labelled" :title="t('Accept the change (or the changes selected) and go to the next')" :disabled="!changes.total" @click="resolveCurrent(true)">
            <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3 8.5l3.2 3.2L13 4.8" stroke="#2e7d32" stroke-width="1.8" fill="none" /></svg><span class="label">{{ t('Accept') }}</span>
          </button>
          <button ref="acceptMenuButton" type="button" class="arrow" :title="t('More ways to accept')" :aria-label="t('More ways to accept')" aria-haspopup="menu" :aria-expanded="changeMenu.open === 'accept'" :disabled="!changes.total" @click="openChangeMenu('accept')">▾</button>
        </span>
        <span class="split">
          <button type="button" class="labelled" :title="t('Reject the change (or the changes selected) and go to the next')" :disabled="!changes.total" @click="resolveCurrent(false)">
            <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4 4l8 8M12 4l-8 8" stroke="#b3261e" stroke-width="1.8" fill="none" /></svg><span class="label">{{ t('Reject') }}</span>
          </button>
          <button ref="rejectMenuButton" type="button" class="arrow" :title="t('More ways to reject')" :aria-label="t('More ways to reject')" aria-haspopup="menu" :aria-expanded="changeMenu.open === 'reject'" :disabled="!changes.total" @click="openChangeMenu('reject')">▾</button>
        </span>
        </fieldset>
        <span class="badge" :class="{ zero: !changes.total }" :title="changesTitle">{{ changes.total }}</span>
      </span>
      <select data-tab="review view" class="markup-choice" :value="markup" :aria-label="t('Show markup')" :title="t('What the tracked changes look like, as Word’s Display for Review: all markup, simple markup (a line in the margin beside each change), no markup (the text as it will be once they are accepted) or the original (the text before them, read-only). Only how the page shows them: the changes stay in the file')" @change="setMarkup($event.target.value)">
        <option value="all">{{ t('All markup') }}</option>
        <option value="simple">{{ t('Simple markup') }}</option>
        <option value="none">{{ t('No markup') }}</option>
        <option value="original">{{ t('Original') }}</option>
      </select>
      <span class="sep" data-tab="review" />
      <!-- Proofing, as on Word's Review tab: the thesaurus (AutoCorrect is beside AutoFormat, in the status bar). -->
      <button data-tab="review" type="button" :title="t('Thesaurus (Shift+F7): other words for the word at the caret, on this device')" :aria-label="t('Thesaurus')" @click="openThesaurus">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 2.5h4.2c.8 0 1.3.5 1.3 1.3V13.5c0-.7-.6-1.2-1.3-1.2H2.5zM13.5 2.5H9.3c-.8 0-1.3.5-1.3 1.3V13.5c0-.7.6-1.2 1.3-1.2h4.2z" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round" fill="none" /></svg><span class="short-label">{{ t('Thesaurus') }}</span>
      </button>
      <!-- Hidden text, for a test's answer key: as Word, a font setting; here beside the review of what the file carries. -->
      <button data-tab="review" type="button" :title="t('Hidden text (Ctrl+Shift+H): for a test’s answers, shown here and printed only when asked')" :aria-label="t('Hidden text')" :class="{ on: state.hidden }" :aria-pressed="state.hidden" @click="toggleHidden">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M1.5 8s2.4-4.2 6.5-4.2S14.5 8 14.5 8s-2.4 4.2-6.5 4.2S1.5 8 1.5 8z" stroke="currentColor" stroke-width="1.2" fill="none" /><circle cx="8" cy="8" r="1.9" stroke="currentColor" stroke-width="1.2" fill="none" /><path d="M2.5 14L13.5 2" stroke="currentColor" stroke-width="1.3" /></svg>
      </button>
      <!-- Redact: the selected text blacked out for good, as a page is before it is released. -->
      <button data-tab="review" type="button" :title="t('Redact: black out the selected text for good, such as a source’s name – never tracked, so it is gone from the file')" :aria-label="t('Redact')" :disabled="textLocked" @click="redactSelection">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="1.5" y="5.5" width="13" height="5" rx="1" fill="currentColor" /><path d="M1.5 3h13M1.5 13h7" stroke="currentColor" stroke-width="1.3" /></svg><span class="short-label">{{ t('Redact') }}</span>
      </button>
      <button data-tab="review" type="button" class="review-toggle" :title="t('Review your changes, what the file carries and its accessibility – any time, not only when you download')" @click="openReview()">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 4l1.2 1.2L5.5 3M2 9l1.2 1.2L5.5 8M7.5 4.2h6.5M7.5 9.2h6.5M7.5 13h6.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" fill="none" /></svg>
        <span class="label">{{ t('Review') }}</span>
      </button>
      <button data-tab="review" type="button" class="new-comment ph rv" :title="canComment ? 'New comment (Ctrl+Alt+M)' : `New comment – ${lockMessage}`" :disabled="!canComment" @click="startComment">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 3h11v7.5H7L4 13v-2.5H2.5z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /><path d="M8 4.8v4M6 6.8h4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" /></svg>
      </button>
      <button data-tab="review" type="button" :title="t('Compare documents: this document against another version of it, on this device – what changed in the text')" :aria-label="t('Compare documents')" @click="startCompare">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 2.5h5v11H2zM9 2.5h5v11H9z" stroke="currentColor" stroke-width="1.3" fill="none" /><path d="M3.5 5h2M3.5 7.5h2M10.5 5h2M10.5 7.5h2M10.5 10h2" stroke="currentColor" stroke-width="1.1" /></svg><span class="short-label">{{ t('Compare') }}</span>
      </button>
      <button data-tab="review" type="button" :title="reading.on ? t('Stop reading aloud') : t('Read aloud (Ctrl+Alt+Space): from the caret, or the selected text, with a voice on this device – nothing is sent anywhere')" :aria-label="reading.on ? t('Stop reading aloud') : t('Read aloud')" :class="{ on: reading.on }" :aria-pressed="reading.on" @click="toggleReadAloud">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 6h2.5L8 3v10L4.5 10H2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M10.5 5.5a3.5 3.5 0 0 1 0 5M12.5 3.5a6.3 6.3 0 0 1 0 9" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" fill="none" /></svg><span class="short-label">{{ reading.on ? t('Stop') : t('Read aloud') }}</span>
      </button>
      <button v-if="reading.on" data-tab="review" type="button" :aria-label="reading.paused ? t('Resume reading') : t('Pause reading')" :title="reading.paused ? t('Resume reading') : t('Pause reading')" @click="pauseReadAloud">
        <svg v-if="reading.paused" viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4.5 3l8 5-8 5z" fill="currentColor" /></svg>
        <svg v-else viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4.5 3h2.5v10H4.5zM9 3h2.5v10H9z" fill="currentColor" /></svg>
      </button>
      <span data-tab="review" v-if="threads.length" class="comment-nav rv" role="group" :aria-label="t('Go to a comment')">
        <button type="button" :title="t('Previous comment')" :aria-label="t('Previous comment')" @click="gotoComment(-1)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M5 2.5h9v6.5H9.5L7 11.5V9H5z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M4 9.5L1.5 12 4 14.5" stroke="currentColor" stroke-width="1.5" fill="none" /></svg>
        </button>
        <button type="button" :title="t('Next comment')" :aria-label="t('Next comment')" @click="gotoComment(1)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 2.5h9v6.5H9L6.5 11.5V9H2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><path d="M12 9.5l2.5 2.5-2.5 2.5" stroke="currentColor" stroke-width="1.5" fill="none" /></svg>
        </button>
      </span>
      <button data-tab="review" v-if="threads.length" type="button" class="comments-toggle rv" :title="`Comments (${threads.length})`" :aria-label="`Comments (${threads.length})`" :class="{ on: commentsOpen }" :aria-pressed="commentsOpen" @click="commentsOpen = !commentsOpen">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 3h11v7.5H7L4 13v-2.5H2.5z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" fill="none" /></svg>
        <span class="badge">{{ threads.length }}</span>
      </button>
      <!-- The View tab, as Word's: the Navigation pane, zoom, and how the page shows (formatting marks
           and markup are in their own tabs too). Its own tools show in the ribbon's View tab only. -->
      <button data-tab="view" type="button" class="nav-pane-toggle" :title="t('Navigation pane: move around the document by its headings, and move a heading with its text')" :aria-label="t('Navigation pane')" :class="{ on: navOpen }" :aria-pressed="navOpen" @click="toggleNav">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="1.8" y="2.5" width="12.4" height="11" rx="1" stroke="currentColor" stroke-width="1.3" fill="none" /><path d="M6 2.5v11M3 5h1.8M3 7.2h1.8M3 9.4h1.8" stroke="currentColor" stroke-width="1.1" stroke-linecap="round" /></svg><span class="short-label">{{ t('Navigation') }}</span>
      </button>
      <select data-tab="view" class="zoom-choice" :value="String(zoomLevel)" :aria-label="t('Zoom')" :title="t('Zoom: the page shown bigger or smaller; the document is unchanged')" @change="setZoom(Number($event.target.value))">
        <option v-for="z in ZOOMS" :key="z" :value="String(z)">{{ z }}%</option>
      </select>
      <button data-tab="view" type="button" :title="t('100%: the page at its own size')" :aria-label="t('Zoom to 100%')" :disabled="zoomLevel === 100" @click="setZoom(100)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="2" y="2" width="12" height="12" rx="1" stroke="currentColor" stroke-width="1.2" fill="none" /><text x="8" y="10.6" text-anchor="middle" font-size="6.5" font-family="system-ui, sans-serif" fill="currentColor">1:1</text></svg><span class="short-label">100%</span>
      </button>
      <button v-if="!ribbon" type="button" class="ribbon-mode" data-tab="classic" :title="t('Show the tools in tabs (File, Home, Insert, Layout, References, Review, View), as Word’s ribbon')" @click="setRibbon(true)">{{ t('Tabs') }}</button>
      <button type="button" class="help-toggle" :title="t('Help: keyboard shortcuts and checks (Ctrl+/)')" :aria-label="t('Help')" aria-haspopup="dialog" @click="helpOpen = true">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="8" cy="8" r="6.3" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M6.2 6.3a1.9 1.9 0 1 1 2.6 1.7c-.5.2-.8.6-.8 1.1v.4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" fill="none" /><circle cx="8" cy="11.6" r=".85" fill="currentColor" /></svg>
      </button>
      </div>
    </div>

    <FindBar v-if="findOpen" ref="findBar" :root="page" :roots="searchRoots" :edit="applyEdit" :tracked="trackedBy" @close="onFindClosed" @reveal="revealRange" />
    <FillInDialog v-if="fillIn.open" ref="fillInDialog" :style="{ left: `${fillIn.x}px`, top: `${fillIn.y}px` }" :roots="() => [page]" :edit="applyEdit" :tracked="trackedBy" @close="closeFillIn" @reveal="revealRange" @filled="onFilled" @mailmerge="openMailMerge" />

    <div class="workspace">
      <NavigationPane v-if="navOpen" :headings="navHeadings" :current="navCurrent" :movable="!textLocked" @close="toggleNav" @go="goToHeading" @move="moveHeadingBy" @move-before="moveHeadingBefore" />
      <div class="canvas" @mousedown.self="focusEnd">
        <div class="docx-sheet" :style="sheetStyle" :data-markup="markup === 'all' ? null : markup" @mousedown.self="focusEnd">
          <div v-if="watermark.text" class="watermark-layer" aria-hidden="true">
            <span v-for="(y, i) in watermark.centers" :key="i" class="watermark-mark" :class="{ diagonal: watermark.diagonal }" :style="{ top: `${y}px`, color: watermark.color, fontFamily: `'${watermark.font}', Calibri, Carlito, sans-serif`, fontSize: `${watermark.size}px` }">{{ watermark.text }}</span>
          </div>
          <div ref="marksEl" class="marks-layer" aria-hidden="true" />
          <div v-if="hasLineNumbers" ref="lineNumbersEl" class="line-numbers-layer" aria-hidden="true" />
          <div ref="pageBreaksEl" class="page-breaks-layer" aria-hidden="true" />
          <div v-if="picture.el" class="picture-frame" :class="{ floating: picture.wrap !== 'inline', cropping: picture.cropping }" :style="{ left: `${picture.x}px`, top: `${picture.y}px`, width: `${picture.w}px`, height: `${picture.h}px` }">
            <span v-for="h in HANDLES" :key="h" class="handle" :class="h" :title="picture.cropping ? 'Drag to crop' : h.length === 2 ? 'Drag to resize' : 'Drag to stretch'" @pointerdown.prevent.stop="picture.cropping ? startCrop($event, h) : startResize($event, h)" />
            <span v-if="!picture.cropping" class="handle rotate" :title="t('Drag to rotate (Shift: in steps of 15°)')" @pointerdown.prevent.stop="startRotate" />
          </div>
          <button v-if="picture.el" ref="layoutOptionsButton" type="button" class="layout-options" :title="t('Layout Options: wrapping, position, size and alt text')" aria-haspopup="dialog" :aria-expanded="picturePanel.open" :style="{ left: `${picture.x + picture.w + 6}px`, top: `${picture.y}px` }" @mousedown.prevent @click="togglePicturePanel">
            <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="2" y="3" width="6" height="5" stroke="currentColor" stroke-width="1.3" fill="none" /><path d="M10 4h4M10 7h4M2 11h12M2 14h12" stroke="currentColor" stroke-width="1.3" /></svg>
          </button>
          <div v-if="variantTabs.header.length > 1" class="hf-tabs header" role="tablist" :aria-label="t('Headers')">
            <button v-for="{ v, i } in variantTabs.header" :key="v.rid" type="button" role="tab" :aria-selected="hf.header === i" @mousedown.prevent @click="showVariant('header', i)">{{ v.label }}</button>
          </div>
          <div
            ref="page"
            class="docx-page"
            :class="{ painting: painter.active }"
            contenteditable="true"
            spellcheck="true"
            :style="pageStyle"
            @beforeinput="onBeforeInput"
            @input="onInput"
            @keydown="onKeydown"
            @paste="onPaste"
            @drop="onDrop"
            @mousedown="onPageMousedown"
            @mouseup="onPageMouseup"
            @contextmenu="onContextMenu"
            @compositionend="onCompositionEnd"
            @click="onClick"
          />
          <div v-if="variantTabs.footer.length > 1" class="hf-tabs footer" role="tablist" :aria-label="t('Footers')">
            <button v-for="{ v, i } in variantTabs.footer" :key="v.rid" type="button" role="tab" :aria-selected="hf.footer === i" @mousedown.prevent @click="showVariant('footer', i)">{{ v.label }}</button>
          </div>
        </div>
      </div>
      <CommentsPanel
        v-if="threads.length || composer.open"
        v-show="commentsOpen || composer.open"
        :threads="threads"
        :active="activeComment"
        :composing="composer.open"
        :ask-name="!named"
        :read-only="!canComment"
        @select="selectComment"
        @close="commentsOpen = false; cancelComment()"
        @post="postComment"
        @cancel="cancelComment"
        @delete="deleteComment"
        @reply="replyComment"
        @edit="editComment"
        @resolve="resolveComment"
      />
    </div>

    <div v-if="chooser.open" ref="chooserEl" class="chooser" :role="chooser.kind === 'date' ? 'dialog' : 'listbox'" :aria-label="t('Choose a value')" :style="{ left: `${chooser.x}px`, top: `${chooser.y}px` }" @keydown.esc="closeChooser(true)" @keydown="onChooserKeydown">
      <input v-if="chooser.kind === 'date'" ref="dateInput" type="date" :value="chooser.date" :aria-label="t('Date')" @keydown="onDateKeydown" @change="onDateChange" />
      <template v-else>
        <button v-for="c in chooser.choices" :key="c.value" type="button" role="option" :aria-selected="c.text === chooser.current" @click="choose(c)">{{ c.text }}</button>
        <p v-if="!chooser.choices.length" class="empty">{{ t('No choices') }}</p>
      </template>
    </div>

    <div v-if="changeMenu.open" ref="changeMenuEl" class="change-menu" role="menu" :aria-label="changeMenu.open === 'accept' ? 'Accept' : 'Reject'" :style="{ left: `${changeMenu.x}px`, top: `${changeMenu.y}px` }" @mousedown.prevent @keydown.esc="closeChangeMenu">
      <template v-if="changeMenu.open === 'accept'">
        <button type="button" role="menuitem" @click="menuResolve(true, 'one')">{{ t('Accept this change') }}</button>
        <button type="button" role="menuitem" @click="menuResolve(true, 'stay')">{{ t('Accept, and stay here') }}</button>
        <button type="button" role="menuitem" @click="menuResolve(true, 'all')">{{ t('Accept all changes') }}</button>
      </template>
      <template v-else>
        <button type="button" role="menuitem" @click="menuResolve(false, 'one')">{{ t('Reject this change') }}</button>
        <button type="button" role="menuitem" @click="menuResolve(false, 'stay')">{{ t('Reject, and stay here') }}</button>
        <button type="button" role="menuitem" @click="menuResolve(false, 'all')">{{ t('Reject all changes') }}</button>
      </template>
    </div>

    <form v-if="picturePanel.open" ref="picturePanelEl" class="picture-panel" role="dialog" :aria-label="t('Layout Options')" :style="{ left: `${picturePanel.x}px`, top: `${picturePanel.y}px` }" @submit.prevent @keydown.esc="closePicturePanel">
      <fieldset class="wraps">
        <legend>{{ t('Text wrapping') }}</legend>
        <label v-for="[value, label] in WRAPS" :key="value"><input type="radio" name="wrap" :value="value" :checked="picturePanel.wrap === value" @change="applyWrap(value)" /> {{ label }}</label>
      </fieldset>
      <fieldset :disabled="picturePanel.wrap === 'inline'">
        <legend>{{ t('Position') }}</legend>
        <label>{{ t('Horizontal') }}
          <select v-model="picturePanel.hAlign" @change="applyPosition('h')">
            <option value="left">{{ t('Left of the margins') }}</option>
            <option value="center">{{ t('Centred between the margins') }}</option>
            <option value="right">{{ t('Right of the margins') }}</option>
            <option value="offset">{{ t('At a distance from the left margin') }}</option>
            <option v-if="picturePanel.hAlign === 'word'" value="word">{{ t('As set in Word') }}</option>
          </select>
        </label>
        <label v-if="picturePanel.hAlign === 'offset'">{{ t('Distance from the margin') }} ({{ picturePanel.unit }}) <input v-model.number="picturePanel.hOffset" type="number" step="0.1" @change="applyPosition('h')" /></label>
        <label>{{ t('Vertical') }}
          <select v-model="picturePanel.vAlign" @change="applyPosition('v')">
            <option value="offset">{{ t('Below its paragraph (moves with text)') }}</option>
            <option value="top">{{ t('Top of the margins') }}</option>
            <option value="center">{{ t('Middle of the margins') }}</option>
            <option value="bottom">{{ t('Bottom of the margins') }}</option>
            <option v-if="picturePanel.vAlign === 'word'" value="word">{{ t('As set in Word') }}</option>
          </select>
        </label>
        <label v-if="picturePanel.vAlign === 'offset'">{{ t('Distance below the paragraph') }} ({{ picturePanel.unit }}) <input v-model.number="picturePanel.vOffset" type="number" step="0.1" @change="applyPosition('v')" /></label>
      </fieldset>
      <fieldset>
        <legend>{{ t('Size') }}</legend>
        <div class="size-row">
          <label>{{ t('Width') }} ({{ picturePanel.unit }}) <input v-model.number="picturePanel.width" type="number" min="0.1" step="0.1" @change="applySize('width')" /></label>
          <label>{{ t('Height') }} ({{ picturePanel.unit }}) <input v-model.number="picturePanel.height" type="number" min="0.1" step="0.1" @change="applySize('height')" /></label>
        </div>
        <label class="inline"><input v-model="picturePanel.lock" type="checkbox" /> {{ t('Keep proportions') }}</label>
        <button v-if="!picture.shape" type="button" @click="resetPictureSize">{{ t('Reset size') }}</button>
      </fieldset>
      <!-- A line, arrow or shape: its wrapping, position, size and rotation, not a picture's crop or effects. -->
      <details v-if="!picture.shape">
        <summary>{{ t('Crop') }}</summary>
        <div class="size-row">
          <label>Left (%) <input v-model.number="picturePanel.cropL" type="number" step="1" @change="applyCrop" /></label>
          <label>Right (%) <input v-model.number="picturePanel.cropR" type="number" step="1" @change="applyCrop" /></label>
          <label>Top (%) <input v-model.number="picturePanel.cropT" type="number" step="1" @change="applyCrop" /></label>
          <label>Bottom (%) <input v-model.number="picturePanel.cropB" type="number" step="1" @change="applyCrop" /></label>
        </div>
        <button type="button" :aria-pressed="picture.cropping" @click="toggleCropping">{{ picture.cropping ? 'Done cropping' : 'Crop with the handles' }}</button>
      </details>
      <details>
        <summary>{{ t('Rotate and flip') }}</summary>
        <div class="buttons">
          <button type="button" @click="rotateBy(90)">{{ t('Rotate right 90°') }}</button>
          <button type="button" @click="rotateBy(-90)">{{ t('Rotate left 90°') }}</button>
          <button type="button" :aria-pressed="picturePanel.flipH" @click="flip('flipH')">{{ t('Flip horizontal') }}</button>
          <button type="button" :aria-pressed="picturePanel.flipV" @click="flip('flipV')">{{ t('Flip vertical') }}</button>
        </div>
        <label>Rotation (°) <input v-model.number="picturePanel.rot" type="number" step="1" @change="applyRotation" /></label>
      </details>
      <details v-if="!picture.shape">
        <summary>{{ t('Effects') }}</summary>
        <div class="size-row">
          <label>{{ t('Border') }}
            <select v-model="picturePanel.border" @change="applyEffects">
              <option value="none">{{ t('No border') }}</option>
              <option value="thin">{{ t('Thin') }}</option>
              <option value="medium">{{ t('Medium') }}</option>
              <option value="thick">{{ t('Thick') }}</option>
            </select>
          </label>
          <label>{{ t('Border colour') }} <input v-model="picturePanel.borderColor" type="color" @change="applyEffects" /></label>
        </div>
        <label class="inline"><input v-model="picturePanel.shadow" type="checkbox" @change="applyEffects" /> {{ t('Shadow') }}</label>
        <label class="inline"><input v-model="picturePanel.glow" type="checkbox" @change="applyEffects" /> {{ t('Glow') }}</label>
        <label class="inline"><input v-model="picturePanel.softEdge" type="checkbox" @change="applyEffects" /> {{ t('Soft edges') }}</label>
        <label class="inline"><input v-model="picturePanel.reflection" type="checkbox" @change="applyEffects" /> {{ t('Reflection') }}</label>
        <label>{{ t('Colour') }}
          <select v-model="picturePanel.colour" @change="applyEffects">
            <option value="normal">{{ t('As it is') }}</option>
            <option value="greyscale">{{ t('Greyscale') }}</option>
          </select>
        </label>
        <div class="size-row">
          <label>Brightness (%) <input v-model.number="picturePanel.brightness" type="number" min="-100" max="100" step="5" @change="applyEffects" /></label>
          <label>Contrast (%) <input v-model.number="picturePanel.contrast" type="number" min="-100" max="100" step="5" @change="applyEffects" /></label>
        </div>
      </details>
      <details v-if="!picture.shape">
        <summary>{{ t('Style and 3-D') }}</summary>
        <label>{{ t('Picture style') }}
          <select v-model="picturePanel.style" @change="applyStyle">
            <option value="">{{ t('Choose a style…') }}</option>
            <option value="none">{{ t('No style (plain picture)') }}</option>
            <option v-for="st in PICTURE_STYLES" :key="st.name" :value="st.name">{{ st.name }}</option>
          </select>
        </label>
        <label>{{ t('Shape') }}
          <select v-model="picturePanel.shape" @change="applyShape">
            <option v-for="[value, label] in SHAPES" :key="value" :value="value">{{ label }}</option>
            <option v-if="!SHAPES.some(([v]) => v === picturePanel.shape)" :value="picturePanel.shape" disabled>As set in Word</option>
          </select>
        </label>
        <fieldset class="sub">
          <legend>{{ t('3-D format') }}</legend>
          <label>{{ t('Bevel') }}
            <select v-model="picturePanel.bevel" @change="apply3D">
              <option value="">{{ t('None') }}</option>
              <option v-for="[value, label] in BEVELS" :key="value" :value="value">{{ label }}</option>
              <option v-if="picturePanel.bevel && !BEVELS.some(([v]) => v === picturePanel.bevel)" :value="picturePanel.bevel" disabled>As set in Word</option>
            </select>
          </label>
          <div class="size-row">
            <label>Bevel width (pt) <input v-model.number="picturePanel.bevelW" type="number" min="0" step="0.5" :disabled="!picturePanel.bevel" @change="apply3D" /></label>
            <label>Bevel height (pt) <input v-model.number="picturePanel.bevelH" type="number" min="0" step="0.5" :disabled="!picturePanel.bevel" @change="apply3D" /></label>
          </div>
          <label>{{ t('Material') }}
            <select v-model="picturePanel.material" @change="apply3D">
              <option value="">{{ t('As Word chooses (Warm Matte)') }}</option>
              <option v-for="[value, label] in MATERIALS" :key="value" :value="value">{{ label }}</option>
            </select>
          </label>
          <div class="size-row">
            <label>{{ t('Lighting') }}
              <select v-model="picturePanel.lightRig" @change="apply3D">
                <option v-for="[value, label] in LIGHTS" :key="value" :value="value">{{ label }}</option>
                <option v-if="!LIGHTS.some(([v]) => v === picturePanel.lightRig)" :value="picturePanel.lightRig" disabled>As set in Word</option>
              </select>
            </label>
            <label>Light angle (°) <input v-model.number="picturePanel.lightAngle" type="number" step="5" @change="apply3D" /></label>
          </div>
        </fieldset>
        <fieldset class="sub">
          <legend>{{ t('3-D rotation') }}</legend>
          <label>{{ t('Preset') }}
            <select v-model="picturePanel.camera" @change="applyCameraPreset">
              <option value="">{{ t('None') }}</option>
              <option v-for="[value, label] in CAMERAS" :key="value" :value="value">{{ label }}</option>
              <option v-if="picturePanel.camera && !CAMERAS.some(([v]) => v === picturePanel.camera)" :value="picturePanel.camera" disabled>As set in Word</option>
            </select>
          </label>
          <div class="size-row">
            <label>X rotation (°) <input v-model.number="picturePanel.rotX" type="number" step="5" @change="apply3D('angles')" /></label>
            <label>Y rotation (°) <input v-model.number="picturePanel.rotY" type="number" step="5" @change="apply3D('angles')" /></label>
            <label>Z rotation (°) <input v-model.number="picturePanel.rotZ" type="number" step="5" @change="apply3D('angles')" /></label>
            <label>Perspective (°) <input v-model.number="picturePanel.perspective" type="number" min="0" max="120" step="5" @change="apply3D('angles')" /></label>
          </div>
        </fieldset>
      </details>
      <details>
        <summary>{{ t('Artistic effects') }}</summary>
        <label>{{ t('Artistic effect') }}
          <select v-model="picturePanel.artistic" :disabled="picturePanel.drawing" @change="applyArtistic">
            <option value="">{{ t('None') }}</option>
            <option v-for="[value, label] in ARTISTIC" :key="value" :value="value">{{ label }}</option>
          </select>
        </label>
        <label>{{ t('Amount') }} <input v-model.number="picturePanel.amount" type="range" min="0" max="100" step="5" :disabled="!picturePanel.artistic || picturePanel.drawing" @change="applyArtistic" /></label>
        <p class="note">{{ t('Drawn into the picture, which Word shows as it is here; the original is kept, so the effect can be changed or taken off here.') }}</p>
      </details>
      <label class="alt">Alt text (describe the picture for people who can't see it)
        <textarea v-model="picturePanel.alt" rows="2" @change="applyAlt" />
      </label>
      <div class="actions"><button type="button" @click="closePicturePanel">{{ t('Close') }}</button></div>
    </form>

    <div v-if="saveMenu.open" ref="saveMenuEl" class="change-menu save-menu" role="menu" :aria-label="t('Save')" :style="{ left: `${saveMenu.x}px`, top: `${saveMenu.y}px` }" @mousedown.prevent @keydown.esc="closeSaveMenu(true)" @keydown="onMenuArrows">
      <button type="button" role="menuitem" @click="saveMenuAction('docx')">{{ fileHandle ? `Save to ${fileHandle.name}` : `Word document (.${fileExt})` }}<span class="shortcut">{{ t('Ctrl+S') }}</span></button>
      <button v-if="fileHandle" type="button" role="menuitem" @click="saveMenuAction('copy')">Download a copy<span class="shortcut">{{ t('To Downloads') }}</span></button>
      <button v-if="canSaveAs" type="button" role="menuitem" @click="saveMenuAction('saveAs')">Save as…<span class="shortcut">{{ t('A file on this device, saved to after') }}</span></button>
      <button type="button" role="menuitem" @click="saveMenuAction('pdf')">PDF…<span class="shortcut">{{ t('Ctrl+P') }}</span></button>
      <button v-if="canShareFiles" type="button" role="menuitem" @click="saveMenuAction('share')">Share…<span class="shortcut">{{ t('Mail, messages…') }}</span></button>
      <button type="button" role="menuitem" @click="saveMenuAction('web')">{{ t('Copy for the web') }}<span class="shortcut">{{ t('A website’s editor (CMS), an e-mail') }}</span></button>
      <button type="button" role="menuitem" @click="saveMenuAction('markdown')">{{ t('Copy as Markdown') }}<span class="shortcut">{{ t('Markdown editors, GitHub, Ghost') }}</span></button>
      <button type="button" role="menuitem" @click="saveMenuAction('merge')">Mail merge…<span class="shortcut">{{ t('A letter for each row of a list') }}</span></button>
      <button type="button" role="menuitem" aria-haspopup="dialog" @click="saveMenuAction('password')">{{ passwordProtected ? t('Change or remove password…') : t('Protect with password…') }}<span class="shortcut">{{ t('Encrypt the file, as Word does') }}</span></button>
    </div>

    <form v-if="pdfDialog.open" ref="pdfDialogEl" class="chooser author-dialog pdf-dialog" role="dialog" :aria-label="t('Save as PDF')" aria-describedby="pdf-how" :style="{ left: `${pdfDialog.x}px`, top: `${pdfDialog.y}px` }" @submit.prevent="printToPdf" @keydown.esc="closePdfDialog">
      <h2>{{ t('Save as PDF') }}</h2>
      <p id="pdf-how">Your browser makes the PDF, on this device: in the print window, choose <strong>{{ t('Save as PDF') }}</strong> (or <strong>{{ t('Microsoft Print to PDF') }}</strong>) as the printer or destination, then save. It suggests the name <strong>{{ baseName || 'document' }}.pdf</strong>.</p>
      <ul class="pdf-notes">
        <li v-if="pageCount">About {{ pageCount }} {{ pageCount === 1 ? 'page' : 'pages' }} on {{ paperName() }}{{ pdfBands.every ? `, with the ${pdfBands.every} on each` : '' }}.</li>
        <li v-if="pdfBands.once">The {{ pdfBands.once }} {{ pdfBands.once.includes(' and ') ? 'print' : 'prints' }} once, {{ pdfBands.onceWhere }}, as {{ pdfBands.once.includes(' and ') ? 'they show' : 'it shows' }} here: {{ pdfBands.why }}</li>
        <li v-if="changes.total" class="warn">{{ changes.total === 1 ? 'A tracked change prints' : `${changes.total} tracked changes print` }} marked, as {{ changes.total === 1 ? 'it shows' : 'they show' }} here. For a clean copy, accept or reject {{ changes.total === 1 ? 'it' : 'them' }} first.</li>
        <li v-if="hasHighlight" class="warn">Highlighted text prints highlighted.</li>
        <li v-if="hasHidden"><label class="check"><input v-model="printHidden" type="checkbox" /> {{ t('Print hidden text (a test’s answers, for its answer key)') }}</label></li>
        <li>Comments and the marks drawn over the page here aren’t in the PDF.</li>
      </ul>
      <p class="pdf-word">{{ t('The Word file isn’t saved by this: download it as well to keep your changes.') }}</p>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('Open print window') }}</button>
        <button type="button" @click="closePdfDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <div v-if="tableMenu.open" ref="tableMenuEl" class="change-menu" role="menu" :aria-label="t('Table layout')" :style="{ left: `${tableMenu.x}px`, top: `${tableMenu.y}px` }" @mousedown.prevent @keydown.esc="closeTableMenu" @keydown="onMenuArrows">
      <button type="button" role="menuitem" @click="tableAction('rowsAbove')">{{ tableMenu.rows > 1 ? `Insert ${tableMenu.rows} rows above` : 'Insert row above' }}</button>
      <button type="button" role="menuitem" @click="tableAction('rowsBelow')">{{ tableMenu.rows > 1 ? `Insert ${tableMenu.rows} rows below` : 'Insert row below' }}</button>
      <button type="button" role="menuitem" @click="tableAction('columnsLeft')">{{ tableMenu.newCols > 1 ? `Insert ${tableMenu.newCols} columns to the left` : 'Insert column to the left' }}</button>
      <button type="button" role="menuitem" @click="tableAction('columnsRight')">{{ tableMenu.newCols > 1 ? `Insert ${tableMenu.newCols} columns to the right` : 'Insert column to the right' }}</button>
      <hr />
      <button type="button" role="menuitem" :disabled="!tableMenu.multi" :title="tableMenu.multi ? '' : 'Select the cells to merge first'" @click="tableAction('merge')">{{ t('Merge cells') }}</button>
      <button type="button" role="menuitem" @click="openSplitDialog">{{ t('Split cells…') }}</button>
      <hr />
      <button type="button" role="menuitem" aria-haspopup="dialog" @click="openCellFormat">{{ t('Shading and borders…') }}</button>
      <hr />
      <button type="button" role="menuitem" aria-haspopup="dialog" @click="openSortDialog">{{ t('Sort…') }}</button>
      <button type="button" role="menuitem" aria-haspopup="dialog" @click="openFormulaDialog">{{ t('Formula…') }}</button>
      <hr />
      <button type="button" role="menuitem" @click="tableAction('deleteRows')">{{ tableMenu.rows > 1 ? `Delete ${tableMenu.rows} rows` : 'Delete row' }}</button>
      <button type="button" role="menuitem" @click="tableAction('deleteColumns')">{{ tableMenu.cols > 1 ? `Delete ${tableMenu.cols} columns` : 'Delete column' }}</button>
      <button type="button" role="menuitem" @click="tableAction('deleteTable')">{{ t('Delete table') }}</button>
    </div>

    <!-- Word's Shading and Borders buttons (Table Design), for the selected cells. -->
    <div v-if="cellFormat.open" ref="cellFormatEl" class="palette cell-format" role="dialog" :aria-label="t('Shading and borders')" :style="{ left: `${cellFormat.x}px`, top: `${cellFormat.y}px` }" @mousedown="keepSelection" @keydown.esc="closeCellFormat">
      <p class="palette-label">{{ t('Borders') }}</p>
      <div class="border-choices">
        <button v-for="b in BORDER_CHOICES" :key="b.value" type="button" :title="b.title" @click="formatCells('borders', b.value)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 2.5h11v11h-11zM8 2.5v11M2.5 8h11" stroke="currentColor" stroke-width=".8" stroke-dasharray="1 1.2" fill="none" opacity=".55" /><path :d="b.icon" stroke="currentColor" stroke-width="1.6" stroke-linecap="square" fill="none" /></svg>
          <span>{{ b.label }}</span>
        </button>
      </div>
      <p class="palette-label">{{ t('Shading') }}</p>
      <button type="button" class="no-color" @click="formatCells('shading', null)">{{ t('No colour') }}</button>
      <div class="swatches">
        <button v-for="c in shadingChoices.theme" :key="c.hex + c.label" type="button" class="swatch" :style="{ background: `#${c.hex}` }" :title="c.label" :aria-label="c.label" @click="formatCells('shading', c.hex)" />
      </div>
      <template v-if="shadingChoices.used.length">
        <p class="palette-label">{{ t('In this document') }}</p>
        <div class="swatches">
          <button v-for="c in shadingChoices.used" :key="c.hex" type="button" class="swatch" :style="{ background: `#${c.hex}` }" :title="c.label" :aria-label="c.label" @click="formatCells('shading', c.hex)" />
        </div>
      </template>
      <MoreColors @pick="(hex) => formatCells('shading', hex)" />
    </div>

    <form v-if="splitDialog.open" ref="splitDialogEl" class="chooser author-dialog split-dialog" role="dialog" :aria-label="t('Split cells')" :style="{ left: `${splitDialog.x}px`, top: `${splitDialog.y}px` }" @submit.prevent="splitCells" @keydown.esc="closeSplitDialog">
      <label>{{ t('Number of columns') }} <input v-model.number="splitDialog.cols" type="number" min="1" max="20" required /></label>
      <label>{{ t('Number of rows') }} <input v-model.number="splitDialog.rows" type="number" min="1" max="50" required /></label>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('Split') }}</button>
        <button type="button" @click="closeSplitDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <!-- Word's Line and Paragraph Spacing menu, and its Paragraph dialog (indents and spacing). -->
    <div v-if="spacingMenu.open" ref="spacingMenuEl" class="change-menu spacing-menu" role="menu" :aria-label="t('Line and paragraph spacing')" :style="{ left: `${spacingMenu.x}px`, top: `${spacingMenu.y}px` }" @mousedown.prevent @keydown.esc="closeSpacingMenu" @keydown="onMenuArrows">
      <button v-for="[v, label] in LINE_SPACINGS" :key="label" type="button" role="menuitemradio" :aria-checked="spacingMenu.line === Math.round(v * 240)" @click="setLineSpacing(v)">{{ label }}</button>
      <hr />
      <button type="button" role="menuitem" @click="toggleSpace('before')">{{ spacingMenu.before ? 'Remove space before paragraph' : 'Add space before paragraph' }}</button>
      <button type="button" role="menuitem" @click="toggleSpace('after')">{{ spacingMenu.after ? 'Remove space after paragraph' : 'Add space after paragraph' }}</button>
      <hr />
      <button type="button" role="menuitem" @click="openParagraphDialog">{{ t('Paragraph…') }}</button>
    </div>

    <form v-if="pageDialog.open" ref="pageDialogEl" class="chooser author-dialog para-dialog page-dialog" role="dialog" :aria-label="t('Page setup')" :style="{ left: `${pageDialog.x}px`, top: `${pageDialog.y}px`, maxHeight: `calc(100vh - ${pageDialog.y + 8}px)` }" novalidate @submit.prevent="applyPageSetup" @keydown.esc="closePageSetup">
      <label class="apply-to">{{ t('Apply to') }}
        <select v-model="pageDialog.scope" :aria-label="t('Apply paper and margins to')">
          <option value="whole">{{ t('Whole document') }}</option>
          <option v-if="pageDialog.sections > 1" value="section">{{ t('This section ({n} of {total})', { n: pageDialog.sectionNo, total: pageDialog.sections }) }}</option>
          <option value="forward">{{ t('This point forward') }}</option>
          <option v-if="pageDialog.selected" value="selected">{{ t('Selected text') }}</option>
        </select>
      </label>
      <p v-if="pageDialog.scope === 'forward' || pageDialog.scope === 'selected'" class="note">{{ pageBreakNote() }}</p>
      <fieldset>
        <legend>{{ t('Paper') }}</legend>
        <label>{{ t('Size') }}
          <select v-model="pageDialog.paper">
            <option v-for="p in PAPERS" :key="p.name" :value="p.name">{{ p.name }} ({{ p.label }})</option>
            <option value="custom">{{ t('Custom size') }}</option>
          </select>
        </label>
        <template v-if="pageDialog.paper === 'custom'">
          <label>{{ t('Width') }} ({{ pageDialog.unit }}) <input v-model.number="pageDialog.width" type="number" min="0" step="0.1" /></label>
          <label>{{ t('Height') }} ({{ pageDialog.unit }}) <input v-model.number="pageDialog.height" type="number" min="0" step="0.1" /></label>
        </template>
        <label>{{ t('Orientation') }}
          <select v-model="pageDialog.orient">
            <option value="portrait">{{ t('Portrait') }}</option>
            <option value="landscape">{{ t('Landscape') }}</option>
          </select>
        </label>
      </fieldset>
      <fieldset>
        <legend>{{ t('Margins') }}</legend>
        <label>{{ t('Preset') }}
          <select :value="marginPreset" @change="pickMargins($event.target.value)">
            <option v-for="m in MARGIN_PRESETS" :key="m.name" :value="m.name">{{ m.name === 'Normal' ? t('Normal margins') : t(m.name) }}</option>
            <option v-if="marginPreset === 'custom'" value="custom">{{ t('Custom') }}</option>
          </select>
        </label>
        <label>{{ t('Top') }} ({{ pageDialog.unit }}) <input v-model.number="pageDialog.top" type="number" min="0" step="0.1" /></label>
        <label>{{ t('Bottom') }} ({{ pageDialog.unit }}) <input v-model.number="pageDialog.bottom" type="number" min="0" step="0.1" /></label>
        <label>{{ t('Left') }} ({{ pageDialog.unit }}) <input v-model.number="pageDialog.left" type="number" min="0" step="0.1" /></label>
        <label>{{ t('Right') }} ({{ pageDialog.unit }}) <input v-model.number="pageDialog.right" type="number" min="0" step="0.1" /></label>
        <label>{{ t('Units') }}
          <select :value="pageDialog.unit" :aria-label="t('Units')" :title="t('Centimetres or inches, for the margins here and the other dialogs’ lengths (remembered in this browser)')" @change="setPageUnit($event.target.value)">
            <option value="cm">{{ t('Centimetres') }}</option>
            <option value="in">{{ t('Inches') }}</option>
          </select>
        </label>
      </fieldset>
      <fieldset>
        <legend>{{ t('Columns') }}</legend>
        <label>{{ t('Number of columns') }}
          <select v-model.number="pageDialog.cols">
            <option v-for="n in columnChoices()" :key="n" :value="n">{{ n === 1 ? t('One') : n === 2 ? t('Two') : n === 3 ? t('Three') : n }}</option>
          </select>
        </label>
        <template v-if="pageDialog.cols > 1">
          <label>{{ t('Spacing') }} ({{ pageDialog.unit }}) <input v-model.number="pageDialog.colSpace" type="number" min="0" step="0.1" :aria-label="t('Spacing between columns ({unit})', { unit: pageDialog.unit })" /></label>
          <label class="check"><input v-model="pageDialog.colSep" type="checkbox" /> {{ t('Line between') }}</label>
        </template>
        <p v-if="pageDialog.colUnequal" class="note">{{ t('These columns have widths of their own: changed here, they become equal.') }}</p>
      </fieldset>
      <fieldset>
        <legend>{{ t('Headers and footers') }}</legend>
        <label class="check"><input v-model="pageDialog.titlePg" type="checkbox" /> {{ t('Different first page') }}</label>
        <label class="check"><input v-model="pageDialog.evenAndOdd" type="checkbox" /> {{ t('Different odd and even pages') }}</label>
      </fieldset>
      <fieldset>
        <legend>{{ t('Watermark') }}</legend>
        <label>{{ t('Text') }}
          <select v-model="pageDialog.mark" :aria-label="t('Watermark text')">
            <option value="">{{ t('(none)') }}</option>
            <option v-for="m in WATERMARKS" :key="m" :value="m">{{ m }}</option>
            <option value="custom">{{ t('Custom…') }}</option>
          </select>
        </label>
        <label v-if="pageDialog.mark === 'custom'">{{ t('Custom text') }} <input v-model="pageDialog.markText" type="text" maxlength="60" autocomplete="off" :aria-label="t('Watermark’s own text')" /></label>
        <label v-if="pageDialog.mark">{{ t('Layout') }}
          <select v-model="pageDialog.markLayout" :aria-label="t('Watermark layout')">
            <option value="diagonal">{{ t('Diagonal') }}</option>
            <option value="horizontal">{{ t('Horizontal') }}</option>
          </select>
        </label>
      </fieldset>
      <p class="note">{{ t('Headers and footers and the watermark: for the whole document') }}</p>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('OK') }}</button>
        <button type="button" @click="closePageSetup">{{ t('Cancel') }}</button>
      </div>
    </form>

    <form v-if="paraDialog.open" ref="paraDialogEl" class="chooser author-dialog para-dialog" role="dialog" :aria-label="t('Paragraph')" :style="{ left: `${paraDialog.x}px`, top: `${paraDialog.y}px` }" novalidate @submit.prevent="applyParagraphDialog" @keydown.esc="closeParagraphDialog">
      <fieldset>
        <legend>{{ t('Indentation') }}</legend>
        <label>{{ t('Left') }} ({{ paraDialog.unit }}) <input v-model.number="paraDialog.left" type="number" step="0.1" /></label>
        <label>{{ t('Right') }} ({{ paraDialog.unit }}) <input v-model.number="paraDialog.right" type="number" step="0.1" /></label>
        <label>{{ t('Special') }}
          <select v-model="paraDialog.special">
            <option value="none">{{ t('(none)') }}</option>
            <option value="firstLine">{{ t('First line') }}</option>
            <option value="hanging">{{ t('Hanging') }}</option>
          </select>
        </label>
        <label>{{ t('By') }} ({{ paraDialog.unit }}) <input v-model.number="paraDialog.by" type="number" min="0" step="0.1" :disabled="paraDialog.special === 'none'" /></label>
        <label>{{ t('Units') }}
          <select :value="paraDialog.unit" :aria-label="t('Units')" :title="t('Centimetres or inches, for the indents and tab stops here and the other dialogs’ lengths (remembered in this browser)')" @change="setParaUnit($event.target.value)">
            <option value="cm">{{ t('Centimetres') }}</option>
            <option value="in">{{ t('Inches') }}</option>
          </select>
        </label>
      </fieldset>
      <fieldset>
        <legend>{{ t('Spacing') }}</legend>
        <label>Before (pt) <input v-model.number="paraDialog.before" type="number" min="0" step="6" /></label>
        <label>After (pt) <input v-model.number="paraDialog.after" type="number" min="0" step="6" /></label>
        <label>{{ t('Line spacing') }}
          <select v-model="paraDialog.rule" @change="onLineRuleChange">
            <option value="single">{{ t('Single') }}</option>
            <option value="1.5">{{ t('1.5 lines') }}</option>
            <option value="double">{{ t('Double') }}</option>
            <option value="atLeast">{{ t('At least') }}</option>
            <option value="exact">{{ t('Exactly') }}</option>
            <option value="multiple">{{ t('Multiple') }}</option>
          </select>
        </label>
        <label v-if="!LINE_PRESETS[paraDialog.rule]">{{ t('At') }} ({{ paraDialog.rule === 'multiple' ? 'lines' : 'pt' }}) <input v-model.number="paraDialog.at" type="number" min="0" :step="paraDialog.rule === 'multiple' ? 0.05 : 1" /></label>
      </fieldset>
      <fieldset class="tab-stops">
        <legend>{{ t('Tab stops') }}</legend>
        <ul v-if="paraDialog.tabs.length" :aria-label="t('Tab stops')">
          <li v-for="stop in paraDialog.tabs" :key="stop.pos">
            <span class="pos">{{ formatLength(stop.pos) }} {{ paraDialog.unit }}</span>
            <span>{{ TAB_ALIGNS[stop.align] || stop.align }}</span>
            <span class="leader">{{ TAB_LEADERS[stop.leader]?.label || '' }}</span>
            <button type="button" :aria-label="`Remove the tab stop at ${formatLength(stop.pos)} ${paraDialog.unit}`" :title="t('Remove this tab stop')" @click="removeTabStop(stop.pos)">×</button>
          </li>
        </ul>
        <p v-else class="hint">None: tabs go to Word’s default stops, every {{ formatLength(doc.defaultTabStop) }} {{ paraDialog.unit }}.</p>
        <div class="tab-add">
          <label>{{ t('At') }} ({{ paraDialog.unit }}) <input v-model.number="paraDialog.tabPos" type="number" min="0" step="0.25" :aria-label="`Tab stop position (${paraDialog.unit})`" @keydown.enter.prevent="addTabStop" /></label>
          <label>{{ t('Alignment') }}
            <select v-model="paraDialog.tabAlign">
              <option v-for="(label, value) in TAB_ALIGNS" :key="value" :value="value">{{ label }}</option>
            </select>
          </label>
          <label>{{ t('Leader') }}
            <select v-model="paraDialog.tabLeader">
              <option v-for="(l, value) in TAB_LEADERS" :key="value" :value="value">{{ l.label }}</option>
            </select>
          </label>
          <div class="tab-buttons">
            <button type="button" @click="addTabStop">{{ t('Set') }}</button>
            <button type="button" :disabled="!paraDialog.tabs.length" @click="paraDialog.tabs = []">{{ t('Clear all') }}</button>
          </div>
        </div>
        <p v-if="paraDialog.tabError" class="error" role="alert">{{ paraDialog.tabError }}</p>
      </fieldset>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('OK') }}</button>
        <button type="button" @click="closeParagraphDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <form v-if="modifyDialog.open" ref="modifyDialogEl" class="chooser author-dialog para-dialog modify-dialog" role="dialog" :aria-label="`Modify style ${modifyDialog.name}`" :style="{ left: `${modifyDialog.x}px`, top: `${modifyDialog.y}px` }" novalidate @submit.prevent="applyModifyStyle" @keydown.esc="closeModifyStyle">
      <h2>Modify style: {{ modifyDialog.name }}</h2>
      <fieldset>
        <legend>{{ t('Text') }}</legend>
        <label>{{ t('Font') }}
          <select v-model="modifyDialog.font" :aria-label="t('Style font')">
            <option v-for="f in modifyFonts" :key="f" :value="f">{{ f }}</option>
          </select>
        </label>
        <label>Size (pt) <input v-model.number="modifyDialog.size" type="number" min="1" max="1638" step="0.5" :aria-label="t('Style font size (pt)')" /></label>
        <div class="checks">
          <label class="check"><input v-model="modifyDialog.bold" type="checkbox" /> <b>Bold</b></label>
          <label class="check"><input v-model="modifyDialog.italic" type="checkbox" /> <i>Italic</i></label>
        </div>
        <label>{{ t('Colour') }}
          <span class="color-pick">
            <label class="check"><input v-model="modifyDialog.auto" type="checkbox" /> {{ t('Automatic') }}</label>
            <input v-model="modifyDialog.color" type="color" :disabled="modifyDialog.auto" :aria-label="t('Style colour')" />
          </span>
        </label>
      </fieldset>
      <fieldset>
        <legend>{{ t('Paragraph') }}</legend>
        <label>{{ t('Alignment') }}
          <select v-model="modifyDialog.align" :aria-label="t('Style alignment')">
            <option value="left">{{ t('Left') }}</option>
            <option value="center">{{ t('Centred') }}</option>
            <option value="right">{{ t('Right') }}</option>
            <option value="both">{{ t('Justified') }}</option>
          </select>
        </label>
        <label>Before (pt) <input v-model.number="modifyDialog.before" type="number" min="0" step="6" :aria-label="t('Style space before (pt)')" /></label>
        <label>After (pt) <input v-model.number="modifyDialog.after" type="number" min="0" step="6" :aria-label="t('Style space after (pt)')" /></label>
        <label>{{ t('Line spacing') }}
          <select v-model="modifyDialog.line" :aria-label="t('Style line spacing')">
            <option v-if="!LINE_CHOICES.some(([v]) => String(v) === modifyDialog.line)" :value="modifyDialog.line">{{ modifyDialog.line ? `${Math.round(Number(modifyDialog.line) / 2.4) / 100} lines` : 'As set (exactly or at least)' }}</option>
            <option v-for="[v, label] in LINE_CHOICES" :key="v" :value="String(v)">{{ label }}</option>
          </select>
        </label>
      </fieldset>
      <fieldset>
        <legend>{{ t('Indentation') }}</legend>
        <label>{{ t('Left') }} ({{ modifyDialog.unit }}) <input v-model.number="modifyDialog.left" type="number" step="0.1" :aria-label="t('Style left indent')" /></label>
        <label>{{ t('Right') }} ({{ modifyDialog.unit }}) <input v-model.number="modifyDialog.right" type="number" step="0.1" :aria-label="t('Style right indent')" /></label>
        <label>{{ t('Special') }}
          <select v-model="modifyDialog.special" :aria-label="t('Style special indent')">
            <option value="none">{{ t('(none)') }}</option>
            <option value="firstLine">{{ t('First line') }}</option>
            <option value="hanging">{{ t('Hanging') }}</option>
          </select>
        </label>
        <label>{{ t('By') }} ({{ modifyDialog.unit }}) <input v-model.number="modifyDialog.by" type="number" min="0" step="0.1" :disabled="modifyDialog.special === 'none'" :aria-label="t('Style special indent by')" /></label>
        <label>{{ t('Units') }}
          <select :value="modifyDialog.unit" :aria-label="t('Style units')" :title="t('Centimetres or inches, for the indents here and the other dialogs’ lengths (remembered in this browser)')" @change="setModifyUnit($event.target.value)">
            <option value="cm">{{ t('Centimetres') }}</option>
            <option value="in">{{ t('Inches') }}</option>
          </select>
        </label>
      </fieldset>
      <p class="preview" :style="{ fontFamily: `'${modifyDialog.font}'`, fontSize: `${Math.min(28, Number(modifyDialog.size) || 11)}pt`, fontWeight: modifyDialog.bold ? 700 : 400, fontStyle: modifyDialog.italic ? 'italic' : 'normal', color: modifyDialog.auto ? '#000' : modifyDialog.color, textAlign: modifyDialog.align === 'both' ? 'justify' : modifyDialog.align, ...modifyPreviewIndent }"><span>{{ modifyDialog.name }}: the quick brown fox jumps over the lazy dog, and the first line shows its indent</span></p>
      <p class="note">{{ t('Every paragraph in this style follows, here and in Word.') }}</p>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('OK') }}</button>
        <button type="button" @click="closeModifyStyle">{{ t('Cancel') }}</button>
      </div>
    </form>

    <div v-if="tablePicker.open" ref="tablePickerEl" class="table-picker" role="dialog" :aria-label="t('Insert table')" :style="{ left: `${tablePicker.x}px`, top: `${tablePicker.y}px` }" @keydown.esc="closeTablePicker">
      <p class="size" aria-hidden="true">{{ tablePicker.cols }} × {{ tablePicker.rows }} table</p>
      <div class="grid" @mouseleave="tablePicker.cols = tablePicker.rows = 1">
        <template v-for="r in PICKER_ROWS" :key="r">
          <button
            v-for="c in PICKER_COLS"
            :key="c"
            type="button"
            class="cell"
            :class="{ on: r <= tablePicker.rows && c <= tablePicker.cols }"
            :tabindex="r === tablePicker.rows && c === tablePicker.cols ? 0 : -1"
            :aria-label="`${c} by ${r} table`"
            @mousedown.prevent
            @mouseenter="Object.assign(tablePicker, { rows: r, cols: c })"
            @keydown="onPickerKeydown"
            @click="insertTable(r, c)"
          />
        </template>
      </div>
      <form class="custom" @submit.prevent="insertTable(tablePicker.customRows, tablePicker.customCols)">
        <label>{{ t('Columns') }} <input v-model.number="tablePicker.customCols" type="number" min="1" :max="MAX_COLS" required /></label>
        <label>{{ t('Rows') }} <input v-model.number="tablePicker.customRows" type="number" min="1" :max="MAX_ROWS" required /></label>
        <button type="submit">{{ t('Insert') }}</button>
      </form>
    </div>

    <div v-if="palette.open" ref="paletteEl" class="palette" role="dialog" :aria-label="t('Highlight colour')" :style="{ left: `${palette.x}px`, top: `${palette.y}px` }" @mousedown.prevent @keydown.esc="closePalette">
      <div class="swatches">
        <button v-for="[name, label] in HIGHLIGHT_COLORS" :key="name" type="button" class="swatch" :class="{ current: name === highlightColor }" :style="{ background: HIGHLIGHT[name] }" :title="label" :aria-label="label" @click="pickHighlight(name)" />
      </div>
      <button type="button" class="no-color" @click="pickHighlight(null)">{{ t('No colour') }}</button>
    </div>

    <div v-if="colorPalette.open" ref="colorPaletteEl" class="palette color-palette" role="dialog" :aria-label="t('Font colour')" :style="{ left: `${colorPalette.x}px`, top: `${colorPalette.y}px` }" @mousedown="keepSelection" @keydown.esc="closeColorPalette">
      <button type="button" class="no-color" @click="pickFontColor(null)">{{ t('Automatic') }}</button>
      <p class="palette-label">{{ t('Theme colours') }}</p>
      <div class="swatches">
        <button v-for="c in colorChoices.theme" :key="c.theme" type="button" class="swatch" :class="{ current: fontColor?.hex === c.hex && fontColor?.theme === c.theme }" :style="{ background: `#${c.hex}` }" :title="c.label" :aria-label="c.label" @click="pickFontColor(c)" />
      </div>
      <template v-if="colorChoices.used.length">
        <p class="palette-label">{{ t('In this document') }}</p>
        <div class="swatches">
          <button v-for="c in colorChoices.used" :key="c.hex" type="button" class="swatch" :class="{ current: fontColor?.hex === c.hex && !fontColor?.theme }" :style="{ background: `#${c.hex}` }" :title="c.label" :aria-label="c.label" @click="pickFontColor(c)" />
        </div>
      </template>
      <MoreColors :current="fontColor && !fontColor.theme ? fontColor.hex : null" @pick="(hex) => pickFontColor({ hex, label: `#${hex}` })" />
    </div>

    <!-- A password: to protect the document with (or change, or take off), or to read a version to compare. -->
    <form v-if="pwDialog.open" class="chooser author-dialog drop-offer password-dialog" role="dialog" :aria-label="pwDialog.mode === 'compare' ? t('Password for {name}', { name: pwDialog.file.name }) : passwordProtected ? t('Change or remove password') : t('Protect with password')" @submit.prevent="confirmPasswordDialog" @keydown.esc="closePasswordDialog">
      <p class="pw-title"><strong>{{ pwDialog.mode === 'compare' ? t('Password for {name}', { name: pwDialog.file.name }) : passwordProtected ? t('Change or remove password') : t('Protect with password') }}</strong></p>
      <p id="pw-note" class="pw-note">{{ pwDialog.mode === 'compare'
        ? t('The other version is encrypted. Enter its password to compare it: it is checked on this device and never sent anywhere.')
        : t('Anyone who opens the downloaded file, in Word or here, needs this password. Keep it safe: a lost password can’t be recovered, and nor can the document. It stays on this device.') }}</p>
      <label for="pw-new">{{ pwDialog.mode === 'compare' ? t('Password') : passwordProtected ? t('New password') : t('Password') }}</label>
      <input id="pw-new" ref="pwInput" v-model="pwDialog.value" type="password" maxlength="255" :autocomplete="pwDialog.mode === 'compare' ? 'off' : 'new-password'" spellcheck="false" aria-describedby="pw-note pw-error" :aria-invalid="pwDialog.error ? 'true' : 'false'" :disabled="pwDialog.busy" />
      <template v-if="pwDialog.mode === 'set'">
        <label for="pw-again">{{ t('Type it again') }}</label>
        <input id="pw-again" v-model="pwDialog.again" type="password" maxlength="255" autocomplete="new-password" spellcheck="false" :disabled="pwDialog.busy" />
      </template>
      <p id="pw-error" class="pw-error" role="alert">{{ pwDialog.error }}</p>
      <div class="actions">
        <button type="submit" class="confirm" :disabled="pwDialog.busy">{{ pwDialog.mode === 'compare' ? t('Compare') : passwordProtected ? t('Change password') : t('Protect') }}</button>
        <button v-if="pwDialog.mode === 'set' && passwordProtected" type="button" :disabled="pwDialog.busy" @click="removePassword">{{ t('Remove password') }}</button>
        <button type="button" @click="closePasswordDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <form v-if="dropOffer.file" ref="dropOfferEl" class="chooser author-dialog drop-offer" role="dialog" :aria-label="t('Open in a new tab')" @submit.prevent="openDroppedInNewTab" @keydown.esc="dropOffer.file = null">
      <p>“{{ dropOffer.file.name }}” opens in a new tab, beside this document, which stays open here.</p>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('Open in a new tab') }}</button>
        <button type="button" @click="dropOffer.file = null">{{ t('Cancel') }}</button>
      </div>
    </form>

    <HelpDialog v-if="helpOpen" @close="closeHelp" @check="checkFromHelp" />
    <ThesaurusDialog v-if="thesaurusData" :word="thesaurusData.word" :lang="thesaurusData.lang" :participle="thesaurusData.participle" @close="closeThesaurus" @insert="replaceWithSynonym" />
    <AutoCorrectDialog v-if="autoCorrectDialog" :settings="autoCorrectSettings" :selected="autoCorrectDialog.selected" @close="closeAutoCorrect" @save="saveAutoCorrectSettings" />
    <WordCountDialog v-if="wordCountData" :words="words" :characters="chars" :pages="pageCount" :stats="wordCountData" :target="lengthTarget" :headline="headline" @close="closeWordCount" @save="setLengthTarget" @show-long="showLongSentence" />

    <ReviewDialog v-if="reviewData" :action="saveAction" :focus="reviewFocus" @release="releaseAction" @personal="setRemovePersonal" @personal-default="setPersonalDefault" @show-issue="showIssue" @title="setTitle" @update-properties="setUpdateProperties" :review="{ ...reviewData, resolved }" :skip="skipReview" @close="closeReview" @download="downloadReviewed" @skip="setSkipReview" @update-fields="(on) => { doc.updateFields = on; reviewData.fields.updateOnOpen = on; markDirty(); }" />

    <form v-if="authorDialog.open" ref="authorEl" class="chooser author-dialog" role="dialog" :aria-label="t('Your name on tracked changes')" :style="{ left: `${authorDialog.x}px`, top: `${authorDialog.y}px` }" @submit.prevent="confirmAuthor" @keydown.esc="closeAuthorDialog">
      <label for="author-name">{{ t('Your name on tracked changes') }}</label>
      <!-- No browser autocomplete: Firefox's list of earlier entries opens over the buttons below and
           takes the click meant for them. The name is remembered by Intact itself. -->
      <input id="author-name" ref="authorInput" v-model="authorDialog.name" type="text" maxlength="100" autocomplete="off" required />
      <div class="actions">
        <button type="submit" class="confirm">{{ tracking ? 'Save' : 'Start tracking' }}</button>
        <button type="button" @click="closeAuthorDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <form v-if="linkDialog.open" ref="linkEl" class="chooser link-dialog" role="dialog" :aria-label="t('Link')" :style="{ left: `${linkDialog.x}px`, top: `${linkDialog.y}px` }" @submit.prevent="confirmLink" @keydown.esc="closeLinkDialog">
      <label for="link-address">{{ t('Address: web page, e-mail or DOI') }}</label>
      <input id="link-address" ref="linkInput" v-model="linkDialog.address" type="text" autocomplete="off" spellcheck="false" :placeholder="t('https://… or 10.1000/xyz')" :aria-invalid="!!linkDialog.error" />
      <template v-if="linkDialog.askText">
        <label for="link-text">{{ t('Text to show') }}</label>
        <input id="link-text" v-model="linkDialog.text" type="text" autocomplete="off" :placeholder="linkDialog.address" />
      </template>
      <p v-if="linkDialog.error" class="error" role="alert">{{ linkDialog.error }}</p>
      <div class="actions">
        <button type="submit" class="confirm">{{ linkDialog.editing ? 'Update' : 'Insert' }}</button>
        <button v-if="linkDialog.editing" type="button" @click="removeLink">{{ t('Remove link') }}</button>
        <button type="button" @click="closeLinkDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <div v-if="insertMenu.open" ref="insertMenuEl" class="change-menu insert-menu" role="menu" :aria-label="t('Insert')" :style="{ left: `${insertMenu.x}px`, top: `${insertMenu.y}px` }" @mousedown.prevent @keydown.esc="closeInsertMenu(true)" @keydown="onMenuArrows">
      <button type="button" role="menuitem" @click="insertCheckbox">Check box<span class="hint">{{ t('☐ ticked with a click') }}</span></button>
      <p class="menu-label" role="presentation">{{ t('Today’s date') }}</p>
      <button v-for="d in insertMenu.dates" :key="d.style" type="button" role="menuitem" @click="insertDate(d)">{{ d.text }}</button>
      <button type="button" role="menuitemcheckbox" class="auto-date" :aria-checked="String(dateUpdates)" :title="t('As Word’s Date and Time dialog: the date as a field Word changes to that day’s date each time it opens or prints the file, or as text that stays as it is')" @click="toggleDateUpdates"><span class="check" aria-hidden="true">{{ dateUpdates ? '☑' : '☐' }}</span>Update automatically<span class="hint">{{ dateUpdates ? 'Word shows the day it’s opened' : 'Stays as today’s date' }}</span></button>
      <p class="menu-label" role="presentation">{{ t('Lines and shapes') }}</p>
      <button type="button" role="menuitem" @click="insertAnswerLine">{{ t('Answer line') }}<span class="hint">{{ t('Name: ______ to the margin') }}</span></button>
      <button type="button" role="menuitem" @click="insertShape('line')">{{ t('Line') }}<span class="hint">{{ t('In front of the text: drag it into place') }}</span></button>
      <button type="button" role="menuitem" @click="insertShape('arrow')">{{ t('Arrow') }}<span class="hint">{{ t('To label a diagram or a picture') }}</span></button>
      <button type="button" role="menuitem" @click="insertShape('textbox')">{{ t('Text box') }}<span class="hint">{{ t('In line with the text, to type in') }}</span></button>
      <p class="menu-label" role="presentation">{{ t('References') }}</p>
      <button type="button" role="menuitem" @click="insertTableOfContents">Table of contents<span class="hint">{{ t('From the headings') }}</span></button>
      <button type="button" role="menuitem" @click="openCaption">Caption…<span class="hint">{{ t('Figure 1, Table 1') }}</span></button>
      <button type="button" role="menuitem" @click="openCrossRef">Cross-reference…<span class="hint">{{ t('To a heading or caption') }}</span></button>
      <button type="button" role="menuitem" @click="openRefDialog('equation')">Equation…<span class="hint">{{ t('Typed as in LaTeX') }}</span></button>
      <button type="button" role="menuitem" @click="openRefDialog('citation')">Citation…<span class="hint">{{ t('APA or MLA') }}</span></button>
      <button type="button" role="menuitem" @click="insertBibliography">Bibliography<span class="hint">{{ doc.sources?.style === 'MLA' ? 'Works Cited' : 'References' }}</span></button>
    </div>

    <!-- As Word's Insert Page Number (Top of Page, Bottom of Page) and Edit Header / Edit Footer. -->
    <div v-if="pageNumberMenu.open" ref="pageNumberMenuEl" class="change-menu insert-menu page-number-menu" role="menu" :aria-label="t('Page number and header')" :style="{ left: `${pageNumberMenu.x}px`, top: `${pageNumberMenu.y}px` }" @mousedown.prevent @keydown.esc="closePageNumberMenu(true)" @keydown="onMenuArrows">
      <p class="menu-label" role="presentation">{{ t('Top of page (header)') }}</p>
      <button v-for="a in PAGE_NUMBER_ALIGNS" :key="`header-${a.value}`" type="button" role="menuitem" :aria-label="t('Page number at the top, {where}', { where: t(a.label).toLowerCase() })" @click="pageNumberAt('header', a.value)">{{ t(a.label) }}<span v-if="a.value === 'right'" class="hint">{{ t('Surname 1 (MLA)') }}</span></button>
      <p class="menu-label" role="presentation">{{ t('Bottom of page (footer)') }}</p>
      <button v-for="a in PAGE_NUMBER_ALIGNS" :key="`footer-${a.value}`" type="button" role="menuitem" :aria-label="t('Page number at the bottom, {where}', { where: t(a.label).toLowerCase() })" @click="pageNumberAt('footer', a.value)">{{ t(a.label) }}</button>
      <button type="button" role="menuitem" class="menu-sep" @click="openPageNumberFormat">{{ t('Format page numbers…') }}<span class="hint">{{ t('Number format, and where numbering starts') }}</span></button>
      <button type="button" role="menuitem" class="menu-sep" @click="editHeaderFooter('header')">{{ t('Edit header') }}<span class="hint">{{ t('Text on every page, at the top') }}</span></button>
      <button type="button" role="menuitem" @click="editHeaderFooter('footer')">{{ t('Edit footer') }}<span class="hint">{{ t('Text on every page, at the bottom') }}</span></button>
    </div>

    <!-- As Word's Format Page Numbers: the number format and where numbering starts, for a section
         (or, from the caret on, a new one), written in its w:pgNumType. -->
    <form v-if="numDialog.open" ref="numDialogEl" class="chooser author-dialog para-dialog page-dialog" role="dialog" :aria-label="t('Page number format')" :style="{ left: `${numDialog.x}px`, top: `${numDialog.y}px`, maxHeight: `calc(100vh - ${numDialog.y + 8}px)` }" novalidate @submit.prevent="applyPageNumberFormat" @keydown.esc="closePageNumberFormat">
      <label class="apply-to">{{ t('Apply to') }}
        <select v-model="numDialog.scope" :aria-label="t('Apply page number format to')">
          <option v-if="numDialog.sections > 1" value="section">{{ t('This section ({n} of {total})', { n: numDialog.sectionNo, total: numDialog.sections }) }}</option>
          <option value="whole">{{ t('Whole document') }}</option>
          <option v-if="numDialog.canSplit" value="forward">{{ t('This point forward') }}</option>
        </select>
      </label>
      <p v-if="numDialog.scope === 'forward'" class="note">{{ t('A section break (next page) goes before the paragraph with the caret: the pages from there on are numbered as below, the ones before keep theirs.') }}</p>
      <label class="apply-to">{{ t('Number format') }}
        <select v-model="numDialog.fmt" :aria-label="t('Number format')">
          <option v-for="f in PAGE_NUMBER_FORMATS" :key="f.value" :value="f.value">{{ f.label }}</option>
        </select>
      </label>
      <fieldset>
        <legend>{{ t('Page numbering') }}</legend>
        <label class="check"><input v-model="numDialog.restart" type="radio" name="page-numbering" :value="false" /> {{ t('Continue from previous section') }}</label>
        <label class="check"><input v-model="numDialog.restart" type="radio" name="page-numbering" :value="true" /> {{ t('Start at') }}</label>
        <label>{{ t('Start at page number') }} <input v-model.number="numDialog.start" type="number" min="0" step="1" :disabled="!numDialog.restart" /></label>
      </fieldset>
      <p class="note">{{ t('Page numbers in Word and when printed; here a page number shows as #.') }}</p>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('OK') }}</button>
        <button type="button" @click="closePageNumberFormat">{{ t('Cancel') }}</button>
      </div>
    </form>

    <!-- Word's Numbering Library and Multilevel List, with Indonesia's legal numbering (Pasal, ayat, huruf). -->
    <div v-if="listLibrary.open" ref="listLibraryEl" class="change-menu insert-menu list-library" role="menu" :aria-label="t('Numbering library')" :style="{ left: `${listLibrary.x}px`, top: `${listLibrary.y}px` }" @mousedown.prevent @keydown.esc="closeListLibrary(true)" @keydown="onMenuArrows">
      <button v-for="p in LIST_LIBRARY" :key="p.id" type="button" role="menuitem" @click="applyListPreset(p.id)">{{ t(p.label) }}<span class="hint">{{ t(p.hint) }}</span></button>
    </div>

    <!-- Word's Sort (Table Layout > Sort): the rows below the heading row, by a column's text, number or date. -->
    <form v-if="sortDialog.open" ref="sortDialogEl" class="chooser author-dialog para-dialog" role="dialog" :aria-label="t('Sort')" :style="{ left: `${sortDialog.x}px`, top: `${sortDialog.y}px` }" novalidate @submit.prevent="sortTable" @keydown.esc="closeSortDialog">
      <label>{{ t('Sort by') }}
        <select v-model.number="sortDialog.col" :aria-label="t('Sort by')">
          <option v-for="c in sortDialog.columns" :key="c.value" :value="c.value">{{ c.label }}</option>
        </select>
      </label>
      <label>{{ t('Type') }}
        <select v-model="sortDialog.type" :aria-label="t('Type')">
          <option value="text">{{ t('Text') }}</option>
          <option value="number">{{ t('Number') }}</option>
          <option value="date">{{ t('Date') }}</option>
        </select>
      </label>
      <label class="check"><input v-model="sortDialog.desc" type="radio" name="sort-order" :value="false" /> {{ t('Ascending') }}</label>
      <label class="check"><input v-model="sortDialog.desc" type="radio" name="sort-order" :value="true" /> {{ t('Descending') }}</label>
      <label class="check"><input v-model="sortDialog.header" type="checkbox" @change="refreshSortColumns" /> {{ t('My table has a header row') }}</label>
      <div class="actions">
        <button type="submit" class="confirm">{{ t('OK') }}</button>
        <button type="button" @click="closeSortDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <!-- Word's Formula (Table Layout > Formula): an = field in the cell, worked out now and on F9. -->
    <form v-if="formulaDialog.open" ref="formulaDialogEl" class="chooser author-dialog para-dialog formula-dialog" role="dialog" :aria-label="t('Formula')" :style="{ left: `${formulaDialog.x}px`, top: `${formulaDialog.y}px` }" novalidate @submit.prevent="insertFormula" @keydown.esc="closeFormulaDialog">
      <label>{{ t('Formula') }} <input v-model="formulaDialog.formula" type="text" spellcheck="false" autocomplete="off" :aria-label="t('Formula')" /></label>
      <label>{{ t('Number format') }}
        <select v-model="formulaDialog.picture" :aria-label="t('Number format')">
          <option value="">{{ t('As worked out') }}</option>
          <option v-for="f in formulaPictures" :key="f" :value="f">{{ f }}</option>
        </select>
      </label>
      <label>{{ t('Paste function') }}
        <select :aria-label="t('Paste function')" @change="pasteFunction($event)">
          <option value="" />
          <option v-for="f in FORMULA_FUNCTIONS" :key="f" :value="f">{{ f }}</option>
        </select>
      </label>
      <p class="note" :class="{ error: formulaPreview.error }" role="status">{{ formulaPreview.text }}</p>
      <p class="note">{{ t('ABOVE, LEFT, BELOW and RIGHT take the numbers beside the cell, up to the first cell that isn’t one; cells are A1, B2… Update fields (F9) works it out again.') }}</p>
      <div class="actions">
        <button type="submit" class="confirm" :disabled="formulaPreview.error">{{ t('OK') }}</button>
        <button type="button" @click="closeFormulaDialog">{{ t('Cancel') }}</button>
      </div>
    </form>

    <CompareDialog v-if="compare.result" :result="compare.result" :other="compare.other" :author="author" :protected="passwordProtected" @close="closeCompare" @go="goToChange" @redline="sendRedline" />
    <input ref="compareInput" type="file" hidden accept=".docx,.dotx,.docm,.dotm,application/vnd.openxmlformats-officedocument.wordprocessingml.document" @change="onCompareFile" />

    <div v-if="caseMenu.open" ref="caseMenuEl" class="change-menu case-menu" role="menu" :aria-label="t('Change case')" :style="{ left: `${caseMenu.x}px`, top: `${caseMenu.y}px` }" @mousedown.prevent @keydown.esc="closeCaseMenu(true)" @keydown="onMenuArrows">
      <button v-for="c in CASES" :key="c.id" type="button" role="menuitem" @click="changeCaseTo(c.id)">{{ c.label }}</button>
    </div>

    <!-- The right-click menu, as Word's: what can be done where it was opened. Shift+right-click
         gives the browser's own menu (with its spelling suggestions). -->
    <div v-if="ctxMenu.open" ref="ctxMenuEl" class="change-menu context-menu" role="menu" :aria-label="t('Context menu')" :style="{ left: `${ctxMenu.x}px`, top: `${ctxMenu.y}px` }" @mousedown.prevent @contextmenu.prevent @keydown.esc="closeContextMenu(true)" @keydown="onMenuArrows">
      <template v-if="!ctxMenu.locked">
        <button type="button" role="menuitem" :disabled="!ctxMenu.selected" @click="contextCut">Cut<span class="key">{{ t('Ctrl+X') }}</span></button>
      </template>
      <button type="button" role="menuitem" :disabled="!ctxMenu.selected" @click="contextCopy">Copy<span class="key">{{ t('Ctrl+C') }}</span></button>
      <button v-if="ctxMenu.selected" type="button" role="menuitem" @click="contextRun(() => copyForWeb('html', true))">{{ t('Copy for the web') }}</button>
      <template v-if="!ctxMenu.locked">
        <button type="button" role="menuitem" @click="contextPaste(false)">Paste<span class="key">{{ t('Ctrl+V') }}</span></button>
        <button type="button" role="menuitem" @click="contextPaste(true)">{{ t('Paste as text') }}</button>
      </template>
      <template v-if="ctxMenu.change && !lockMode">
        <hr />
        <button type="button" role="menuitem" @click="contextRun(() => resolveCurrent(true, 'stay'))">Accept change</button>
        <button type="button" role="menuitem" @click="contextRun(() => resolveCurrent(false, 'stay'))">Reject change</button>
      </template>
      <hr v-if="!ctxMenu.locked || canComment" />
      <template v-if="!ctxMenu.locked">
        <template v-if="ctxMenu.link">
          <button type="button" role="menuitem" @click="contextOpenLink">{{ t('Open link') }}</button>
          <button type="button" role="menuitem" @click="contextRun(openLinkDialog, true)">{{ t('Edit link…') }}</button>
          <button type="button" role="menuitem" @click="contextRun(removeLinkAtCaret)">{{ t('Remove link') }}</button>
        </template>
        <button v-else type="button" role="menuitem" @click="contextRun(openLinkDialog, true)">Link…<span class="key">{{ t('Ctrl+K') }}</span></button>
      </template>
      <button v-if="canComment" type="button" role="menuitem" @click="contextRun(startComment)">New comment<span class="key">{{ t('Ctrl+Alt+M') }}</span></button>
      <template v-if="!ctxMenu.locked">
        <button v-if="ctxMenu.selected" type="button" role="menuitem" @click="contextRun(toggleHidden)">{{ t('Hidden text') }}<span class="key">{{ t('Ctrl+Shift+H') }}</span></button>
        <button v-if="ctxMenu.selected" type="button" role="menuitem" @click="contextRun(redactSelection)">{{ t('Redact') }}</button>
        <button type="button" role="menuitem" @click="contextRun(openThesaurus)">{{ t('Thesaurus…') }}<span class="key">{{ t('Shift+F7') }}</span></button>
        <button type="button" role="menuitem" @click="contextRun(openParagraphDialog, true)">{{ t('Paragraph…') }}</button>
        <button type="button" role="menuitem" @click="contextRun(openModifyStyle)">{{ t('Modify style…') }}</button>
        <template v-if="ctxMenu.table">
          <hr />
          <button type="button" role="menuitem" @click="contextRun(() => tableAction('rowsAbove'))">Insert row above</button>
          <button type="button" role="menuitem" @click="contextRun(() => tableAction('rowsBelow'))">Insert row below</button>
          <button type="button" role="menuitem" @click="contextRun(() => tableAction('columnsLeft'))">Insert column to the left</button>
          <button type="button" role="menuitem" @click="contextRun(() => tableAction('columnsRight'))">Insert column to the right</button>
          <button v-if="ctxMenu.multi" type="button" role="menuitem" @click="contextRun(() => tableAction('merge'))">Merge cells</button>
          <button type="button" role="menuitem" @click="contextRun(openSplitDialog, true)">{{ t('Split cells…') }}</button>
          <button type="button" role="menuitem" @click="contextRun(openCellFormat, true)">{{ t('Shading and borders…') }}</button>
          <button type="button" role="menuitem" @click="contextRun(() => tableAction('deleteRows'))">Delete row</button>
          <button type="button" role="menuitem" @click="contextRun(() => tableAction('deleteColumns'))">Delete column</button>
          <button type="button" role="menuitem" @click="contextRun(() => tableAction('deleteTable'))">Delete table</button>
        </template>
      </template>
      <p class="menu-note" role="presentation">{{ t('Shift+right-click: the browser’s menu, with spelling suggestions') }}</p>
    </div>

    <div v-if="shareReady.file" ref="shareReadyEl" class="chooser share-ready" role="dialog" :aria-label="t('Share')" @keydown.esc="shareReady.file = null">
      <p>{{ shareReady.file.name }} is ready.</p>
      <div class="actions">
        <button type="button" class="confirm" @click="shareFile(shareReady.file)">{{ t('Share…') }}</button>
        <button type="button" @click="shareReady.file = null">{{ t('Cancel') }}</button>
      </div>
    </div>

    <SignatureDialog v-if="signatureOpen" @close="closeSignature" @insert="insertSignature" />
    <SymbolDialog v-if="symbolOpen" @close="closeSymbols" @insert="insertSymbol" />
    <MailMergeDialog v-if="mailMerge.open" :placeholders="mailMerge.placeholders" :single-blocked="mailMerge.singleBlocked" :merge="runMailMerge" @close="closeMailMerge" />
    <EquationDialog v-if="refDialog === 'equation'" @close="closeRefDialog" @insert="insertEquation" />
    <CitationDialog v-if="refDialog === 'citation'" :sources="doc.sources.list" :style="doc.sources.style" @close="closeRefDialog" @insert="insertCitation" />
    <CaptionDialog v-if="refDialog === 'caption'" :counts="captionInfo.counts" :target="captionInfo.target" :table="captionInfo.table" @close="closeRefDialog" @insert="insertCaption" />
    <CrossRefDialog v-if="refDialog === 'xref'" :targets="xrefTargets" @close="closeRefDialog" @insert="insertCrossRef" />

    <footer class="statusbar">
      <span :title="`${chars.withSpaces} characters with spaces, ${chars.withoutSpaces} without`"><button type="button" class="word-count" :class="{ over: overTarget }" :aria-label="t('Word count, length and readability')" @click="openWordCount">{{ t('{words} words · {chars} characters', { words: selectedWords != null ? t('{n} of {total}', { n: selectedWords, total: words }) : lengthTarget.words ? `${words} / ${lengthTarget.words}` : words, chars: selectedChars != null ? t('{n} of {total}', { n: selectedChars, total: chars.withSpaces }) : chars.withSpaces }) }}</button></span>
      <button v-if="lengthTarget.headline && headline" type="button" class="headline-count" :class="{ over: headline.chars > lengthTarget.headline }" :title="headlineTitle" @click="openWordCount">{{ t('Headline {chars} / {limit}', { chars: headline.chars, limit: lengthTarget.headline }) }}</button>
      <button type="button" class="nav-toggle" :aria-pressed="navOpen" :title="t('Headings: move around the document by its headings, as Word’s Navigation pane')" @click="toggleNav">{{ t('Headings') }}</button>
      <button v-if="pageCount" type="button" class="page-count" :aria-pressed="showPageBreaks" :title="pageCountTitle" :aria-label="`About ${pageCount} ${pageCount === 1 ? 'page' : 'pages'}: show where the pages break`" @click="togglePageBreaks">≈ {{ pageCount }} {{ pageCount === 1 ? t('page') : t('pages') }}</button>
      <button type="button" class="smart-typing" :aria-pressed="autoFormat" :title="autoFormat ? 'As you type, as in Word: &quot; and \' give curly quotes, -- a dash, ... an ellipsis, (c) ©, --&gt; an arrow, 1. or * then a space at the start of a paragraph a list, and a web or e-mail address then a space a link – click to type them as they are' : 'Quotes, dashes, ellipses, list numbers and addresses are typed as they are – click to format them as you type, as in Word'" @click="toggleAutoFormat">AutoFormat: {{ autoFormat ? 'on' : 'off' }}</button>
      <!-- Word's AutoCorrect Options, which hold AutoFormat As You Type too: your entries and common typos. -->
      <button type="button" class="smart-typing" :aria-label="t('AutoCorrect options')" :title="t('AutoCorrect: your own words to replace as you type (cllr → Councillor), and common typos')" @click="openAutoCorrect">{{ t('AutoCorrect') }}</button>
      <span v-if="fontsMissing.length" class="fonts-missing" :title="`Not on this device, so shown in another font: ${fontsMissing.join(', ')}. The document keeps them: the text looks as it should in Word on a computer that has them.`">{{ fontsMissing.length === 1 ? `Font shown with a substitute: ${fontsMissing[0]}` : `${fontsMissing.length} fonts shown with substitutes` }}</span>
      <button v-if="tracking" type="button" class="tracking-status" :title="t('Change your name on tracked changes')" @click="openAuthorDialog">{{ t('Tracking changes as {name}', { name: author }) }}</button>
      <span v-if="draftKept" class="draft-status" :title="draftTitle">{{ t('Changes kept on this device') }}</span>
      <span role="status" aria-live="polite" :title="message || null">{{ message }}</span>
      <span class="zoom" role="group" :aria-label="t('Zoom')">
        <button type="button" :title="t('Zoom out (Ctrl+mouse wheel)')" :aria-label="t('Zoom out')" :disabled="zoomLevel <= ZOOMS[0]" @click="zoomBy(-1)">−</button>
        <button type="button" class="zoom-level" :title="zoomLevel === 100 ? 'Zoom: the page at its size' : 'Back to 100%'" :aria-label="`Zoom ${zoomLevel}%${zoomLevel === 100 ? '' : ', back to 100%'}`" @click="setZoom(100)">{{ zoomLevel }}%</button>
        <button type="button" :title="t('Zoom in (Ctrl+mouse wheel)')" :aria-label="t('Zoom in')" :disabled="zoomLevel >= ZOOMS[ZOOMS.length - 1]" @click="zoomBy(1)">+</button>
      </span>
    </footer>
  </div>
</template>

<script setup>
import { ref, shallowRef, reactive, computed, watch, nextTick, onMounted, onBeforeUnmount } from 'vue';
import { canPickSaveFile, pickFileToSave, mayWrite, writeFile } from '../editor/fileHandles.js';
import { useRouter, onBeforeRouteLeave } from 'vue-router';
import AppHeader from '../components/AppHeader.vue';
import FindBar from '../components/FindBar.vue';
import FillInDialog from '../components/FillInDialog.vue';
import MoreColors from '../components/MoreColors.vue';
import SymbolDialog from '../components/SymbolDialog.vue';
import MailMergeDialog from '../components/MailMergeDialog.vue';
import JSZip from 'jszip';
import { mailMerge as mergeDocuments, singleBlocker } from '../docx/merge.js';
import EquationDialog from '../components/EquationDialog.vue';
import CitationDialog from '../components/CitationDialog.vue';
import CaptionDialog from '../components/CaptionDialog.vue';
import CrossRefDialog from '../components/CrossRefDialog.vue';
import NavigationPane from '../components/NavigationPane.vue';
import { tocXml, captionXml, bibliographyXml, citationXml } from '../docx/references.js';
import { t, lang } from '../i18n.js';
import CompareDialog from '../components/CompareDialog.vue';
import { paragraphsOfXml, compareParagraphs } from '../editor/compare.js';
import { redlinePackage } from '../docx/redline.js';
import { evaluateFormula, formatResult, formulaInstr, parseCellNumber, separators, FORMULA_FUNCTIONS } from '../docx/formulas.js';
import { CASES, nextCase, changeCase, caseLabel } from '../editor/case.js';
import { nameOf } from '../editor/symbols.js';
import { elementChildren } from '../docx/xml.js';
import { columnsWrapper } from '../docx/DocxDocument.js';
import { wordDatePicture, dateFieldInstruction } from '../editor/dates.js';
import { findPlaceholders } from '../editor/placeholders.js';
import CommentsPanel from '../components/CommentsPanel.vue';
import ReviewDialog from '../components/ReviewDialog.vue';
import SignatureDialog from '../components/SignatureDialog.vue';
import HelpDialog from '../components/HelpDialog.vue';
import WordCountDialog from '../components/WordCountDialog.vue';
import ThesaurusDialog from '../components/ThesaurusDialog.vue';
import AutoCorrectDialog from '../components/AutoCorrectDialog.vue';
import { thesaurusLanguage, matchCase } from '../editor/thesaurus.js';
import { participleAfter } from '../editor/inflect.js';
import { loadAutoCorrect, saveAutoCorrect, autoCorrectWord } from '../editor/autoCorrectList.js';
import { webContent, toHtml, toText, toMarkdown, copyFormats } from '../editor/webCopy.js';
import { wordCount, readability, headlineOf, blockText, textRange } from '../editor/readability.js';
import { findMatches, matchRange } from '../editor/find.js';
import { currentDoc, setDocument } from '../store.js';
import { saveDraft, deleteDraft, newDraftId, MAX_AGE } from '../drafts.js';
import { openStartInNewTab, openFileInNewTab, holdDraft } from '../tabs.js';
import { createHistory } from '../editor/history.js';
import { isEmptyBlock, isAtBlockStart, adjacentCell, selectCell } from '../editor/structure.js';
import { controlAt, controlBox, setControlText, clearPlaceholder, fieldAtSelection } from '../editor/controls.js';
import { parseClipboard, runNodes, insertAtCaret, newLink, tablesAsText } from '../editor/paste.js';
import { normalizeLink } from '../docx/links.js';
import {
  removesContent, rangesRemovedBy, trackedChangeAt, refusalMessage, commentMarkers, restoreCommentMarkers,
  onlyHiddenMarkers, beyondHiddenMarkers,
} from '../editor/protect.js';
import { readOnlyAt, readOnlyMessage, isEditingKey, caretBeside } from '../editor/readonly.js';
import { PROTECTION_LABELS, sniffFile } from '../docx/package.js';
import { decrypt } from '../docx/encryption.js';
import { checkAccessibility } from '../editor/accessibility.js';
import { drawLineNumbers } from '../editor/lineNumbers.js';
import { estimatePages, drawPageBreaks, UNIT_SELECTOR } from '../editor/pages.js';
import { printCss, basicPrintCss, printModes } from '../editor/printBands.js';
import { PERSONAL } from '../docx/properties.js';
import { checkScreening } from '../editor/screening.js';
import { shortcutFor } from '../editor/shortcuts.js';
import { HIGHLIGHT_COLORS, setVertAlign, setHighlight, formatState, portions } from '../editor/inline.js';
import { drawMarks } from '../editor/marks.js';
import { rovingToolbar, fitLabels } from '../editor/toolbar.js';
import { fitPictures, pictureZoom } from '../editor/fitPictures.js';
import { changeItems, changesAt, adjacentChange, itemRange, mergeGroup } from '../editor/changes.js';
import { HIGHLIGHT } from '../docx/styles.js';
import { WRAPS, PICTURE_STYLES, SHAPES, BEVELS, CAMERAS, ARTISTIC, MATERIALS, LIGHTS } from '../docx/pictures.js';
import { CAMERA_ANGLES } from '../docx/lighting.js';
import { renderArtistic } from '../editor/artistic.js';
import { pictureOutline } from '../editor/outline.js';
import { missingFonts } from '../docx/fonts.js';
import { createReader } from '../editor/readAloud.js';
import { restyle, noteRuns, trackRuns, setFont, setSize, setColor, setHidden, setUnderline, setCharacterStyle, clearDirect, changeSize, nextFontSize } from '../editor/runs.js';
import { smartQuote, smartReplacement, listTrigger, autoCorrect, linkTrigger } from '../editor/smartTyping.js';
import { formatNumber } from '../docx/numbering.js';
import { updateFields, parseInstr } from '../docx/fieldUpdates.js';
import { UNIT_NAMES, EMU_PER, unitFor, saveUnit, fromTwips, toTwips, switchUnit } from '../editor/units.js';
import {
  revisionDate, typeTracked, typeUntracked, insertTracked, deleteTracked, unitRange, markSplit, wrapInsertedBreak, wrapComposedText,
  inField, replaceTracked,
} from '../editor/track.js';

const router = useRouter();
const doc = currentDoc.value;
const page = ref(null);
const dirty = ref(!!doc.restored); // a restored draft holds changes not yet downloaded
const saving = ref(false);
const message = ref('');
const words = ref(0);
const selectedWords = ref(null); // words in the selection, as Word's status bar counts them; null if none
const chars = reactive({ withSpaces: 0, withoutSpaces: 0 }); // characters in the text, as Word counts them
const selectedChars = ref(null);
// A length to write to, as a newsroom sets one: a target in words and a limit for the headline in
// characters (the Word count dialog). Kept with the document while it is open here, and in its draft,
// never in the file.
const lengthTarget = reactive({ words: doc.restoredFrom?.length?.words ?? null, headline: doc.restoredFrom?.length?.headline ?? null });
const headline = ref(null); // { text, chars } of the text's first paragraph, when it is a heading
const wordCountData = ref(null); // the Word count dialog's statistics, while it is open
const baseName = ref(doc.name.replace(/\.(docx|dotx|docm|dotm)$/i, ''));
// Templates and files with macros are saved as what they are: Word refuses a .docx holding macros.
// A template can be made into a new document instead (see templateBar below).
const fromTemplate = ref(doc.fromTemplate);
const fileExt = computed(() => (fromTemplate.value, doc.format.ext));
const extTitle = computed(() => ({ docx: 'Word document', dotx: 'Word template', docm: 'Word document with macros', dotm: 'Word template with macros' }[fileExt.value]));
// The document's own paragraph styles, offered in the style menu after the built-in ones.
const styleChoices = ref([]);
const knownBlock = computed(() => ['normal', 'title', 'h1', 'h2', 'h3'].includes(state.block)
  || styleChoices.value.some((c) => `style:${c.id}` === state.block));
const state = reactive({ bold: false, italic: false, underline: false, strike: false, sup: false, sub: false, hidden: false, highlight: null, block: 'normal', align: 'left', list: null, table: false, font: '', size: 0, charStyle: '' });
const chooser = reactive({ open: false, kind: '', x: 0, y: 0, choices: [], current: '', date: '' });
const chooserEl = ref(null);
const threads = ref([]);
const commentsOpen = ref(false);
const activeComment = ref(null);
const findOpen = ref(false);

// Track Changes: on when the document says so (as in Word), or when switched on here. The author's
// name is remembered in this browser only.
const AUTHOR_KEY = 'intact.author';
const stored = (() => { try { return localStorage.getItem(AUTHOR_KEY) || ''; } catch { return ''; } })();
const tracking = ref(doc.tracking);
const author = ref(stored || 'Author');
const named = ref(!!stored); // the user has given a name
const authorDialog = reactive({ open: false, name: '', x: 0, y: 0, start: false });
const trackButton = ref(null);
const authorEl = ref(null);
const authorInput = ref(null);
watch(tracking, (on) => { doc.tracking = on; });

// ------------------------------------------------------------------ fill in placeholders

// "[Client name]", "{{date}}", «Amount» and <<Amount>> in the document (see editor/placeholders.js),
// counted with the words, and filled in from a dialog, each once for everywhere it appears.
const placeholderCount = ref(0);
const fillIn = reactive({ open: false, x: 0, y: 0 });
const fillInButton = ref(null);
const fillInDialog = ref(null);

function countPlaceholders() {
  placeholderCount.value = page.value ? findPlaceholders(page.value).filter((p) => p.editable).length : 0;
}

function openFillIn() {
  if (textLocked.value) return explainLock();
  findOpen.value = false; // both paint the same highlights
  const rect = fillInButton.value?.getBoundingClientRect() ?? { left: window.innerWidth, bottom: 96 };
  Object.assign(fillIn, { open: true, x: Math.max(8, rect.left - 180), y: rect.bottom + 4 });
  nextTick(() => keepInView(fillIn, fillInDialog.value?.$el));
}

function closeFillIn() {
  fillIn.open = false;
  page.value.focus();
}

function onFilled(text) {
  countPlaceholders();
  nextTick(() => { message.value = text; }); // after the edit's own status bar update
}

// ------------------------------------------------------------------ templates

/** Makes the template opened into a new document, or goes back to saving the template itself. */
function setFromTemplate(on) {
  doc.newFromTemplate(on);
  fromTemplate.value = doc.fromTemplate;
  markDirty();
  message.value = on
    ? `New document: it downloads as ${baseName.value.trim() || 'document'}.${fileExt.value}, and the template stays as it is`
    : `Editing the template: it downloads as ${baseName.value.trim() || 'document'}.${fileExt.value}`;
}

// ------------------------------------------------------------------ restrictions

// Word's Restrict Editing, digital signatures and "read-only recommended" (see docx/package.js).
// The document opens as Word would let you edit it: read-only, comments only, filling in its fields
// only, or with every edit tracked. "Edit anyway" lifts that for this session, after saying what it
// means; the restriction itself stays in the file, and a signature goes when the file changes.
const unlocked = ref(!!doc.restoredFrom?.unlocked); // as it was when the draft was kept
const count = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;
/** How the document may be edited here: readOnly, comments, forms, trackedChanges or null (freely). */
const lockMode = computed(() => {
  if (unlocked.value) return null;
  const mode = doc.protection?.mode;
  if (mode === 'readOnly' || doc.signatures || doc.writeProtection) return 'readOnly';
  return mode || null;
});
// How tracked changes show, as Word's Display for Review (see setMarkup): 'all', 'simple', 'none' or
// 'original'. Declared here, as showing the original locks the text (below).
const markup = ref('all');
/** The text can't be changed (only comments, or only fields, at most), or the original is shown. */
const textLocked = computed(() => ['readOnly', 'comments', 'forms'].includes(lockMode.value) || markup.value === 'original');
const canComment = computed(() => !textLocked.value || lockMode.value === 'comments');
const restrictions = computed(() => {
  const out = [];
  const p = doc.protection;
  if (p) {
    out.push({
      title: 'Restrict Editing is on.',
      text: `Word allows ${{ readOnly: 'no changes (read only)', comments: 'comments only', forms: 'filling in forms only', trackedChanges: 'tracked changes only' }[p.mode]}${p.password ? ', and asks for a password to stop the protection' : ''}.`,
    });
  }
  if (doc.signatures) {
    out.push({ title: 'Digitally signed.', text: `The document has ${count(doc.signatures.count, 'digital signature')}; changing it removes ${doc.signatures.count === 1 ? 'it' : 'them'}, as in Word.` });
  }
  if (doc.format.macros) {
    out.push({ title: 'Contains macros.', text: 'Intact never runs them: they are kept in the file as they are. Enable them in Word only if you trust where the file came from.' });
  }
  const w = doc.writeProtection;
  if (w) {
    out.push({
      title: w.recommended ? 'Read-only recommended.' : 'Password to modify.',
      text: w.recommended ? 'The author recommends opening it read-only unless you need to change it.' : 'Word asks for a password to save changes to this file; edits can be saved as a copy.',
    });
  }
  return out;
});
const lockText = computed(() => ({
  readOnly: 'Opened read-only.',
  comments: 'You can add and answer comments; the text is read-only.',
  forms: 'You can fill in its fields; the rest is read-only.',
  trackedChanges: 'Your edits are tracked, and tracked changes can’t be accepted or rejected.',
}[lockMode.value] || ''));
const unlockedText = computed(() => [
  'Editing anyway.',
  doc.protection && 'The restriction stays in the file you download, so Word still applies it.',
  doc.signatures && 'If you change anything, the downloaded file has no signature.',
].filter(Boolean).join(' '));
/** Why the document can't be changed here, for the status bar and tooltips. */
const lockMessage = computed(() => {
  const why = doc.protection && doc.protection.mode === lockMode.value
    ? `Restrict Editing in Word allows ${{ readOnly: 'no changes', comments: 'comments only', forms: 'filling in forms only', trackedChanges: 'tracked changes only' }[lockMode.value]}`
    : doc.signatures ? 'the document is digitally signed' : 'the author recommends opening it read-only';
  return lockMode.value === 'trackedChanges' ? `${why}` : `${why} – choose Edit anyway above to change it`;
});

// ------------------------------------------------------------------ sensitivity label

// The sensitivity label (Microsoft Purview) or classification marking the document carries, shown
// beside its name, as Office shows it. Never changed here: the file keeps it as it is.
const labelChip = computed(() => {
  const { label, marking } = doc.properties || {};
  if (label) {
    const how = { privileged: 'chosen by a person', standard: 'applied by default or automatically' }[(label.method || '').toLowerCase()];
    const when = label.date && !Number.isNaN(Date.parse(label.date)) ? new Date(label.date).toLocaleDateString(undefined, { dateStyle: 'medium' }) : '';
    const marks = label.markings.filter((m) => m !== 'encryption');
    return {
      text: label.name || 'Sensitivity label',
      title: [
        `Sensitivity label: ${label.name || 'set (its name isn’t stored in the file)'}`,
        how || when ? ` – ${[how, when && `on ${when}`].filter(Boolean).join(', ')}` : '',
        '.',
        marks.length ? ` In Office it marks the document with a ${marks.join(', ').replace(/, ([^,]*)$/, ' and $1')}.` : '',
        ' Intact keeps the label as it is: change it in Word.',
      ].join(''),
    };
  }
  if (marking) return { text: marking.value, title: `Classification (${marking.name}): ${marking.value}. Intact keeps it as it is: change it in Word.` };
  return null;
});

// ------------------------------------------------------------------ thesaurus

// Word's Thesaurus (Shift+F7): the word at the caret, or selected, looked up; a word chosen goes in
// its place (tracked while tracking), in its case. See components/ThesaurusDialog.vue.
const thesaurusData = ref(null); // { word, lang } while it is open
let thesaurusRange = null; // the word it replaces

function openThesaurus() {
  restoreSelection();
  let range = selectionInPage() || savedRange;
  if (range && !page.value.contains(range.commonAncestorContainer)) range = null;
  if (range?.collapsed) range = wordAt(range);
  let word = range ? range.toString().replace(/\s+/g, ' ').trim() : '';
  if (word.length > 60 || !/[\p{L}]/u.test(word)) word = '';
  thesaurusRange = word ? range.cloneRange() : null;
  const node = range?.startContainer || savedRange?.startContainer;
  const el = node?.nodeType === 3 ? node.parentElement : node;
  const tag = el?.closest?.('[lang]')?.lang || doc.styles.language() || (lang.value === 'id' ? 'id' : 'en');
  // "has approved", "was approved": an -ed form there is a past participle (see editor/inflect.js).
  let participle = false;
  const block = word && closestBlock(range.startContainer);
  if (block) {
    const lead = document.createRange();
    lead.setStart(block, 0);
    lead.setEnd(range.startContainer, range.startOffset);
    participle = participleAfter(lead.toString());
  }
  thesaurusData.value = { word, lang: thesaurusLanguage(tag), participle };
}

function closeThesaurus() {
  thesaurusData.value = null;
  refocus();
}

/** Puts the word chosen in the thesaurus in place of the word looked up (or at the caret). */
function replaceWithSynonym(word) {
  const was = thesaurusData.value?.word || '';
  thesaurusData.value = null;
  if (textLocked.value) return explainLock();
  page.value.focus();
  if (thesaurusRange && page.value.contains(thesaurusRange.startContainer)) selectRange(thesaurusRange);
  else restoreSelection();
  const text = was ? matchCase(word, was) : word;
  typeText(text);
  undoStack.commit();
  markDirty();
  const tracked = tracking.value ? t(' (tracked)') : '';
  message.value = was ? t('“{old}” replaced with “{new}”{tracked}', { old: was, new: text, tracked }) : t('“{new}” inserted{tracked}', { new: text, tracked });
}

// ------------------------------------------------------------------ AutoCorrect entries

// Word's AutoCorrect "Replace text as you type": your entries and common English typos, as the word
// is ended by a space, punctuation or Enter (see editor/autoCorrectList.js). Kept in this browser.
const autoCorrectSettings = ref(loadAutoCorrect());
const autoCorrectDialog = ref(null); // { selected } while it is open
const ENDS_WORD = /^[\s.,;:!?)\]}"'”’»]$/u;

function openAutoCorrect() {
  const range = selectionInPage() || savedRange;
  const selected = range && !range.collapsed && page.value.contains(range.commonAncestorContainer) ? range.toString().replace(/\s+/g, ' ').trim().slice(0, 255) : '';
  autoCorrectDialog.value = { selected };
}

function closeAutoCorrect() {
  autoCorrectDialog.value = null;
  refocus();
}

function saveAutoCorrectSettings(settings) {
  autoCorrectSettings.value = settings;
  const kept = saveAutoCorrect(settings);
  autoCorrectDialog.value = null;
  refocus();
  const n = settings.entries.length;
  message.value = !settings.on ? t('AutoCorrect is off: text is typed as it is')
    : `${t(n === 1 ? 'AutoCorrect: 1 entry of yours' : 'AutoCorrect: {n} entries of yours', { n })}${kept ? '' : t(' – for this session only: the browser doesn’t keep them')}`;
}

/** The AutoCorrect replacement for the word just before the caret (`range`, collapsed), or null. */
function correctionAt(range, typed = ' ') {
  const node = range?.startContainer;
  if (!range?.collapsed || node?.nodeType !== 3 || inPlainFieldQuiet()) return null;
  if (node.parentElement.closest('a, .raw-inline, [data-raw], .field, [data-tc="del"]')) return null;
  const block = closestBlock(node);
  if (!block) return null;
  const lead = document.createRange();
  lead.setStart(block, 0);
  lead.setEnd(node, range.startOffset);
  const tag = node.parentElement.closest('[lang]')?.lang || doc.styles.language() || (lang.value === 'id' ? 'id' : 'en');
  const hit = autoCorrectWord(lead.toString(), autoCorrectSettings.value, /^en\b/i.test(tag), typed);
  // The word typed in one piece of text, as it is.
  if (!hit || range.startOffset < hit.from.length || node.data.slice(range.startOffset - hit.from.length, range.startOffset) !== hit.from) return null;
  // What goes after it: a quote typed is a curly one where AutoFormat makes them.
  const quote = autoFormat.value && smartQuote(lead.toString(), typed, tag);
  return { ...hit, after: quote || typed };
}

/** Replaces the word of `hit` (see correctionAt): `after` characters typed after it stay after the caret. */
function applyCorrection(hit, after = '') {
  const now = selectionInPage();
  const at = now?.startContainer;
  const typed = hit.from + after;
  // Found again where it is now: typing the character after it may have moved it to another node.
  if (!now?.collapsed || at?.nodeType !== 3 || at.data.slice(Math.max(0, now.startOffset - typed.length), now.startOffset) !== typed) return false;
  const word = document.createRange();
  word.setStart(at, now.startOffset - typed.length);
  word.setEnd(at, now.startOffset - after.length);
  selectRange(word);
  typeText(hit.to);
  // The caret goes back after what was typed after the word.
  const sel = window.getSelection();
  for (let i = 0; i < after.length; i++) sel.modify('move', 'forward', 'character');
  undoStack.commit();
  markDirty();
  return true;
}

/**
 * A space or punctuation ending a word with an entry: the character is typed first, as an undo step
 * of its own, then the word replaced, so Ctrl+Z straight after gives back what was typed, as in Word.
 */
function autoCorrectTyped(e) {
  if (!ENDS_WORD.test(e.data)) return false;
  const hit = correctionAt(selectionInPage(), e.data);
  if (!hit) return false;
  e.preventDefault();
  typeText(hit.after);
  undoStack.commit();
  applyCorrection(hit, hit.after);
  return true;
}

/** Enter after a word with an entry: the word is replaced (an undo step of its own), then the paragraph is split. */
function autoCorrectEnter(e) {
  const hit = correctionAt(selectionInPage(), '\n');
  if (!hit) return false;
  e.preventDefault();
  undoStack.flush();
  applyCorrection(hit);
  // Then Enter, as it would have been: tracked while tracking.
  const opts = trackedBy(selectionInPage()?.startContainer);
  if (opts) pendingSplit = opts;
  document.execCommand('insertParagraph'); // fires input
  return true;
}

// ------------------------------------------------------------------ copy for the web

/**
 * Copy for the web (Download's menu, or the right-click menu for the selection): the text with
 * headings, bold, italic, links, lists and tables only, as a website's editor (CMS) or an e-mail
 * takes it, with plain text beside it for fields that take only text; or as Markdown. See
 * editor/webCopy.js.
 */
function copyForWeb(kind, selection = false) {
  const range = selection ? (selectionInPage() || savedRange) : null;
  if (selection && (!range || range.collapsed)) {
    message.value = t('Select the text to copy first');
    return;
  }
  const { blocks, left } = webContent(page.value, { range, levelOf: (el) => doc.listOf(el)?.ilvl ?? Number(el.dataset.ilvl || 0) });
  if (!blocks.length) {
    message.value = t('Nothing to copy: the text is empty');
    return;
  }
  const formats = kind === 'markdown' ? { 'text/plain': toMarkdown(blocks) } : { 'text/html': toHtml(blocks), 'text/plain': toText(blocks) };
  if (!copyFormats(formats)) {
    message.value = t('The browser didn’t copy: try again from the menu');
    return;
  }
  const what = [left.pictures && t(left.pictures === 1 ? '1 picture' : '{n} pictures', { n: left.pictures }), left.textBoxes && t(left.textBoxes === 1 ? '1 text box' : '{n} text boxes', { n: left.textBoxes }),
    left.notes && t(left.notes === 1 ? '1 note' : '{n} notes', { n: left.notes }), left.equations && t(left.equations === 1 ? '1 equation' : '{n} equations', { n: left.equations })].filter(Boolean);
  const counted = t(blocks.length === 1 ? '1 paragraph' : '{n} paragraphs', { n: blocks.length });
  message.value = `${kind === 'markdown' ? t('Copied as Markdown: {what}', { what: counted }) : t('Copied for the web: {what}, with headings, bold, italic, links and lists – paste it into your site’s editor or an e-mail', { what: counted })}${what.length ? t(' · left out: {what}', { what: what.join(', ') }) : ''}`;
}

// ------------------------------------------------------------------ word count and length

const overTarget = computed(() => selectedWords.value == null && !!lengthTarget.words && words.value > lengthTarget.words);
const headlineTitle = computed(() => {
  const h = headline.value;
  if (!h || !lengthTarget.headline) return '';
  const d = lengthTarget.headline - h.chars;
  return `${t('Headline: {chars} characters, with spaces – “{text}”', { chars: h.chars, text: h.text })}. ${d >= 0 ? t('{n} to spare', { n: d }) : t('{n} over the limit', { n: -d })}`;
});
let longSentences = []; // the long sentences found when Word count last opened: [{ block, start, end, words }]
let longShown = -1; // the one shown last, to show the next each time

/** Word count, as Word's (a click on the status bar's count): counts, length target and readability. */
function openWordCount() {
  countWords();
  updatePages();
  const language = doc.styles.language() || '';
  const stats = readability(page.value, { english: !language || /^en\b/i.test(language) });
  longSentences = stats.long; // the page's elements stay here, out of Vue's reach
  wordCountData.value = { ...stats, long: stats.long.map(({ words: n }) => ({ words: n })) };
}

function closeWordCount() {
  wordCountData.value = null;
  refocus();
}

/** Sets (or clears, with null) the target in words and the headline's limit in characters. */
function setLengthTarget({ words: target, headline: limit }) {
  Object.assign(lengthTarget, { words: target, headline: limit });
  wordCountData.value = null;
  scheduleDraft();
  refocus();
  const parts = [target && t('a target of {n} words', { n: target }), limit && t('a headline limit of {n} characters', { n: limit })].filter(Boolean);
  message.value = parts.length ? t('Length: {what}, shown in the status bar', { what: parts.join(t(' and ')) }) : t('No length target');
}

/** Selects the next long sentence (from Word count), so it can be cut or split. */
function showLongSentence() {
  wordCountData.value = null;
  const live = longSentences.filter((s) => s.block.isConnected);
  if (!live.length) return;
  longShown = (longShown + 1) % live.length;
  const s = live[longShown];
  const range = textRange(blockText(s.block), s.start, s.end);
  page.value.focus();
  selectRange(range);
  s.block.scrollIntoView({ block: 'center' });
  message.value = t('Long sentence: {n} words ({i} of {count}) – Word count shows the next', { n: s.words, i: longShown + 1, count: live.length });
}

// ------------------------------------------------------------------ redaction

// Redact: the selected text replaced with █, one for each character, as a page is blacked out before
// it is released (a source's name, an address in a leaked document). Never tracked, even while
// tracking: a tracked deletion keeps the text in the file, for anyone to read in Word. A link in it
// goes, address and all (see doc.forgetLink). What was redacted is remembered in this session only
// (never in a draft or the file), so "Before you release" can say where it still is.
const REDACTION = '█';
const redactedTexts = new Set();
const DELETIONS = '[data-tc="del"], .raw-inline.deleted, [data-tracked]';

/** A tracked change in `range` (anyone's), whose earlier wording would stay in the file, or null. */
function trackedIn(range) {
  return [...page.value.querySelectorAll(DELETIONS)].find((el) => range.intersectsNode(el)) || null;
}

/** Read-only content with text inside `range` (a field, an equation, kept content), which can't be redacted here. */
function lockedTextIn(range) {
  if (readOnlyAt(range.startContainer, page.value)) return readOnlyAt(range.startContainer, page.value);
  if (readOnlyAt(range.endContainer, page.value)) return readOnlyAt(range.endContainer, page.value);
  return [...page.value.querySelectorAll('[contenteditable="false"]')].find((el) => range.intersectsNode(el)
    && !el.contains(range.commonAncestorContainer) && !el.matches('.note-ref, .comment-ref, .picture, .picture *')
    && el.textContent.trim()) || null;
}

/** Blacks out the editable text of `range` (see portions); returns how many characters. */
function redactRange(range) {
  const pieces = portions(range);
  const links = new Set();
  let n = 0;
  for (let i = pieces.length - 1; i >= 0; i--) {
    const { node, s, e } = pieces[i];
    node.data = node.data.slice(0, s) + node.data.slice(s, e).replace(/[^\t\n\r]/gu, REDACTION) + node.data.slice(e);
    n += e - s;
    const a = node.parentElement.closest('a[href]');
    if (a && page.value.contains(a)) links.add(a);
  }
  for (const a of links) {
    doc.forgetLink(a);
    a.replaceWith(...a.childNodes);
  }
  return n;
}

/** Redact (Review tab, right-click menu): blacks out the selected text for good. */
function redactSelection() {
  if (textLocked.value) return explainLock();
  restoreSelection();
  const range = selectionInPage() || savedRange;
  if (!range || range.collapsed || !page.value.contains(range.commonAncestorContainer)) {
    message.value = t('Select the text to redact first, such as a source’s name');
    page.value.focus();
    return;
  }
  const tracked = trackedIn(range);
  if (tracked) {
    message.value = t('The selection holds tracked changes: accept or reject them first (Review › Accept), so no earlier wording stays in the file');
    flash(tracked);
    return;
  }
  const locked = lockedTextIn(range);
  if (locked) return explainReadOnly(locked, true);
  const text = range.toString().replace(/\s+/g, ' ').trim();
  const pictures = [...page.value.querySelectorAll('.picture')].some((el) => range.intersectsNode(el));
  undoStack.flush();
  const end = [range.endContainer, range.endOffset];
  const n = redactRange(range);
  if (!n) return;
  page.value.normalize();
  undoStack.commit();
  markDirty();
  if (text.replaceAll(REDACTION, '').trim().length >= 2) redactedTexts.add(text);
  try { getSelection().collapse(...end); } catch { /* the end moved: the caret stays where it is */ }
  const left = redactedLeftovers().find((l) => l.text === text);
  const more = left ? left.redactable + left.tracked + left.locked : 0;
  message.value = [
    t('Redacted{tracked} – Undo (Ctrl+Z) brings it back', { tracked: tracking.value ? t(' (not tracked: a tracked deletion would keep the text in the file)') : '' }),
    pictures && t('pictures in it stay: delete them'),
    more && t('“{text}” is in the document {n} more times: Before you release (Review) redacts them all', { text: text.length > 40 ? `${text.slice(0, 40)}…` : text, n: more }),
  ].filter(Boolean).join(' · ');
}

/**
 * Where what was redacted this session still is: [{ text, redactable, tracked, locked, comments,
 * properties }], counts of the places in the text (and its headers, footers and notes) that can be
 * redacted here, in tracked changes, in read-only content; comments and properties that hold it.
 */
function redactedLeftovers() {
  const comments = threads.value.flatMap((th) => [th, ...th.replies]).map((c) => (c.paragraphs || []).join('\n'));
  const p = doc.properties || {};
  const values = [...(p.core || []), ...(p.app || []), ...(p.custom || [])].map((x) => String(x.value ?? ''));
  const out = [];
  for (const text of redactedTexts) {
    const where = { text, redactable: 0, tracked: 0, locked: 0, comments: 0, properties: 0 };
    for (const m of findMatches(page.value, text)) {
      const el = m.segment.nodes[0].node.parentElement;
      if (el.closest(DELETIONS)) where.tracked++;
      else if (m.editable && !readOnlyAt(el, page.value)) where.redactable++;
      else where.locked++;
    }
    const lower = text.toLowerCase();
    where.comments = comments.filter((c) => c.toLowerCase().includes(lower)).length;
    where.properties = values.filter((v) => v.toLowerCase().includes(lower)).length;
    if (where.redactable || where.tracked || where.locked || where.comments || where.properties) out.push(where);
  }
  return out;
}

/** Before you release › Redact them all: every place a redacted text is still in editable text. */
function redactLeftovers() {
  undoStack.flush();
  let n = 0;
  for (const text of redactedTexts) {
    const ranges = findMatches(page.value, text)
      .filter((m) => m.editable && !m.segment.nodes[0].node.parentElement.closest(DELETIONS) && !readOnlyAt(m.segment.nodes[0].node.parentElement, page.value))
      .map(matchRange);
    // From the last back, so that the earlier ones keep their offsets.
    for (const range of ranges.reverse()) n += redactRange(range) ? 1 : 0;
  }
  page.value.normalize();
  undoStack.commit();
  markDirty();
  countWords();
  message.value = t('Redacted {n} more places{tracked} – Undo (Ctrl+Z) brings them back', { n, tracked: tracking.value ? t(' (not tracked)') : '' });
}

/** Redacted passages in the page: runs of █ (also from an earlier session, or another tool). */
function redactedPassages() {
  return (page.value.textContent.match(/█+/g) || []).length;
}

// ------------------------------------------------------------------ before you release

/** Hidden text (w:vanish) in the page: passages, how many can be removed here, and the first one. */
function hiddenText() {
  const spans = [...page.value.querySelectorAll('span[data-hidden]')]
    .filter((sp) => !sp.parentElement.closest('span[data-hidden]') && sp.textContent.trim());
  const removable = spans.filter((sp) => !readOnlyAt(sp, page.value) && !sp.closest('[data-tracked], [data-raw]') && !sp.querySelector('[data-tracked], [data-raw]'));
  const passages = (list) => list.filter((sp) => !(sp.previousSibling?.nodeType === 1 && sp.previousSibling.matches('span[data-hidden]'))).length;
  return { spans, removable, count: passages(spans), removableCount: passages(removable), sample: spans[0]?.textContent.trim() || '' };
}

/** What the review's "Before you release" lists (see ReviewDialog), beyond doc.inspect(). */
function inspection() {
  const comments = threads.value.flatMap((t) => [t, ...t.replies]);
  const hidden = hiddenText();
  const found = {
    ...doc.inspect(),
    comments: { count: comments.length, authors: [...new Set(comments.map((c) => c.author).filter(Boolean))] },
    hidden: { count: hidden.count, removable: hidden.removableCount, sample: hidden.sample },
    redacted: { passages: redactedPassages(), leftovers: redactedLeftovers() },
    canChange: !textLocked.value,
  };
  return {
    ...found,
    // Removing personal information: chosen by the default (see choosePersonalDefault), which is
    // offered where it would lose nothing; and your name, to say when the author isn't you.
    removePersonalAuto: personalAuto && doc.removePersonal,
    personalDefault: personalDefault.value,
    personalDefaultOffered: namesSomeone() && nothingLost(found),
    you: named.value ? author.value : '',
  };
}

/** Removes every comment and its replies, as one undoable step (as Word's Delete All Comments). */
function deleteAllComments() {
  const ids = threads.value.map((t) => t.id);
  if (!ids.length) return;
  undoStack.flush();
  for (const id of ids) {
    const marks = page.value.querySelectorAll(`[data-comment-start="${CSS.escape(id)}"], [data-comment-end="${CSS.escape(id)}"], [data-comment-ref="${CSS.escape(id)}"]`);
    for (const el of marks) el.remove();
    if (!marks.length && doc.storyOf(id)) doc.deleteStoryComment(id); // in a header, footer or note
  }
  undoStack.commit();
  markDirty();
  refreshComments();
  message.value = `Removed ${count(ids.length, 'comment')} and their replies – Undo (Ctrl+Z) brings them back`;
}

/** Removes the hidden text that can be changed here, as one undoable step; not tracked, as in Word's Document Inspector. */
function removeHiddenText() {
  const { removable } = hiddenText();
  if (!removable.length) return;
  undoStack.flush();
  for (const sp of removable) sp.remove();
  page.value.normalize();
  undoStack.commit();
  markDirty();
  countWords();
  message.value = `Removed hidden text${tracking.value ? ' (not tracked)' : ''} – Undo (Ctrl+Z) brings it back`;
}

/** An action from "Before you release" in the review; the review then shows the document as it is now. */
async function releaseAction(what) {
  if (textLocked.value) return explainLock();
  if (what === 'comments') deleteAllComments();
  else if (what === 'changes') resolveAll(true);
  else if (what === 'hidden') removeHiddenText();
  else if (what === 'redact') redactLeftovers();
  await openReview();
}

/** Remove personal information from the downloaded file (or not): see doc.writePersonal. */
async function setRemovePersonal(on) {
  if (textLocked.value) return explainLock();
  personalChosen = true;
  personalAuto = false;
  doc.removePersonal = on;
  markDirty();
  await openReview('release'); // the part ticked in stays open, though nothing in it needs a look now
}

// Personal information in the file's properties often isn't the sender's: a CV or letter made from
// someone else's file or template carries their name as its author. Removing it can be made the
// default, by choice, in the review ("Do this for every file I change…", remembered in this
// browser): then it is chosen for each file where nothing would be lost by it: one you have changed
// (a file downloaded as it was opened stays as it was), whose properties name someone (not a
// placeholder such as "Author"), with no comments or tracked changes (whose names would become
// "Author"), no restriction, and its properties recording the edit. The review says so; unticked
// for a file, it stays unticked for it.
const PERSONAL_KEY = 'intact.removePersonal';
const personalDefault = ref((() => { try { return localStorage.getItem(PERSONAL_KEY) === '1'; } catch { return false; } })());
let personalChosen = false; // ticked or unticked for this file
let personalAuto = false; // ticked for this file by the default
const PLACEHOLDER = /^(author|unknown|user|admin(istrator)?)$/i;

/** Whether the file's properties name someone. */
function namesSomeone() {
  return [...(doc.properties.core || []), ...(doc.properties.app || [])]
    .some((x) => PERSONAL.has(x.label) && x.value && !PLACEHOLDER.test(x.value.trim()));
}
/** Whether removing personal information loses nothing here (no names on comments or changes). */
function nothingLost(found) {
  return !found.comments.count && !found.revisions && !textLocked.value;
}
function choosePersonalDefault(found) {
  if (personalChosen || doc.removePersonalSet) return;
  const auto = personalDefault.value && namesSomeone() && nothingLost(found) && dirty.value && doc.updateProperties;
  if (auto !== personalAuto) {
    personalAuto = auto;
    doc.removePersonal = auto;
  }
}
async function setPersonalDefault(on) {
  personalDefault.value = on;
  try { localStorage.setItem(PERSONAL_KEY, on ? '1' : '0'); } catch { /* this session only */ }
  if (on && !doc.removePersonal) return setRemovePersonal(true);
  await openReview('release');
}

function explainLock() {
  if (markup.value === 'original' && !['readOnly', 'comments', 'forms'].includes(lockMode.value)) {
    message.value = 'Showing the original, before the tracked changes: choose All markup (Review) to edit';
    return;
  }
  message.value = `${{ readOnly: 'Read-only', comments: 'Only comments can be added', forms: 'Only the fields can be filled in', trackedChanges: 'Only tracked changes' }[lockMode.value]}: ${lockMessage.value}`;
}

const LOCK_HOSTS = '[contenteditable="true"], [data-locked-host]';
/** Fields that stay editable while only filling in forms is allowed. */
const OPEN_FIELD = '.sdt:not(.sdt-locked), .sdt-row:not(.sdt-locked), .sdt-cell:not(.sdt-locked), .ff-text:not([data-ff-off])';
/** Legacy form fields (see docx/formFields.js) filled in by a click or a key: check boxes and lists. */
const FORM_CHOICE = '.ff-check, .ff-drop';

/**
 * Makes the page's editing hosts (the text, cells, notes, headers, footers, text boxes) read-only
 * while the text is locked, so neither the browser nor the editor can change them; in "filling in
 * forms" the fields stay editable. Locked hosts are marked, to be put back exactly.
 */
function applyLock(on = textLocked.value) {
  const root = page.value;
  if (!root) return;
  for (const el of [root, ...root.querySelectorAll(LOCK_HOSTS)]) {
    const lock = on && !(lockMode.value === 'forms' && el !== root && el.closest(OPEN_FIELD));
    if (lock && el.contentEditable === 'true') {
      el.contentEditable = 'false';
      el.dataset.lockedHost = '';
    } else if (!lock && 'lockedHost' in el.dataset) {
      el.contentEditable = 'true';
      delete el.dataset.lockedHost;
    }
  }
  // A legacy text form field is text in a paragraph: while only filling in forms is allowed, it is
  // made an editing host of its own, as Word lets only its fields be typed in.
  for (const el of root.querySelectorAll('.ff-text')) {
    const open = on && lockMode.value === 'forms' && !('ffOff' in el.dataset) && !el.parentElement.closest('[contenteditable="false"]:not([data-locked-host])');
    if (open && !('ffHost' in el.dataset)) {
      el.contentEditable = 'true';
      el.dataset.ffHost = '';
    } else if (!open && 'ffHost' in el.dataset) {
      el.removeAttribute('contenteditable');
      delete el.dataset.ffHost;
    }
  }
}

/**
 * Runs `read`, which reads the page into the document, with the page's hosts as they really are:
 * saving tells read-only content by contenteditable="false". `read` reads synchronously.
 */
function withHostsUnlocked(read) {
  if (!textLocked.value) return read();
  applyLock(false);
  try { return read(); } finally { applyLock(true); }
}

function editAnyway() {
  const what = [
    doc.protection && `Word restricts editing of this document (${PROTECTION_LABELS[doc.protection.mode]}). Editing it here goes around that restriction: do so only if you are allowed to. The restriction stays in the file, so Word still applies it.`,
    doc.signatures && `Changing the document removes its ${doc.signatures.count === 1 ? 'digital signature' : `${doc.signatures.count} digital signatures`} from the file you download, as editing a signed document in Word does: ${doc.signatures.count === 1 ? 'it' : 'they'} would no longer match it.`,
  ].filter(Boolean);
  if (what.length && !window.confirm(`${what.join('\n\n')}\n\nEdit anyway?`)) return;
  unlocked.value = true;
  applyLock(false);
  message.value = 'Editing anyway';
  page.value.focus();
}
const findBar = ref(null);
const dateInput = ref(null);
let chooserTarget = null;
let savedRange = null;
let styleEl = null;
let undoStack = null;

const BLOCK_SELECTOR = 'p,h1,h2,h3,h4,h5,h6,li';
const aligns = [
  { value: 'left', label: 'Align left', title: 'Align left (Ctrl+L)', icon: 'M2 3h12M2 6.5h8M2 10h12M2 13.5h8' },
  { value: 'center', label: 'Center', title: 'Center (Ctrl+E)', icon: 'M2 3h12M4 6.5h8M2 10h12M4 13.5h8' },
  { value: 'right', label: 'Align right', title: 'Align right (Ctrl+R)', icon: 'M2 3h12M6 6.5h8M2 10h12M6 13.5h8' },
  { value: 'justify', label: 'Justify', title: 'Justify (Ctrl+J)', icon: 'M2 3h12M2 6.5h12M2 10h12M2 13.5h12' },
];

// The sheet is the page: header (in the top margin), the editable body, footer (bottom margin). As in
// Word, the body starts at the top margin or below the header, whichever is lower. Headers and
// footers are in the page itself (see doc.renderBand), placed by CSS from these sizes.
const footerSpace = ref(0); // the footer's height, kept free below the text
/** Whether the page has a header (footer) band: one made here and taken away again by Undo hasn't. */
const hasBand = (which) => (page.value ? !!page.value.querySelector(`:scope > .docx-hf-band.${which}`) : !!(which === 'header' ? doc.headers : doc.footers).length);
const pageTick = ref(0); // the page's geometry (Page setup) or its bands (a footer added) changed
const sheetStyle = computed(() => (pageTick.value, { width: `${doc.page.width}px`, minHeight: `${doc.page.minHeight}px`, zoom: zoomLevel.value === 100 ? null : String(sheetZoom.value) }));
const pageStyle = computed(() => {
  void pageTick.value;
  const p = doc.page;
  return {
    padding: `${hasBand('header') ? 0 : p.top}px ${p.right}px ${hasBand('footer') ? Math.max(p.bottom, footerSpace.value) : p.bottom}px ${p.left}px`,
    '--hf-header': `${p.header}px`,
    '--hf-top': `${Math.max(0, p.top - p.header)}px`,
    '--hf-footer': `${p.footer}px`,
    '--hf-bottom': `${Math.max(0, p.bottom - p.footer)}px`,
    '--hf-left': `${p.left}px`,
    '--hf-right': `${p.right}px`,
  };
});

/** Keeps the footer's height free at the bottom of the page, as the text grows above it. */
function measureFooter() {
  const band = page.value?.querySelector(':scope > .docx-hf-band.footer');
  const h = band ? band.offsetHeight : 0;
  if (h !== footerSpace.value) footerSpace.value = h;
}

// Which header / footer is shown, when the document has several (first page, even pages, sections):
// a stylesheet hides the others, so the choice survives Undo replacing the page's content.
const hf = reactive({ header: 0, footer: 0 });
let variantStyle = null;
let printStyle = null;
function refreshPrintPage() {
  if (printStyle) printStyle.textContent = basicPrintCss(doc.page);
}

// Printing, and so saving as PDF: just before the browser lays the pages out, headers and footers
// are set to print on every page (see editor/printBands.js), and the page's title becomes the file's
// name, which browsers suggest as the PDF's name. Both go back afterwards.
let titleBeforePrint = null;
function onBeforePrint() {
  if (!page.value || !printStyle) return;
  printStyle.textContent = printCss(page.value, doc, sheetZoom.value);
  titleBeforePrint ??= document.title; // once, should the browser say so twice
  document.title = baseName.value || titleBeforePrint;
}
function onAfterPrint() {
  refreshPrintPage();
  if (titleBeforePrint != null) document.title = titleBeforePrint;
  titleBeforePrint = null;
}
onMounted(() => {
  window.addEventListener('beforeprint', onBeforePrint);
  window.addEventListener('afterprint', onAfterPrint);
});
onBeforeUnmount(() => {
  window.removeEventListener('beforeprint', onBeforePrint);
  window.removeEventListener('afterprint', onAfterPrint);
  if (titleBeforePrint != null) document.title = titleBeforePrint;
})
// The header (footer) tabs: the variants the document uses (a first page's only with a different
// first page…) whose host is in the page, with their index (their hosts' data-variant).
const variantTabs = computed(() => {
  void pageTick.value;
  const tabs = (which) => (which === 'header' ? doc.headers : doc.footers)
    .map((v, i) => ({ v, i }))
    .filter(({ v, i }) => doc.variantShown(v) && (!page.value || page.value.querySelector(`.docx-hf-part[data-which="${which}"][data-variant="${i}"]`)));
  return { header: tabs('header'), footer: tabs('footer') };
});

/** Shows a header and footer the document uses, if the one shown no longer is (or none was chosen). */
function keepVariantsShown() {
  for (const which of ['header', 'footer']) {
    const tabs = variantTabs.value[which];
    if (tabs.length && !tabs.some((t) => t.i === hf[which])) hf[which] = tabs[0].i;
  }
}

function refreshVariants() {
  if (!variantStyle) return;
  variantStyle.textContent = ['header', 'footer']
    .map((w) => `.docx-page .docx-hf-part[data-which="${w}"]:not([data-variant="${hf[w]}"]) { display: none; }`).join('\n');
  nextTick(measureFooter);
}
watch(() => ({ ...hf }), refreshVariants);

/** Shows another header or footer (first page, other pages, a section's) and puts the caret in it. */
function showVariant(which, i) {
  undoStack?.flush();
  hf[which] = i;
  const host = page.value.querySelector(`.docx-hf-part[data-which="${which}"][data-variant="${i}"]`);
  if (!host) return;
  nextTick(() => {
    host.focus({ preventScroll: true });
    placeCaret(host.querySelector(BLOCK_SELECTOR) || host, true);
  });
}

// Fonts the document uses that aren't on this device (nor embedded in it): named in the status bar.
const fontsMissing = ref([]);

function refreshCss() {
  styleEl.textContent = doc.css('.docx-sheet');
}

onMounted(() => {
  styleEl = document.createElement('style');
  document.head.append(styleEl);
  refreshCss();
  page.value.replaceChildren(doc.render());
  refreshSections();
  nextTick(keepVariantsShown);
  // The document's language, so the browser checks spelling (and hyphenates) in it where it can.
  const lang = doc.styles.language();
  if (lang && /^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$/i.test(lang)) page.value.closest('.docx-sheet').lang = lang;
  styleChoices.value = doc.styleChoices();
  loadChoices();
  variantStyle = document.createElement('style');
  document.head.append(variantStyle);
  // Printing: the document's own paper size and margins (see the print styles below).
  printStyle = document.createElement('style');
  document.head.append(printStyle);
  refreshPrintPage();
  refreshVariants();
  measureFooter();
  requestAnimationFrame(() => { if (page.value) fontsMissing.value = missingFonts(page.value); });
  threads.value = commentThreads();
  commentsOpen.value = threads.value.length > 0 && window.innerWidth >= 1100;
  paintComments();
  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, false);
  document.addEventListener('selectionchange', onSelectionChange);
  document.addEventListener('mousedown', onDocumentMousedown, true);
  window.addEventListener('keydown', onWindowKeydown, true);
  window.addEventListener('beforeunload', onBeforeUnload);
  countWords();
  refreshChanges();
  refreshHeadings();
  page.value.focus();
  placeCaret(bodyBlocks()[0] || page.value, true);
  doc.markPristine(page.value); // before locking: saving reads the page as it really is
  applyLock();
  if (lockMode.value === 'trackedChanges') tracking.value = true; // as Word does under this restriction
  // Comments in headers, footers and notes aren't in the page: their changes are undone with it.
  // Undo also covers what isn't in the page: comments in headers, footers and notes, and Page setup.
  undoStack = createHistory(page.value, {
    state: {
      get: () => JSON.stringify([doc.storyState(), doc.pageSetup, doc.hfFlags, doc.styles.styleState()]),
      set: (v) => {
        const [story, setup, flags, styles] = JSON.parse(v);
        doc.setStoryState(story);
        doc.setPageSetup(setup);
        doc.hfFlags = { ...flags };
        doc.relabelVariants(page.value);
        doc.styles.setStyleState(styles);
        refreshCss();
      },
    },
  });
  undoStack.commit();
  // A starter: its placeholders are what to do first, so Fill in opens with them.
  if (doc.starter) {
    countPlaceholders();
    if (placeholderCount.value && !textLocked.value) nextTick(openFillIn);
  }
});

// ------------------------------------------------------------------ formatting marks

// Shown or not is remembered in this browser, as Word remembers it.
const MARKS_KEY = 'intact.marks';
const showMarks = ref((() => { try { return localStorage.getItem(MARKS_KEY) === '1'; } catch { return false; } })());
const marksEl = ref(null);
let marksFrame = 0;
let marksObserver = null;

/** Redraws the formatting marks on the next frame (once, however often it is asked). */
function scheduleMarks() {
  if (!showMarks.value || marksFrame) return;
  marksFrame = requestAnimationFrame(() => {
    marksFrame = 0;
    if (showMarks.value && page.value && marksEl.value) drawMarks(page.value, marksEl.value, page.value.closest('.canvas').getBoundingClientRect(), sheetZoom.value);
  });
}

function toggleMarks() {
  showMarks.value = !showMarks.value;
  try { localStorage.setItem(MARKS_KEY, showMarks.value ? '1' : '0'); } catch { /* this session only */ }
  watchMarks();
}

/** Keeps the marks up to date while they are shown: the page changing, scrolling or resizing. */
function watchMarks() {
  const canvas = page.value?.closest('.canvas');
  if (showMarks.value) {
    marksObserver ||= new MutationObserver(scheduleMarks);
    marksObserver.observe(page.value, { childList: true, subtree: true, characterData: true, attributes: true });
    canvas?.addEventListener('scroll', scheduleMarks, { passive: true });
    window.addEventListener('resize', scheduleMarks);
    scheduleMarks();
  } else {
    marksObserver?.disconnect();
    canvas?.removeEventListener('scroll', scheduleMarks);
    window.removeEventListener('resize', scheduleMarks);
    marksEl.value?.replaceChildren();
  }
}

// ------------------------------------------------------------------ line numbers

// Sections with line numbering show them in the margin, redrawn as the text changes (see
// editor/lineNumbers.js); like the formatting marks, they are never part of the page.
const lineSections = doc.lineNumbering();
const hasLineNumbers = lineSections.some(Boolean);
const lineNumbersEl = ref(null);
let lineFrame = 0;
let lineObserver = null;
function scheduleLineNumbers() {
  if (lineFrame) return;
  lineFrame = requestAnimationFrame(() => {
    lineFrame = 0;
    if (page.value && lineNumbersEl.value) drawLineNumbers(page.value, lineNumbersEl.value, lineSections, sheetZoom.value);
  });
}
// ------------------------------------------------------------------ pages, about

// The sheet is one long page; Word decides where pages break when it opens the file. How many pages
// the document would print on, and about where each begins, is worked out from the paper size and
// margins (see editor/pages.js): counted in the status bar and drawn over the page as dashed lines,
// never in it. Shown or not is remembered in this browser. Not on a phone, where the page takes the
// phone's width and margins, so its lines aren't the paper's.
const PAGES_KEY = 'intact.pageBreaks';
const showPageBreaks = ref((() => { try { return localStorage.getItem(PAGES_KEY) !== '0'; } catch { return true; } })());
const pageBreaksEl = ref(null);
const pageCount = ref(0); // 0: not worked out (a phone's width)
let pagesTimer = 0;
let pagesObserver = null;
const narrow = typeof matchMedia === 'function' ? matchMedia('screen and (max-width: 860px)') : null;

/** The text's height on a page, in the page's own pixels: the paper less its margins, as Word fills it. */
function textHeight() {
  const p = doc.page;
  const z = sheetZoom.value;
  const shown = (which) => {
    const part = [...page.value.querySelectorAll(`:scope > .docx-hf-band.${which} .docx-hf-part`)].find((el) => el.getClientRects().length);
    return part ? part.getBoundingClientRect().height / z : 0;
  };
  // A header or footer taller than its margin pushes the text away from it, as in Word.
  const top = Math.max(p.top, hasBand('header') ? p.header + shown('header') : 0);
  const bottom = Math.max(p.bottom, hasBand('footer') ? p.footer + shown('footer') : 0);
  return p.minHeight - top - bottom;
}

/**
 * Where the pages break at 100%, as printed: zoomed, the browser lays the text out at that size,
 * which moves lines, so the page is copied (hidden, at 100%) and measured there, and each break is
 * put back on the page shown at the same place in the same paragraph or row.
 */
function pagesAt100() {
  const z = sheetZoom.value;
  if (z === 1) return estimatePages(page.value, textHeight(), 1);
  const sheet = page.value.closest('.docx-sheet');
  const ghost = sheet.cloneNode(false); // the sheet's classes, for the page's styles
  ghost.removeAttribute('style');
  ghost.setAttribute('aria-hidden', 'true');
  ghost.style.cssText = `position: absolute; left: -100000px; top: 0; visibility: hidden; pointer-events: none; width: ${doc.page.width}px; zoom: 1;`;
  const copy = page.value.cloneNode(true);
  copy.removeAttribute('contenteditable');
  ghost.append(copy);
  sheet.parentElement.append(ghost);
  try {
    const { count, breaks } = estimatePages(copy, textHeight(), 1);
    const inCopy = [...copy.querySelectorAll(UNIT_SELECTOR)];
    const shown = [...page.value.querySelectorAll(UNIT_SELECTOR)];
    return {
      count,
      breaks: breaks.map(({ y, block }) => {
        const target = shown[inCopy.indexOf(block)];
        if (!target) return null;
        const from = block.getBoundingClientRect();
        const to = target.getBoundingClientRect();
        const at = to.top + (from.height ? (y - from.top) / from.height : 0) * to.height;
        return { y: snapToLine(target, at), block: target };
      }).filter(Boolean),
    };
  } finally {
    ghost.remove();
  }
}

/** The top of the line of `block` (shown) nearest client y `at`: a break goes between lines. */
function snapToLine(block, at) {
  const range = document.createRange();
  range.selectNodeContents(block);
  let best = at;
  let gap = Infinity;
  for (const r of range.getClientRects()) {
    if (r.height && Math.abs(r.top - at) < gap) { gap = Math.abs(r.top - at); best = r.top; }
  }
  return best;
}

function updatePages() {
  pagesTimer = 0;
  if (!page.value) return;
  if (narrow?.matches) {
    pageCount.value = 0;
    pageBreaksEl.value?.replaceChildren();
    showPageCount();
    refreshWatermark();
    return;
  }
  const { count, breaks } = pagesAt100();
  refreshWatermark(breaks.map((b) => b.y));
  pageCount.value = count;
  showPageCount();
  if (pageBreaksEl.value) {
    if (showPageBreaks.value) drawPageBreaks(pageBreaksEl.value, breaks.map((b) => b.y), sheetZoom.value);
    else pageBreaksEl.value.replaceChildren();
  }
}
/**
 * The number of pages fields (NUMPAGES) show about how many pages there are, as the status bar
 * does; not worked out (a phone), the number in the file. Only data attributes change, which
 * nothing saves or watches.
 */
function showPageCount() {
  const show = pageCount.value ? String(pageCount.value) : '';
  for (const f of page.value.querySelectorAll('.field.page-field[data-instr^="NUMPAGES" i], .field.page-field[data-instr^="SECTIONPAGES" i]')) {
    if (f.dataset.show !== show) f.dataset.show = show;
    if (f.dataset.digits !== String(show.length)) f.dataset.digits = String(show.length);
  }
}
/** Works the pages out again a moment after the last change (typing goes on undisturbed). */
function schedulePages(delay = 250) {
  clearTimeout(pagesTimer);
  pagesTimer = setTimeout(updatePages, delay);
}
function togglePageBreaks() {
  showPageBreaks.value = !showPageBreaks.value;
  try { localStorage.setItem(PAGES_KEY, showPageBreaks.value ? '1' : '0'); } catch { /* this session only */ }
  nextTick(updatePages);
}
const pageCountTitle = computed(() => void pageTick.value || `About ${pageCount.value} ${pageCount.value === 1 ? 'page' : 'pages'} on ${paperName()}, worked out from the paper size and margins. Word lays out the pages itself when it opens the file, with its own fonts, so they may break a little differently. Click to ${showPageBreaks.value ? 'hide' : 'show'} where the pages break.`);
function paperName() {
  const p = doc.page;
  const [w, h] = [p.width, p.minHeight].map((v) => Math.round((v * 25.4) / 96));
  const sizes = { '210x297': 'A4', '216x279': 'Letter', '216x356': 'Legal', '148x210': 'A5', '297x420': 'A3', '184x267': 'Executive', '215x330': 'F4', '216x330': 'Folio' };
  const name = sizes[`${Math.min(w, h)}x${Math.max(w, h)}`];
  return name ? `${name} paper` : `${w} × ${h} mm paper`;
}
onMounted(() => {
  pagesObserver = new MutationObserver(() => schedulePages());
  pagesObserver.observe(page.value, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['style', 'class'] });
  window.addEventListener('resize', schedulePages);
  narrow?.addEventListener?.('change', schedulePages);
  document.fonts?.ready.then(() => schedulePages(0));
  schedulePages(0);
  // Declared further down (zoom), so watched once everything is.
  watch([sheetZoom, pageTick, () => ({ ...hf })], () => nextTick(() => schedulePages(0)));
});
onBeforeUnmount(() => {
  pagesObserver?.disconnect();
  clearTimeout(pagesTimer);
  window.removeEventListener('resize', schedulePages);
  narrow?.removeEventListener?.('change', schedulePages);
});

onMounted(() => {
  if (!hasLineNumbers) return;
  lineObserver = new MutationObserver(scheduleLineNumbers);
  lineObserver.observe(page.value, { childList: true, subtree: true, characterData: true });
  window.addEventListener('resize', scheduleLineNumbers);
  document.fonts?.ready.then(scheduleLineNumbers);
  scheduleLineNumbers();
});
onBeforeUnmount(() => {
  lineObserver?.disconnect();
  cancelAnimationFrame(lineFrame);
  window.removeEventListener('resize', scheduleLineNumbers);
});

onMounted(() => {
  if (showMarks.value) watchMarks();
  document.fonts?.ready.then(scheduleMarks); // web fonts change where the text is
});

// ------------------------------------------------------------------ tabs where Word puts them

// Each tab ends where Word ends it, with its leader (editor/tabLayout.js): laid out again when the
// text changes (the text before a right tab moves it), once a frame. Only the page's styles change,
// so watching the text (not attributes) can't loop.
let tabsFrame = 0;
let tabsObserver = null;
function scheduleTabs() {
  if (tabsFrame) return;
  tabsFrame = requestAnimationFrame(() => {
    tabsFrame = 0;
    if (page.value?.textContent.includes('\t')) doc.applyTabStops(page.value);
  });
}
onMounted(() => {
  if (!page.value) return;
  tabsObserver = new MutationObserver(scheduleTabs);
  tabsObserver.observe(page.value, { childList: true, subtree: true, characterData: true });
  scheduleTabs();
  document.fonts?.ready.then(scheduleTabs);
});
onBeforeUnmount(() => {
  tabsObserver?.disconnect();
  cancelAnimationFrame(tabsFrame);
});

onBeforeUnmount(() => {
  variantStyle?.remove();
  printStyle?.remove();
  marksObserver?.disconnect();
  cancelAnimationFrame(marksFrame);
  window.removeEventListener('resize', scheduleMarks);

  document.removeEventListener('selectionchange', onSelectionChange);
  document.removeEventListener('mousedown', onDocumentMousedown, true);
  window.removeEventListener('keydown', onWindowKeydown, true);
  window.removeEventListener('beforeunload', onBeforeUnload);
  undoStack?.dispose();
  styleEl?.remove();
  if (canHighlight) {
    CSS.highlights.delete('comment');
    CSS.highlights.delete('comment-active');
    CSS.highlights.delete('field-near');
  }
});

onBeforeRouteLeave(() => {
  if (dirty.value && !window.confirm('You have unsaved changes. Leave without downloading? They will be discarded.')) return false;
  discardDraft(); // leaving on purpose: nothing to bring back
  return true;
});

// ------------------------------------------------------------------ toolbar keyboard

const toolbarEl = ref(null);
let toolbarNav = null;
let toolbarLabels = null;

// ------------------------------------------------------------------ the ribbon

const RIBBON_TABS = [
  { id: 'file', label: 'File', title: 'File: download or save, print, PDF, share, mail merge, protect with a password, and inspect the document before it goes' },
  { id: 'home', label: 'Home', title: 'Home: styles, fonts, formatting, paragraphs and lists, Find and formatting marks' },
  { id: 'insert', label: 'Insert', title: 'Insert: a link, check box, date, reference, symbol, picture, signature or table, and filling in placeholders' },
  { id: 'layout', label: 'Layout', title: 'Layout: page setup (paper, margins, sections, headers and footers, watermark), page numbers and table layout' },
  { id: 'references', label: 'References', title: 'References: footnotes, endnotes and updating fields' },
  { id: 'review', label: 'Review', title: 'Review: Track Changes, accepting and rejecting, comments and checking before you send' },
  { id: 'view', label: 'View', title: 'View: the Navigation pane, zoom, formatting marks and how tracked changes show' },
];
const RIBBON_KEY = 'intact.ribbon';
const RIBBON_TAB_KEY = 'intact.ribbonTab';
const remembered = (key, fallback) => { try { return localStorage.getItem(key) ?? fallback; } catch { return fallback; } };
const ribbon = ref(remembered(RIBBON_KEY, '1') !== '0');
const ribbonTab = ref(RIBBON_TABS.some((t) => t.id === remembered(RIBBON_TAB_KEY, 'home')) ? remembered(RIBBON_TAB_KEY, 'home') : 'home');

function setRibbonTab(id) {
  ribbonTab.value = id;
  try { localStorage.setItem(RIBBON_TAB_KEY, id); } catch { /* this session only */ }
  nextTick(() => toolbarLabels?.refit());
}

/** The tools in tabs (the ribbon), or all of them in one toolbar; remembered in this browser. */
function setRibbon(on) {
  ribbon.value = on;
  try { localStorage.setItem(RIBBON_KEY, on ? '1' : '0'); } catch { /* this session only */ }
  nextTick(() => {
    toolbarLabels?.refit();
    document.querySelector(on ? '.ribbon-tabs [aria-selected="true"]' : '.toolbar .ribbon-mode')?.focus();
  });
}

/** The tabs from the keyboard, as WAI-ARIA's tabs: ← and → (Home, End) choose the next tab. */
function onRibbonKeys(e) {
  const at = RIBBON_TABS.findIndex((t) => t.id === ribbonTab.value);
  const to = { ArrowRight: at + 1, ArrowLeft: at - 1, Home: 0, End: RIBBON_TABS.length - 1 }[e.key];
  if (to == null || !e.target.matches('[role="tab"]')) return;
  e.preventDefault();
  const id = RIBBON_TABS[(to + RIBBON_TABS.length) % RIBBON_TABS.length].id;
  setRibbonTab(id);
  nextTick(() => document.getElementById(`ribbon-tab-${id}`)?.focus());
}
onMounted(() => {
  toolbarNav = rovingToolbar(toolbarEl.value, { onEscape: refocus });
  toolbarLabels = fitLabels(toolbarEl.value);
  document.fonts?.ready.then(() => toolbarLabels?.refit()); // labels' widths in the interface font
});
// Controls that come and go (Comments, a document's character styles) change what fits. (Watched
// once mounted: those are declared further down.)
onMounted(() => watch(() => [threads.value.length, characterStyles.value.length, placeholderCount.value > 0], () => nextTick(() => toolbarLabels?.refit())));
onMounted(detectShare);
onBeforeUnmount(() => { toolbarNav?.dispose(); toolbarLabels?.dispose(); });

// ------------------------------------------------------------------ zoom

// Zoom, as Word's status bar has it: the page shown bigger or smaller (CSS zoom on the sheet), the
// document unchanged. Remembered in this browser. What is drawn over the page (formatting marks,
// line numbers, a picture's frame) is placed in the page's own pixels, so it divides by the zoom.
const ZOOMS = [50, 67, 75, 80, 90, 100, 110, 125, 150, 175, 200];
const ZOOM_KEY = 'intact.zoom';
const zoomLevel = ref((() => {
  try { const z = Number(localStorage.getItem(ZOOM_KEY)); return ZOOMS.includes(z) ? z : 100; } catch { return 100; }
})());
const sheetZoom = computed(() => zoomLevel.value / 100);

function setZoom(z) {
  if (!ZOOMS.includes(z) || z === zoomLevel.value) return;
  zoomLevel.value = z;
  try { localStorage.setItem(ZOOM_KEY, String(z)); } catch { /* this session only */ }
  nextTick(() => {
    scheduleMarks();
    if (hasLineNumbers) scheduleLineNumbers();
    if (picture.el) updatePictureFrame();
    const range = selectionInPage();
    if (range) revealRect(range.getBoundingClientRect());
  });
}

/** The next zoom up (1) or down (-1). */
function zoomBy(dir) {
  const next = dir > 0 ? ZOOMS.find((z) => z > zoomLevel.value) : [...ZOOMS].reverse().find((z) => z < zoomLevel.value);
  if (next) setZoom(next);
}

// Ctrl+mouse wheel (and a trackpad's pinch, which browsers send as one) on the page zooms it, as in
// Word, rather than the whole app. A step at each notch of a wheel; a pinch's small steps add up.
let wheelDelta = 0;
function onCanvasWheel(e) {
  if (!e.ctrlKey && !e.metaKey) return;
  e.preventDefault();
  wheelDelta += e.deltaY;
  if (Math.abs(wheelDelta) < 40) return;
  zoomBy(wheelDelta < 0 ? 1 : -1);
  wheelDelta = 0;
}
onMounted(() => page.value.closest('.canvas').addEventListener('wheel', onCanvasWheel, { passive: false }));
onBeforeUnmount(() => page.value?.closest('.canvas')?.removeEventListener('wheel', onCanvasWheel));

// Wide pictures in line with the text fit a phone's screen (see editor/fitPictures.js).
let unfitPictures = null;
onMounted(() => { unfitPictures = fitPictures(page.value); });
onBeforeUnmount(() => unfitPictures?.());

// ------------------------------------------------------------------ draft

// While there are changes not downloaded, a copy of the document is kept in this browser (see
// drafts.js), a moment after each edit, so that a reload, a crash or a closed tab doesn't lose them;
// the start page offers it back. It goes when the document is downloaded or closed.
const DRAFT_DELAY = 1500;
const draftId = (doc.draftId ||= newDraftId());
const draftKept = ref(false);
const draftTitle = `Your unsaved changes are kept in this browser, on this device only, so a reload, a crash or a closed tab doesn’t lose them. The copy goes when you download or close the document, and after ${Math.round(MAX_AGE / 86400000)} days.`;
let draftTimer = 0;
let drafting = null; // the draft being written
let composing = false; // IME input: the page is left alone until it ends
let leaving = false;

function scheduleDraft(delay = DRAFT_DELAY) {
  clearTimeout(draftTimer);
  draftTimer = setTimeout(keepDraft, delay);
}

async function keepDraft() {
  draftTimer = 0;
  if (!dirty.value || leaving || !page.value) return;
  if (saving.value || reviewData.value || composing || drafting) { scheduleDraft(); return; }
  drafting = (async () => {
    const blob = await withHostsUnlocked(() => doc.draft(page.value));
    if (!dirty.value || leaving) return; // downloaded or closed meanwhile
    if (!blob) { // every change undone: nothing to keep
      draftKept.value = false;
      await deleteDraft(draftId);
      return;
    }
    draftKept.value = await saveDraft({
      id: draftId, name: `${baseName.value.trim() || 'document'}.${fileExt.value}`, blob,
      firstNewId: doc.firstNewId, unlocked: unlocked.value, length: { ...lengthTarget }, protected: !!doc.encryption,
    });
  })();
  try {
    await drafting;
  } catch (e) {
    console.warn('Draft not kept:', e);
  } finally {
    drafting = null;
  }
}

/** Waits for a draft being written, so that saving and the review never write at the same time. */
async function draftSettled() {
  while (drafting) await drafting.catch(() => {});
}

/** The draft is no longer needed: the document was downloaded, or is being left on purpose. */
function discardDraft({ closing = true } = {}) {
  clearTimeout(draftTimer);
  draftTimer = 0;
  if (closing) leaving = true;
  draftKept.value = false;
  draftSettled().then(() => deleteDraft(draftId));
}

/** Keeps the draft now rather than later when the tab is hidden: it may not come back. */
function onHidden() {
  if (document.visibilityState === 'hidden' && draftTimer) keepDraft();
}
onMounted(() => {
  document.addEventListener('visibilitychange', onHidden);
  window.addEventListener('pagehide', onHidden);
  page.value.addEventListener('compositionstart', () => { composing = true; });
  page.value.addEventListener('compositionend', () => { composing = false; });
  if (doc.restoredFrom) {
    const when = new Date(doc.restoredFrom.savedAt).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
    message.value = `Restored your unsaved changes from ${when}`;
    draftKept.value = true;
  }
});
onBeforeUnmount(() => {
  document.removeEventListener('visibilitychange', onHidden);
  window.removeEventListener('pagehide', onHidden);
  clearTimeout(draftTimer);
});

function onBeforeUnload(e) {
  if (dirty.value) e.preventDefault();
}

// ------------------------------------------------------------------ selection & toolbar state

function selectionInPage() {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  return page.value?.contains(range.commonAncestorContainer) ? range : null;
}

/**
 * The paragraph a selection starts in. Firefox starts a Select All at the page itself (offset 0)
 * rather than inside the first paragraph, so look down into the node the offset points at.
 */
function startBlock(range) {
  const { startContainer: c, startOffset: o } = range;
  const at = c.nodeType === 1 ? c.childNodes[o] || c : c;
  return closestBlock(at) || (at.nodeType === 1 ? at.querySelector(BLOCK_SELECTOR) : null);
}

function closestBlock(node) {
  const el = node.nodeType === 1 ? node : node.parentElement;
  const block = el?.closest(BLOCK_SELECTOR);
  return block && page.value.contains(block) ? block : null;
}

function onSelectionChange() {
  const range = selectionInPage();
  if (!range) return;
  if (caretOutOfMarker(range)) return; // the selection changes again, and this runs again
  countSelectedWords(range);
  updateNavCurrent();
  lastReadOnly = null;
  const pic = pictureInSelection(range);
  if (pic !== picture.el) {
    if (pic) showPicture(pic);
    else if (!picturePanel.open) picture.el = null;
  }
  // Clicking into placeholder text selects all of it, so typing replaces it (as in Word).
  const control = range.collapsed ? controlAt(range.startContainer, page.value) : null;
  if (control?.classList.contains('showing-placeholder') && control.querySelector(':scope > .sdt-content')) {
    selectCell(controlBox(control));
    return;
  }
  savedRange = range.cloneRange();
  markNearField(range);
  state.bold = document.queryCommandState('bold');
  state.italic = document.queryCommandState('italic');
  state.underline = document.queryCommandState('underline');
  state.strike = document.queryCommandState('strikeThrough');
  Object.assign(state, (({ superscript, subscript, highlight }) => ({ sup: superscript, sub: subscript, highlight }))(formatState(range)));
  state.hidden = !!(range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('[data-hidden]');
  const block = startBlock(range);
  state.block = block ? doc.blockValue(block.dataset.style) : 'normal';
  const align = block ? getComputedStyle(block).textAlign : 'left';
  // As shown: in a right-to-left paragraph, start is at the right.
  state.align = (block?.dir === 'rtl' ? { start: 'right', end: 'left' } : { start: 'left', end: 'right' })[align] || align;
  state.list = block ? doc.listKind(doc.listOf(block)) : null;
  state.table = !!tableSelection();
  const at = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  const cs = getComputedStyle(at);
  state.font = cs.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
  state.size = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
  const styled = at.closest('[data-rstyle]');
  state.charStyle = styled && block?.contains(styled) ? styled.dataset.rstyle : '';
}

/** Back to the page with the selection it had (after a menu, a dialog or the file picker). */
function refocus() {
  const range = savedRange;
  page.value.focus();
  if (range && page.value.contains(range.commonAncestorContainer)) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

function restoreSelection() {
  if (selectionInPage()) return;
  page.value.focus();
  if (savedRange) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }
}

function keepFocus(e) {
  // Keep the caret in the document when clicking toolbar buttons (but let <select> open).
  if (e.target.closest('button')) e.preventDefault();
}

function placeCaret(el, atStart = false) {
  const range = document.createRange();
  range.selectNodeContents(el);
  range.collapse(atStart);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

/**
 * Firefox puts the caret inside non-editable text at the edge of a line, where typing and deleting
 * do nothing: in an empty marker (a comment's mark, a hidden comment range or bookmark, a note's
 * number), or at the start or end of read-only text such as someone's tracked change ending a
 * paragraph (End puts it there). Puts it just outside instead, after an empty marker or on the side
 * of the text it is at: true if it moved. A caret inside such text, away from its edges, stays.
 */
function caretOutOfMarker(range) {
  if (!range.collapsed) return false;
  let el = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  let atom = null;
  for (; el && el !== page.value && !el.isContentEditable; el = el.parentElement) atom = el;
  if (!atom || !el?.isContentEditable || atom.tagName !== 'SPAN' || atom.querySelector('img, math')) return false;
  let before = false;
  if (atom.textContent) {
    const textOn = (side) => {
      const r = document.createRange();
      r.selectNodeContents(atom);
      if (side === 'before') r.setEnd(range.startContainer, range.startOffset);
      else r.setStart(range.startContainer, range.startOffset);
      return r.toString();
    };
    if (!textOn('after')) before = false;
    else if (!textOn('before')) before = true;
    else return false; // within the text: where the reader put it
  }
  const after = document.createRange();
  if (before) after.setStartBefore(atom);
  else after.setStartAfter(atom);
  after.collapse(true);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(after);
  return true;
}

function focusEnd() {
  page.value.focus();
  const blocks = bodyBlocks();
  placeCaret(blocks[blocks.length - 1] || page.value);
}

// ------------------------------------------------------------------ editing commands

const CHAR_FORMATS = new Set(['bold', 'italic', 'underline', 'strikeThrough']);

/** Plain-text fields (as in Word) take no character formatting. */
function inPlainField() {
  const range = selectionInPage();
  const box = range && controlAt(range.startContainer, page.value)?.querySelector(':scope > [data-plain]');
  if (box) message.value = 'Plain text fields can’t be formatted';
  return !!box;
}

function exec(cmd, value = null) {
  restoreSelection();
  if (CHAR_FORMATS.has(cmd) && inPlainField()) return;
  undoStack.flush();
  const after = withTrackedRuns(selectionInPage(), () => document.execCommand(cmd, false, value));
  if (after instanceof Range) selectRange(after);
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

let ownFormatting = false; // a formatting command of Intact's is running (its input events are ours)

/**
 * Runs `command`, a change to the formatting of the text in `range`. While tracking, the text whose
 * formatting it changed records its old formatting as a tracked change (see editor/runs.js), and
 * the range over that text is returned; otherwise what `command` returns.
 */
function withTrackedRuns(range, command) {
  const opts = range && !range.collapsed ? trackedBy(range.startContainer) : null;
  const noted = opts ? noteRuns(doc, range) : null;
  ownFormatting = true;
  let result;
  try {
    result = command();
  } finally {
    ownFormatting = false;
  }
  return noted ? trackRuns(doc, noted, opts) || result : result;
}

function selectRange(range) {
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

/**
 * Runs `change`, a change to the style, alignment or list of `blocks`. While tracking, each records
 * its old paragraph formatting as a tracked change. `after(block)`: the element a block became.
 */
function withTrackedParagraphs(blocks, change, after = (b) => b) {
  const opts = blocks.length ? trackedBy(blocks[0]) : null;
  const olds = opts ? blocks.map((b) => doc.blockPPr(b)) : null;
  const result = change();
  if (opts && result !== false) {
    blocks.forEach((b, i) => {
      const el = after(b);
      if (el?.isConnected && !inField(el)) doc.trackParagraph(el, olds[i], opts);
    });
  }
  return result;
}

/**
 * Aligns the selected paragraphs. Done here rather than with the browser's justify commands, which
 * Firefox carries out with an align attribute instead of the text-align style that is saved.
 */
function alignBlocks(value) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
  const blocks = selectedBlocks(range);
  withTrackedParagraphs(blocks, () => {
    for (const block of blocks) {
      // The buttons align as shown; a right-to-left paragraph stores it logically (see jcToCss).
      block.style.textAlign = block.dir === 'rtl' ? ({ left: 'end', right: 'start' }[value] || value) : value;
      block.removeAttribute('align');
    }
  });
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

function afterHistory(changed) {
  if (!changed) return;
  savedRange = null;
  pageChanged();
  refreshChanges();
  refreshComments(); // undo/redo replaced the DOM the comment ranges pointed into
  onSelectionChange();
  markDirty();
}
const undo = () => { closeChooser(); afterHistory(undoStack.undo()); };
const redo = () => { closeChooser(); afterHistory(undoStack.redo()); };

function selectedBlocks(range) {
  const all = [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((b) => range.intersectsNode(b) && !b.closest('.sdt-locked'));
  if (all.length) return all;
  const one = closestBlock(range.startContainer);
  return one ? [one] : [];
}

/**
 * Applies a paragraph style from the style menu to the selected paragraphs: a built-in one
 * ('normal', 'title', 'h1'–'h3', added to the file if missing) or one of the document's own
 * ('style:<id>').
 */
function applyBlockStyle(kind) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
  const own = kind.startsWith('style:') ? kind.slice(6) : null;
  const styleId = own || doc.styleIdFor(kind);
  refreshCss(); // a missing builtin style may just have been added
  const tag = own ? doc.tagFor(own) : kind === 'normal' ? 'p' : kind === 'title' ? 'h1' : kind;
  const { startContainer, startOffset, endContainer, endOffset } = range;
  const replaced = new Map();
  const blocks = selectedBlocks(range);
  withTrackedParagraphs(blocks, () => {
    for (const block of blocks) {
      let el = block;
      if (block.tagName.toLowerCase() !== tag) {
        el = document.createElement(tag);
        for (const attr of block.attributes) el.setAttribute(attr.name, attr.value);
        while (block.firstChild) el.append(block.firstChild);
        block.replaceWith(el);
        replaced.set(block, el);
      }
      if (styleId) el.dataset.style = styleId;
      else delete el.dataset.style;
    }
  }, (b) => replaced.get(b) || b);
  const r = document.createRange();
  r.setStart(replaced.get(startContainer) || startContainer, startOffset);
  r.setEnd(replaced.get(endContainer) || endContainer, endOffset);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

// ------------------------------------------------------------------ lists

/** Runs a change to list membership as one undoable step. */
function changeLists(change) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
  const blocks = selectedBlocks(range);
  if (withTrackedParagraphs(blocks, () => change(blocks)) === false) return;
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

/** The list next to `block` (before or after it), if it is of `kind`. */
function adjacentList(block, before, kind) {
  let el = before ? block.previousElementSibling : block.nextElementSibling;
  while (el?.classList.contains('raw-hidden')) el = before ? el.previousElementSibling : el.nextElementSibling;
  const list = el?.matches(BLOCK_SELECTOR) ? doc.listOf(el) : null;
  return list && doc.listKind(list) === kind ? list : null;
}

/**
 * Bullets / Numbering, as in Word: turns the selected paragraphs into a list of that kind (joining
 * a list of the same kind right before or after them), or takes them out of it if they already are.
 */
function toggleList(kind) {
  changeLists((blocks) => {
    if (!blocks.length) return false;
    const lists = blocks.map((b) => doc.listOf(b));
    if (lists.every((l) => doc.listKind(l) === kind)) {
      blocks.forEach((b) => doc.removeList(b));
      return true;
    }
    const join = adjacentList(blocks[0], true, kind) || adjacentList(blocks[blocks.length - 1], false, kind);
    const numId = join?.numId ?? doc.newList(kind);
    blocks.forEach((b, i) => {
      const ilvl = lists[i]?.ilvl ?? join?.ilvl ?? 0;
      doc.setList(b, numId, doc.hasListLevel(numId, ilvl) ? ilvl : 0);
    });
    return true;
  });
}

/** Moves the selected list items one level down (+1) or up (-1). */
function changeLevel(delta) {
  changeLists((blocks) => {
    const items = blocks.map((b) => [b, doc.listOf(b)]).filter(([, l]) => l && doc.hasListLevel(l.numId, l.ilvl + delta));
    if (!items.length) {
      if (delta > 0 && blocks.some((b) => doc.listOf(b))) message.value = 'This list has no further levels';
      return false;
    }
    items.forEach(([b, l]) => doc.setList(b, l.numId, l.ilvl + delta));
    return true;
  });
}

/** Enter on an empty list item moves it up a level, and out of the list from the top level. */
function endListItem(block) {
  const list = doc.listOf(block);
  changeLists(() => {
    if (list.ilvl > 0) doc.setList(block, list.numId, list.ilvl - 1);
    else doc.removeList(block);
  });
}

/** Tab / Shift+Tab: next or previous table cell, else list level, else a tab character. */
function onTab(e) {
  const range = selectionInPage();
  if (!range) return;
  const start = range.startContainer;
  const cell = (start.nodeType === 1 ? start : start.parentElement).closest('td');
  if (cell && page.value.contains(cell)) {
    const target = adjacentCell(cell, e.shiftKey ? -1 : 1);
    if (target) selectCell(target);
    else if (!e.shiftKey) addRow(cell); // Tab in the last cell, as in Word
    return;
  }
  const block = closestBlock(range.startContainer);
  const inList = range.collapsed
    ? block && doc.listOf(block) && isAtBlockStart(range, block)
    : selectedBlocks(range).some((b) => doc.listOf(b));
  if (inList) changeLevel(e.shiftKey ? -1 : 1);
  else if (!e.shiftKey) {
    const range = selectionInPage();
    if (range && !range.collapsed && refuseRemoving([range])) return;
    if (range && !range.collapsed) deleteSelection();
    document.execCommand('insertText', false, '\t');
  }
}

// ------------------------------------------------------------------ events

let wordTimer = 0;

/** Counts the words in the text, its notes and text boxes, as Word does: not in headers and footers. */
function countWords() {
  if (!page.value) return;
  countPlaceholders();
  const bands = [...page.value.querySelectorAll(':scope > .docx-hf-band')].map((b) => b.innerText || '');
  words.value = wordCount(page.value.innerText || '') - bands.reduce((n, t) => n + wordCount(t), 0);
  const count = (text) => [charCount(text, true), charCount(text, false)];
  const [all, allBare] = count(page.value.innerText || '');
  const [hf, hfBare] = bands.map((t) => count(t)).reduce((a, b) => [a[0] + b[0], a[1] + b[1]], [0, 0]);
  chars.withSpaces = all - hf;
  chars.withoutSpaces = allBare - hfBare;
  const head = headlineOf(page.value);
  const text = head && blockText(head).text.replace(/\s+/g, ' ').trim();
  headline.value = text ? { text, chars: charCount(text, true) } : null;
}

/** Characters, as Word counts them: line and paragraph breaks aren't; spaces are, unless `spaces` is false. */
function charCount(text, spaces) {
  return [...text.replace(spaces ? /[\n\r]/g : /\s/g, '')].length;
}

/** The words in the selection (null when nothing is selected), shown as "12 of 340 words". */
function countSelectedWords(range) {
  selectedWords.value = range && !range.collapsed ? wordCount(range.toString()) : null;
  selectedChars.value = range && !range.collapsed ? charCount(range.toString(), true) : null;
}

function markDirty() {
  measureFooter();
  dirty.value = true;
  scheduleDraft();
  message.value = '';
  clearTimeout(wordTimer);
  wordTimer = setTimeout(() => {
    countWords();
    refreshChanges();
    refreshHeadings();
  }, 300);
}

function onInput(e) {
  restoreCommentMarkers(markersBefore, page.value);
  markersBefore = [];
  if (pendingJoin) {
    joinLeftParagraphs(pendingJoin);
    pendingJoin = null;
  }
  if (pendingSplit) {
    markSplit(pendingSplit); // Enter while tracking: the new paragraph break is an insertion
    pendingSplit = null;
  }
  // Firefox, typing over whole paragraphs (as after Select All), can put the text outside them.
  if (!e?.isComposing) for (const n of [...page.value.childNodes]) if (n.nodeType === 3 && n.data) moveIntoParagraph(n);
  // Never leave the page (the text above the notes) empty: keep at least one paragraph to type into.
  let first = page.value.firstElementChild;
  if (first?.matches('.docx-hf-band.header')) first = first.nextElementSibling;
  if (!first || first.matches('.docx-notes, .docx-hf-band')) {
    const p = document.createElement('p');
    p.append(document.createElement('br'));
    page.value.insertBefore(p, first);
    placeCaret(p, true);
  }
  const range = selectionInPage();
  const control = range && controlAt(range.startContainer, page.value);
  if (control) clearPlaceholder(control, (id) => doc.isPlaceholderStyle(id));
  const cell = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('td.showing-placeholder');
  if (cell) {
    cell.classList.remove('showing-placeholder');
    for (const s of cell.querySelectorAll('[data-rstyle]')) if (doc.isPlaceholderStyle(s.dataset.rstyle)) delete s.dataset.rstyle;
  }
  doc.renumber(page.value);
  undoStack.schedule();
  markDirty();
  refreshComments();
  if (e?.inputType?.startsWith('format') && !ownFormatting) noteUntrackedFormatting();
  // Typing in a citation or bibliography from Zotero, EndNote or Mendeley: say it will be rewritten.
  const ref = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('.field[data-ref-manager]');
  if (ref) message.value = `${ref.dataset.refManager}: ${ref.dataset.refManager.split(' ')[0]} replaces this text when it refreshes in Word – change it there instead`;
}

/** Moves text typed directly into the page into the paragraph after it (or a new one), caret after it. */
function moveIntoParagraph(text) {
  let next = text.nextSibling;
  while (next && next.nodeType !== 1) next = next.nextSibling;
  let p = next?.matches(BLOCK_SELECTOR) ? next : null;
  if (p) {
    if (!p.textContent && p.lastChild?.tagName === 'BR') p.lastChild.remove();
    p.prepend(text);
  } else {
    p = document.createElement('p');
    text.before(p);
    p.append(text);
  }
  const range = document.createRange();
  range.setStart(text, text.length);
  range.collapse(true);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

function onBeforeInput(e) {
  // Undo/redo from the browser's Edit or context menu: use our history, not the native stack.
  if (e.inputType === 'historyUndo' || e.inputType === 'historyRedo') {
    e.preventDefault();
    if (e.inputType === 'historyUndo') undo();
    else redo();
    return;
  }
  if (replaceFormText(e)) return;
  if (!removesContent(e)) return;
  // Typing, deleting or pasting while the selection is in read-only content: say what it is.
  const locked = readOnlySelection();
  if (locked) {
    e.preventDefault();
    explainReadOnly(locked, true);
    return;
  }
  if (smartTyping(e)) return;
  if (trackedBeforeInput(e)) return;
  const ranges = rangesRemovedBy(e);
  // Backspace or Delete next to hidden markers (comment ranges, bookmarks, field codes) would remove
  // just the invisible marker: delete what is beyond it instead.
  const sel = window.getSelection();
  const step = { deleteContentBackward: true, deleteContentForward: false }[e.inputType];
  if (step !== undefined && sel.isCollapsed && ranges.length === 1 && onlyHiddenMarkers(ranges[0])) {
    e.preventDefault();
    const target = beyondHiddenMarkers(ranges[0], step);
    if (!target || refuseRemoving([target])) return;
    markersBefore = commentMarkers(page.value);
    sel.removeAllRanges();
    sel.addRange(target);
    document.execCommand('delete'); // fires input, so onInput puts the markers back
    return;
  }
  // A deletion across paragraphs must join them; see joinLeftParagraphs.
  const first = ranges[0] && closestBlock(ranges[0].startContainer);
  const last = ranges[0] && closestBlock(ranges[0].endContainer);
  pendingJoin = first && last && first !== last ? { first, last, breaks: new Set(first.querySelectorAll('br')) } : null;
  // Typing over, deleting, cutting or dragging away a tracked change would drop it silently.
  if (refuseRemoving(ranges)) e.preventDefault();
  // Comment range ends deleted with the text are put back in onInput, at the caret: where the text
  // was deleted. (After a drag the caret is at the drop instead; the cleanup on save covers that.)
  else if (e.inputType !== 'deleteByDrag') markersBefore = commentMarkers(page.value);
}

let markersBefore = [];
let pendingJoin = null;

/**
 * The browser won't merge into a paragraph left holding only a hidden marker (a comment range or
 * bookmark start), so deleting from its start into the next paragraph left an invisible empty
 * paragraph behind. Join what the deletion left of the two, as the browser's own merge would.
 */
function joinLeftParagraphs({ first, last, breaks }) {
  if (!first.isConnected) return;
  let next = first.nextElementSibling;
  while (next?.classList.contains('raw-hidden')) next = next.nextElementSibling;
  if (last.isConnected && next === last && first.parentElement === last.parentElement) {
    const placeholder = (el) => el.lastChild?.tagName === 'BR' && !el.textContent;
    if (placeholder(first)) first.lastChild.remove();
    const kids = [...last.childNodes].filter((n) => !(n.tagName === 'BR' && n === last.lastChild && !last.textContent));
    first.append(...kids);
    last.remove();
  }
  // A browser fills a paragraph it has emptied with a placeholder <br> (Firefox keeps it when the
  // next paragraph joins on): one that now has text after it would be saved as a line break.
  for (const br of first.querySelectorAll('br')) {
    if (!breaks.has(br) && hasContentAfter(br, first)) br.remove();
  }
}

/** True if anything visible comes after `node` in `block`. */
function hasContentAfter(node, block) {
  const range = document.createRange();
  range.setStartAfter(node);
  range.setEnd(block, block.childNodes.length);
  const frag = range.cloneContents();
  return !!frag.textContent || !!frag.querySelector('br, img, .raw-inline:not(.raw-hidden), .sdt');
}

// ------------------------------------------------------------------ table rows

/** A new row at the end of the table (tracked while tracking), with the caret in its first cell. */
function addRow(cell) {
  const table = cell.closest('table');
  undoStack.flush();
  const { row, reason } = doc.addTableRow(table, trackedBy(cell));
  if (!row) {
    message.value = reason;
    return;
  }
  undoStack.commit();
  selectCell(row.querySelector('td[contenteditable="true"]') || row.cells[0]);
  markDirty();
}

// ------------------------------------------------------------------ picture layout

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const PX = 9525; // EMU per pixel
const picture = reactive({ el: null, x: 0, y: 0, w: 0, h: 0, wrap: 'inline', cropping: false, shape: false });
const picturePanel = reactive({
  unit: 'cm', // lengths in centimetres or inches (see editor/units.js)
  open: false, x: 0, y: 0, wrap: 'inline', hAlign: 'left', hOffset: 0, vAlign: 'offset', vOffset: 0, width: 0, height: 0, lock: true, alt: '',
  cropL: 0, cropT: 0, cropR: 0, cropB: 0, rot: 0, flipH: false, flipV: false,
  border: 'none', borderColor: '#000000', shadow: false, glow: false, softEdge: false, reflection: false, colour: 'normal', brightness: 0, contrast: 0,
  style: '', shape: 'rect', bevel: '', camera: '', artistic: '', drawing: false,
  bevelW: 6, bevelH: 2, material: '', lightRig: 'threePt', lightAngle: 0, rotX: 0, rotY: 0, rotZ: 0, perspective: 0, amount: 50,
});
const BORDERS = { thin: 0.75, medium: 1.5, thick: 3 };
const picturePanelEl = ref(null);
const layoutOptionsButton = ref(null);
const round = (n, d = 2) => Math.round(n * 10 ** d) / 10 ** d;

/** The picture a selection is on: exactly one picture and nothing else (as a click selects it). */
function pictureInSelection(range) {
  if (!range || range.collapsed) return null;
  const frag = range.cloneContents();
  if (frag.textContent.trim() || frag.querySelectorAll('.picture').length !== 1) return null;
  return [...page.value.querySelectorAll('.picture')].find((p) => range.intersectsNode(p)) || null;
}

/** Selects a picture, as a click does in Word: the frame with its handles shows around it. */
function selectPicture(span) {
  const range = document.createRange();
  range.selectNode(span);
  page.value.focus({ preventScroll: true });
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  savedRange = range.cloneRange();
  showPicture(span);
}

function showPicture(span) {
  if (span && lockMode.value) {
    // Resizing or moving a picture isn't tracked: under "tracked changes only" it can't be done either.
    picture.el = null;
    return;
  }
  if (picture.el !== span) {
    picturePanel.open = false;
    picture.cropping = false;
  }
  picture.el = span;
  picture.shape = span?.dataset.shape != null;
  updatePictureFrame();
  if (span) message.value = picture.shape ? 'Shape: drag it to move it, a handle to resize it, or open Layout Options to turn it' : 'Picture: drag a corner to resize, or open Layout Options for wrapping and position';
}

/** Puts the frame over the selected picture (in the sheet's coordinates). */
function updatePictureFrame() {
  const span = picture.el;
  if (!span?.isConnected) {
    picture.el = null;
    picturePanel.open = false;
    return;
  }
  const sheet = page.value.closest('.docx-sheet').getBoundingClientRect();
  const r = span.getBoundingClientRect(); // the picture's frame (what shows of it, cropped)
  const z = sheetZoom.value; // the frame is drawn in the page's own pixels
  Object.assign(picture, { x: (r.left - sheet.left) / z, y: (r.top - sheet.top) / z, w: r.width / z, h: r.height / z, wrap: span.dataset.wrap || 'inline' });
}

/** Applies a change to the selected picture as one undoable step, keeping it selected. */
function changePicture(change, done) {
  const span = picture.el;
  if (!span?.isConnected) return;
  undoStack.flush();
  change(span);
  undoStack.commit();
  markDirty();
  nextTick(() => {
    if (span.isConnected) selectPicture(span);
    fillPicturePanel();
    // A line, arrow or shape is changed as a picture is, and called what it is.
    if (done) message.value = picture.shape ? done.replace(/^Picture\b/, 'Shape') : done;
  });
}

/** Drag a handle: corners keep the proportions, sides stretch, as in Word. */
function startResize(e, handle) {
  const span = picture.el;
  const info = span && doc.pictureLayout(span);
  if (!info) return;
  const look = doc.pictureLookOf(span);
  const start = { x: e.clientX, y: e.clientY, w: info.cx / PX, h: info.cy / PX };
  const zoom = pictureZoom(span) * sheetZoom.value; // shown smaller (or zoomed): the pointer moves in shown pixels
  const target = e.currentTarget;
  target.setPointerCapture(e.pointerId);
  let size = { w: start.w, h: start.h };
  const move = (ev) => {
    const dx = ((ev.clientX - start.x) / zoom) * (handle.includes('w') ? -1 : handle.includes('e') ? 1 : 0);
    const dy = ((ev.clientY - start.y) / zoom) * (handle.includes('n') ? -1 : handle.includes('s') ? 1 : 0);
    let w = start.w + dx;
    let h = start.h + dy;
    if (handle.length === 2) {
      const scale = Math.max(w / start.w, h / start.h);
      w = start.w * scale;
      h = start.h * scale;
    }
    size = { w: Math.max(8, w), h: Math.max(8, h) };
    doc.stylePicture(span, { ...info, cx: size.w * PX, cy: size.h * PX }, look); // a preview until it is let go
    updatePictureFrame();
  };
  const up = () => {
    target.removeEventListener('pointermove', move);
    target.removeEventListener('pointerup', up);
    target.removeEventListener('pointercancel', up);
    if (size.w === start.w && size.h === start.h) return;
    const unit = docUnit();
    changePicture((s) => doc.setPictureSize(s, size.w * PX, size.h * PX), `Picture resized to ${round(size.w * PX / EMU_PER[unit], 1)} × ${round(size.h * PX / EMU_PER[unit], 1)} ${unit}`);
  };
  target.addEventListener('pointermove', move);
  target.addEventListener('pointerup', up);
  target.addEventListener('pointercancel', up);
}

/**
 * A press on a picture selects it; dragging a selected floating picture moves it, its new place
 * measured from the left margin and from its paragraph (so it moves with the text, as by default
 * in Word).
 */
function onPictureMousedown(e, span) {
  e.preventDefault();
  const wasSelected = picture.el === span;
  selectPicture(span);
  if (!wasSelected || (span.dataset.wrap || 'inline') === 'inline' || e.button !== 0) return;
  const start = { x: e.clientX, y: e.clientY };
  let moved = false;
  const move = (ev) => {
    const dx = ev.clientX - start.x;
    const dy = ev.clientY - start.y;
    if (!moved && Math.hypot(dx, dy) < 4) return;
    moved = true;
    const z = sheetZoom.value; // moved in shown pixels, placed in the page's own
    span.style.transform = `translate(${dx / z}px, ${dy / z}px) ${span.dataset.turn || ''}`;
    updatePictureFrame();
  };
  const up = (ev) => {
    window.removeEventListener('pointermove', move);
    window.removeEventListener('pointerup', up);
    if (!moved) return;
    const r = span.getBoundingClientRect(); // where it was dropped
    span.style.transform = span.dataset.turn || '';
    const sheet = page.value.closest('.docx-sheet').getBoundingClientRect();
    const block = span.closest(BLOCK_SELECTOR) || span.parentElement;
    const z = sheetZoom.value;
    const left = (r.left - sheet.left) / z - doc.page.left;
    const top = (r.top - block.getBoundingClientRect().top) / z;
    changePicture((s) => doc.setPicturePosition(s, {
      h: { relativeFrom: 'margin', offset: Math.round(left) * PX },
      v: { relativeFrom: 'paragraph', offset: Math.round(top) * PX },
    }), 'Picture moved');
    void ev;
  };
  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', up);
}

/** Fills the Layout Options panel from the selected picture. */
function fillPicturePanel() {
  if (!picture.el) return;
  const info = doc.pictureLayout(picture.el);
  if (!info) return;
  const page = doc.page;
  const unit = docUnit();
  const PER_UNIT = EMU_PER[unit]; // EMU per unit shown: a centimetre or an inch
  const h = info.h || { relativeFrom: 'margin', align: 'left' };
  const v = info.v || { relativeFrom: 'paragraph', offset: 0 };
  let hAlign = 'word';
  let hOffset = 0;
  if (h.align && ['left', 'center', 'right'].includes(h.align) && ['margin', 'column'].includes(h.relativeFrom)) hAlign = h.align;
  else if (h.align == null && ['margin', 'column', 'page'].includes(h.relativeFrom)) {
    hAlign = 'offset';
    hOffset = round(((h.offset || 0) / PX - (h.relativeFrom === 'page' ? page.left : 0)) * PX / PER_UNIT);
  }
  let vAlign = 'word';
  if (v.align && ['top', 'center', 'bottom'].includes(v.align) && v.relativeFrom === 'margin') vAlign = v.align;
  else if (v.align == null && ['paragraph', 'line'].includes(v.relativeFrom)) vAlign = 'offset';
  const look = doc.pictureLookOf(picture.el);
  const border = look.border ? (look.border.width <= 1 ? 'thin' : look.border.width <= 2 ? 'medium' : 'thick') : 'none';
  Object.assign(picturePanel, {
    unit, wrap: info.wrap, hAlign, hOffset, vAlign, vOffset: vAlign === 'offset' ? round((v.offset || 0) / PER_UNIT) : 0,
    width: round(info.cx / PER_UNIT), height: round(info.cy / PER_UNIT), alt: info.descr,
    cropL: round(look.crop.l * 100, 1), cropT: round(look.crop.t * 100, 1), cropR: round(look.crop.r * 100, 1), cropB: round(look.crop.b * 100, 1),
    rot: round(look.rot, 1), flipH: look.flipH, flipV: look.flipV,
    border, borderColor: `#${look.border?.color || '000000'}`.toLowerCase(), shadow: !!look.shadow, glow: !!look.glow, softEdge: !!look.softEdge,
    reflection: !!look.reflection, colour: look.greyscale ? 'greyscale' : 'normal', brightness: round(look.brightness * 100), contrast: round(look.contrast * 100),
    style: '', shape: look.geometry, bevel: look.bevel?.prst || '', camera: look.camera || '', artistic: look.artistic?.effect || '',
    amount: look.artistic?.amount ?? 50,
    bevelW: round((look.bevel?.w ?? 76200) / 12700, 1), bevelH: round((look.bevel?.h ?? 25400) / 12700, 1),
    material: look.material || '', lightRig: look.light?.rig || 'threePt', lightAngle: round(look.light?.angle || 0),
    ...(() => {
      const [x, y, z, fov] = look.rotation3d ? [look.rotation3d.x, look.rotation3d.y, look.rotation3d.z, look.rotation3d.perspective] : CAMERA_ANGLES[look.camera] || [0, 0, 0, 0];
      return { rotX: round(x, 1), rotY: round(y, 1), rotZ: round(z, 1), perspective: round(fov, 1) };
    })(),
  });
}

function togglePicturePanel() {
  if (picturePanel.open) return closePicturePanel();
  fillPicturePanel();
  const r = layoutOptionsButton.value.getBoundingClientRect();
  Object.assign(picturePanel, { open: true, x: Math.max(8, Math.min(r.right + 6, window.innerWidth - 310)), y: Math.max(8, r.top) });
  nextTick(() => {
    const el = picturePanelEl.value;
    if (!el) return;
    picturePanel.y = Math.max(8, Math.min(picturePanel.y, window.innerHeight - el.offsetHeight - 8)); // all of it on screen
    el.querySelector('input:checked, input')?.focus();
  });
}

function closePicturePanel() {
  picturePanel.open = false;
  if (picture.el?.isConnected) selectPicture(picture.el);
  else page.value.focus();
}

const WRAP_NAMES = Object.fromEntries(WRAPS);

function applyWrap(wrap) {
  // Tight and Through: the text wraps to the picture's outline (what isn't transparent), as in Word.
  const outline = (wrap === 'tight' || wrap === 'through') && picture.el
    ? pictureOutline(picture.el.querySelector('img'), doc.pictureLookOf(picture.el).crop) : null;
  changePicture((s) => doc.setPictureWrap(s, wrap, outline), `Text wrapping: ${WRAP_NAMES[wrap]}${outline ? ', around its outline' : ''}`);
}

function applyPosition(which) {
  const p = picturePanel;
  const PER_UNIT = EMU_PER[p.unit];
  const h = which === 'h' && p.hAlign !== 'word'
    ? (p.hAlign === 'offset' ? { relativeFrom: 'margin', offset: (Number(p.hOffset) || 0) * PER_UNIT } : { relativeFrom: 'margin', align: p.hAlign })
    : null;
  const v = which === 'v' && p.vAlign !== 'word'
    ? (p.vAlign === 'offset' ? { relativeFrom: 'paragraph', offset: (Number(p.vOffset) || 0) * PER_UNIT } : { relativeFrom: 'margin', align: p.vAlign })
    : null;
  if (!h && !v) return;
  changePicture((s) => doc.setPicturePosition(s, { h, v }), 'Picture positioned');
}

function applySize(which) {
  const info = doc.pictureLayout(picture.el);
  if (!info) return;
  const PER_UNIT = EMU_PER[picturePanel.unit];
  let w = Math.max(0.1, Number(picturePanel.width) || 0.1) * PER_UNIT;
  let h = Math.max(0.1, Number(picturePanel.height) || 0.1) * PER_UNIT;
  if (picturePanel.lock) {
    if (which === 'width') h = (w * info.cy) / info.cx;
    else w = (h * info.cx) / info.cy;
  }
  changePicture((s) => doc.setPictureSize(s, w, h), `Picture resized to ${round(w / PER_UNIT, 1)} × ${round(h / PER_UNIT, 1)} ${picturePanel.unit}`);
}

/** Back to the picture's own size (at 96 pixels an inch), narrowed to fit between the margins. */
function resetPictureSize() {
  const img = picture.el?.querySelector('img');
  if (!img?.naturalWidth) return;
  const max = doc.page.width - doc.page.left - doc.page.right;
  const scale = Math.min(1, max / img.naturalWidth);
  changePicture((s) => doc.setPictureSize(s, img.naturalWidth * scale * PX, img.naturalHeight * scale * PX), 'Picture back to its own size');
}

/** Crop with the handles, as Word's Crop button: dragging a handle cuts (or uncuts) that side. */
function toggleCropping() {
  picture.cropping = !picture.cropping;
  message.value = picture.cropping ? 'Crop: drag the handles to cut the picture; Esc or Done cropping when finished' : 'Picture cropped';
}

function startCrop(e, handle) {
  const span = picture.el;
  const info = span && doc.pictureLayout(span);
  if (!info) return;
  const look = doc.pictureLookOf(span);
  const c0 = look.crop;
  const fullW = info.cx / PX / (1 - c0.l - c0.r);
  const fullH = info.cy / PX / (1 - c0.t - c0.b);
  const start = { x: e.clientX, y: e.clientY };
  const zoom = pictureZoom(span) * sheetZoom.value;
  const target = e.currentTarget;
  target.setPointerCapture(e.pointerId);
  let crop = c0;
  const move = (ev) => {
    const dx = (ev.clientX - start.x) / zoom / fullW;
    const dy = (ev.clientY - start.y) / zoom / fullH;
    crop = {
      l: handle.includes('w') ? c0.l + dx : c0.l,
      r: handle.includes('e') ? c0.r - dx : c0.r,
      t: handle.includes('n') ? c0.t + dy : c0.t,
      b: handle.includes('s') ? c0.b - dy : c0.b,
    };
    for (const side of ['l', 'r', 't', 'b']) crop[side] = Math.max(-0.5, Math.min(0.9, crop[side]));
    if (crop.l + crop.r > 0.95 || crop.t + crop.b > 0.95) return;
    const w = fullW * (1 - crop.l - crop.r);
    const h = fullH * (1 - crop.t - crop.b);
    doc.stylePicture(span, { ...info, cx: w * PX, cy: h * PX }, { ...look, crop });
    updatePictureFrame();
  };
  const up = () => {
    target.removeEventListener('pointermove', move);
    target.removeEventListener('pointerup', up);
    if (crop === c0) return;
    changePicture((s) => doc.setPictureCrop(s, crop), 'Picture cropped');
    nextTick(() => { picture.cropping = true; });
  };
  target.addEventListener('pointermove', move);
  target.addEventListener('pointerup', up);
}

/** Drag the round handle above a picture to rotate it (Shift: in steps of 15°), as in Word. */
function startRotate(e) {
  const span = picture.el;
  const info = span && doc.pictureLayout(span);
  if (!info) return;
  const look = doc.pictureLookOf(span);
  const r = span.getBoundingClientRect();
  const center = { x: r.left + r.width / 2, y: r.top + r.height / 2 };
  const target = e.currentTarget;
  target.setPointerCapture(e.pointerId);
  let rot = look.rot;
  const move = (ev) => {
    let deg = (Math.atan2(ev.clientY - center.y, ev.clientX - center.x) * 180) / Math.PI + 90;
    if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
    rot = ((Math.round(deg) % 360) + 360) % 360;
    doc.stylePicture(span, info, { ...look, rot });
    updatePictureFrame();
  };
  const up = () => {
    target.removeEventListener('pointermove', move);
    target.removeEventListener('pointerup', up);
    if (rot === look.rot) return;
    changePicture((s) => doc.setPictureRotation(s, { rot, flipH: look.flipH, flipV: look.flipV }), `Picture rotated to ${rot}°`);
  };
  target.addEventListener('pointermove', move);
  target.addEventListener('pointerup', up);
}

function applyCrop() {
  const p = picturePanel;
  const crop = { l: (Number(p.cropL) || 0) / 100, t: (Number(p.cropT) || 0) / 100, r: (Number(p.cropR) || 0) / 100, b: (Number(p.cropB) || 0) / 100 };
  changePicture((s) => doc.setPictureCrop(s, crop), 'Picture cropped');
}

function applyRotation() {
  const p = picturePanel;
  changePicture((s) => doc.setPictureRotation(s, { rot: Number(p.rot) || 0, flipH: p.flipH, flipV: p.flipV }), `Picture rotated to ${((Math.round(Number(p.rot) || 0) % 360) + 360) % 360}°`);
}

function rotateBy(deg) {
  picturePanel.rot = (((Number(picturePanel.rot) || 0) + deg) % 360 + 360) % 360;
  applyRotation();
}

function flip(axis) {
  picturePanel[axis] = !picturePanel[axis];
  const p = picturePanel;
  changePicture((s) => doc.setPictureRotation(s, { rot: Number(p.rot) || 0, flipH: p.flipH, flipV: p.flipV }), axis === 'flipH' ? 'Picture flipped horizontally' : 'Picture flipped vertically');
}

function applyEffects() {
  const p = picturePanel;
  const look = {
    border: p.border === 'none' ? null : { color: p.borderColor.replace('#', ''), width: BORDERS[p.border] },
    shadow: p.shadow, glow: p.glow ? 'FFC000' : null, softEdge: p.softEdge, reflection: p.reflection,
    greyscale: p.colour === 'greyscale', brightness: (Number(p.brightness) || 0) / 100, contrast: (Number(p.contrast) || 0) / 100,
  };
  changePicture((s) => doc.setPictureEffects(s, look), 'Picture effects changed');
}

/** A picture style from Word's gallery: its shape, border, effects and 3-D format replace the picture's. */
function applyStyle() {
  const name = picturePanel.style;
  if (!name) return;
  const style = name === 'none' ? null : PICTURE_STYLES.find((st) => st.name === name);
  changePicture((s) => doc.setPictureStyle(s, style), style ? `Picture style: ${style.name}` : 'Picture style taken off');
}

function applyShape() {
  const label = SHAPES.find(([v]) => v === picturePanel.shape)?.[1] || picturePanel.shape;
  changePicture((s) => doc.setPictureGeometry(s, picturePanel.shape), `Picture shape: ${label}`);
}

/**
 * The 3-D format and rotation from the panel, as Word's 3-D Format and 3-D Rotation set them. The
 * angles (X, Y, Z rotation, perspective) are written when changed (`how` 'angles') or already set;
 * a preset alone stays a preset.
 */
function apply3D(how) {
  const p = picturePanel;
  const look = doc.pictureLookOf(picture.el);
  const bevel = p.bevel ? { prst: p.bevel, w: Math.max(0, Number(p.bevelW) || 0) * 12700, h: Math.max(0, Number(p.bevelH) || 0) * 12700 } : null;
  const angles = how === 'angles' || look.rotation3d;
  const rotation = angles ? { x: Number(p.rotX) || 0, y: Number(p.rotY) || 0, z: Number(p.rotZ) || 0, perspective: Math.max(0, Number(p.perspective) || 0) } : null;
  changePicture((s) => doc.setPicture3D(s, {
    bevel, camera: p.camera || null, material: p.material || null, rotation,
    light: { rig: p.lightRig || 'threePt', dir: look.light?.dir || 't', angle: Number(p.lightAngle) || 0 },
  }), how === 'angles' ? '3-D rotation changed' : '3-D format changed');
}

/** A 3-D rotation preset: its angles replace any set by hand. */
function applyCameraPreset() {
  const p = picturePanel;
  const look = doc.pictureLookOf(picture.el);
  const bevel = p.bevel ? { prst: p.bevel, w: Math.max(0, Number(p.bevelW) || 0) * 12700, h: Math.max(0, Number(p.bevelH) || 0) * 12700 } : null;
  const label = CAMERAS.find(([v]) => v === p.camera)?.[1];
  changePicture((s) => doc.setPicture3D(s, {
    bevel, camera: p.camera || null, material: p.material || null,
    light: { rig: p.lightRig || 'threePt', dir: look.light?.dir || 't', angle: Number(p.lightAngle) || 0 },
  }), label ? `3-D rotation: ${label}` : '3-D rotation taken off');
}

/** An artistic effect: the picture is drawn with it here (the original kept), or taken off. */
async function applyArtistic() {
  const span = picture.el;
  const effect = picturePanel.artistic;
  if (!span) return;
  const label = ARTISTIC.find(([v]) => v === effect)?.[1];
  if (!effect) {
    changePicture((s) => doc.setPictureArtistic(s, null), 'Artistic effect taken off');
    return;
  }
  picturePanel.drawing = true;
  message.value = `Drawing ${label}…`;
  try {
    const amount = Number(picturePanel.amount) || 0;
    const png = await renderArtistic(doc.pictureOriginal(span).url, effect, amount);
    changePicture((s) => doc.setPictureArtistic(s, effect, png, amount), `Artistic effect: ${label}`);
  } catch (err) {
    message.value = err.message || 'The effect couldn’t be drawn';
  } finally {
    picturePanel.drawing = false;
  }
}

function applyAlt() {
  changePicture((s) => doc.setPictureDescription(s, picturePanel.alt.trim()), picturePanel.alt.trim() ? 'Alt text saved' : 'Alt text removed');
}

// ------------------------------------------------------------------ table layout

const tableMenu = reactive({ open: false, x: 0, y: 0, rows: 1, cols: 1, newCols: 1, multi: false });
const tableMenuEl = ref(null);
const layoutButton = ref(null);

const cellOf = (node) => (node?.nodeType === 1 ? node : node?.parentElement)?.closest('td') || null;

/**
 * The table cells selected (or the one with the caret): the table, the cell the selection starts
 * in, and the rows and grid columns covered ([first, last]). Null outside a table.
 */
function tableSelection() {
  const range = selectionInPage() || savedRange;
  // A read-only cell just clicked holds no caret: it is the cell meant (see clickReadOnly).
  const clicked = !selectionInPage() && lastReadOnly?.isConnected && lastReadOnly.closest('td');
  if (clicked && page.value.contains(clicked) && clicked.closest('table').dataset.raw != null) {
    return { table: clicked.closest('table'), cell: clicked, rows: [0, 0], cols: [0, 0], newCols: 1, multi: false };
  }
  const a = range && cellOf(range.startContainer);
  if (!a || !page.value.contains(a)) return null;
  const table = a.closest('table');
  if (table.dataset.raw == null) return null;
  let b = cellOf(range.endContainer);
  // A selection ending at the very start of a cell (as when Tab selects a whole cell) ends before it.
  const end = range.cloneRange();
  end.collapse(false);
  if (b && b !== a && isAtBlockStart(end, b) && !isEmptyBlock(b)) {
    const cells = [...table.querySelectorAll(':scope > tbody > tr > td, :scope > tr > td')];
    b = cells[cells.indexOf(b) - 1] || a;
  }
  if (!b || b.closest('table') !== table) b = a;
  const pa = doc.cellPosition(a);
  const pb = doc.cellPosition(b);
  if (!pa || !pb) return { table, cell: a, rows: [0, 0], cols: [0, 0], newCols: 1, multi: false };
  const cols = [Math.min(pa.start, pb.start), Math.max(pa.start + pa.span, pb.start + pb.span) - 1];
  return {
    table,
    cell: a,
    newCols: a === b ? 1 : cols[1] - cols[0] + 1, // columns to insert: one for a single (even spanning) cell
    multi: a !== b,
    rows: [Math.min(pa.row, pb.row), Math.max(pa.row, pb.row)],
    cols: [Math.min(pa.start, pb.start), Math.max(pa.start + pa.span, pb.start + pb.span) - 1],
  };
}

function toggleTableMenu() {
  if (tableMenu.open) return closeTableMenu();
  const sel = tableSelection();
  if (!sel) return;
  const rect = layoutButton.value.getBoundingClientRect();
  Object.assign(tableMenu, {
    open: true,
    x: Math.max(8, Math.min(rect.left, window.innerWidth - 240)),
    y: rect.bottom + 4,
    rows: sel.rows[1] - sel.rows[0] + 1,
    cols: sel.cols[1] - sel.cols[0] + 1,
    newCols: sel.newCols,
    multi: sel.multi,
  });
  nextTick(() => {
    keepInView(tableMenu, tableMenuEl.value);
    tableMenuEl.value?.querySelector('button')?.focus();
  });
}

function closeTableMenu() {
  tableMenu.open = false;
  page.value.focus();
  restoreSelection();
}

/** Up and Down (Home, End) move through a menu's items. */
function onMenuArrows(e) {
  if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
  e.preventDefault();
  const items = [...e.currentTarget.querySelectorAll('button:not(:disabled)')];
  const i = items.indexOf(document.activeElement);
  const next = { ArrowDown: (i + 1) % items.length, ArrowUp: (i - 1 + items.length) % items.length, Home: 0, End: items.length - 1 }[e.key];
  items[next].focus();
}

/** The rendered cell of `row` that covers grid column `col` (or its last cell). */
function cellAtColumn(row, col) {
  const cells = [...row.cells];
  return cells.find((td) => {
    const p = doc.cellPosition(td);
    return p && p.start <= col && col < p.start + p.span;
  }) || cells[cells.length - 1] || null;
}

const plural = (n, one, many = `${one}s`) => (n === 1 ? one : `${n} ${many}`);

/** Insert or delete rows, columns or the table, from the Table layout menu (as Word's Layout tab). */
function tableAction(kind) {
  tableMenu.open = false;
  const before = tableSelection(); // before refocusing, which may put back an older selection
  refocus();
  const sel = before?.cell.isConnected ? before : tableSelection();
  if (!sel) {
    message.value = 'Click in a table first';
    return;
  }
  const locked = readOnlyAt(sel.cell, page.value);
  if (locked) return explainReadOnly(locked, true);
  const { table, cell } = sel;
  const [r0, r1] = sel.rows;
  const [c0, c1] = sel.cols;
  const rowCount = r1 - r0 + 1;
  const colCount = c1 - c0 + 1;
  const tracked = trackedBy(cell);
  const neighbours = { before: table.previousElementSibling, after: table.nextElementSibling, parent: table.parentElement };
  const row = cell.parentElement;
  undoStack.flush();
  let res;
  let focus = null;
  let done = '';
  if (kind === 'rowsAbove' || kind === 'rowsBelow') {
    const below = kind === 'rowsBelow';
    res = doc.insertTableRows(table, below ? r1 : r0, rowCount, below, tracked);
    focus = res.rows && cellAtColumn(res.rows[0], c0);
    done = `${rowCount === 1 ? 'Row' : `${rowCount} rows`} inserted ${below ? 'below' : 'above'}`;
  } else if (kind === 'columnsLeft' || kind === 'columnsRight') {
    const right = kind === 'columnsRight';
    res = doc.insertTableColumns(table, c0, colCount, sel.newCols, right, tracked);
    if (!res.reason) focus = cellAtColumn(row, right ? c1 + 1 : c0);
    done = `${sel.newCols === 1 ? 'Column' : `${sel.newCols} columns`} inserted to the ${right ? 'right' : 'left'}${tracked ? ' (tracked)' : ''}`;
  } else if (kind === 'deleteRows') {
    res = doc.deleteTableRows(table, r0, r1, tracked);
    done = tracked ? `${plural(rowCount, 'Row', 'rows')} marked as deleted (tracked)` : `${plural(rowCount, 'Row', 'rows')} deleted – Undo (Ctrl+Z) brings ${rowCount === 1 ? 'it' : 'them'} back`;
    if (!res.reason && !res.removed) {
      const rows = doc.tableRows(table);
      focus = row.isConnected ? cellAtColumn(row, c0) : rows[Math.min(r0, rows.length - 1)] && cellAtColumn(rows[Math.min(r0, rows.length - 1)], c0);
    }
  } else if (kind === 'deleteColumns') {
    res = doc.deleteTableColumns(table, c0, c1, tracked);
    done = tracked && !res.removed && table.querySelector('td[data-rev-cell="del"]')
      ? `${plural(colCount, 'Column', 'columns')} marked as deleted (tracked)`
      : `${plural(colCount, 'Column', 'columns')} deleted – Undo (Ctrl+Z) brings ${colCount === 1 ? 'it' : 'them'} back`;
    if (!res.reason && !res.removed && row.isConnected) focus = cellAtColumn(row, Math.max(0, c0 - 1));
  } else if (kind === 'merge') {
    res = doc.mergeTableCells(table, sel.rows, sel.cols, tracked);
    if (res.cell) focus = res.cell;
    done = tracked ? 'Cells merged (tracked)' : 'Cells merged';
  } else if (kind === 'split') {
    const { cols, rows } = splitDialog;
    res = doc.splitTableCell(table, cell, cols, rows, tracked);
    focus = cell.isConnected ? cell : null;
    done = `Cell split into ${cols} × ${rows}${tracked ? ' (tracked)' : ''}`;
  } else if (kind === 'deleteTable') {
    res = doc.deleteTable(table, tracked);
    done = tracked ? 'Table marked as deleted (tracked)' : 'Table deleted – Undo (Ctrl+Z) brings it back';
  }
  if (res.reason) {
    message.value = res.reason;
    return;
  }
  if (res.removed) {
    // The caret goes where the table was: the paragraph after it (or before it).
    if (!page.value.firstElementChild) page.value.append(doc.emptyParagraph());
    const next = [neighbours.after, neighbours.before].find((n) => n?.isConnected);
    const target = next?.matches(BLOCK_SELECTOR) ? next : next?.querySelector(BLOCK_SELECTOR) || neighbours.parent.querySelector(BLOCK_SELECTOR) || page.value.querySelector(BLOCK_SELECTOR);
    if (target) placeCaret(target, next !== neighbours.before);
  } else if (focus) {
    placeCaret(focus.querySelector(BLOCK_SELECTOR) || focus, true);
  } else if (cell.isConnected) {
    placeCaret(cell.querySelector(BLOCK_SELECTOR) || cell, true);
  }
  doc.renumber(page.value);
  undoStack.commit();
  markDirty();
  refreshChanges();
  refreshComments();
  onSelectionChange();
  message.value = done;
}

// ------------------------------------------------------------------ cell shading and borders

// Word's Borders button's most used choices, for the selected cells as a block.
const BORDER_CHOICES = [
  { value: 'all', label: 'All borders', title: 'All borders: a line around every selected cell', icon: 'M2.5 2.5h11v11h-11zM8 2.5v11M2.5 8h11' },
  { value: 'outside', label: 'Outside borders', title: 'Outside borders: a line around the selected cells as a block', icon: 'M2.5 2.5h11v11h-11z' },
  { value: 'bottom', label: 'Bottom border', title: 'Bottom border: a line under the selected cells (a heading row, a total)', icon: 'M2.5 13.5h11' },
  { value: 'none', label: 'No border', title: 'No border: no lines around the selected cells, not even the table style’s', icon: '' },
];
const cellFormat = reactive({ open: false, x: 0, y: 0, sel: null });
const cellFormatEl = ref(null);

/** A colour mixed with white (`amount` 0–1 of white), as Word's lighter tints of the theme colours. */
function tint(hex, amount) {
  return [0, 2, 4].map((i) => Math.round(parseInt(hex.slice(i, i + 2), 16) + (255 - parseInt(hex.slice(i, i + 2), 16)) * amount).toString(16).padStart(2, '0')).join('').toUpperCase();
}

// Shading: the theme's colours lightened, as Word's palette offers them for a fill (text stays
// readable on them), then the theme's own; and the colours the document's text uses.
const shadingChoices = computed(() => {
  const theme = colorChoices.value.theme;
  return {
    theme: [
      ...theme.map((c) => ({ hex: tint(c.hex, 0.8), label: `${c.label}, lighter 80%` })),
      ...theme.map((c) => ({ hex: tint(c.hex, 0.4), label: `${c.label}, lighter 40%` })),
      ...theme,
    ],
    used: colorChoices.value.used,
  };
});

function openCellFormat() {
  const sel = tableSelection();
  tableMenu.open = false;
  if (!sel) {
    message.value = 'Click in a table first';
    return;
  }
  loadChoices();
  const rect = anchorRect(layoutButton.value);
  Object.assign(cellFormat, { open: true, sel, x: Math.max(8, rect.left - 60), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(cellFormat, cellFormatEl.value);
    cellFormatEl.value?.querySelector('button')?.focus();
  });
}

function closeCellFormat() {
  cellFormat.open = false;
  page.value.focus();
  restoreSelection();
}

/** Shading (`value` a colour, or null) or borders (`value` 'all', 'outside', 'bottom', 'none') for the selected cells. */
function formatCells(kind, value) {
  const sel = cellFormat.sel?.cell.isConnected ? cellFormat.sel : tableSelection();
  cellFormat.open = false;
  refocus();
  if (!sel) {
    message.value = 'Click in a table first';
    return;
  }
  const locked = readOnlyAt(sel.cell, page.value);
  if (locked) return explainReadOnly(locked, true);
  const tracked = trackedBy(sel.cell);
  // The selection's ends, by their text: the cells are put back in their rows, which would collapse a Range.
  const was = selectionInPage();
  const ends = was && [was.startContainer, was.startOffset, was.endContainer, was.endOffset];
  undoStack.flush();
  const res = doc.formatTableCells(sel.table, sel.rows, sel.cols, kind, value, tracked);
  if (res.reason) {
    message.value = res.reason;
    return;
  }
  // The selection stays as it was (its cells are kept), as Word keeps it, for the next choice.
  if (ends && ends[0].isConnected && ends[2].isConnected) {
    const range = document.createRange();
    range.setStart(ends[0], ends[1]);
    range.setEnd(ends[2], ends[3]);
    const s = window.getSelection();
    s.removeAllRanges();
    s.addRange(range);
  } else if (sel.cell.isConnected) placeCaret(sel.cell.querySelector(BLOCK_SELECTOR) || sel.cell, true);
  undoStack.commit();
  markDirty();
  refreshChanges();
  onSelectionChange();
  const cells = sel.multi ? 'the selected cells' : 'the cell';
  message.value = (kind === 'shading'
    ? (value ? `Shading #${value} on ${cells}` : `Shading taken off ${cells}`)
    : { all: `All borders on ${cells}`, outside: `Outside borders on ${cells}`, bottom: `Bottom border on ${cells}`, none: `No borders on ${cells}` }[value])
    + (tracked ? ' (tracked)' : '');
}

const splitDialog = reactive({ open: false, x: 0, y: 0, cols: 2, rows: 1 });
const splitDialogEl = ref(null);

/** Split cells…: asks how many columns and rows, as Word does. */
function openSplitDialog() {
  tableMenu.open = false;
  const rect = anchorRect(layoutButton.value);
  Object.assign(splitDialog, { open: true, cols: 2, rows: 1, x: Math.max(8, Math.min(rect.left, window.innerWidth - 240)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(splitDialog, splitDialogEl.value);
    splitDialogEl.value?.querySelector('input')?.select();
  });
}

function closeSplitDialog() {
  splitDialog.open = false;
  refocus();
}

function splitCells() {
  splitDialog.cols = Math.min(20, Math.max(1, Math.floor(Number(splitDialog.cols) || 1)));
  splitDialog.rows = Math.min(50, Math.max(1, Math.floor(Number(splitDialog.rows) || 1)));
  splitDialog.open = false;
  if (splitDialog.cols === 1 && splitDialog.rows === 1) return refocus();
  tableAction('split');
}

// ------------------------------------------------------------------ line and paragraph spacing

const LINE_SPACINGS = [[1, '1.0'], [1.15, '1.15'], [1.5, '1.5'], [2, '2.0'], [2.5, '2.5'], [3, '3.0']];
const LINE_PRESETS = { single: 240, '1.5': 360, double: 480 }; // the Paragraph dialog's, in 240ths of a line
const INDENT_STEP = 720; // Word's Increase Indent: to the next half inch
const SPACE_ADDED = 240; // Word's "Add space before/after paragraph": 12 pt
const spacingMenu = reactive({ open: false, x: 0, y: 0, line: null, before: 0, after: 0 });
const spacingMenuEl = ref(null);
const spacingButton = ref(null);

/** The paragraphs selected (or the one with the caret). */
function selectedParagraphs() {
  const range = selectionInPage() || savedRange;
  return range && page.value.contains(range.commonAncestorContainer) ? selectedBlocks(range) : [];
}

function toggleSpacingMenu() {
  if (spacingMenu.open) return closeSpacingMenu();
  const [block] = selectedParagraphs();
  if (!block) {
    message.value = 'Click in a paragraph first';
    return;
  }
  const f = doc.paragraphFormat(block);
  const rect = spacingButton.value.getBoundingClientRect();
  Object.assign(spacingMenu, {
    open: true,
    x: Math.max(8, Math.min(rect.left, window.innerWidth - 240)),
    y: rect.bottom + 4,
    line: f.lineRule === 'auto' ? f.line : null,
    before: f.before,
    after: f.after,
  });
  nextTick(() => {
    keepInView(spacingMenu, spacingMenuEl.value);
    spacingMenuEl.value?.querySelector('button')?.focus();
  });
}

function closeSpacingMenu() {
  spacingMenu.open = false;
  page.value.focus();
  restoreSelection();
}

/**
 * Changes the selected paragraphs' spacing or indentation, as one undoable step (tracked while
 * tracking): `change(format)` gives what to set for a paragraph with that format (see
 * doc.paragraphFormat), or null to leave it as it is.
 */
function formatParagraphs(change) {
  spacingMenu.open = false;
  if (textLocked.value) return explainLock(); // "filling in forms": fields only, not paragraphs
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  const blocks = selectedBlocks(range);
  if (!blocks.length) return;
  undoStack.flush();
  const changed = withTrackedParagraphs(blocks, () => {
    let any = false;
    for (const b of blocks) {
      const c = change(doc.paragraphFormat(b));
      if (!c || !Object.keys(c).length) continue;
      doc.setParagraphFormat(b, c);
      any = true;
    }
    return any;
  });
  if (!changed) return;
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

/** Line spacing in lines (1, 1.15, 1.5…), as Word's menu and Ctrl+1, Ctrl+5, Ctrl+2 set it. */
function setLineSpacing(lines) {
  const line = Math.round(lines * 240);
  formatParagraphs((f) => (f.lineRule === 'auto' && f.line === line ? null : { line, lineRule: 'auto' }));
}

/** Word's "Add / Remove space before (after) paragraph": 12 pt, or none (Ctrl+0 for before). */
function toggleSpace(which) {
  const [first] = selectedParagraphs();
  if (!first) return;
  const add = !(doc.paragraphFormat(first)[which] > 0);
  formatParagraphs((f) => (add === !(f[which] > 0) ? { [which]: add ? SPACE_ADDED : 0 } : null));
}

/**
 * Increase / Decrease Indent (Ctrl+M, Ctrl+Shift+M), as in Word: a list item changes its level;
 * another paragraph's left indent moves to the next (or previous) half inch, within the margins.
 */
function indentBlocks(delta) {
  restoreSelection();
  const range = selectionInPage();
  if (range && selectedBlocks(range).some((b) => doc.listOf(b))) return changeLevel(delta);
  const widest = Math.max(0, (doc.page.width - doc.page.left - doc.page.right) * 15 - INDENT_STEP); // px to twips
  formatParagraphs((f) => {
    const steps = delta > 0 ? Math.floor(f.left / INDENT_STEP) + 1 : Math.ceil(f.left / INDENT_STEP) - 1;
    const left = Math.min(widest, Math.max(0, steps * INDENT_STEP));
    return left === f.left ? null : { left };
  });
}

// The Paragraph dialog: values as shown (cm, pt, lines), and as they were when it opened, so only
// what is changed in it is written (fields left alone keep coming from the paragraph's style).
const paraDialog = reactive({
  open: false, x: 0, y: 0, unit: 'cm', left: 0, right: 0, special: 'none', by: 0, before: 0, after: 0, rule: 'single', at: 1,
  tabs: [], tabPos: '', tabAlign: 'left', tabLeader: 'none', tabError: '',
});
let paraInitialTabs = '[]';
// Word's Tabs dialog: alignments and leaders, as w:tab writes them.
const TAB_ALIGNS = { left: 'Left', center: 'Centre', right: 'Right', decimal: 'Decimal', bar: 'Bar' };
const TAB_LEADERS = { none: { label: 'None' }, dot: { label: '……… dots' }, hyphen: { label: '-------- dashes' }, underscore: { label: '_____ line' } };
const MAX_TAB = 31680; // Word's furthest tab stop: 22 inches
const formatLength = (twips) => String(fromTwips(twips, paraDialog.unit));

/** The unit lengths are shown in, for this document (see editor/units.js). */
function docUnit() {
  const now = doc.pageSetupNow();
  return unitFor(now.w, now.h);
}

/** Sets a tab stop at the position typed (one already there is replaced), as Word's Tabs dialog's Set. */
function addTabStop() {
  paraDialog.tabError = '';
  const at = Number(paraDialog.tabPos);
  if (paraDialog.tabPos === '' || !Number.isFinite(at) || at < 0) {
    paraDialog.tabError = `Type where the tab stop goes, in ${UNIT_NAMES[paraDialog.unit]} from the left margin`;
    return;
  }
  const pos = toTwips(at, paraDialog.unit);
  if (pos > MAX_TAB) {
    paraDialog.tabError = `A tab stop can be at most ${formatLength(MAX_TAB)} ${paraDialog.unit} from the margin`;
    return;
  }
  const stop = { pos, align: paraDialog.tabAlign, leader: paraDialog.tabLeader };
  paraDialog.tabs = [...paraDialog.tabs.filter((t) => Math.abs(t.pos - pos) > 5), stop].sort((a, b) => a.pos - b.pos);
  paraDialog.tabPos = '';
}

function removeTabStop(pos) {
  paraDialog.tabs = paraDialog.tabs.filter((t) => t.pos !== pos);
}
const paraDialogEl = ref(null);
let paraInitial = null;
let paraTwips = null; // the indents as they were, exactly, for switching units
const toPt = (twips) => Math.round(twips / 2) / 10;

function dialogValues(f, unit) {
  const len = (twips) => fromTwips(twips, unit);
  let rule = 'multiple';
  let at = Math.round((f.line / 240) * 100) / 100;
  if (f.lineRule === 'exact' || f.lineRule === 'atLeast') {
    rule = f.lineRule;
    at = toPt(f.line);
  } else {
    rule = Object.keys(LINE_PRESETS).find((k) => LINE_PRESETS[k] === f.line) || 'multiple';
  }
  // As Word's dialog: with a hanging indent, Left is where the first line starts (w:left less the
  // hanging), so "Left 0, Hanging 0.5 in" is the usual hanging indent, w:left="720" w:hanging="720".
  return {
    left: len(f.left - f.hanging), right: len(f.right),
    special: f.hanging ? 'hanging' : f.firstLine ? 'firstLine' : 'none', by: len(f.hanging || f.firstLine),
    before: toPt(f.before), after: toPt(f.after), rule, at,
  };
}

function openParagraphDialog() {
  spacingMenu.open = false;
  const [block] = selectedParagraphs();
  if (!block) return;
  const unit = docUnit();
  const format = doc.paragraphFormat(block);
  paraInitial = dialogValues(format, unit);
  paraTwips = { left: format.left - format.hanging, right: format.right, by: format.hanging || format.firstLine };
  const { stops } = doc.tabStops(block);
  paraInitialTabs = JSON.stringify(stops);
  const rect = anchorRect(spacingButton.value);
  Object.assign(paraDialog, paraInitial, {
    unit, tabs: stops, tabPos: '', tabAlign: 'left', tabLeader: 'none', tabError: '',
    open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)), y: rect.bottom + 4,
  });
  nextTick(() => {
    keepInView(paraDialog, paraDialogEl.value);
    paraDialogEl.value?.querySelector('input')?.select();
  });
}

/** The Paragraph dialog's lengths in another unit (remembered for the next dialogs). */
function setParaUnit(unit) {
  saveUnit(unit);
  switchUnit(paraDialog, paraInitial, paraTwips, ['left', 'right', 'by', 'tabPos'], unit);
}

function closeParagraphDialog() {
  paraDialog.open = false;
  refocus();
}

/** Choosing "Multiple", "At least" or "Exactly" offers a value in its unit (lines or points). */
function onLineRuleChange() {
  const { rule } = paraDialog;
  const points = (r) => r === 'atLeast' || r === 'exact';
  if (rule === 'multiple') paraDialog.at = paraInitial.rule === 'multiple' ? paraInitial.at : 1.15;
  else if (points(rule)) paraDialog.at = points(paraInitial.rule) ? paraInitial.at : 12;
}

function applyParagraphDialog() {
  const d = { ...paraDialog };
  const i = paraInitial;
  paraDialog.open = false;
  const num = (v) => (Number.isFinite(Number(v)) && v !== '' ? Number(v) : 0);
  const change = {};
  const leftChanged = num(d.left) !== i.left;
  const specialChanged = d.special !== i.special || (d.special !== 'none' && num(d.by) !== i.by);
  // By as it was when it wasn't changed (exactly: what is shown is rounded).
  const by = num(d.by) === i.by && d.special === i.special ? paraTwips.by : Math.max(0, toTwips(num(d.by), d.unit));
  // With a hanging indent, Left is where the first line starts: w:left is Left and the hanging
  // together, so it is written again when either changes (from its exact value when Left wasn't).
  const hangs = (special) => special === 'hanging';
  if (leftChanged || (specialChanged && (hangs(d.special) || hangs(i.special)))) {
    const left = leftChanged ? toTwips(num(d.left), d.unit) : paraTwips.left;
    change.left = left + (hangs(d.special) ? by : 0);
  }
  if (num(d.right) !== i.right) change.right = toTwips(num(d.right), d.unit);
  if (specialChanged) {
    if (d.special === 'hanging') change.hanging = by;
    else change.firstLine = d.special === 'firstLine' ? by : 0;
  }
  if (num(d.before) !== i.before) change.before = Math.max(0, Math.round(num(d.before) * 20));
  if (num(d.after) !== i.after) change.after = Math.max(0, Math.round(num(d.after) * 20));
  if (d.rule !== i.rule || (!LINE_PRESETS[d.rule] && num(d.at) !== i.at)) {
    if (LINE_PRESETS[d.rule]) Object.assign(change, { line: LINE_PRESETS[d.rule], lineRule: 'auto' });
    else if (d.rule === 'multiple') Object.assign(change, { line: Math.max(1, Math.round(num(d.at) * 240)), lineRule: 'auto' });
    else Object.assign(change, { line: Math.max(20, Math.round(num(d.at) * 20)), lineRule: d.rule });
  }
  // A tab position typed but not Set yet is set, as Word's OK does.
  if (d.tabPos !== '' && Number.isFinite(Number(d.tabPos))) {
    paraDialog.tabs = d.tabs;
    addTabStop();
    d.tabs = paraDialog.tabs;
  }
  if (JSON.stringify(d.tabs) !== paraInitialTabs) change.tabs = d.tabs.map((t) => ({ ...t }));
  refocus();
  if (Object.keys(change).length) formatParagraphs(() => change);
}

// ------------------------------------------------------------------ inserting tables

const PICKER_ROWS = 8;
const PICKER_COLS = 10;
const MAX_COLS = 63; // as many as Word allows
const MAX_ROWS = 500;
/**
 * Keeps a popup that opens below the toolbar wholly inside the window: once it has rendered, its
 * own size (not a guess at it) moves it left or up as far as it must, 8px from the edges.
 */
function keepInView(state, el) {
  if (!el) return;
  const { width, height } = el.getBoundingClientRect();
  state.x = Math.max(8, Math.min(state.x, window.innerWidth - width - 8));
  state.y = Math.max(8, Math.min(state.y, window.innerHeight - height - 8));
}

const tablePicker = reactive({ open: false, x: 0, y: 0, rows: 1, cols: 1, customRows: 2, customCols: 3 });
const tablePickerEl = ref(null);
const tableButton = ref(null);

function toggleTablePicker() {
  if (tablePicker.open) return closeTablePicker();
  const rect = tableButton.value.getBoundingClientRect();
  Object.assign(tablePicker, { open: true, rows: 1, cols: 1, x: Math.max(8, Math.min(rect.left - 80, window.innerWidth - 250)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(tablePicker, tablePickerEl.value);
    tablePickerEl.value?.querySelector('.cell')?.focus();
  });
}

function closeTablePicker() {
  tablePicker.open = false;
  page.value.focus();
  restoreSelection();
}

/** Arrow keys grow or shrink the table in the picker, as in Word; Enter inserts it. */
function onPickerKeydown(e) {
  const step = { ArrowRight: [0, 1], ArrowLeft: [0, -1], ArrowDown: [1, 0], ArrowUp: [-1, 0] }[e.key];
  if (!step) return;
  e.preventDefault();
  tablePicker.rows = Math.min(PICKER_ROWS, Math.max(1, tablePicker.rows + step[0]));
  tablePicker.cols = Math.min(PICKER_COLS, Math.max(1, tablePicker.cols + step[1]));
  nextTick(() => tablePickerEl.value?.querySelectorAll('.cell')[(tablePicker.rows - 1) * PICKER_COLS + tablePicker.cols - 1]?.focus());
}

/**
 * Inserts an empty table of `rows` × `cols` at the caret, as Word does: before the paragraph if the
 * caret is at its start (or it is empty), otherwise the paragraph is split there and the table goes
 * between. While tracking, its rows are tracked insertions. The caret goes into its first cell.
 */
function insertTable(rows, cols) {
  tablePicker.open = false;
  refocus();
  rows = Math.min(MAX_ROWS, Math.max(1, Math.floor(Number(rows) || 1)));
  cols = Math.min(MAX_COLS, Math.max(1, Math.floor(Number(cols) || 1)));
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text where the table should go';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (controlAt(range.startContainer, page.value)) {
    message.value = 'A table can’t go in a field';
    return;
  }
  if (refuseInNote(range, 'A table')) return;
  const tracked = trackedBy(range.startContainer);
  undoStack.flush();
  range.collapse(true);
  let block = closestBlock(range.startContainer);
  if (!block) {
    block = doc.emptyParagraph();
    page.value.append(block);
  } else if (!isEmptyBlock(block) && !isAtBlockStart(range, block)) {
    document.execCommand('insertParagraph');
    if (tracked) markSplit(tracked);
    block = closestBlock(selectionInPage().startContainer) || block;
  }
  const table = doc.newTable(rows, cols, tracked);
  block.before(table);
  refreshCss(); // the Table Grid style may be new
  placeCaret(table.rows[0].cells[0].querySelector(BLOCK_SELECTOR) || table.rows[0].cells[0], true);
  doc.renumber(page.value);
  undoStack.commit();
  markDirty();
  refreshChanges();
  onSelectionChange();
  message.value = `Table inserted: ${cols} × ${rows}${tracked ? ' (tracked)' : ''}`;
}

// ------------------------------------------------------------------ inserting pictures

const pictureInput = ref(null);

// ------------------------------------------------------------------ signature

const signatureOpen = ref(false);
let signatureRange = null; // where the signature goes: the selection when the dialog opened

/** Opens the signature pad, for the caret's place (or in place of the selection, a "____" line). */
function openSignature() {
  if (textLocked.value) return explainLock();
  restoreSelection();
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text where your signature should go, or select the line to sign on';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (refuseInNote(range, 'A signature')) return;
  if (controlAt(range.startContainer, page.value) || formTextAt(range.startContainer)) {
    message.value = 'A signature can’t go in a field: put the caret after it';
    return;
  }
  signatureRange = range.cloneRange();
  signatureOpen.value = true;
}

function closeSignature() {
  signatureOpen.value = false;
  refocusAt(signatureRange);
}

function refocusAt(range) {
  page.value.focus();
  if (range && page.value.contains(range.commonAncestorContainer)) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

async function insertSignature({ file, width }) {
  signatureOpen.value = false;
  const at = signatureRange && page.value.contains(signatureRange.commonAncestorContainer) ? signatureRange : null;
  await insertPictures([file], at, { width, descr: 'Signature', what: 'Signature' });
}

// ------------------------------------------------------------------ insert: check box, today's date

const insertMenu = reactive({ open: false, x: 0, y: 0, dates: [] });
const insertMenuEl = ref(null);
const insertButton = ref(null);
let insertRange = null; // where the check box or date goes: the selection when the menu opened

/** Today's date in a few of the forms Word's Date and Time dialog offers, in the text's language. */
function todaysDates(lang) {
  const today = new Date();
  const as = (options) => {
    try { return new Intl.DateTimeFormat(lang, options).format(today); } catch { return new Intl.DateTimeFormat('en', options).format(today); }
  };
  const iso = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  const SHORT = { year: 'numeric', month: '2-digit', day: '2-digit' };
  const dates = [
    { style: 'long', text: as({ dateStyle: 'long' }), picture: wordDatePicture(lang, { dateStyle: 'long' }) },
    { style: 'full', text: as({ dateStyle: 'full' }), picture: wordDatePicture(lang, { dateStyle: 'full' }) },
    { style: 'short', text: as(SHORT), picture: wordDatePicture(lang, SHORT) },
    { style: 'iso', text: iso, picture: 'yyyy-MM-dd' },
  ];
  return dates.filter((d, i) => dates.findIndex((e) => e.text === d.text) === i);
}

function toggleInsertMenu() {
  if (insertMenu.open) return closeInsertMenu();
  restoreSelection();
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text where the check box or date should go';
    return;
  }
  insertRange = range.cloneRange();
  const lang = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('[lang]')?.lang
    || doc.styles.language() || navigator.language || 'en';
  const rect = insertButton.value.getBoundingClientRect();
  Object.assign(insertMenu, {
    open: true,
    x: Math.max(8, Math.min(rect.left, window.innerWidth - 248)),
    y: rect.bottom + 4,
    dates: todaysDates(lang),
  });
  nextTick(() => {
    keepInView(insertMenu, insertMenuEl.value);
    insertMenuEl.value?.querySelector('button')?.focus();
  });
}

function closeInsertMenu(refocusText = false) {
  insertMenu.open = false;
  if (refocusText) refocusAt(insertRange);
}

// "Update automatically", as in Word's Date and Time dialog: off, the date is text; on, a DATE
// field in the same format, which Word updates to the day it opens (or prints) the file. Remembered
// in this browser, as Word remembers it.
const DATE_UPDATES_KEY = 'intact.dateUpdates';
const dateUpdates = ref((() => { try { return localStorage.getItem(DATE_UPDATES_KEY) === '1'; } catch { return false; } })());

function toggleDateUpdates() {
  dateUpdates.value = !dateUpdates.value;
  try { localStorage.setItem(DATE_UPDATES_KEY, dateUpdates.value ? '1' : '0'); } catch { /* this session only */ }
}

/**
 * Today's date at the caret (or in place of the selection): as typed, or (Update automatically) as
 * Word's DATE field with today's date as its result; tracked while tracking.
 */
function insertDate({ text, picture }) {
  closeInsertMenu();
  refocusAt(insertRange);
  let range = selectionInPage();
  if (!range) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  const field = dateUpdates.value;
  if (field && inField(range.startContainer)) {
    message.value = 'A date that updates can’t go inside a field: put the caret after it';
    return;
  }
  if (field && !range.collapsed && refuseRemoving([range])) return;
  undoStack.flush();
  if (!field) {
    document.execCommand('insertText', false, text);
  } else {
    if (!range.collapsed) {
      const opts = trackedBy(range.startContainer);
      if (opts) deleteTracked(range, false, opts);
      else deleteSelection();
      range = selectionInPage();
    }
    insertNodes(doc.newField(dateFieldInstruction(picture), text), range);
    onInput();
  }
  undoStack.commit();
  markDirty();
  message.value = field ? `Date inserted: ${text} – Word updates it to the day it opens the file` : `Date inserted: ${text}`;
}

/**
 * A check box at the caret, as Word's Developer tab inserts one: a content control (w14:checkbox)
 * that a click ticks, here and in Word.
 */
function insertCheckbox() {
  closeInsertMenu();
  refocusAt(insertRange);
  const range = selectionInPage();
  if (!range) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (refuseInNote(range, 'A check box')) return;
  if (controlAt(range.startContainer, page.value) || formTextAt(range.startContainer)) {
    message.value = 'A check box can’t go in a field: put the caret after it';
    return;
  }
  if (!range.collapsed && refuseRemoving([range])) return;
  const box = doc.newCheckbox();
  undoStack.flush();
  const tracked = trackedBy(range.startContainer);
  if (!range.collapsed) {
    if (tracked) deleteTracked(range, false, tracked);
    else deleteSelection();
  }
  if (tracked) insertTracked([box], tracked);
  else insertAtCaret([box]);
  undoStack.commit();
  markDirty();
  refreshChanges();
  message.value = `Check box inserted${tracked ? ' (tracked)' : ''} – click it to tick it`;
}

// ------------------------------------------------------------------ references: contents, captions, citations

// The References dialogs open from the Insert menu, for the place it was opened at (insertRange).
const refDialog = ref(null); // 'equation' | 'citation' | 'caption' | 'xref'
const captionInfo = reactive({ counts: {}, target: 'the paragraph', table: false });
const xrefTargets = ref({});
let xrefEls = {}; // what each cross-reference target is in the page: { Heading: [{ block }], Figure: [{ block, field }] }
const capitalized = (t) => t[0].toUpperCase() + t.slice(1);

function openRefDialog(kind) {
  closeInsertMenu();
  if (textLocked.value) return explainLock();
  refDialog.value = kind;
}

function closeRefDialog() {
  refDialog.value = null;
  refocusAt(insertRange);
}

/**
 * The caret where the Insert menu was opened, back in the page, if `what` can go there: in line
 * with the text (not in a field), or with `block` in a paragraph of the text itself (not a table,
 * note, header, footer or text box), and not while tracking (its paragraphs aren't tracked).
 * Returns the range ({ range, block } with `block`), or null after saying why not.
 */
function insertSpot(what, { block: asBlock = false } = {}) {
  if (textLocked.value) {
    explainLock();
    return null;
  }
  refocusAt(insertRange);
  const range = selectionInPage();
  if (!range) {
    message.value = `Click in the text where ${what} should go`;
    return null;
  }
  const locked = readOnlySelection();
  if (locked) {
    explainReadOnly(locked, true);
    return null;
  }
  if (refuseInNote(range, capitalized(what))) return null;
  if (inField(range.startContainer) || controlAt(range.startContainer, page.value) || formTextAt(range.startContainer)) {
    message.value = `${capitalized(what)} can’t go in a field: put the caret after it`;
    return null;
  }
  if (!asBlock) return range;
  const block = closestBlock(range.startContainer);
  if (!block || block.parentElement !== page.value) {
    message.value = `${capitalized(what)} goes in the text itself: not in a table, header, footer or text box`;
    return null;
  }
  if (trackedBy(range.startContainer)) {
    message.value = `${capitalized(what)} can’t be tracked here: turn off Track changes to insert it`;
    return null;
  }
  return { range, block };
}

/** Puts new paragraphs where `spot` says: in place of an empty paragraph, before one at its start, else after it. */
function placeBlocks({ range, block }, nodes) {
  if (isEmptyBlock(block)) block.replaceWith(...nodes);
  else if (isAtBlockStart(range, block)) block.before(...nodes);
  else block.after(...nodes);
  // Something to type in after them, at the end of the text.
  const last = nodes[nodes.length - 1];
  let next = last.nextElementSibling;
  if (!next || next.matches('.docx-notes, .docx-hf-band')) {
    next = doc.emptyParagraph();
    last.after(next);
  }
  placeCaret(next, true);
}

/** Puts `nodes` at the caret (in place of the selection), as typing would: tracked while tracking. */
function putInline(range, nodes) {
  if (!range.collapsed && refuseRemoving([range])) return false;
  const tracked = trackedBy(range.startContainer);
  if (!range.collapsed) {
    if (tracked) deleteTracked(range, false, tracked);
    else deleteSelection();
  }
  if (tracked) insertTracked(nodes, tracked);
  else insertAtCaret(nodes);
  return true;
}

/** Ends an insertion: one undoable step, the page's numbering and marks up to date, and what was done said. */
function finishInsert(said) {
  refreshCss();
  doc.renumber(page.value);
  onInput();
  undoStack.commit();
  refreshChanges();
  onSelectionChange();
  message.value = said;
}

/**
 * A table of contents, as Word's Automatic Table: "Contents" and the TOC field over headings 1–3,
 * built here from the headings (linked to them) at once. Word fills in the page numbers.
 */
function insertTableOfContents() {
  closeInsertMenu();
  const spot = insertSpot('a table of contents', { block: true });
  if (!spot) return;
  undoStack.flush();
  placeBlocks(spot, doc.renderXmlParagraphs(tocXml(doc.styles.ensureBuiltin('tocHeading'))));
  refreshCss();
  const t = updateFields(doc, page.value, { only: ['TOC'] });
  doc.updateFields = true; // page numbers are Word's: it fills them in when it opens the file
  finishInsert(t.entries
    ? `Table of contents inserted: ${t.entries} ${t.entries === 1 ? 'heading' : 'headings'} – Word fills in the page numbers when it opens the file; Update (F9) adds headings written later`
    : 'Table of contents inserted: no headings yet – give paragraphs a heading style, then Update (F9)');
}

/** The captions in the text, by label: { figure: [{ block, field, label }] } (label as written). */
function captionsInText() {
  const out = {};
  for (const field of page.value.querySelectorAll('span.field[data-instr]')) {
    if (field.closest('.docx-notes, .docx-hf-band, .textbox')) continue;
    const { code, args } = parseInstr(field.dataset.instr);
    if (code !== 'SEQ' || !args[0]) continue;
    const block = field.closest(BLOCK_SELECTOR);
    if (block) (out[args[0].toLowerCase()] ||= []).push({ block, field, label: args[0] });
  }
  return out;
}

/** What a caption at the caret would be for: a table the caret is in, the picture selected, or its paragraph. */
function captionAnchor(range) {
  const at = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  const table = at.closest('table');
  if (table && table.parentElement === page.value) return { el: table, target: 'the table', table: true };
  const pic = picture.el?.isConnected ? closestBlock(picture.el) : null;
  const block = pic || closestBlock(range.startContainer);
  if (!block || block.parentElement !== page.value) return null;
  return { el: block, target: pic ? 'the picture' : 'the paragraph', table: false };
}

function openCaption() {
  closeInsertMenu();
  if (textLocked.value) return explainLock();
  refocusAt(insertRange);
  const range = selectionInPage();
  const anchor = range && captionAnchor(range);
  if (!anchor) {
    message.value = 'Click in the text, a table or a picture the caption is for';
    return;
  }
  const counts = Object.fromEntries(Object.entries(captionsInText()).map(([k, v]) => [k, v.length]));
  Object.assign(captionInfo, { counts, target: anchor.target, table: anchor.table });
  refDialog.value = 'caption';
}

/** Insert Caption: "Figure 1: text" in Word's Caption style, above or below what it is for, numbered with the others. */
function insertCaption({ label, text, position }) {
  refDialog.value = null;
  refocusAt(insertRange);
  const range = selectionInPage();
  const anchor = range && captionAnchor(range);
  if (!anchor) return;
  if (trackedBy(range.startContainer)) {
    message.value = 'A caption can’t be tracked here: turn off Track changes to insert it';
    return;
  }
  undoStack.flush();
  const [p] = doc.renderXmlParagraphs([captionXml(doc.styles.ensureBuiltin('caption'), label, 1, text)]);
  if (position === 'above') anchor.el.before(p);
  else anchor.el.after(p);
  updateFields(doc, page.value, { only: ['SEQ'] }); // its number among the others
  placeCaret(p);
  const n = p.querySelector('span.field')?.textContent || '';
  finishInsert(`Caption inserted: ${label} ${n} – Update (F9) renumbers captions after one is moved`);
}

/** The headings and captions a cross-reference can point to, for the dialog (and the page's elements, kept). */
function openCrossRef() {
  closeInsertMenu();
  if (textLocked.value) return explainLock();
  refocusAt(insertRange);
  const headings = [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((el) => !el.closest('.docx-notes, .docx-hf-band, .textbox, table')
    && (doc.outlineLevel(el) ?? 9) < 9 && el.textContent.trim());
  xrefEls = { Heading: headings.map((block) => ({ block })) };
  const targets = { Heading: headings.map((el) => ({ text: el.textContent.trim(), level: doc.outlineLevel(el) + 1 })) };
  for (const list of Object.values(captionsInText())) {
    const label = capitalized(list[0].label);
    xrefEls[label] = list;
    targets[label] = list.map((c) => ({ text: c.block.textContent.trim() }));
  }
  for (const l of ['Figure', 'Table']) targets[l] ||= [];
  xrefTargets.value = targets;
  refDialog.value = 'xref';
}

/** A new _Ref bookmark around `block`'s text, from its start to its end (or to just after `upTo`), as Word adds one to refer to. */
function refBookmark(block, upTo = null) {
  const name = `_Ref${Math.floor(100000000 + Math.random() * 900000000)}`;
  const [start, end] = doc.newBookmark(name);
  block.insertBefore(start, block.firstChild);
  if (upTo) upTo.after(end);
  else if (block.lastChild?.tagName === 'BR') block.insertBefore(end, block.lastChild);
  else block.append(end);
  return name;
}

/** Insert Cross-reference: a REF field to a bookmark on the heading or caption, showing its text now. */
function insertCrossRef({ kind, index, as }) {
  refDialog.value = null;
  const target = xrefEls[kind]?.[index];
  const range = insertSpot('a cross-reference');
  if (!range || !target?.block.isConnected) return;
  undoStack.flush();
  const { block, field } = target;
  const number = kind !== 'Heading' && as === 'number';
  const name = refBookmark(block, number ? field : null);
  let text = number ? `${target.label} ${field.textContent.trim()}` : block.textContent.trim();
  let instr = ` REF ${name} \\h `;
  if (as === 'position') {
    text = block.compareDocumentPosition(range.startContainer) & Node.DOCUMENT_POSITION_FOLLOWING ? 'above' : 'below';
    instr = ` REF ${name} \\p \\h `;
  }
  if (!putInline(selectionInPage(), doc.newField(instr, text))) return;
  finishInsert(`Cross-reference inserted: “${text}” – Update (F9) brings it up to date`);
}

/** Insert Equation: Word's Office Math at the caret, on its own line in an empty paragraph if asked. */
function insertEquation({ xml, display }) {
  refDialog.value = null;
  const range = insertSpot('an equation');
  if (!range) return;
  undoStack.flush();
  const block = closestBlock(range.startContainer);
  const own = display && block && isEmptyBlock(block);
  const math = own ? xml : xml.replace(/^<m:oMathPara>|<\/m:oMathPara>$/g, '');
  if (block && isEmptyBlock(block)) for (const br of block.querySelectorAll('br')) br.remove();
  if (!putInline(selectionInPage() || range, doc.renderXmlRuns(math))) return;
  finishInsert(`Equation inserted${own ? ' on its own line' : ''} – Word edits it as its own`);
}

/**
 * Insert Citation: a CITATION field for a source the document has, or a new one added to its
 * sources. A new source, or another style, brings the citations and bibliographies up to date.
 */
function insertCitation({ tag, source, page: cited, style }) {
  refDialog.value = null;
  const range = insertSpot('a citation');
  if (!range) return;
  undoStack.flush();
  const restyled = doc.sources.style !== style;
  doc.setCitationStyle(style);
  const src = tag ? doc.sources.list.find((x) => x.tag === tag) : doc.addSource(source);
  if (!src || !putInline(range, doc.renderXmlRuns(citationXml(src, style, cited)))) return;
  if (restyled || !tag) updateFields(doc, page.value, { only: ['CITATION', 'BIBLIOGRAPHY'] });
  finishInsert(`Citation inserted: ${src.authors[0]?.last || src.corporate || src.title}${restyled ? ` – citations now in ${style} style` : ''}`);
}

/** Insert Bibliography: "References" (APA) or "Works Cited" (MLA) and the document's sources, as Word's gallery inserts it. */
function insertBibliography() {
  closeInsertMenu();
  const spot = insertSpot('a bibliography', { block: true });
  if (!spot) return;
  undoStack.flush();
  const style = doc.sources.style === 'MLA' ? 'MLA' : 'APA';
  placeBlocks(spot, doc.renderXmlParagraphs(bibliographyXml({
    sources: doc.sources.list, style, headingStyle: doc.styles.ensureBuiltin('h1'), entryStyle: doc.styles.ensureBuiltin('bibliography'),
  })));
  const n = doc.sources.list.length;
  finishInsert(n
    ? `Bibliography inserted: ${n} ${n === 1 ? 'source' : 'sources'} in ${style} style – Update (F9) adds sources cited later`
    : 'Bibliography inserted: no sources yet – insert citations, then Update (F9)');
}

// ------------------------------------------------------------------ navigation pane

// The Navigation pane's Headings, as Word's: shown or hidden (remembered in this browser).
const NAV_KEY = 'intact.navigationPane';
const navOpen = ref((() => { try { return localStorage.getItem(NAV_KEY) === '1'; } catch { return false; } })());
const navHeadings = ref([]);
const navCurrent = ref(-1);
let navEls = [];

/** The headings of the text, for the pane (not those in tables, notes, headers, footers or text boxes). */
function refreshHeadings() {
  if (!navOpen.value || !page.value) return;
  navEls = [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((el) => !el.closest('.docx-notes, .docx-hf-band, .textbox, table')
    && (doc.outlineLevel(el) ?? 9) < 9 && el.textContent.trim());
  navHeadings.value = navEls.map((el, index) => ({ index, level: doc.outlineLevel(el) + 1, text: el.textContent.trim().replace(/\s+/g, ' ') }));
  updateNavCurrent();
}

/** The heading the caret is under. */
function updateNavCurrent() {
  if (!navOpen.value) return;
  const range = selectionInPage();
  const block = range && closestBlock(range.startContainer);
  if (!block) return;
  let at = -1;
  navEls.forEach((el, i) => {
    if (el === block || el.compareDocumentPosition(block) & Node.DOCUMENT_POSITION_FOLLOWING) at = i;
  });
  navCurrent.value = at;
}

function toggleNav() {
  navOpen.value = !navOpen.value;
  try { localStorage.setItem(NAV_KEY, navOpen.value ? '1' : '0'); } catch { /* this session only */ }
  nextTick(refreshHeadings);
}

/** Goes to heading `i`: in view, the caret at its start. */
function goToHeading(i) {
  const el = navEls[i];
  if (!el?.isConnected) return refreshHeadings();
  page.value.focus();
  placeCaret(el, true);
  el.scrollIntoView({ block: 'start' });
  navCurrent.value = i;
}

// ------------------------------------------------------------------ moving a heading with its text

// As Word's Navigation pane: a heading moved (dragged, or ▲ ▼, Alt+Shift+↑ ↓) takes everything under
// it with it, up to the next heading of its level or above: a chapter with its scenes. The page's
// own paragraphs, tables and the like are moved as they are, so the file keeps them exactly; list
// numbers, note numbers and comments follow. Refused, with the reason, where a move would break
// something Word relies on: while tracking (Word would record a move; Intact doesn't), across a
// section break, or where a bookmark, comment, field or editing permission starts on one side of
// the part's edge and ends on the other.

/** Heading `i`'s part, the page's own children from it to the next heading of its level or above: { nodes } or { why }. */
function headingPart(i) {
  const el = navEls[i];
  if (!el?.isConnected) return { why: 'the headings changed: try again' };
  if (el.parentElement !== page.value) return { why: 'it is inside a field, list or other content that is kept as a whole' };
  const level = navHeadings.value[i].level;
  const k = navHeadings.value.findIndex((h, j) => j > i && h.level <= level);
  const next = k >= 0 ? navEls[k] : null;
  const nodes = [];
  for (let n = el; n; n = n.nextElementSibling) {
    if (n !== el && (n === next || n.matches('.docx-notes, .docx-hf-band'))) break;
    if (next && n !== next && n.contains(next)) return { why: 'the next heading is inside content kept as a whole' };
    nodes.push(n);
  }
  return { nodes, end: k };
}

/** Why `nodes` can't be moved without breaking the file, or ''. */
function unmovable(nodes) {
  if (nodes.some((n) => n.matches('[data-section-end]') || n.querySelector('[data-section-end]'))) {
    return 'it holds a section break (another page setup): move the text in Word';
  }
  // Ranges that start and end apart (bookmarks, comments, permissions, tracked moves) and fields:
  // each one in the part must start and end in it.
  const open = new Map();
  let fields = 0;
  const pairs = { bookmarkStart: 'bookmark', bookmarkEnd: 'bookmark', commentRangeStart: 'comment', commentRangeEnd: 'comment', permStart: 'perm', permEnd: 'perm', moveFromRangeStart: 'moveFrom', moveFromRangeEnd: 'moveFrom', moveToRangeStart: 'moveTo', moveToRangeEnd: 'moveTo' };
  const visit = (x) => {
    const kind = pairs[x.localName];
    if (kind) {
      const key = `${kind}:${wAttrOf(x, 'id')}`;
      open.set(key, (open.get(key) || 0) + (x.localName.endsWith('Start') ? 1 : -1));
    } else if (x.localName === 'fldChar') {
      const t = wAttrOf(x, 'fldCharType');
      if (t === 'begin') fields += 1;
      else if (t === 'end') fields -= 1;
      if (fields < 0) fields = -Infinity;
    }
  };
  const els = nodes.flatMap((n) => [n, ...n.querySelectorAll('[data-raw], [data-ppr]')]).filter((e) => e.dataset.raw != null);
  for (const e of els) {
    const raw = doc.raw[Number(e.dataset.raw)];
    if (!raw || raw.nodeType !== 1) continue;
    visit(raw);
    for (const x of raw.getElementsByTagName('*')) visit(x);
  }
  for (const n of nodes) {
    for (const e of [n, ...n.querySelectorAll('[data-comment-start], [data-comment-end]')]) {
      if (e.dataset.commentStart != null && !e.dataset.raw) open.set(`comment:${e.dataset.commentStart}`, (open.get(`comment:${e.dataset.commentStart}`) || 0) + 1);
      if (e.dataset.commentEnd != null && !e.dataset.raw) open.set(`comment:${e.dataset.commentEnd}`, (open.get(`comment:${e.dataset.commentEnd}`) || 0) - 1);
    }
  }
  if (fields !== 0) return 'a field (a table of contents, an index…) starts on one side of its edge and ends on the other';
  const split = [...open].find(([, v]) => v !== 0);
  if (split) {
    const what = { bookmark: 'a bookmark', comment: 'a comment', perm: 'a part others may edit', moveFrom: 'a tracked move', moveTo: 'a tracked move' }[split[0].split(':')[0]];
    return `${what} starts on one side of its edge and ends on the other`;
  }
  return '';
}

/** The w: attribute `name` of XML element `x`. */
function wAttrOf(x, name) {
  return x.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', name) ?? x.getAttribute(`w:${name}`);
}

/** Moves heading `i`'s part so it comes before `before` (a page child; null: after `after`). */
function moveHeadingPart(i, place, describe, fromKeyboard) {
  const name = navHeadings.value[i]?.text || 'the heading';
  if (textLocked.value) return explainLock();
  if (tracking.value) {
    message.value = `“${name}” wasn’t moved: moving text isn’t tracked here – turn off Track Changes to move it, or cut and paste it`;
    return;
  }
  const part = headingPart(i);
  const why = part.why || unmovable(part.nodes);
  if (why) {
    message.value = `“${name}” can’t be moved here: ${why}`;
    return;
  }
  const { nodes } = part;
  const anchor = place(nodes);
  if (anchor === undefined) return; // nothing to move past: place said so
  if (anchor && nodes.includes(anchor)) return;
  undoStack.flush();
  const frag = document.createDocumentFragment();
  for (const n of nodes) frag.append(n);
  // At the end of the text: before the notes and the footer, where the page has them.
  const end = anchor || page.value.querySelector(':scope > .docx-notes, :scope > .docx-hf-band.footer');
  if (end) end.before(frag);
  else page.value.append(frag);
  doc.renumber(page.value);
  refreshComments();
  onInput();
  undoStack.commit();
  refreshChanges();
  pageChanged();
  schedulePages(0);
  refreshHeadings();
  const at = navEls.indexOf(nodes[0]);
  navCurrent.value = at;
  page.value.focus({ preventScroll: true });
  placeCaret(nodes[0], true);
  nodes[0].scrollIntoView({ block: 'start' });
  if (fromKeyboard) nextTick(() => document.querySelectorAll('.nav-pane li .go')[at]?.focus());
  message.value = `Moved “${name}” ${describe}, with its text – Undo (Ctrl+Z) puts it back`;
}

/** ▲ ▼ (Alt+Shift+↑ ↓): heading `i` past the heading of its level before or after it, under the same heading. */
function moveHeadingBy(i, dir) {
  const fromKeyboard = document.activeElement?.closest?.('.nav-pane') != null;
  const heads = navHeadings.value;
  const level = heads[i]?.level;
  // Its neighbour of the same level, not past a heading above it (its chapter's, for a scene).
  let j = i + dir;
  while (j >= 0 && j < heads.length && heads[j].level > level) j += dir;
  const name = heads[i]?.text || 'the heading';
  if (j < 0 || j >= heads.length || heads[j].level < level) {
    message.value = `“${name}” is already the ${dir < 0 ? 'first' : 'last'}${level > 1 ? ' under its heading' : ''}`;
    if (fromKeyboard) nextTick(() => document.querySelectorAll('.nav-pane li .go')[i]?.focus());
    return;
  }
  const other = heads[j].text;
  if (dir < 0) {
    moveHeadingPart(i, () => navEls[j], `above “${other}”`, fromKeyboard);
  } else {
    moveHeadingPart(i, () => {
      const { nodes, why } = headingPart(j);
      if (why) {
        message.value = `“${name}” can’t be moved past “${other}”: ${why}`;
        return undefined;
      }
      return nodes[nodes.length - 1].nextElementSibling?.matches('.docx-notes, .docx-hf-band') ? null : nodes[nodes.length - 1].nextElementSibling;
    }, `below “${other}”`, fromKeyboard);
  }
}

/** A heading dragged onto heading `to`: it goes before it, with its text. */
function moveHeadingBefore(from, to) {
  const target = navEls[to];
  const name = navHeadings.value[from]?.text || 'the heading';
  if (!target?.isConnected || target.parentElement !== page.value) {
    message.value = `“${name}” can’t go there: that heading is inside content kept as a whole`;
    return;
  }
  moveHeadingPart(from, (nodes) => {
    if (nodes.includes(target)) {
      message.value = `“${name}” can’t go inside its own text`;
      return undefined;
    }
    return target;
  }, `above “${navHeadings.value[to].text}”`, false);
}

function pickPicture() {
  pictureInput.value.value = '';
  pictureInput.value.click();
}

function onPicturesPicked(e) {
  const files = [...e.target.files];
  e.target.value = '';
  if (files.length) insertPictures(files);
}

/**
 * Inserts pictures at the caret (or `at`, a range), in line with the text and replacing what is
 * selected, as Word does; tracked while tracking.
 */
async function insertPictures(files, at = null, { width, descr, what = 'Picture' } = {}) {
  if (at) {
    page.value.focus();
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(at);
  } else {
    refocus();
  }
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text where the picture should go';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  // A logo goes in a header or footer, with the part's own picture files; not in notes or text boxes.
  // (The innermost story counts: a text box in a footer is a text box.)
  const story = noteAt(range.startContainer);
  const part = story?.classList.contains('docx-hf-part') ? doc.partOf(story) : null;
  if (!part && refuseInNote(range, 'A picture')) return;
  if (controlAt(range.startContainer, page.value)) {
    message.value = 'A picture can’t go in a field: insert it in Word';
    return;
  }
  if (!range.collapsed && refuseRemoving([range])) return;
  const where = range.cloneRange();
  const made = [];
  const reasons = [];
  for (const file of files) {
    const res = await doc.addPicture(file, { width, descr, part });
    if (res.span) made.push(res.span);
    else reasons.push(res.reason);
  }
  if (!made.length) {
    message.value = reasons[0] || 'No picture was inserted';
    return;
  }
  page.value.focus();
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(where);
  undoStack.flush();
  const tracked = trackedBy(where.startContainer);
  if (!where.collapsed) {
    if (tracked) deleteTracked(where, false, tracked);
    else deleteSelection();
  }
  if (tracked) insertTracked(made, tracked);
  else insertAtCaret(made);
  page.value.normalize();
  undoStack.commit();
  markDirty();
  refreshChanges();
  onSelectionChange();
  message.value = `${made.length === 1 ? what : `${made.length} pictures`} inserted${tracked ? ' (tracked)' : ''}${reasons.length ? ` – ${reasons[0]}` : ''}`;
}

// ------------------------------------------------------------------ adding comments

const composer = reactive({ open: false });
let composerRange = null; // where the new comment goes

/** With nothing selected, a comment goes on the word at the caret, as in Word. */
function wordAround(range) {
  const { startContainer: node, startOffset: offset } = range;
  if (!range.collapsed || node.nodeType !== 3) return range;
  const letter = /[\p{L}\p{N}_'’-]/u;
  let s = offset;
  let e = offset;
  while (s > 0 && letter.test(node.data[s - 1])) s--;
  while (e < node.length && letter.test(node.data[e])) e++;
  if (s === e) return range;
  const word = document.createRange();
  word.setStart(node, s);
  word.setEnd(node, e);
  return word;
}

/** New comment (Ctrl+Alt+M): opens a box for it in the comments panel, on the selected text. */
function startComment() {
  if (!canComment.value) return explainLock();
  restoreSelection();
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text, or select some, to comment on it';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (refuseInNote(range, 'A new comment')) return;
  const hostOf = (n) => (n.nodeType === 1 ? n : n.parentElement).closest('[contenteditable="true"]');
  if (hostOf(range.startContainer) !== hostOf(range.endContainer)) {
    message.value = 'A comment can’t run from a table cell into the text around it: select within one';
    return;
  }
  composerRange = wordAround(range).cloneRange();
  composer.open = true;
  commentsOpen.value = true;
  paintComments();
}

function postComment({ text, name }) {
  if (name) {
    author.value = name;
    named.value = true;
    try { localStorage.setItem(AUTHOR_KEY, name); } catch { /* this session only */ }
  }
  undoStack.flush();
  const comment = doc.addComment(composerRange, { author: author.value, text, date: revisionDate() });
  composer.open = false;
  composerRange = null;
  undoStack.commit();
  markDirty();
  refreshComments();
  selectComment(comment.id, false);
  page.value.focus();
  message.value = 'Comment added';
}

function cancelComment() {
  if (!composer.open) return;
  composer.open = false;
  composerRange = null;
  paintComments();
  page.value.focus();
  restoreSelection();
}

/** Runs a change to comments as one undoable step, then refreshes the panel. */
function changeComments(change, done) {
  undoStack.flush();
  change();
  undoStack.commit();
  markDirty();
  refreshComments();
  message.value = done;
}

/** A reply to thread `id`, as in Word (written as a real reply, linked to its comment). */
function replyComment({ id, text, name }) {
  if (name) {
    author.value = name;
    named.value = true;
    try { localStorage.setItem(AUTHOR_KEY, name); } catch { /* this session only */ }
  }
  let reply = null;
  changeComments(() => {
    const opts = { author: author.value, text, date: revisionDate() };
    reply = doc.addReply(id, opts, page.value) || doc.replyInStory(id, opts); // in the text, or in a header, footer or note
  }, 'Reply added');
  if (!reply) message.value = 'This comment isn’t in the text here: reply to it in Word';
}

/** A comment's new text, kept on its mark until saved, so Undo covers it. */
function editComment({ id, text }) {
  const mark = commentMark(id);
  if (!text) return;
  if (mark) changeComments(() => { mark.dataset.edit = text; }, 'Comment edited');
  else if (doc.storyOf(id)) changeComments(() => doc.editStoryComment(id, text), 'Comment edited');
}

/** Resolves or reopens a thread, kept on its mark until saved, so Undo covers it. */
function resolveComment({ id, done }) {
  const mark = commentMark(id);
  const text = done ? 'Comment resolved' : 'Comment reopened';
  if (mark) changeComments(() => { mark.dataset.done = done ? '1' : '0'; }, text);
  else if (doc.storyOf(id)) changeComments(() => doc.resolveStoryComment(id, done), text);
}

/** Deletes a comment and its replies (their marks go from the text); Undo brings them back. */
function deleteComment(id) {
  undoStack.flush();
  const marks = page.value.querySelectorAll(`[data-comment-start="${CSS.escape(id)}"], [data-comment-end="${CSS.escape(id)}"], [data-comment-ref="${CSS.escape(id)}"]`);
  for (const el of marks) el.remove();
  if (!marks.length && doc.storyOf(id)) doc.deleteStoryComment(id); // in a header, footer or note
  undoStack.commit();
  markDirty();
  refreshComments();
  message.value = 'Comment deleted – Undo (Ctrl+Z) brings it back';
}

// ------------------------------------------------------------------ accepting and rejecting changes

const changes = reactive({ total: 0 });

// On a phone, a document with tracked changes or comments puts the review's tools first in the
// toolbar (see the styles: .review-first), and says so once when it opens.
const reviewFirst = computed(() => changes.total > 0 || threads.value.length > 0);
const onPhone = () => window.matchMedia?.('(max-width: 720px)').matches;
let reviewHinted = false;
onMounted(() => watch(() => [changes.total, threads.value.length], ([n, c]) => {
  if (reviewHinted || !(n || c)) return;
  reviewHinted = true;
  if (!onPhone() || message.value) return;
  const what = [n && `${n} tracked ${n === 1 ? 'change' : 'changes'}`, c && `${c} ${c === 1 ? 'comment' : 'comments'}`].filter(Boolean).join(' and ');
  message.value = n ? `${what} to review: Accept and Reject are first in the toolbar` : `${what} to review: they are first in the toolbar`;
}, { immediate: true }));
const resolved = reactive({ accepted: 0, rejected: 0 }); // this session, for the review
const changeMenu = reactive({ open: null, x: 0, y: 0 });
const changeMenuEl = ref(null);
const acceptMenuButton = ref(null);
const rejectMenuButton = ref(null);
const changesTitle = computed(() => {
  if (!changes.total) return 'No tracked changes';
  const parts = [`${changes.total} tracked ${changes.total === 1 ? 'change' : 'changes'} to accept or reject here`];
  return parts.join('; ');
});

function refreshChanges() {
  if (!page.value) return;
  changes.total = changeItems(page.value).length;
}

const revisionOfItem = (item) => doc.revisionOf(item.el, ['run', 'para', 'row', 'cells', 'merge', 'tprops', 'rprops', 'cprops', 'locked'].includes(item.kind) ? item.kind : null);

/** "the insertion by Alice", "your deletion", "the formatting change by Eve", "the inserted row by Dan". */
function describeChange(item) {
  const rev = revisionOfItem(item);
  if (!rev) return 'the change';
  const who = rev.own ? 'your' : 'the';
  const by = rev.own || !rev.author ? '' : ` by ${rev.author}`;
  if (item.kind === 'run') return `the formatting change${by}`;
  if (item.kind === 'para') return `the paragraph formatting change${by}`;
  if (item.kind === 'row') return `the ${rev.kind === 'ins' ? 'inserted' : 'deleted'} row${by}`;
  if (item.kind === 'cells') return `the ${rev.kind === 'ins' ? 'inserted' : 'deleted'} column${by}`;
  if (item.kind === 'merge') return `the ${item.el.dataset.revMerge === 'merge' || mergeGroup(item.el).some((c) => c.dataset.revMerge === 'merge') ? 'merged' : 'split'} cells${by}`;
  if (item.kind === 'locked') return `the ${rev.count === 1 ? 'tracked change' : `${rev.count} tracked changes`} in ${rev.what}${by}`;
  if (item.kind === 'tprops') return `the table formatting change${by}`;
  if (item.kind === 'rprops') return `the row formatting change${by}`;
  if (item.kind === 'cprops') return `the cell formatting change${by}`;
  if (rev.mark) return `${who} paragraph break ${rev.kind === 'ins' ? 'insertion' : 'deletion'}${by}`;
  return `${who} ${rev.kind === 'ins' ? 'insertion' : 'deletion'}${by}`;
}

/** A caret left between paragraphs (or before a table) moves into the text next to it. */
function settleCaret(range) {
  const { startContainer: node, startOffset: offset } = range;
  if (node.nodeType !== 1 || node.closest('[contenteditable="true"]') !== node) return range;
  let next = node.childNodes[offset] || node.childNodes[offset - 1];
  if (next?.tagName === 'TABLE') next = next.querySelector('td[contenteditable="true"] :is(p, h1, h2, h3, h4, h5, h6)') || next;
  if (!next || next.nodeType !== 1 || next.tagName === 'TABLE') return range;
  const r = document.createRange();
  r.setStart(next, 0);
  r.collapse(true);
  return r;
}

/** Selects a change (an inline one, or the end of a paragraph whose break it is) and shows it. */
// The change Next or Previous last went to: a browser drops a selection in read-only content (a
// locked cell) when a toolbar button is clicked, so Accept and Reject fall back on it.
let selectedChange = null;

function selectChange(item) {
  selectedChange = item;
  const range = itemRange(item);
  page.value.focus();
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  savedRange = range.cloneRange();
  item.el.scrollIntoView?.({ block: 'nearest' });
  flash(item.el);
  const rev = revisionOfItem(item);
  message.value = `${describeChange(item).replace(/^./, (c) => c.toUpperCase())}${rev?.date ? `, ${rev.date.slice(0, 10)}` : ''}`;
}

function gotoChange(dir) {
  showAllMarkup();
  const items = changeItems(page.value);
  const item = adjacentChange(selectionInPage() || savedRange, items, dir);
  if (item) selectChange(item);
  else message.value = 'There are no tracked changes to go to';
}

/**
 * Accepts or rejects the changes at the selection (see changesAt), as one undoable step. `then`:
 * 'next' goes on to the next change, as Word's Accept and Reject buttons do; 'stay' doesn't.
 */
function resolveCurrent(accept, then = 'next') {
  showAllMarkup();
  const range = selectionInPage() || savedRange;
  const items = changeItems(page.value);
  let targets = changesAt(range, items);
  if (!targets.length && selectedChange?.el.isConnected) targets = items.filter((i) => i.el === selectedChange.el && i.kind === selectedChange.kind);
  if (!targets.length) {
    message.value = 'Put the cursor on a tracked change (or select some), or use Next change';
    return;
  }
  const what = targets.length === 1 ? describeChange(targets[0]) : `${targets.length} changes`;
  // Afterwards, the caret goes where the first change was.
  const first = targets[0];
  const at = document.createRange();
  if (first.kind === 'mark') at.setStart(first.el, first.el.childNodes.length);
  else at.setStartBefore(['row', 'cells', 'merge', 'tprops', 'rprops', 'cprops'].includes(first.kind) ? first.el.closest('table') : first.el);
  at.collapse(true);
  resolveItems(targets, accept);
  const sel = window.getSelection();
  if (at.startContainer.isConnected) {
    const caret = settleCaret(at);
    sel.removeAllRanges();
    sel.addRange(caret);
    savedRange = caret.cloneRange();
  }
  message.value = `${accept ? 'Accepted' : 'Rejected'} ${what}`;
  if (then === 'next' && targets.length === 1) {
    const next = adjacentChange(selectionInPage(), changeItems(page.value), 1);
    if (next) {
      selectChange(next);
      message.value = `${accept ? 'Accepted' : 'Rejected'} ${what}. Next: ${describeChange(next)}`;
    } else {
      message.value = `${accept ? 'Accepted' : 'Rejected'} ${what}. No more tracked changes`;
    }
  }
}

function resolveAll(accept) {
  const items = changeItems(page.value);
  if (!items.length) return;
  resolveItems(items, accept);
  message.value = `${accept ? 'Accepted' : 'Rejected'} all ${items.length} ${items.length === 1 ? 'change' : 'changes'}`;
}

/** Resolves changes, last first so earlier ones stay where they are, as one undoable step. */
function resolveItems(items, accept) {
  closeChooser();
  undoStack.flush();
  for (const item of [...items].reverse()) {
    if (item.el.isConnected) doc.resolveRevision(item, accept);
  }
  resolved[accept ? 'accepted' : 'rejected'] += items.length;
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
  refreshChanges();
  refreshComments();
}

function openChangeMenu(which) {
  if (changeMenu.open === which) return closeChangeMenu();
  const button = which === 'accept' ? acceptMenuButton.value : rejectMenuButton.value;
  const rect = button.getBoundingClientRect();
  Object.assign(changeMenu, { open: which, x: Math.max(8, Math.min(rect.left - 150, window.innerWidth - 220)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(changeMenu, changeMenuEl.value);
    changeMenuEl.value?.querySelector('button')?.focus();
  });
}

function closeChangeMenu() {
  changeMenu.open = null;
  page.value.focus();
  restoreSelection();
}

function menuResolve(accept, how) {
  changeMenu.open = null;
  page.value.focus();
  restoreSelection();
  if (how === 'all') resolveAll(accept);
  else resolveCurrent(accept, how === 'one' ? 'next' : 'stay');
}

// ------------------------------------------------------------------ superscript, subscript, highlight

/** Runs a change to character formatting of the selection as one undoable step; false if refused. */
function formatSelection(change) {
  restoreSelection();
  const range = selectionInPage();
  if (!range || inPlainField()) return false;
  const locked = readOnlySelection();
  if (locked) {
    explainReadOnly(locked, true);
    return false;
  }
  undoStack.flush();
  const after = withTrackedRuns(range, () => change(range));
  if (after) selectRange(after);
  undoStack.commit();
  onSelectionChange();
  markDirty();
  return true;
}

/** Superscript or subscript on or off, as Word's buttons: one turns the other off. */
function toggleVertAlign(kind) {
  formatSelection((range) => {
    if (range.collapsed) {
      document.execCommand(kind); // for what is typed next
      return null;
    }
    const on = formatState(range)[kind];
    return setVertAlign(range, on ? null : kind);
  });
}

// The highlight colour last chosen, remembered in this browser only.
const HIGHLIGHT_KEY = 'intact.highlight';
const highlightColor = ref((() => {
  try {
    const c = localStorage.getItem(HIGHLIGHT_KEY);
    return HIGHLIGHT_COLORS.some(([name]) => name === c) ? c : 'yellow';
  } catch { return 'yellow'; }
})());
const palette = reactive({ open: false, x: 0, y: 0 });
const paletteEl = ref(null);
const paletteButton = ref(null);
const colorLabel = (name) => HIGHLIGHT_COLORS.find(([n]) => n === name)?.[1] || name;

/** Highlights the selection in `color`, or removes highlight (null). `toggle`: off if it is that colour. */
function applyHighlight(color, toggle = false) {
  const range = selectionInPage() || savedRange;
  if (!range || range.collapsed) {
    message.value = 'Select the text to highlight first';
    return;
  }
  formatSelection((r) => {
    const off = !color || (toggle && formatState(r).highlight === color);
    return setHighlight(r, off ? null : color);
  });
}

function togglePalette() {
  if (palette.open) return closePalette();
  const rect = paletteButton.value.getBoundingClientRect();
  Object.assign(palette, { open: true, x: Math.max(8, Math.min(rect.left - 120, window.innerWidth - 200)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(palette, paletteEl.value);
    paletteEl.value?.querySelector('.swatch.current, .swatch')?.focus();
  });
}

function closePalette() {
  palette.open = false;
  page.value.focus();
  restoreSelection();
}

function pickHighlight(color) {
  palette.open = false;
  if (color) {
    highlightColor.value = color;
    try { localStorage.setItem(HIGHLIGHT_KEY, color); } catch { /* this session only */ }
  }
  page.value.focus();
  restoreSelection();
  applyHighlight(color);
}

// ------------------------------------------------------------------ keyboard shortcuts

/** Runs one of Word's shortcuts (see editor/shortcuts.js). */
function runShortcut(name) {
  if (['left', 'center', 'right', 'justify'].includes(name)) return alignBlocks(name);
  if (name === 'growFont') return growFont(1);
  if (name === 'shrinkFont') return growFont(-1);
  if (name === 'growFontPoint') return growFont(1, true);
  if (name === 'shrinkFontPoint') return growFont(-1, true);
  if (['bold', 'italic', 'underline'].includes(name)) return exec(name);
  if (name === 'link') return openLinkDialog();
  if (name === 'normal') return applyBlockStyle('normal');
  if (name.startsWith('heading')) return applyBlockStyle(`h${name.slice(-1)}`);
  if (name === 'bullets') return toggleList('bullet');
  if (name === 'indent' || name === 'outdent') return indentBlocks(name === 'indent' ? 1 : -1);
  if (name === 'lineSingle') return setLineSpacing(1);
  if (name === 'lineOneHalf') return setLineSpacing(1.5);
  if (name === 'lineDouble') return setLineSpacing(2);
  if (name === 'spaceBefore') return toggleSpace('before');
  if (name === 'clearFormatting') return clearCharacterFormatting();
  if (name === 'superscript' || name === 'subscript') return toggleVertAlign(name);
  if (name === 'highlight') return applyHighlight(highlightColor.value, true);
  if (name === 'formattingMarks') return toggleMarks();
  if (name === 'comment') return startComment();
  if (name === 'nonBreakingSpace') return typeText('\u00a0');
  if (name === 'pageBreak') return insertPageBreak();
  if (name === 'footnote') return insertFootnote();
  if (name === 'endnote') return insertEndnote();
  if (name === 'copyFormat') return startPainter(false, true);
  if (name === 'pasteFormat') return paintSelection();
  if (name === 'hidden') return toggleHidden();
  if (name === 'readAloud') return toggleReadAloud();
}

// ------------------------------------------------------------------ answer lines

/**
 * Insert > Answer line: a line from the caret to the paragraph's right edge, to write an answer or a
 * name on ("Name: ________"), as teachers make one in Word: an underlined tab to a right tab stop at
 * the margin (added to the paragraph if it hasn't one there). The tab is a run of its own, so Word
 * and Intact draw the line alike and print it (see editor/tabLayout.js).
 */
function insertAnswerLine() {
  closeInsertMenu();
  refocusAt(insertRange);
  let range = selectionInPage();
  if (!range) {
    message.value = t('Click in the text where the answer line should go');
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  const block = closestBlock(range.startContainer);
  if (!block || !page.value.contains(block)) return;
  // The right edge of the paragraph's text, from the margin, in twips.
  const cs = getComputedStyle(block);
  const width = block.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
  const { left } = doc.paragraphFormat(block);
  const pos = Math.max(left + 720, left + Math.floor(width * 15) - 15);
  const { stops } = doc.tabStops(block);
  if (!stops.some((s) => s.align === 'right' && Math.abs(s.pos - pos) <= 60)) {
    formatParagraphs(() => ({ tabs: [...stops.filter((s) => Math.abs(s.pos - pos) > 60), { pos, align: 'right', leader: 'none' }].sort((a, b) => a.pos - b.pos) }));
  }
  range = selectionInPage() || range;
  if (!range.collapsed) {
    range.collapse(false);
    selectRange(range);
  }
  typeText('\t');
  // The tab just typed, underlined: a run of its own.
  const sel = window.getSelection();
  const node = sel.focusNode;
  const at = sel.focusOffset;
  if (node?.nodeType !== 3 || node.data[at - 1] !== '\t') return;
  const tab = document.createRange();
  tab.setStart(node, at - 1);
  tab.setEnd(node, at);
  selectRange(tab);
  restyleSelection(setUnderline);
  // The caret after the line's run, so what is typed next isn't underlined.
  const after = window.getSelection().rangeCount ? window.getSelection().getRangeAt(0) : null;
  let runEl = after?.endContainer.nodeType === 3 ? after.endContainer.parentElement : after?.endContainer;
  while (runEl && runEl.parentElement && runEl.parentElement !== block && !runEl.parentElement.matches('[contenteditable], td, th')) runEl = runEl.parentElement;
  if (runEl && runEl !== block && block.contains(runEl)) {
    const caret = document.createRange();
    caret.setStartAfter(runEl);
    caret.collapse(true);
    selectRange(caret);
  } else window.getSelection().collapseToEnd();
  // As Word's Ctrl+U, Tab, Ctrl+U: underline off for what is typed next.
  if (document.queryCommandState('underline')) document.execCommand('underline');
  scheduleTabs();
  message.value = t('Answer line added: an underlined tab to a right tab stop at the margin, as Word makes one');
}

// ------------------------------------------------------------------ lines, arrows and text boxes

const SHAPE_NAMES = { line: 'Line', arrow: 'Arrow', textbox: 'Text box' };

/**
 * Insert > Line, Arrow or Text box, as Word's Insert > Shapes makes them. A line or an arrow, 4 cm
 * long, goes in front of the text where the caret is (anchored to its paragraph), selected, to be
 * dragged into place, resized and turned as a picture is; a text box (5 × 1.5 cm, with a thin
 * border) goes in line with the text at the caret, to type in.
 */
async function insertShape(kind) {
  closeInsertMenu();
  refocusAt(insertRange);
  const range = selectionInPage();
  if (!range) {
    message.value = t('Click in the text where it should go');
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (noteAt(range.startContainer)) {
    message.value = t('{what} goes in the text itself, not in a header, footer, note or text box', { what: t(SHAPE_NAMES[kind]) });
    return;
  }
  if (controlAt(range.startContainer, page.value) || formTextAt(range.startContainer)) {
    message.value = t('{what} can’t go in a field: put the caret after it', { what: t(SHAPE_NAMES[kind]) });
    return;
  }
  const where = range.cloneRange();
  where.collapse(false);
  let opts;
  if (kind === 'textbox') opts = { cx: 1800000, cy: 540000 };
  else {
    // At the caret, from where its paragraph's text starts (EMU), halfway down its line.
    const block = closestBlock(where.startContainer);
    const caret = where.getClientRects()[0] || where.startContainer.parentElement?.getBoundingClientRect();
    const box = block?.getBoundingClientRect();
    const zoom = sheetZoom.value || 1;
    const x = caret && box ? Math.max(0, (caret.left - box.left - parseFloat(getComputedStyle(block).paddingLeft) * zoom) / zoom) : 0;
    const y = caret && box ? Math.max(0, (caret.top + caret.height / 2 - box.top) / zoom) : 0;
    // In whole pixels, as a picture dropped somewhere is placed.
    opts = { cx: 1440000, cy: 0, offset: { x: Math.round(x) * 9525, y: Math.round(y) * 9525 } };
  }
  const span = await doc.addShape(kind, opts);
  page.value.focus();
  selectRange(where);
  undoStack.flush();
  const tracked = trackedBy(where.startContainer);
  if (tracked) insertTracked([span], tracked);
  else insertAtCaret([span]);
  page.value.normalize();
  undoStack.commit();
  markDirty();
  refreshChanges();
  if (kind === 'textbox') {
    // In it, ready to type.
    // Its paragraph (not the box around it, which a selector list would find first).
    const p = span.querySelector('.textbox-content p') || span.querySelector('.textbox-content');
    const caret = document.createRange();
    caret.selectNodeContents(p);
    caret.collapse(true);
    p.closest('[contenteditable="true"]')?.focus();
    selectRange(caret);
    message.value = t('Text box inserted: type in it{tracked}', { tracked: tracked ? t(' (tracked)') : '' });
  } else {
    selectPicture(span);
    message.value = t('{what} inserted{tracked}: drag it into place, a handle to make it longer, or Layout Options to turn it', { what: t(SHAPE_NAMES[kind]), tracked: tracked ? t(' (tracked)') : '' });
  }
  onSelectionChange();
}

// ------------------------------------------------------------------ read aloud

// Read Aloud, as Word's (Review > Read Aloud, Ctrl+Alt+Space), with voices on this device only: see
// editor/readAloud.js. The paragraph and the word being read are highlighted, never changed.
const reading = reactive({ on: false, paused: false });
let reader = null;

/** The language a paragraph is read in: its text's, else the document's. */
const langOf = (block) => block.querySelector('[lang]')?.lang || block.closest('[lang]')?.lang || doc.styles.language() || navigator.language || 'en';

/** A language's name for messages ("Indonesian", not its region: any voice of the language reads it), or its tag. */
function languageName(tag) {
  const primary = tag.split(/[-_]/)[0];
  try { return new Intl.DisplayNames([lang.value === 'id' ? 'id' : 'en'], { type: 'language' }).of(primary) || tag; } catch { return tag; }
}

function readerState({ reading: on, paused, error }) {
  reading.on = on;
  reading.paused = !!paused;
  if (!error) return;
  if (error === 'unsupported') message.value = t('This browser can’t read aloud');
  else if (error === 'no-local-voices') message.value = t('No voice on this device to read with: add one in the system’s speech settings. Intact reads only with voices on this device, so the text isn’t sent anywhere');
  else if (error.startsWith('no-voice:')) message.value = t('No voice on this device reads {language}: add one in the system’s speech settings. Intact reads only with voices on this device, so the text isn’t sent anywhere', { language: languageName(error.slice(9) || 'en') });
  else message.value = t('Reading aloud stopped: the browser’s speech failed');
}

/** Reads aloud from the caret's paragraph to the end, or the selected paragraphs; stops if reading. */
function toggleReadAloud() {
  reader ||= createReader(page.value, { onState: readerState, langOf });
  if (reading.on) {
    reader.stop();
    message.value = t('Reading aloud stopped');
    return;
  }
  const range = selectionInPage() || savedRange;
  const from = range ? closestBlock(range.startContainer) : null;
  const only = range && !range.collapsed ? [...page.value.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li')].filter((b) => range.intersectsNode(b)) : null;
  message.value = only ? t('Reading the selected text aloud: Esc or Stop to stop') : t('Reading aloud from the caret: Esc or Stop to stop');
  reader.start(from, only);
}

// Closing the document stops reading it.
onBeforeUnmount(() => reader?.stop());

function pauseReadAloud() {
  if (!reader) return;
  if (reading.paused) reader.resume();
  else reader.pause();
}

/**
 * Hidden text on or off for the selected text, as Word's Ctrl+Shift+H: hidden unless all of it is
 * already. A test's answers, hidden, make its answer key: shown here, printed only when asked.
 */
function toggleHidden() {
  const range = selectionInPage() || savedRange;
  if (!range || range.collapsed) {
    message.value = t('Select the text to hide first, such as a test’s answers');
    page.value.focus();
    return;
  }
  const texts = portions(range).map((p) => p.node);
  const allHidden = texts.length > 0 && texts.every((n) => n.parentElement?.closest('[data-hidden]'));
  restyleSelection(setHidden(!allHidden));
  state.hidden = !allHidden;
  message.value = allHidden ? t('Hidden text shown again: it prints') : t('Hidden text: shown here with a dotted line, and printed only when Print hidden text is chosen');
}

/**
 * Ctrl+Space (and Clear formatting): takes the selected text's direct formatting and character style
 * off, as Word does: bold, italic, underline, font, size, colour… Highlight stays. At the caret, it
 * turns bold, italic, underline, strikethrough, superscript and subscript off for what is typed next.
 */
function clearCharacterFormatting() {
  formatSelection((range) => {
    if (!range.collapsed) return restyle(doc, range, clearDirect((id) => doc.styles.byId.get(id)?.name.toLowerCase() === 'hyperlink'));
    for (const cmd of [...CHAR_FORMATS, 'superscript', 'subscript']) {
      // A mixed selection may first turn fully on; the second toggle then turns it off.
      for (let i = 0; i < 2 && document.queryCommandState(cmd); i++) document.execCommand(cmd);
    }
    return null;
  });
}

// ------------------------------------------------------------------ change case

const caseMenu = reactive({ open: false, x: 0, y: 0 });
const caseMenuEl = ref(null);
const caseButton = ref(null);
let caseRange = null; // the text to change: the selection when the menu opened

function toggleCaseMenu() {
  if (caseMenu.open) return closeCaseMenu();
  restoreSelection();
  const range = selectionInPage();
  if (!range) {
    message.value = 'Select the text to change the case of';
    return;
  }
  caseRange = range.cloneRange();
  const rect = caseButton.value.getBoundingClientRect();
  Object.assign(caseMenu, { open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 220)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(caseMenu, caseMenuEl.value);
    caseMenuEl.value?.querySelector('button')?.focus();
  });
}

function closeCaseMenu(refocusText = false) {
  caseMenu.open = false;
  if (refocusText) refocusAt(caseRange);
}

/** The word at a collapsed `range` (letters and digits around the caret), or null. */
function wordAt(range) {
  const node = range.startContainer;
  if (node.nodeType !== 3) return null;
  const isWord = (ch) => /[\p{L}\p{N}'’]/u.test(ch || '');
  let start = range.startOffset;
  let end = range.startOffset;
  while (start > 0 && isWord(node.data[start - 1])) start--;
  while (end < node.length && isWord(node.data[end])) end++;
  if (start === end) return null;
  const word = document.createRange();
  word.setStart(node, start);
  word.setEnd(node, end);
  return word;
}

/**
 * Change Case, as Word's Aa menu and Shift+F3: the selected text (or the word at the caret) in
 * `kind`'s case (see editor/case.js), or, with `kind` null, the next in Shift+F3's cycle (UPPERCASE,
 * lowercase, Capitalize Each Word). Each run keeps its formatting. While tracking, each changed run
 * is a tracked deletion and insertion, as Word records a change of case. Deleted text isn't changed.
 */
function changeCaseTo(kind) {
  if (caseMenu.open) closeCaseMenu(true);
  else restoreSelection();
  if (textLocked.value) return explainLock();
  let range = selectionInPage();
  if (!range) return;
  if (range.collapsed) {
    range = wordAt(range);
    if (!range) {
      message.value = 'Select the text to change the case of, or put the caret in a word';
      return;
    }
    selectRange(range);
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  const pieces = portions(range).filter((p) => !p.node.parentElement.closest('[data-tc="del"]'));
  if (!pieces.length) return;
  const block = closestBlock(range.startContainer) || closestBlock(pieces[0].node);
  const before = document.createRange();
  if (block) {
    before.setStart(block, 0);
    before.setEnd(pieces[0].node, pieces[0].s);
  }
  const lang = pieces[0].node.parentElement.closest('[lang]')?.lang || doc.styles.language() || undefined;
  const texts = pieces.map((p) => p.node.data.slice(p.s, p.e));
  kind ||= nextCase(texts.join(''), lang);
  if (!kind) return;
  const changed = changeCase(texts, kind, block ? before.toString() : '', lang);
  if (changed.every((t, i) => t === texts[i])) {
    message.value = `The text is in ${caseLabel(kind)} already`;
    return;
  }
  const tracked = trackedBy(range.startContainer);
  if (tracked && refuseRemoving([range])) return;
  undoStack.flush();
  let first = null;
  let last = null;
  if (tracked) {
    // From the last piece back, so that the pieces before keep their offsets.
    const date = revisionDate();
    for (let i = pieces.length - 1; i >= 0; i--) {
      const { node, s, e } = pieces[i];
      if (changed[i] === texts[i]) {
        first = [node, s];
        last ||= [node, e];
        continue;
      }
      const r = document.createRange();
      r.setStart(node, s);
      r.setEnd(node, e);
      const after = replaceTracked(r, changed[i], { ...tracked, date });
      first = [after.node, after.offset - changed[i].length];
      last ||= [after.node, after.offset];
    }
    refreshChanges();
  } else {
    for (let i = pieces.length - 1; i >= 0; i--) {
      const { node, s, e } = pieces[i];
      node.data = node.data.slice(0, s) + changed[i] + node.data.slice(e);
      first = [node, s];
      last ||= [node, s + changed[i].length];
    }
  }
  const sel = document.createRange();
  sel.setStart(...first);
  sel.setEnd(...last);
  selectRange(sel); // still selected, as in Word: Shift+F3 again goes on to the next case
  undoStack.commit();
  markDirty();
  onSelectionChange();
  message.value = `${caseLabel(kind)}${tracked ? ' (tracked)' : ''}`;
}

// ------------------------------------------------------------------ format painter

// Word's Format Painter: a click copies the character formatting where the caret is (or of the
// selection's first character) and paints it on the text selected next; a double-click keeps it on
// for several places, until Esc or another click on it. Selecting whole paragraphs paints the
// paragraph's style too. Ctrl+Shift+C and Ctrl+Shift+V copy and paint without the button.
const painter = reactive({ active: false, sticky: false, copied: false, rPr: null, block: null });

/** The text node whose formatting the selection (or the caret) has: the first selected, or beside the caret. */
function formatSourceNode(range) {
  const first = portions(range)[0];
  if (first) return first.node;
  let node = range.startContainer;
  if (node.nodeType === 3) return node;
  node = node.childNodes[range.startOffset - 1] || node.childNodes[range.startOffset] || node;
  const walker = document.createTreeWalker(closestBlock(node) || node, NodeFilter.SHOW_TEXT);
  return walker.nextNode();
}

/** Copies the formatting at the selection; `active`: and turns the painter on (`sticky`: until stopped). */
function startPainter(sticky = false, keyboard = false) {
  restoreSelection();
  const range = selectionInPage() || savedRange;
  const node = range && formatSourceNode(range);
  if (!node || !page.value.contains(node)) {
    message.value = 'Put the caret in the text whose formatting you want to copy';
    return;
  }
  Object.assign(painter, { copied: true, rPr: doc.rPrOf(node), block: blockKind(closestBlock(node)) });
  if (keyboard) {
    message.value = 'Formatting copied – select text and press Ctrl+Shift+V to paint it on';
    return;
  }
  Object.assign(painter, { active: true, sticky });
  message.value = sticky
    ? 'Format Painter: select (or click) each place to paint the formatting on – Esc when you’re done'
    : 'Format Painter: select (or click) the text to paint the formatting on';
  page.value.focus();
}

function stopPainterQuietly() {
  painter.active = false;
  painter.sticky = false;
}

function stopPainter() {
  painter.active = false;
  painter.sticky = false;
  message.value = 'Format Painter off';
}

function togglePainter() {
  if (painter.active) stopPainter();
  else startPainter(false);
}

/** The paragraph style select's value for a block, as state.block reads it. */
const blockKind = (block) => (block ? doc.blockValue(block.dataset.style) : null);

/** After selecting with the painter on: paints the formatting on the selection (or the word clicked). */
function onPageMouseup() {
  if (!painter.active) return;
  setTimeout(() => {
    if (!painter.active) return;
    let range = selectionInPage();
    if (!range) return;
    if (range.collapsed) {
      range = wordAt(range);
      if (!range) return;
      selectRange(range);
    }
    paintSelection();
    if (!painter.sticky) painter.active = false;
  });
}

/**
 * Paints the copied formatting on the selection: its character formatting replaces the text's (a
 * tracked formatting change while tracking), and selecting whole paragraphs gives them the copied
 * paragraph style too.
 */
function paintSelection() {
  if (!painter.copied) return;
  if (textLocked.value) return explainLock();
  restoreSelection();
  const range = selectionInPage();
  if (!range || range.collapsed) {
    message.value = 'Select the text to paint the formatting on';
    return;
  }
  const blocks = selectedBlocks(range);
  const whole = blocks.length && isAtBlockStart(range, blocks[0]) && atBlockEnd(range, blocks[blocks.length - 1]);
  if (whole && painter.block && blocks.some((b) => blockKind(b) !== painter.block)) applyBlockStyle(painter.block);
  const src = painter.rPr;
  const painted = formatSelection((r) => restyle(doc, r, (rPr) => {
    const change = elementChildren(rPr).find((c) => c.localName === 'rPrChange');
    for (const c of elementChildren(rPr)) if (c !== change) rPr.removeChild(c);
    for (const c of src ? elementChildren(src) : []) if (c.localName !== 'rPrChange') rPr.insertBefore(c.cloneNode(true), change || null);
  }));
  if (painted) message.value = painter.sticky ? 'Formatting painted – select the next place, or Esc when you’re done' : 'Formatting painted';
}

/** True if `range` ends at the end of `block`'s text. */
function atBlockEnd(range, block) {
  const rest = document.createRange();
  rest.setStart(range.endContainer, range.endOffset);
  rest.setEnd(block, block.childNodes.length);
  return !rest.toString().replace(/​/g, '');
}

// ------------------------------------------------------------------ the right-click menu

// Word's right-click menu: Cut, Copy, Paste and Paste as text, a link, a comment, the Paragraph
// dialog, a tracked change's Accept and Reject, and in a table its rows, columns and cells. Opened
// by a right-click, the Menu key or Shift+F10; Shift+right-click (and a long press on a touch
// screen) gives the browser's own menu, which has its spelling suggestions.
const ctxMenu = reactive({ open: false, x: 0, y: 0, selected: false, locked: false, table: false, multi: false, link: false, change: false });
const ctxMenuEl = ref(null);
let contextAnchor = null; // where the menu was, for a dialog opened from it

/** Where a dialog opened from a button goes: below the button, or where the right-click menu was. */
function anchorRect(button) {
  const at = contextAnchor;
  contextAnchor = null;
  return at || button?.getBoundingClientRect() || { left: 8, bottom: 60 };
}

let ctxOpenedAt = 0;

function onContextMenu(e) {
  // Opened from the keyboard a moment ago (see onKeydown): not again.
  if (ctxMenu.open && Date.now() - ctxOpenedAt < 500) {
    e.preventDefault();
    return;
  }
  if (e.shiftKey || e.pointerType === 'touch' || e.sourceCapabilities?.firesTouchEvents) return; // the browser's own
  if (!selectionInPage()) return;
  e.preventDefault();
  openContextMenu(e.clientX, e.clientY, e.target);
}

/** Opens the right-click menu at client (x, y), or at the caret (no place: from the keyboard). */
function openContextMenu(x = 0, y = 0, target = null) {
  const range = selectionInPage();
  if (!range) return;
  stopPainterQuietly();
  ctxOpenedAt = Date.now();
  if (!x && !y) {
    const r = range.getClientRects()[0] || closestBlock(range.startContainer)?.getBoundingClientRect();
    x = r ? r.left : 100;
    y = r ? r.bottom : 100;
  }
  const el = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  const table = !!el?.closest('td') && !!tableSelection();
  // A tracked change under the pointer (the caret can't go into someone else's): its Accept and
  // Reject act on it, as when it is clicked.
  const items = changeItems(page.value);
  const hit = target?.nodeType === 1 ? items.find((i) => i.el === target || i.el.contains(target)) : null;
  if (hit) selectedChange = { el: hit.el, kind: hit.kind };
  Object.assign(ctxMenu, {
    open: true,
    x,
    y,
    selected: !range.collapsed,
    locked: textLocked.value || !!readOnlySelection(),
    table,
    multi: table && !!tableSelection()?.multi,
    link: !!linkAt(range),
    change: !!hit || changesAt(range, items).length > 0,
  });
  nextTick(() => {
    keepInView(ctxMenu, ctxMenuEl.value);
    ctxMenuEl.value?.querySelector('button:not(:disabled)')?.focus();
  });
}

function closeContextMenu(refocusText = false) {
  ctxMenu.open = false;
  if (refocusText) refocus();
}

/** Runs a menu item's command on the text, `dialog`: one that opens where the menu was. */
function contextRun(command, dialog = false) {
  const at = { left: ctxMenu.x, bottom: ctxMenu.y, top: ctxMenu.y, right: ctxMenu.x };
  ctxMenu.open = false;
  refocus();
  if (dialog) contextAnchor = at;
  command();
  if (dialog) nextTick(() => { contextAnchor = null; });
}

function contextCopy() {
  ctxMenu.open = false;
  refocus();
  if (!document.execCommand('copy')) message.value = 'The browser didn’t copy: use Ctrl+C';
}

/** Cut: copies the selection, then deletes it as a key press would (tracked while tracking). */
function contextCut() {
  ctxMenu.open = false;
  refocus();
  const range = selectionInPage();
  if (!range || range.collapsed) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (refuseRemoving([range])) return;
  if (!document.execCommand('copy')) {
    message.value = 'The browser didn’t copy: use Ctrl+X';
    return;
  }
  undoStack.flush();
  const opts = trackedBy(range.startContainer);
  if (opts) deleteTracked(range, false, opts);
  else deleteSelection();
  onInput();
  undoStack.commit();
}

/**
 * Paste (or Paste as text) from the menu: the page reads the clipboard, which the browser may ask
 * to allow, and pastes it as Ctrl+V would (Ctrl+V needs no permission, so the message says so).
 */
async function contextPaste(asText) {
  ctxMenu.open = false;
  refocus();
  const saved = selectionInPage()?.cloneRange();
  const data = new DataTransfer();
  try {
    if (asText || !navigator.clipboard.read) {
      data.setData('text/plain', await navigator.clipboard.readText());
    } else {
      for (const item of await navigator.clipboard.read()) {
        for (const type of item.types) {
          const blob = await item.getType(type);
          if (type === 'text/html' || type === 'text/plain') data.setData(type, await blob.text());
          else if (type.startsWith('image/')) data.items.add(new File([blob], `pasted.${type.slice(6)}`, { type }));
        }
      }
    }
  } catch {
    message.value = `The browser didn’t let the page read the clipboard: use ${asText ? 'Ctrl+Shift+V' : 'Ctrl+V'} to paste`;
    return;
  }
  if (saved) {
    page.value.focus();
    selectRange(saved);
  }
  pasteFrom(data);
}

function contextOpenLink() {
  ctxMenu.open = false;
  const a = linkAt(selectionInPage() || savedRange);
  if (a && /^(https?:|mailto:)/i.test(a.getAttribute('href'))) window.open(a.href, '_blank', 'noopener');
  refocus();
}

/** Remove link, from the menu: the link at the caret becomes text. */
function removeLinkAtCaret() {
  const range = selectionInPage();
  linkTarget = range && linkAt(range);
  linkRange = range?.cloneRange() || null;
  if (linkTarget) removeLink();
}

// ------------------------------------------------------------------ insert symbol

const symbolOpen = ref(false);
let symbolRange = null; // where the symbol goes: the selection when the dialog opened

function openSymbols() {
  restoreSelection();
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text where the symbol should go';
    return;
  }
  symbolRange = range.cloneRange();
  symbolOpen.value = true;
}

function closeSymbols() {
  symbolOpen.value = false;
  refocusAt(symbolRange);
}

/** The symbol chosen, at the caret (or in place of the selection), as a key press would type it. */
function insertSymbol(char) {
  closeSymbols();
  typeText(char);
  message.value = `Symbol inserted: ${char.trim() ? char : nameOf(char)}`;
}

// ------------------------------------------------------------------ fonts, sizes, colours, character styles

// Only what the document already uses is offered: its theme and fonts, its sizes (and Word's
// usual ones), its theme and text colours and its own character styles. See doc.fontChoices etc.
const fontChoices = ref([]);
const sizeChoices = ref({ used: [], others: [] });
const colorChoices = ref({ theme: [], used: [] });
const characterStyles = ref([]);
const fontValue = computed(() => fontChoices.value.find((f) => f.name === state.font)?.value || '');
const sizeValue = computed(() => ([...sizeChoices.value.used, ...sizeChoices.value.others].includes(state.size) ? String(state.size) : ''));

function loadChoices() {
  fontChoices.value = doc.fontChoices();
  sizeChoices.value = doc.sizeChoices();
  colorChoices.value = doc.colorChoices();
  characterStyles.value = doc.characterStyles();
  fontColor.value ||= colorChoices.value.used[0] || colorChoices.value.theme.find((c) => c.theme === 'accent1') || null;
}

/** Changes the selected text's run properties with `mutate` (see editor/runs.js), tracked while tracking. */
function restyleSelection(mutate) {
  const range = selectionInPage() || savedRange;
  if (!range || range.collapsed) {
    message.value = 'Select the text first';
    page.value.focus();
    return;
  }
  formatSelection((r) => restyle(doc, r, mutate));
}

const applyFont = (value) => restyleSelection(setFont(value));
const applyFontSize = (pt) => restyleSelection(setSize(pt));

/** The size, in points, text node `node` shows (its run's, from its style), as the size box reads it. */
function shownSize(node) {
  const el = node.parentElement?.closest(':not(sup, sub)') || node.parentElement;
  return Math.round(parseFloat(getComputedStyle(el).fontSize) * 0.75 * 2) / 2;
}

/**
 * Grow Font / Shrink Font (Ctrl+Shift+> and <), as in Word: each run of the selection to the next
 * size in Word's steps, or by 1 pt (`byPoint`: Ctrl+] and Ctrl+[).
 */
function growFont(dir, byPoint = false) {
  restyleSelection(changeSize((pt) => (byPoint ? pt + dir : nextFontSize(pt, dir)), shownSize));
}
const applyCharacterStyle = (id) => restyleSelection(setCharacterStyle(id || null));

// The font colour last chosen (this document only): the button applies it.
const fontColor = ref(null);
const colorPalette = reactive({ open: false, x: 0, y: 0 });
const colorPaletteEl = ref(null);
const colorButton = ref(null);

function applyFontColor(color) {
  restyleSelection(setColor(color ? { hex: color.hex, theme: color.theme } : null));
}

/**
 * A palette keeps the text's selection while its swatches are clicked; its colour-code box and
 * picker take the focus as fields do (the selection is put back when a colour is applied).
 */
function keepSelection(e) {
  if (!e.target.closest('input')) e.preventDefault();
}

function toggleColorPalette() {
  if (colorPalette.open) return closeColorPalette();
  const rect = colorButton.value.getBoundingClientRect();
  Object.assign(colorPalette, { open: true, x: Math.max(8, Math.min(rect.left - 120, window.innerWidth - 220)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(colorPalette, colorPaletteEl.value);
    colorPaletteEl.value?.querySelector('.swatch.current, .swatch')?.focus();
  });
}

function closeColorPalette() {
  colorPalette.open = false;
  page.value.focus();
  restoreSelection();
}

function pickFontColor(color) {
  colorPalette.open = false;
  fontColor.value = color;
  refocus(); // the selection as it was, also after typing a colour code took the focus
  applyFontColor(color);
}

// ------------------------------------------------------------------ smart quotes and dashes

// On unless turned off, remembered in this browser, as Word's AutoFormat As You Type.
const AUTOFORMAT_KEY = 'intact.smartTyping';
const autoFormat = ref((() => { try { return localStorage.getItem(AUTOFORMAT_KEY) !== '0'; } catch { return true; } })());

function toggleAutoFormat() {
  autoFormat.value = !autoFormat.value;
  try { localStorage.setItem(AUTOFORMAT_KEY, autoFormat.value ? '1' : '0'); } catch { /* this session only */ }
  message.value = autoFormat.value ? 'Quotes, dashes, lists and links are formatted as you type' : 'Quotes, dashes, list numbers and addresses are typed as they are';
  page.value.focus();
}

/**
 * Typing " or ' gives a curly quote (in the text's language), "--" between words an em dash,
 * " - " an en dash and "..." an ellipsis, as Word's AutoFormat does. Typed through typeText, so it
 * is tracked while tracking. True if the key press was dealt with here.
 */
function smartTyping(e) {
  if (e.inputType === 'insertParagraph' && autoCorrectEnter(e)) return true;
  if (e.inputType === 'insertText' && e.data && [...e.data].length === 1 && autoCorrectTyped(e)) return true;
  if (autoFormat.value && e.inputType === 'insertParagraph' && autoLinkEnter(e)) return true;
  if (!autoFormat.value || e.inputType !== 'insertText' || !e.data || [...e.data].length !== 1) return false;
  const range = selectionInPage();
  if (!range || inPlainField() || range.startContainer.parentElement?.closest('.raw-inline, [data-raw]')) return false;
  const block = closestBlock(range.startContainer);
  if (!block) return false;
  const lead = document.createRange();
  lead.setStart(block, 0);
  lead.setEnd(range.startContainer, range.startOffset);
  const before = lead.toString();
  const lang = range.startContainer.parentElement?.closest('[lang]')?.lang || doc.styles.language() || 'en';
  if (range.collapsed && autoCorrectSymbol(e, range, before)) return true;
  if (e.data === ' ' && range.collapsed && autoLinkSpace(e, range, before)) return true;
  if (e.data === ' ' && range.collapsed && autoList(e, block, range, before)) return true;
  if (range.collapsed) {
    const swap = smartReplacement(before, e.data);
    const node = range.startContainer;
    if (swap && node.nodeType === 3 && range.startOffset >= swap.remove) {
      e.preventDefault();
      const sel = window.getSelection();
      sel.setBaseAndExtent(node, range.startOffset - swap.remove, node, range.startOffset);
      typeText(swap.text);
      return true;
    }
  }
  const quote = smartQuote(before, e.data, lang);
  if (!quote) return false;
  e.preventDefault();
  typeText(quote);
  return true;
}

/** Types `text` at the selection, as a key press would: tracked while tracking, never into read-only content. */
function typeText(text) {
  const range = selectionInPage();
  if (!range) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (!range.collapsed && refuseRemoving([range])) return;
  undoStack.flush();
  const opts = trackedBy(range.startContainer);
  if (opts) {
    if (!range.collapsed) deleteTracked(range, false, opts);
    if (typeTracked(text, opts)) onInput();
    else document.execCommand('insertText', false, text);
  } else {
    if (!range.collapsed) deleteSelection();
    if (typeUntracked(text)) onInput();
    else document.execCommand('insertText', false, text);
  }
}

/**
 * Word's AutoCorrect for symbols: "(c)" becomes ©, "-->" an arrow and so on (see autoCorrect), as
 * the last character is typed. That character is typed first, as an undo step of its own, so Ctrl+Z
 * straight after gives back what was typed, as in Word. True if the key press was dealt with here.
 */
function autoCorrectSymbol(e, range, before) {
  const hit = autoCorrect(before, e.data);
  const node = range.startContainer;
  const typedBefore = hit?.from.slice(0, -1);
  if (!hit || node.nodeType !== 3 || node.data.slice(Math.max(0, range.startOffset - typedBefore.length), range.startOffset) !== typedBefore) return false;
  e.preventDefault();
  typeText(e.data);
  undoStack.commit();
  const now = selectionInPage();
  const at = now?.startContainer;
  if (!now?.collapsed || at?.nodeType !== 3 || at.data.slice(Math.max(0, now.startOffset - hit.from.length), now.startOffset) !== hit.from) return true;
  const sel = window.getSelection();
  sel.setBaseAndExtent(at, now.startOffset - hit.from.length, at, now.startOffset);
  typeText(hit.to);
  return true;
}

/**
 * AutoFormat's links (Word's "Internet and network paths with hyperlinks"): a web or e-mail address
 * followed by a space or Enter becomes a link to it (see linkTrigger). The space is typed first, as
 * an undo step of its own, so Ctrl+Z straight after gives back the address as plain text, as in Word.
 * True if the key press was dealt with here.
 */
function autoLinkSpace(e, range, before) {
  if (!linkable(range, before)) return false;
  e.preventDefault();
  typeText(' ');
  undoStack.commit();
  makeTypedLink(1);
  return true;
}

/** Enter after an address: the address becomes a link (an undo step of its own), then the paragraph is split. */
function autoLinkEnter(e) {
  const range = selectionInPage();
  if (!range?.collapsed || inPlainFieldQuiet()) return false;
  const block = closestBlock(range.startContainer);
  if (!block) return false;
  const lead = document.createRange();
  lead.setStart(block, 0);
  lead.setEnd(range.startContainer, range.startOffset);
  if (!linkable(range, lead.toString())) return false;
  e.preventDefault();
  undoStack.flush();
  makeTypedLink(0);
  // Then Enter, as it would have been: tracked while tracking.
  const opts = trackedBy(selectionInPage()?.startContainer);
  if (opts) pendingSplit = opts;
  document.execCommand('insertParagraph'); // fires input
  return true;
}

/** Whether a link can be made of the address just before the caret (`before`: the paragraph's text before it). */
function linkable(range, before) {
  const trigger = linkTrigger(before);
  const node = range.startContainer;
  if (!trigger || node.nodeType !== 3) return false;
  // Typed in one piece of text, not already in a link, a field kept as it is or a plain-text field.
  const end = range.startOffset - trigger.after;
  if (end - trigger.text.length < 0 || node.data.slice(end - trigger.text.length, end) !== trigger.text) return false;
  if (node.parentElement.closest('a, .raw-inline, [data-raw], .field') || inPlainFieldQuiet()) return false;
  return true;
}

function inPlainFieldQuiet() {
  const range = selectionInPage();
  return !!(range && controlAt(range.startContainer, page.value)?.querySelector(':scope > [data-plain]'));
}

/**
 * Makes a link of the address that ends `skip` characters (plus its punctuation) before the caret,
 * as one undo step. Not tracked as such, as making text a link isn't (the text, typed while
 * tracking, is a tracked insertion, and the link goes with it).
 */
function makeTypedLink(skip) {
  const caret = selectionInPage();
  const node = caret?.startContainer;
  if (!caret?.collapsed || node?.nodeType !== 3) return;
  const lead = document.createRange();
  const block = closestBlock(node);
  lead.setStart(block, 0);
  lead.setEnd(node, caret.startOffset - skip);
  const trigger = linkTrigger(lead.toString());
  if (!trigger) return;
  const end = caret.startOffset - skip - trigger.after;
  const start = end - trigger.text.length;
  if (start < 0 || node.data.slice(start, end) !== trigger.text) return;
  undoStack.flush();
  const after = node.splitText(end); // the punctuation and what was typed after the address
  const text = start ? node.splitText(start) : node;
  const a = newLink(trigger.href, []);
  text.replaceWith(a);
  a.append(text);
  const sel = window.getSelection();
  sel.removeAllRanges();
  const r = document.createRange();
  r.setStart(after, Math.min(after.data.length, trigger.after + skip));
  r.collapse(true);
  sel.addRange(r);
  onInput();
  undoStack.commit();
  message.value = `Made a link to ${trigger.href.replace(/^mailto:/, '')} – Ctrl+Z takes it back`;
}

/**
 * AutoFormat's automatic lists: "1. ", "a) ", "(i) ", "* " or "- " typed at the start of a paragraph
 * makes it a list item of that kind (see listTrigger). The space is typed first, as an undo step of
 * its own, so Ctrl+Z gives the text back as typed, as in Word. A number that follows on from a
 * numbered list just above ("2. ") continues it; a bullet under a bulleted list joins it. True if
 * the key press was dealt with here.
 */
function autoList(e, block, range, before) {
  const trigger = listTrigger(before);
  if (!trigger || block.tagName !== 'P' || doc.listOf(block) || block.closest('.docx-hf-band, .textbox') || controlAt(block, page.value)) return false;
  const node = range.startContainer;
  if (node.nodeType !== 3 || node.data.slice(Math.max(0, range.startOffset - before.length), range.startOffset) !== before) return false;
  e.preventDefault();
  typeText(' ');
  undoStack.commit();
  // The typed text, with its space, just before the caret: taken out, as Word does.
  const now = selectionInPage();
  const at = now?.startContainer;
  const typed = `${before} `;
  if (!now?.collapsed || at?.nodeType !== 3 || at.data.slice(Math.max(0, now.startOffset - typed.length), now.startOffset) !== typed) return true;
  const text = document.createRange();
  text.setStart(at, now.startOffset - typed.length);
  text.setEnd(at, now.startOffset);
  const opts = trackedBy(at);
  selectRange(text);
  if (opts) deleteTracked(text, false, opts);
  else deleteSelection();
  const { kind, first } = trigger;
  withTrackedParagraphs([block], () => {
    const above = continuedList(block, trigger);
    doc.setList(block, above ? above.numId : doc.newList(kind, first), above ? above.ilvl : 0);
  });
  if (!block.textContent && !block.querySelector('br')) block.append(document.createElement('br'));
  if (!selectionInPage()) placeCaret(block, true);
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
  return true;
}

/**
 * The list right above `block` that a list typed as `trigger` continues, as in Word, or null: the
 * same bullet, or numbers in the same style whose next one is the number typed ("3." after "2.").
 */
function continuedList(block, { kind, first }) {
  const above = adjacentList(block, true, kind);
  const level = above && doc.listLevel(above.numId, above.ilvl);
  if (!level) return null;
  const dash = (t) => (['-', '–', '—'].includes(t) ? t : null);
  if (kind === 'bullet') return dash(level.text) === (first?.char ?? null) ? above : null;
  let prev = block.previousElementSibling;
  while (prev?.classList.contains('raw-hidden')) prev = prev.previousElementSibling;
  const follows = first.start > 0 && prev?.dataset.marker === first.text.replace('%1', formatNumber(first.start - 1, first.fmt));
  return level.fmt === first.fmt && level.text === first.text && follows ? above : null;
}

// ------------------------------------------------------------------ page numbers, footnotes, page setup

/** Puts `nodes` at the caret as typing would: tracked while tracking (as one insertion). */
function insertNodes(nodes, range) {
  const opts = trackedBy(range.startContainer);
  if (opts) insertTracked(nodes, opts);
  else insertAtCaret(nodes);
}

/** The header or footer (`which`) shown, in the page, or null. */
const shownPart = (which) => page.value.querySelector(`:scope > .docx-hf-band.${which} .docx-hf-part[data-variant="${hf[which]}"]`);

/**
 * A header or footer (`which`) for all pages, for a document without one: one empty paragraph in
 * Word's Header or Footer style, saved only while it is in the page (Undo takes it away). Returns
 * the part to type in.
 */
function addHeaderFooter(which) {
  const band = which === 'header' ? doc.newHeader() : doc.newFooter();
  refreshCss();
  if (which === 'header') page.value.prepend(band);
  else page.value.append(band);
  pageChanged();
  return band.querySelector('.docx-hf-part');
}

/**
 * Page number, as Word's Insert Page Number: Word's PAGE field at the caret in a header or footer.
 * With the caret in the text: into the footer if it is empty, else the caret goes to the footer to
 * say where (a footer's design is the document's); a document without a footer gets one, for all
 * pages, with the number centred. From the menu (`where`: Top or Bottom of Page, `which` 'header'
 * or 'footer', aligned `align`), the same in the header or footer chosen, aligned as chosen. Word
 * numbers the pages: here the field shows 1.
 */
function insertPageNumber(where = null) {
  if (textLocked.value) return explainLock();
  restoreSelection();
  let range = selectionInPage();
  const inPart = !where && range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('.docx-hf-part');
  const which = where?.which || 'footer';
  undoStack.flush();
  if (!inPart) {
    let host = shownPart(which);
    if (host && host.textContent.trim()) {
      placeCaret(host.lastElementChild || host);
      host.scrollIntoView({ block: 'nearest' });
      message.value = where
        ? `The ${which} has text already: click where the number goes in it, then Page number`
        : 'Click where the number goes in the footer (or header), then Page number again';
      return;
    }
    const align = where?.align || 'center';
    if (!host) {
      host = addHeaderFooter(which);
      const p = host.querySelector(BLOCK_SELECTOR);
      // A new footer is centred (as newFooter writes it), a new header isn't aligned.
      if (p && align !== (which === 'footer' ? 'center' : 'left')) p.style.textAlign = align;
    } else {
      const p = host.querySelector(BLOCK_SELECTOR);
      if (p && p.style.textAlign !== align) withTrackedParagraphs([p], () => { p.style.textAlign = align; });
    }
    const target = host.querySelector(BLOCK_SELECTOR) || host;
    if (!target.textContent) for (const br of target.querySelectorAll('br')) br.remove(); // an empty line's placeholder
    placeCaret(target);
    range = selectionInPage();
  } else {
    if (inField(range.startContainer)) {
      message.value = 'A page number can’t go inside a field';
      return;
    }
    const locked = readOnlySelection();
    if (locked) return explainReadOnly(locked, true);
    if (!range.collapsed && refuseRemoving([range])) return;
    if (!range.collapsed) {
      const opts = trackedBy(range.startContainer);
      if (opts) deleteTracked(range, false, opts);
      else deleteSelection();
      range = selectionInPage();
    }
  }
  if (!range) return;
  insertNodes(doc.newPageField(), range);
  onInput();
  undoStack.commit();
  measureFooter();
  message.value = inPart ? 'Page number added: Word numbers the pages' : `Page number added to the ${which}: Word numbers the pages`;
}

// Insert Page Number's positions (Top of Page, Bottom of Page), as Word's galleries offer them.
const PAGE_NUMBER_ALIGNS = [
  { value: 'left', label: 'Left' },
  { value: 'center', label: 'Centre' },
  { value: 'right', label: 'Right' },
];
const pageNumberMenu = reactive({ open: false, x: 0, y: 0 });
const pageNumberMenuEl = ref(null);
const pageNumberButton = ref(null);

function togglePageNumberMenu() {
  if (pageNumberMenu.open) return closePageNumberMenu();
  restoreSelection();
  const rect = pageNumberButton.value.getBoundingClientRect();
  Object.assign(pageNumberMenu, { open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 248)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(pageNumberMenu, pageNumberMenuEl.value);
    pageNumberMenuEl.value?.querySelector('button')?.focus();
  });
}

function closePageNumberMenu(refocusText = false) {
  pageNumberMenu.open = false;
  if (refocusText) refocus();
}

function pageNumberAt(which, align) {
  pageNumberMenu.open = false;
  page.value.focus();
  insertPageNumber({ which, align });
}

/**
 * Edit Header / Edit Footer, as Word's: the caret goes to the end of the header (or footer) shown;
 * a document without one gets one for all pages, empty, to type in (as the page number's).
 */
function editHeaderFooter(which) {
  pageNumberMenu.open = false;
  if (textLocked.value) return explainLock();
  page.value.focus();
  let host = shownPart(which);
  if (host) {
    placeCaret(host.lastElementChild || host);
    host.scrollIntoView({ block: 'nearest' });
    message.value = `Type in the ${which}: it shows on every page it is for`;
    return;
  }
  undoStack.flush();
  host = addHeaderFooter(which);
  placeCaret(host.querySelector(BLOCK_SELECTOR) || host, true);
  host.scrollIntoView({ block: 'nearest' });
  onInput();
  undoStack.commit();
  measureFooter();
  message.value = `A ${which} for every page: type its text (Undo takes it away)`;
}

// ------------------------------------------------------------------ numbering library

// Word's Numbering Library and Multilevel List gallery: its own numbering, and ready-made multilevel
// lists (numbering.js LIST_PRESETS), Indonesia's legal numbering first.
const LIST_LIBRARY = [
  { id: 'pasal', label: 'Pasal 1 → (1) → a. → 1.', hint: 'Articles of a regulation: Pasal, ayat, huruf, angka' },
  { id: 'bab', label: 'BAB I → Pasal 1 → (1) → a.', hint: 'Chapters, with Pasal numbered on through them' },
  { id: 'outline', label: '1. → 1.1. → 1.1.1.', hint: 'Numbered headings and paragraphs, as a report has them' },
  { id: 'number', label: '1. → a. → i.', hint: 'Word’s numbering' },
];
const listLibrary = reactive({ open: false, x: 0, y: 0 });
const listLibraryEl = ref(null);
const listLibraryButton = ref(null);

function toggleListLibrary() {
  if (listLibrary.open) return closeListLibrary();
  restoreSelection();
  const rect = listLibraryButton.value.getBoundingClientRect();
  Object.assign(listLibrary, { open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(listLibrary, listLibraryEl.value);
    listLibraryEl.value?.querySelector('button')?.focus();
  });
}

function closeListLibrary(refocusText = false) {
  listLibrary.open = false;
  if (refocusText) refocus();
}

/**
 * Makes the selected paragraphs a new list of a kind from the library: each keeps its level (a
 * paragraph not in a list takes the first), and Tab and Shift+Tab move it between Pasal, ayat and huruf.
 */
function applyListPreset(id) {
  listLibrary.open = false;
  page.value.focus();
  if (textLocked.value) return explainLock();
  let made = false;
  changeLists((blocks) => {
    if (!blocks.length) return false;
    const numId = id === 'number' ? doc.newList('number') : doc.newPresetList(id);
    blocks.forEach((b) => {
      const list = doc.listOf(b);
      doc.setList(b, numId, list && doc.hasListLevel(numId, list.ilvl) ? list.ilvl : 0);
    });
    made = true;
    return true;
  });
  if (made) message.value = t('Numbered {list}: Tab and Shift+Tab change the level', { list: t(LIST_LIBRARY.find((p) => p.id === id).label) });
}

// ------------------------------------------------------------------ sorting a table

// Word's Sort for a table: the rows below the heading row, by one column's text (as the document's
// language orders it), number (as it writes them) or date; empty cells last. One undoable step.
const sortDialog = reactive({ open: false, x: 0, y: 0, col: 0, type: 'text', desc: false, header: true, columns: [] });
const sortDialogEl = ref(null);
let sortTarget = null;

/** The cell of grid column `col` in row `tr` (a merged cell for each column it spans), or null. */
function cellInColumn(tr, col) {
  for (const td of tr.cells) {
    const pos = doc.cellPosition(td);
    if (pos && col >= pos.start && col < pos.start + pos.span) return td;
  }
  return null;
}

function refreshSortColumns() {
  const rows = sortTarget ? doc.tableRows(sortTarget.table) : [];
  const count = Math.max(0, ...rows.map((r) => [...r.cells].reduce((n, td) => n + (doc.cellPosition(td)?.span || 1), 0)));
  sortDialog.columns = Array.from({ length: count }, (_, c) => {
    const head = sortDialog.header && rows[0] ? cellInColumn(rows[0], c)?.textContent.trim() : '';
    return { value: c, label: head ? head.slice(0, 40) : t('Column {n}', { n: c + 1 }) };
  });
  if (sortDialog.col >= count) sortDialog.col = 0;
}

function openSortDialog() {
  tableMenu.open = false;
  const sel = tableSelection();
  if (!sel) {
    message.value = t('Click in a table first');
    return;
  }
  sortTarget = { table: sel.table };
  const rect = anchorRect(layoutButton.value);
  const rows = doc.tableRows(sel.table);
  // A heading row: a row Word repeats on each page, or a first row that doesn't hold numbers where the next does.
  const pos = doc.cellPosition(sel.cell);
  Object.assign(sortDialog, { open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)), y: rect.bottom + 4, col: pos?.start ?? 0, type: 'text', desc: false, header: rows.length > 1 });
  refreshSortColumns();
  // Numbers in the column: sort them as numbers.
  const seps = separators(doc.styles.language() || navigator.language);
  const values = rows.slice(sortDialog.header ? 1 : 0).map((r) => cellInColumn(r, sortDialog.col)?.textContent.trim()).filter(Boolean);
  if (values.length && values.every((v) => parseCellNumber(v, seps) != null)) sortDialog.type = 'number';
  nextTick(() => {
    keepInView(sortDialog, sortDialogEl.value);
    sortDialogEl.value?.querySelector('select')?.focus();
  });
}

function closeSortDialog() {
  sortDialog.open = false;
  refocus();
}

/** A date written in a cell ("17/08/1945", "17 Agustus 1945", "2026-11-20"), as a time, or null. */
function cellDate(text) {
  const s = text.trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
  m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/.exec(s);
  if (m) {
    const lang = doc.styles.language() || navigator.language || 'en';
    const us = /^en-US$/i.test(lang);
    const [d, mo] = us ? [+m[2], +m[1]] : [+m[1], +m[2]];
    const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return Date.UTC(y, mo - 1, d);
  }
  // Month names in the document's language or English.
  const lang = doc.styles.language() || navigator.language || 'en';
  for (const l of [lang, 'en']) {
    for (let mo = 0; mo < 12; mo++) {
      for (const width of ['long', 'short']) {
        const name = new Intl.DateTimeFormat(l, { month: width }).format(new Date(2001, mo, 1)).replace('.', '').toLowerCase();
        const re = new RegExp(`(\\d{1,2})\\s+${name}\\.?\\s+(\\d{4})|${name}\\.?\\s+(\\d{1,2}),?\\s+(\\d{4})`, 'i');
        m = re.exec(s);
        if (m) return m[1] ? Date.UTC(+m[2], mo, +m[1]) : Date.UTC(+m[4], mo, +m[3]);
      }
    }
  }
  const t2 = Date.parse(s);
  return Number.isNaN(t2) ? null : t2;
}

function sortTable() {
  const d = { ...sortDialog };
  sortDialog.open = false;
  refocus();
  const table = sortTarget?.table;
  if (!table?.isConnected) return;
  const why = doc.tableLocked(table);
  if (why) {
    message.value = why;
    return;
  }
  const locked = readOnlyAt(table, page.value);
  if (locked) return explainReadOnly(locked, true);
  if (tracking.value) {
    message.value = t('Sorting isn’t tracked: turn Track Changes off to sort the table, or move the rows by hand');
    return;
  }
  const rows = doc.tableRows(table);
  const body = rows.filter((r, i) => !(d.header && i === 0));
  if (body.some((r) => r.dataset.revRow || [...r.cells].some((td) => td.classList.contains('vmerge') || td.dataset.revCell || td.dataset.revMerge))) {
    message.value = t('Rows with cells merged down, or with tracked changes to their rows or cells, can’t be sorted: split those cells, or accept or reject the changes first');
    return;
  }
  if (body.length < 2) return;
  const lang = doc.styles.language() || navigator.language || 'en';
  const seps = separators(lang);
  const collator = new Intl.Collator(lang, { numeric: true, sensitivity: 'base' });
  const keyed = body.map((tr, i) => {
    const text = cellInColumn(tr, d.col)?.textContent.trim() || '';
    const key = d.type === 'number' ? parseCellNumber(text, seps) : d.type === 'date' ? cellDate(text) : text || null;
    return { tr, i, key };
  });
  const order = [...keyed].sort((a, b) => {
    if (a.key == null || b.key == null) return a.key == null && b.key == null ? a.i - b.i : a.key == null ? 1 : -1; // empty last
    const c = d.type === 'text' ? collator.compare(a.key, b.key) : a.key - b.key;
    return (d.desc ? -c : c) || a.i - b.i;
  });
  if (order.every((o, i) => o.i === i)) {
    message.value = t('The rows are already in that order');
    return;
  }
  undoStack.flush();
  // Moved as they are, in their new order after the heading row (rows are kept whole, with their cells' properties).
  const after = body[body.length - 1].nextElementSibling;
  const parent = body[0].parentElement;
  for (const { tr } of order) parent.insertBefore(tr, after);
  doc.renumber(page.value);
  onInput();
  undoStack.commit();
  markDirty();
  message.value = t('Sorted {n} rows by {column}, {order}', { n: body.length, column: d.columns[d.col]?.label || '', order: t(d.desc ? 'descending' : 'ascending') });
}

// ------------------------------------------------------------------ table formulas

const formulaDialog = reactive({ open: false, x: 0, y: 0, formula: '=SUM(ABOVE)', picture: '' });
const formulaDialogEl = ref(null);
let formulaRange = null;
const formulaLang = () => doc.styles.language() || navigator.language || 'en';
/** Word's number formats for a formula, written with the document's separators ("#.##0,00" in Indonesian). */
const formulaPictures = computed(() => {
  void formulaDialog.open;
  const { group, decimal } = separators(formulaLang());
  const pic = (s) => s.replace(/[,.]/g, (c) => (c === ',' ? group : decimal));
  const list = ['#,##0', '#,##0.00', '0%', '0.00%'].map(pic);
  return /^id\b/i.test(formulaLang()) ? [...list.slice(0, 2), pic('Rp #,##0'), pic('Rp #,##0.00'), ...list.slice(2)] : [...list.slice(0, 2), pic('$#,##0.00'), ...list.slice(2)];
});
const formulaCell = () => (formulaRange?.startContainer.nodeType === 1 ? formulaRange.startContainer : formulaRange?.startContainer.parentElement)?.closest('td');
const formulaPreview = computed(() => {
  if (!formulaDialog.open) return { text: '', error: false };
  try {
    const value = evaluateFormula(formulaDialog.formula, formulaCell(), formulaLang());
    return { text: t('Result: {value}', { value: formatResult(value, formulaDialog.picture, formulaLang()) }), error: false };
  } catch (e) {
    return { text: e.message === 'zero' ? t('Division by zero') : t('Word wouldn’t understand this formula: start with = and use SUM(ABOVE), B2*C2 and the like'), error: true };
  }
});

function openFormulaDialog() {
  tableMenu.open = false;
  restoreSelection();
  const range = selectionInPage() || savedRange;
  const td = range && cellOf(range.startContainer);
  if (!td || !page.value.contains(td)) {
    message.value = t('Click in the table cell the formula goes in');
    return;
  }
  formulaRange = range.cloneRange();
  formulaRange.collapse(false);
  // As Word's dialog: the numbers above, or else those to the left.
  let formula = '=SUM(ABOVE)';
  try {
    if (!evaluateFormula('=COUNT(ABOVE)', td, formulaLang()) && evaluateFormula('=COUNT(LEFT)', td, formulaLang())) formula = '=SUM(LEFT)';
  } catch { /* the default */ }
  const rect = anchorRect(layoutButton.value);
  Object.assign(formulaDialog, { open: true, formula, picture: '', x: Math.max(8, Math.min(rect.left, window.innerWidth - 340)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(formulaDialog, formulaDialogEl.value);
    formulaDialogEl.value?.querySelector('input')?.focus();
  });
}

function closeFormulaDialog() {
  formulaDialog.open = false;
  refocusAt(formulaRange);
}

/** Paste function: the function, with its brackets, at the end of the formula. */
function pasteFunction(e) {
  const f = e.target.value;
  e.target.value = '';
  if (!f) return;
  formulaDialog.formula = `${formulaDialog.formula.replace(/\s+$/, '') || '='}${f}()`;
  nextTick(() => {
    const input = formulaDialogEl.value?.querySelector('input');
    input?.focus();
    input?.setSelectionRange(input.value.length - 1, input.value.length - 1);
  });
}

/** Inserts the formula as Word's = field at the caret in the cell, its result worked out now. */
function insertFormula() {
  if (formulaPreview.value.error) return;
  const { formula, picture } = formulaDialog;
  formulaDialog.open = false;
  refocusAt(formulaRange);
  const range = selectionInPage();
  if (!range) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (inField(range.startContainer)) {
    message.value = t('A formula can’t go inside a field: put the caret after it');
    return;
  }
  const value = formatResult(evaluateFormula(formula, cellOf(range.startContainer), formulaLang()), picture, formulaLang());
  undoStack.flush();
  insertNodes(doc.newField(formulaInstr(formula, picture), value), range);
  onInput();
  undoStack.commit();
  markDirty();
  message.value = t('Formula {formula}: {value} – Update fields (F9) works it out again', { formula, value });
}

// ------------------------------------------------------------------ comparing documents

const compare = reactive({ result: null, other: '', blocks: [], otherXml: '', redline: null });
const compareInput = ref(null);

function startCompare() {
  compareInput.value.value = '';
  compareInput.value.click();
}

/** The page's paragraphs (not headers, footers or notes), with text, as the XML's are read. */
function pageParagraphs() {
  const blocks = [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((el) => !el.closest('.docx-hf-band, .docx-notes, .textbox') && !el.parentElement.closest(BLOCK_SELECTOR));
  const texts = blocks.map((el) => {
    const copy = el.cloneNode(true);
    for (const x of copy.querySelectorAll('[data-tc="del"], .raw-hidden, .note-ref')) x.remove();
    for (const x of copy.querySelectorAll('.raw-inline[data-raw]')) {
      if (['del', 'moveFrom'].includes(doc.raw[Number(x.dataset.raw)]?.localName)) x.remove();
    }
    return copy.textContent;
  });
  const kept = texts.map((tx, i) => [tx, blocks[i]]).filter(([tx]) => tx.replace(/[\s\u00a0]+/g, ''));
  return { texts: kept.map(([tx]) => tx), blocks: kept.map(([, b]) => b) };
}

/**
 * The page's paragraph for each of the XML's (`texts`), to go to a change: the same one where the
 * page has as many, else the next whose text is the same (without spaces), else the nearest.
 */
function blocksFor(texts) {
  const { texts: shown, blocks } = pageParagraphs();
  if (shown.length === texts.length) return blocks;
  const squash = (x) => x.replace(/[\s\u00a0\u2029⧉]+/g, '');
  const out = [];
  let j = 0;
  for (const text of texts) {
    const k = shown.findIndex((x, i) => i >= j && i < j + 8 && squash(x) === squash(text));
    if (k >= 0) j = k;
    out.push(blocks[Math.min(j, blocks.length - 1)]);
    if (k >= 0) j += 1;
  }
  return out;
}

async function onCompareFile(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  // Encrypted with a password: asked for (see the password dialog), then compared.
  if (sniffFile(bytes)?.kind === 'password') {
    openPasswordDialog('compare', { name: file.name, bytes });
    return;
  }
  await compareWith(file.name, bytes);
}

/**
 * Compares the other version (`bytes`, a .docx package, named `name`) with this document. Both are
 * read from their XML the same way (this one as Download would write it), so a document compared
 * with itself is the same.
 */
async function compareWith(name, bytes) {
  const file = { name };
  try {
    // A .doc, or a file encrypted by rights management: say which, as opening one does.
    const unreadable = sniffFile(bytes);
    if (unreadable) throw new Error(t('{name} can’t be compared: {why}', { name: file.name, why: unreadable.message }));
    const zip = await JSZip.loadAsync(bytes);
    const main = Object.keys(zip.files).find((f) => /^word\/document\d*\.xml$/i.test(f));
    if (!main) throw new Error(t('{name} isn’t a Word document (.docx)', { name: file.name }));
    const otherXml = await zip.file(main).async('string');
    const before = paragraphsOfXml(otherXml);
    await draftSettled();
    await withHostsUnlocked(() => doc.write(page.value));
    const after = paragraphsOfXml(await doc.zip.file(doc.mainPath).async('string'));
    compare.blocks = blocksFor(after);
    compare.other = file.name;
    compare.otherXml = otherXml;
    compare.result = compareParagraphs(before, after);
    // The redline is made now, while the result is read, so that opening it in a new tab follows
    // the click at once (a browser opens a tab only straight after a click).
    compare.redline = compare.result.changes.length ? makeRedline().catch((e) => e) : null;
  } catch (err) {
    message.value = /zip|central directory|End of data/i.test(err.message)
      ? t('{name} can’t be compared: it isn’t a .docx (an older .doc can’t be read here)', { name: file.name })
      : err.message;
  }
}

function closeCompare() {
  compare.result = null;
  compare.redline = null;
  refocus();
}

/** Goes to a change in this document: its paragraph (or, for text taken out, where it was). */
function goToChange(change) {
  const block = compare.blocks[Math.min(change.at, compare.blocks.length - 1)];
  compare.result = null;
  compare.redline = null;
  if (!block?.isConnected) return;
  page.value.focus();
  placeCaret(block, change.at < compare.blocks.length);
  block.scrollIntoView({ block: 'center' });
  message.value = change.kind === 'removed' ? t('Text taken out was here') : change.kind === 'added' ? t('Added in this document') : t('Changed in this document');
}

/**
 * The redline, as Word's Compare makes one: a new document, this one as Download would write it,
 * with the changes since the other version as tracked changes in your name (see docx/redline.js).
 * Encrypted with this document's password, if it has one. { blob, name, accepted }.
 */
async function makeRedline() {
  await draftSettled();
  const saved = await withHostsUnlocked(() => doc.save(page.value, { plain: true }));
  const { zip, accepted } = await redlinePackage(JSZip, await saved.arrayBuffer(), doc.mainPath, compare.otherXml, { author: author.value, date: new Date() });
  const mime = doc.format.mime;
  const plain = await zip.generateAsync({ type: 'blob', mimeType: mime, compression: 'DEFLATE' });
  const blob = await doc.sealed(plain, true);
  const base = baseName.value.trim() || 'document';
  return { blob: new Blob([blob], { type: mime }), name: `${base} (redline).${fileExt.value}`, accepted };
}

/** The redline: opened in a new tab beside this document (`how` 'open'), or downloaded. */
async function sendRedline(how) {
  const made = await compare.redline;
  if (!made || made instanceof Error) {
    message.value = `The redline couldn’t be made: ${made?.message || 'nothing changed'}`;
    return;
  }
  const accepted = made.accepted ? ` – the ${count(made.accepted, 'tracked change')} already in this document ${made.accepted === 1 ? 'is' : 'are'} accepted in it, as Word’s Compare does` : '';
  const locked = passwordProtected.value ? ', encrypted with this document’s password' : '';
  if (how === 'open') {
    const result = openFileInNewTab(new File([made.blob], made.name, { type: made.blob.type }));
    message.value = {
      opened: `The redline is opening in a new tab${locked}${accepted}`,
      blocked: 'The browser didn’t open a new tab: allow pop-ups for this site, or download the redline',
      unsupported: `A new tab is open: download the redline and choose it there${locked}`,
    }[result];
    if (result === 'opened') closeCompare();
    return;
  }
  downloadBlob(made.blob, made.name);
  message.value = `Redline downloaded: ${made.name}${locked}${accepted}`;
  closeCompare();
}

// ------------------------------------------------------------------ password

// Word's Encrypt with Password (File › Info › Protect Document): a document opened with a password
// downloads encrypted with it (see DocxDocument.sealed and encryption.js); any document can be given
// one, or have it changed or taken off. The password is used on this device, to make the key that
// encrypts the file, and isn't kept. The other version to compare may be encrypted too.
const passwordProtected = ref(!!doc.encryption);
const passwordAtOpen = !!doc.encryption;
const pwDialog = reactive({ open: false, mode: 'set', value: '', again: '', error: '', busy: false, file: null });
const pwInput = ref(null);

async function openPasswordDialog(mode = 'set', file = null) {
  closeChooser();
  Object.assign(pwDialog, { open: true, mode, value: '', again: '', error: '', busy: false, file });
  await nextTick();
  pwInput.value?.focus();
}

function closePasswordDialog() {
  Object.assign(pwDialog, { open: false, value: '', again: '', file: null });
  refocus();
}

async function confirmPasswordDialog() {
  if (pwDialog.busy) return;
  const value = pwDialog.value;
  if (!value) {
    pwDialog.error = t('Type a password.');
    return;
  }
  if (pwDialog.mode === 'compare') {
    pwDialog.busy = true;
    try {
      const { zip } = await decrypt(pwDialog.file.bytes, value);
      const { name } = pwDialog.file;
      Object.assign(pwDialog, { open: false, value: '', busy: false, file: null });
      await compareWith(name, zip);
    } catch (e) {
      pwDialog.busy = false;
      if (e.reason === 'wrong') {
        pwDialog.error = t('That password isn’t right. Passwords are case-sensitive: check Caps Lock.');
        await nextTick();
        pwInput.value?.select();
      } else {
        const name = pwDialog.file.name;
        closePasswordDialog();
        message.value = t('{name} can’t be compared: {why}', { name, why: e.message });
      }
    }
    return;
  }
  if (value !== pwDialog.again) {
    pwDialog.error = t('The two passwords aren’t the same. Type the same password in both.');
    return;
  }
  pwDialog.busy = true;
  try {
    const had = passwordProtected.value;
    await doc.setPassword(value);
    passwordProtected.value = true;
    Object.assign(pwDialog, { open: false, value: '', again: '', busy: false });
    markDirty();
    message.value = had ? t('Password changed: the downloaded file opens only with the new one') : t('Protected with a password: the downloaded file opens only with it');
    refocus();
  } catch (e) {
    pwDialog.busy = false;
    pwDialog.error = e.message;
  }
}

async function removePassword() {
  await doc.setPassword(null);
  passwordProtected.value = false;
  Object.assign(pwDialog, { open: false, value: '', again: '', busy: false });
  markDirty();
  message.value = t('Password removed: the downloaded file opens without one');
  refocus();
}

/** What the review says of the password: { kept | added | changed | removed }, or null. */
function passwordReview() {
  if (!passwordAtOpen && !passwordProtected.value) return null;
  if (!passwordProtected.value) return { removed: true };
  if (!passwordAtOpen) return { added: true };
  return doc.encryptionChanged ? { changed: true } : { kept: true };
}

// ------------------------------------------------------------------ page number format

// Word's page number formats (Format Page Numbers' Number format), as w:pgNumType's w:fmt.
const PAGE_NUMBER_FORMATS = [
  { value: 'decimal', label: '1, 2, 3, …' },
  { value: 'numberInDash', label: '- 1 -, - 2 -, - 3 -, …' },
  { value: 'lowerLetter', label: 'a, b, c, …' },
  { value: 'upperLetter', label: 'A, B, C, …' },
  { value: 'lowerRoman', label: 'i, ii, iii, …' },
  { value: 'upperRoman', label: 'I, II, III, …' },
];
const numDialog = reactive({ open: false, x: 0, y: 0, scope: 'whole', sections: 1, sectionNo: 1, canSplit: false, fmt: 'decimal', restart: false, start: 1 });
const numDialogEl = ref(null);
let numInitial = null;

/** "i, ii, iii" for a format, as the dialog shows it. */
const formatLabel = (fmt) => (PAGE_NUMBER_FORMATS.find((f) => f.value === fmt) || PAGE_NUMBER_FORMATS[0]).label.replace(/, …$/, '');

/**
 * Format Page Numbers, as Word's (Insert > Page Number > Format Page Numbers): the number format
 * and the start of the section the caret is in; or the whole document's; or a new section's from
 * the caret on (a section break before its paragraph), for a report's front matter in i, ii, iii
 * and its chapters from 1.
 */
function openPageNumberFormat() {
  pageNumberMenu.open = false;
  if (textLocked.value) return explainLock();
  restoreSelection();
  const range = selectionInPage() || savedRange;
  pageRange = range && !noteAt(range.startContainer) ? range.cloneRange() : null;
  const blocks = pageRange ? blocksIn(pageRange) : [];
  const ends = sectionEnds();
  const end = blocks[0] ? sectionEndOf(blocks[0]) : null;
  const now = end ? doc.sectionSetup(end) : doc.pageSetupNow();
  const all = topBlocks();
  const at = blocks[0] ? all.indexOf(blocks[0]) : -1;
  numInitial = { fmt: now.numFmt || 'decimal', restart: now.numStart != null, start: now.numStart ?? 1 };
  const rect = pageNumberButton.value.getBoundingClientRect();
  Object.assign(numDialog, numInitial, {
    open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)), y: rect.bottom + 4,
    sections: ends.length + 1, sectionNo: end ? ends.indexOf(end) + 1 : ends.length + 1,
    scope: ends.length ? 'section' : 'whole',
    // A new section from the caret: only where there is text before it in its section.
    canSplit: at > 0 && all[at - 1].dataset.sectionEnd == null,
  });
  nextTick(() => {
    keepInView(numDialog, numDialogEl.value);
    numDialogEl.value?.querySelector('select')?.focus();
  });
}

function closePageNumberFormat() {
  numDialog.open = false;
  refocus();
}

/** Applies Format Page Numbers' choices to the section(s), as one undoable step. */
function applyPageNumberFormat() {
  const d = { ...numDialog };
  const i = numInitial;
  numDialog.open = false;
  refocus();
  const start = d.restart ? Math.max(0, Math.round(Number(d.start) || 0)) : null;
  const change = {};
  if (d.fmt !== i.fmt) change.numFmt = d.fmt;
  if (start !== (i.restart ? i.start : null)) change.numStart = start;
  // A new section from the caret is a change of its own, even with the numbering as it was.
  if (!Object.keys(change).length && d.scope !== 'forward') return;
  if (textLocked.value) return explainLock();
  undoStack.flush();
  let where;
  if (d.scope === 'whole') {
    // The format for every section; the start for the first (the others keep where theirs start).
    const ends = [...sectionEnds(), null];
    ends.forEach((end, k) => {
      const set = k === 0 ? { ...change } : {};
      if (change.numFmt !== undefined) set.numFmt = change.numFmt;
      if (!Object.keys(set).length) return;
      if (end) doc.setSectionSetup(end, set);
      else doc.setPageSetup({ ...doc.pageSetup, ...set });
    });
    where = t('for the whole document');
  } else {
    where = applySetupTo(d.scope, { numFmt: d.fmt, numStart: start });
  }
  pageChanged();
  undoStack.commit();
  markDirty();
  message.value = t('Page numbers: {format}{start} {where}', { format: formatLabel(d.fmt), start: start != null ? t(', starting at {n}', { n: start }) : '', where });
}

// The page as it was right after fields were updated here (for the review: "updated here").
let fieldsUpdatedAt = null;

/**
 * Update fields (F9), as Word's for the whole text: caption numbers, note numbers, cross-references
 * and tables of contents (see docx/fieldUpdates.js), as one undoable step. Page numbers are Word's:
 * when a table of contents was updated, the file asks Word to update fields when it opens it.
 */
function updateAllFields() {
  if (textLocked.value) return explainLock();
  undoStack.flush();
  const before = page.value.innerHTML;
  const t = updateFields(doc, page.value);
  refreshCss(); // a TOC level's style may have been added
  const count = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;
  const done = [
    t.toc && (t.toc === 1 ? `the table of contents (${count(t.entries, 'entry', 'entries')})` : `${t.toc} tables of contents`),
    t.ref && count(t.ref, 'cross-reference'),
    t.seq && count(t.seq, 'caption number'),
    t.formula && count(t.formula, 'formula', 'formulas'),
    t.noteref && count(t.noteref, 'note number'),
  ].filter(Boolean);
  const notes = [];
  if (t.missing) notes.push(`${count(t.missing, 'cross-reference')} to text no longer in the document ${t.missing === 1 ? 'was' : 'were'} left as ${t.missing === 1 ? 'it was' : 'they were'}`);
  if (t.skipped) notes.push(`${t.skipped} was left as it was`);
  if (t.toc) {
    doc.updateFields = true;
    notes.push('page numbers are Word’s: the file asks Word to update them when it opens it');
  }
  if (page.value.innerHTML !== before) {
    doc.renumber(page.value);
    onInput();
    undoStack.commit();
  }
  fieldsUpdatedAt = page.value.innerHTML;
  const what = done.length ? `Updated ${done.join(', ').replace(/, ([^,]*)$/, ' and $1')}` : 'The fields are up to date';
  message.value = [what, ...notes].join('; ');
}

/**
 * Insert Footnote (Ctrl+Alt+F) and Insert Endnote (Ctrl+Alt+D), as in Word: a reference at the
 * caret (after the selection), and its note below the text, among the others in the order the text
 * refers to them, with the caret in it to type the note.
 */
function insertNote(kind = 'footnote') {
  const name = kind === 'footnote' ? 'A footnote' : 'An endnote';
  if (textLocked.value) return explainLock();
  restoreSelection();
  let range = selectionInPage();
  if (!range) return;
  if (inField(range.startContainer)) {
    message.value = `${name} can’t go inside a field`;
    return;
  }
  if (refuseInNote(range, name)) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (!range.collapsed) {
    range.collapse(false);
    selectRange(range);
  }
  undoStack.flush();
  const { ref, note } = doc.newNote(kind);
  refreshCss();
  insertNodes([ref], range);
  let section = page.value.querySelector(`:scope > .docx-notes[data-notes="${kind}"]`);
  if (!section) {
    section = doc.notesSection(kind);
    const before = kind === 'footnote' ? ':scope > .docx-notes[data-notes="endnote"], :scope > .docx-hf-band.footer' : ':scope > .docx-hf-band.footer';
    page.value.insertBefore(section, page.value.querySelector(before));
  }
  // Before the note of the next reference in the text that has one.
  const refs = [...page.value.querySelectorAll(`.note-ref[data-note-ref="${kind}"]`)].filter((r) => !r.closest('.docx-notes'));
  const notes = new Map([...section.querySelectorAll(':scope > .docx-note')].map((el) => [el.dataset.noteId, el]));
  const next = refs.slice(refs.indexOf(ref) + 1).map((r) => notes.get(r.dataset.noteId)).find(Boolean);
  section.insertBefore(note, next || null);
  doc.renumberNotes(page.value);
  placeCaret(note.querySelector(BLOCK_SELECTOR) || note);
  note.scrollIntoView({ block: 'nearest' });
  onInput();
  undoStack.commit();
  message.value = `${kind === 'footnote' ? 'Footnote' : 'Endnote'} ${ref.dataset.n} added: type it below the text`;
}
const insertFootnote = () => insertNote('footnote');
const insertEndnote = () => insertNote('endnote');

// Page setup, for the whole document: paper size, orientation and margins (cm or inches), as Word's.
const PAPERS = [
  { name: 'A4', label: '21 × 29.7 cm', w: 11906, h: 16838 },
  { name: 'Letter', label: '8.5 × 11 in', w: 12240, h: 15840 },
  { name: 'Legal', label: '8.5 × 14 in', w: 12240, h: 20160 },
  { name: 'A5', label: '14.8 × 21 cm', w: 8391, h: 11906 },
  { name: 'A3', label: '29.7 × 42 cm', w: 16838, h: 23811 },
  { name: 'Executive', label: '7.25 × 10.5 in', w: 10440, h: 15120 },
  // Indonesia's and much of South-East Asia's office paper: F4 (215 × 330 mm) and Folio (8½ × 13 in),
  // as Word's paper lists name them where printers offer them.
  { name: 'F4', label: '21.5 × 33 cm', w: 12189, h: 18709 },
  { name: 'Folio', label: '8.5 × 13 in', w: 12240, h: 18720 },
];
// A custom paper size's limits, in twips, as Word's Page Setup has them: 0.1 in to 22 in.
const PAPER_MIN = 144;
const PAPER_MAX = 31680;
// Word's margin presets, in twips: 1 inch, ½ inch, 1 and ¾ inch, 1 and 2 inches.
const MARGIN_PRESETS = [
  { name: 'Normal', top: 1440, bottom: 1440, left: 1440, right: 1440 },
  { name: 'Narrow', top: 720, bottom: 720, left: 720, right: 720 },
  { name: 'Moderate', top: 1440, bottom: 1440, left: 1080, right: 1080 },
  { name: 'Wide', top: 1440, bottom: 1440, left: 2880, right: 2880 },
];
const MARGINS = ['top', 'bottom', 'left', 'right'];
const pageDialog = reactive({
  open: false, x: 0, y: 0, unit: 'cm', paper: 'A4', orient: 'portrait', width: 0, height: 0, top: 0, bottom: 0, left: 0, right: 0, mark: '', markText: '', markLayout: 'diagonal',
  scope: 'whole', sections: 1, sectionNo: 1, selected: false, titlePg: false, evenAndOdd: false,
  cols: 1, colSpace: 0, colSep: false, colUnequal: false,
});
let pageRange = null; // the text Page setup applies to, for "This section", "This point forward" and "Selected text"
// Word's Watermark gallery's texts (Design > Watermark).
const WATERMARKS = ['DRAFT', 'CONFIDENTIAL', 'DO NOT COPY', 'SAMPLE', 'URGENT', 'ASAP'];
const pageDialogEl = ref(null);
const pageSetupButton = ref(null);
let pageInitial = null;
let pageTwips = null; // the margins as they were, exactly, for switching units
const marginPreset = computed(() => MARGIN_PRESETS.find((m) => MARGINS.every((k) => Math.abs(Number(pageDialog[k]) - fromTwips(m[k], pageDialog.unit)) < 0.01))?.name || 'custom');

/** Shows a change to the page's size, margins or bands: the sheet, the print page, what is drawn over it. */
function pageChanged() {
  pageTick.value++;
  refreshPrintPage();
  refreshSections();
  nextTick(() => {
    keepVariantsShown();
    measureFooter();
    scheduleMarks();
    if (hasLineNumbers) scheduleLineNumbers();
  });
}

function openPageSetup() {
  if (pageDialog.open) return closePageSetup();
  restoreSelection();
  const range = selectionInPage() || savedRange;
  pageRange = range && !noteAt(range.startContainer) ? range.cloneRange() : null;
  const blocks = pageRange ? blocksIn(pageRange) : [];
  const ends = sectionEnds();
  const end = blocks[0] ? sectionEndOf(blocks[0]) : null;
  const now = end ? doc.sectionSetup(end) : doc.pageSetupNow();
  const [short, long] = [Math.min(now.w, now.h), Math.max(now.w, now.h)];
  const paper = PAPERS.find((p) => Math.abs(p.w - short) <= 20 && Math.abs(p.h - long) <= 20);
  const unit = unitFor(now.w, now.h);
  const cm = (twips) => fromTwips(twips, unit);
  pageTwips = { width: now.w, height: now.h, top: now.top, bottom: now.bottom, left: now.left, right: now.right, colSpace: now.colSpace };
  pageInitial = {
    paper: paper ? paper.name : 'custom', orient: now.orient === 'landscape' || now.w > now.h ? 'landscape' : 'portrait',
    width: cm(now.w), height: cm(now.h),
    top: cm(now.top), bottom: cm(now.bottom), left: cm(now.left), right: cm(now.right), w: now.w, h: now.h,
    cols: now.cols, colSpace: cm(now.colSpace), colSep: now.colSep, colUnequal: now.colUnequal,
    ...watermarkNow(),
    ...doc.headerFooterFlags(),
  };
  const rect = pageSetupButton.value.getBoundingClientRect();
  Object.assign(pageDialog, pageInitial, {
    unit, open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)), y: rect.bottom + 4,
    sections: ends.length + 1, sectionNo: end ? ends.indexOf(end) + 1 : ends.length + 1,
    scope: ends.length ? 'section' : 'whole', selected: !!pageRange && !pageRange.collapsed && blocks.length > 0,
  });
  nextTick(() => {
    keepInView(pageDialog, pageDialogEl.value);
    pageDialogEl.value?.querySelector('select')?.focus();
  });
}

function closePageSetup() {
  pageDialog.open = false;
  refocus();
}

function pickMargins(name) {
  const m = MARGIN_PRESETS.find((x) => x.name === name);
  if (m) for (const k of MARGINS) pageDialog[k] = fromTwips(m[k], pageDialog.unit);
}

/** The Page setup dialog's margins in another unit (remembered for the next dialogs). */
function setPageUnit(unit) {
  saveUnit(unit);
  switchUnit(pageDialog, pageInitial, pageTwips, [...MARGINS, 'width', 'height', 'colSpace'], unit);
}

/** The numbers of columns offered: one to three, as Word's Columns presets, and the section's own if more. */
const columnChoices = () => [...new Set([1, 2, 3, pageInitial?.cols || 1])].sort((a, b) => a - b);

/** Whether only the columns are changed in Page setup (not the paper or margins). */
function onlyColumnsChanged() {
  const d = pageDialog;
  const i = pageInitial;
  if (!i) return false;
  const paperSame = d.paper === i.paper && d.orient === i.orient && (d.paper !== 'custom' || (Number(d.width) === i.width && Number(d.height) === i.height));
  return paperSame && MARGINS.every((k) => Number(d[k]) === i[k]) && columnsChanged(d, i);
}

/** Whether the dialog's columns differ from the section's. */
const columnsChanged = (d, i) => d.cols !== i.cols || (d.cols > 1 && (Number(d.colSpace) !== i.colSpace || d.colSep !== i.colSep));

/** What "This point forward" or "Selected text" adds: continuous breaks for columns alone, as Word's Columns does; else new pages. */
function pageBreakNote() {
  const continuous = onlyColumnsChanged();
  if (pageDialog.scope === 'forward') {
    return continuous
      ? t('A continuous section break goes before the paragraph with the caret: from there on, the text is set in the columns below, on the same page.')
      : t('A section break (next page) goes before the paragraph with the caret: from there on, the pages are set as below.');
  }
  return continuous
    ? t('Continuous section breaks go before and after the selected paragraphs: they are set in the columns below, on the same page as the text around them.')
    : t('Section breaks go before and after the selected paragraphs: their pages are set as below, as for a wide table on a landscape page.');
}

/**
 * Puts the page's sections in columns as they are set now: each section's blocks in a wrapper
 * (see DocxDocument.renderBody) where it has more than one column, none where it has one. The
 * wrappers are only for the page; the caret and selection stay where they were.
 */
function rewrapColumns() {
  const root = page.value;
  if (!root) return;
  const sel = window.getSelection();
  const at = sel.rangeCount && root.contains(sel.anchorNode) ? [sel.anchorNode, sel.anchorOffset, sel.focusNode, sel.focusOffset] : null;
  for (const w of [...root.querySelectorAll(':scope > .docx-columns')]) w.replaceWith(...w.childNodes);
  let group = [];
  const close = (end) => {
    const cols = doc.sectionColumns(end);
    if (cols && group.length) {
      const div = columnsWrapper(cols);
      group[0].before(div);
      div.append(...group);
    }
    group = [];
  };
  for (const block of topBlocks()) {
    group.push(block);
    if (block.dataset.sectionEnd != null) close(block);
  }
  close(null);
  if (at && at.every((x) => x != null) && root.contains(at[0]) && root.contains(at[2])) {
    try { sel.setBaseAndExtent(...at); } catch { /* the caret stays where the browser put it */ }
  }
}

/** Applies what was changed in the Page setup dialog to every section, as one undoable step. */
function applyPageSetup() {
  const d = { ...pageDialog };
  const i = pageInitial;
  pageDialog.open = false;
  const num = (v) => (Number.isFinite(Number(v)) && v !== '' ? Math.max(0, Number(v)) : 0);
  const change = {};
  // A custom size: as typed (a length left as it was keeps its exact value), turned to the orientation.
  const custom = d.paper === 'custom';
  const sized = custom && (num(d.width) !== i.width || num(d.height) !== i.height);
  if (d.paper !== i.paper || d.orient !== i.orient || sized) {
    const paper = PAPERS.find((p) => p.name === d.paper);
    const typed = (k, exact) => (num(d[k]) === i[k] ? exact : toTwips(num(d[k]), d.unit));
    const [a, b] = paper ? [paper.w, paper.h] : [typed('width', i.w), typed('height', i.h)];
    if (custom && [a, b].some((v) => v < PAPER_MIN || v > PAPER_MAX)) {
      message.value = t('The paper’s width and height must be between {min} and {max}: nothing was changed', d.unit === 'in' ? { min: '0.1 in', max: '22 in' } : { min: '0.25 cm', max: '55.88 cm' });
      refocus();
      return;
    }
    const [short, long] = [Math.min(a, b), Math.max(a, b)];
    const landscape = d.orient === 'landscape';
    Object.assign(change, { w: landscape ? long : short, h: landscape ? short : long, orient: d.orient });
  }
  for (const k of MARGINS) if (num(d[k]) !== i[k]) change[k] = toTwips(num(d[k]), d.unit);
  // Columns, as Word's Columns dialog: how many, the space between them and a line between them,
  // all written together (equal columns), with a space left as it was kept exactly.
  const columnsOnly = onlyColumnsChanged();
  if (columnsChanged(d, i)) {
    change.cols = d.cols;
    change.colSpace = num(d.colSpace) === i.colSpace ? pageTwips.colSpace : toTwips(num(d.colSpace), d.unit);
    change.colSep = d.cols > 1 && d.colSep;
  }
  const markText = (d.mark === 'custom' ? d.markText : d.mark).trim();
  const markInitial = (i.mark === 'custom' ? i.markText : i.mark).trim();
  const markChanged = markText !== markInitial || (!!markText && d.markLayout !== i.markLayout);
  const hfChanged = d.titlePg !== i.titlePg || d.evenAndOdd !== i.evenAndOdd;
  refocus();
  const paged = Object.keys(change).length > 0;
  if (!paged && !markChanged && !hfChanged) return;
  if (paged) {
    const width = change.w ?? i.w;
    const left = change.left ?? pageTwips.left;
    const right = change.right ?? pageTwips.right;
    if (width - left - right < 1440) {
      message.value = t('The margins leave less than {room} for the text: nothing was changed', { room: d.unit === 'in' ? '1 in' : '2.5 cm' });
      return;
    }
    // Each column at least half an inch wide, as Word's Columns dialog allows.
    const cols = change.cols ?? i.cols;
    const space = change.colSpace ?? pageTwips.colSpace;
    if (cols > 1 && (width - left - right - (cols - 1) * space) / cols < 720) {
      message.value = t('{n} columns with that spacing leave less than {room} for each: nothing was changed', { n: cols, room: d.unit === 'in' ? '0.5 in' : '1.27 cm' });
      return;
    }
  }
  if ((markChanged || hfChanged || (paged && (d.scope !== 'whole' || sectionEnds().length))) && textLocked.value) return explainLock();
  undoStack.flush();
  const where = paged ? applySetupTo(d.scope, change, columnsOnly ? 'continuous' : 'nextPage') : null;
  if (paged) rewrapColumns(); // new section breaks, or columns changed: the page's sections as they are now
  if (markChanged) setWatermark(markText, d.markLayout !== 'horizontal');
  if (hfChanged) applyHeaderFooterFlags(d.titlePg, d.evenAndOdd);
  pageChanged();
  undoStack.commit();
  markDirty();
  const said = [];
  if (paged) said.push(t('Page setup changed {where}', { where }));
  if (hfChanged) said.push([d.titlePg !== i.titlePg && `${d.titlePg ? 'A' : 'No'} different first page`, d.evenAndOdd !== i.evenAndOdd && `odd and even pages ${d.evenAndOdd ? 'different' : 'alike'}`].filter(Boolean).join(', ') + (d.titlePg && !i.titlePg || d.evenAndOdd && !i.evenAndOdd ? ': their header and footer are in the tabs above and below the page' : ''));
  if (markChanged) said.push(markText ? `Watermark: ${markText}` : 'Watermark removed');
  message.value = said.join(' · ');
}

// ------------------------------------------------------------------ sections

// The body's blocks, top-level: the page's children (and those of a section set in columns), not
// the header, footer or notes. A section ends at a paragraph holding a section break
// (data-section-end), and the final section at the end of the body.
function topBlocks() {
  const out = [];
  for (const c of page.value.children) {
    if (c.matches('.docx-hf-band, .docx-notes')) continue;
    if (c.matches('.docx-columns')) out.push(...c.children);
    else out.push(c);
  }
  return out;
}

/** The top-level blocks a range is in (the one with the caret, for a caret). */
function blocksIn(range) {
  const tops = topBlocks();
  const of = (node) => tops.find((b) => b === node || b.contains(node));
  if (range.collapsed) return [of(range.startContainer)].filter(Boolean);
  const first = of(range.startContainer);
  const last = of(range.endContainer);
  if (!first || !last) return [];
  return tops.slice(tops.indexOf(first), tops.indexOf(last) + 1);
}

/** The paragraphs that end a section, in order (the final section ends with the body). */
function sectionEnds() {
  return topBlocks().filter((b) => b.dataset.sectionEnd != null);
}

/** The paragraph ending the section that top-level `block` is in, or null for the final section. */
function sectionEndOf(block) {
  const tops = topBlocks();
  for (let i = tops.indexOf(block); i >= 0 && i < tops.length; i++) if (tops[i].dataset.sectionEnd != null) return tops[i];
  return null;
}

/** The section properties of the section that top-level `block` is in, as they are now. */
const sectPrOf = (block) => {
  const end = sectionEndOf(block);
  return end ? doc.sectionBreakOf(end) : doc.finalSectPr();
};

/** A paragraph to hold a section break right after (`after`) or before top-level `block`: itself, or a new empty one beside a table. */
function breakHolder(block, after) {
  if (block.matches('p, h1, h2, h3, h4, h5, h6')) return block;
  const p = doc.emptyParagraph();
  if (after) block.after(p);
  else block.before(p);
  return p;
}

/**
 * Applies Page setup's `change` to `scope`, as Word's "Apply to": the whole document (every
 * section), this section, this point forward (a section break before the caret's paragraph), or
 * the selected text (section breaks before and after it). Returns where it applied, for messages.
 */
function applySetupTo(scope, change, breakType = 'nextPage') {
  const setOn = (end, set) => (end ? doc.setSectionSetup(end, set) : doc.setPageSetup({ ...doc.pageSetup, ...set }));
  const setFor = (end) => setOn(end, change);
  const blocks = pageRange?.startContainer.isConnected ? blocksIn(pageRange) : [];
  if (scope === 'whole' || !blocks.length) {
    doc.setPageSetup({ ...doc.pageSetup, ...change });
    for (const end of sectionEnds()) doc.setSectionSetup(end, change);
    return t('for the whole document');
  }
  const first = blocks[0];
  if (scope === 'section') {
    setFor(sectionEndOf(first));
    return t('for section {n}', { n: sectionEnds().indexOf(sectionEndOf(first)) + 1 || sectionEnds().length + 1 });
  }
  let made = 0;
  const tops = () => topBlocks();
  if (scope === 'selected') {
    const last = blocks[blocks.length - 1];
    const all = tops();
    // After the selection: a break, unless the selection ends its section (or the document).
    if (last.dataset.sectionEnd == null && all.indexOf(last) < all.length - 1) {
      const sect = sectPrOf(last);
      const after = sectionEndOf(last); // the section the text after the selection stays in
      doc.addSectionBreak(breakHolder(last, true), sect, breakType);
      setOn(after, { type: breakType }); // it starts after the new break: on a new page, or on the same one
      made++;
    }
  }
  // Before it: a break, unless it starts its section (or the document).
  const all = tops();
  const prev = all[all.indexOf(first) - 1];
  if (prev && prev.dataset.sectionEnd == null) {
    const sect = sectPrOf(first);
    doc.addSectionBreak(breakHolder(prev, true), sect, breakType); // a table before: a new paragraph after it holds the break
    made++;
    setOn(sectionEndOf(first), { ...change, type: breakType }); // the new section starts as the break says
  } else {
    setFor(sectionEndOf(first));
  }
  doc.renumber(page.value);
  return `${scope === 'selected' ? t('for the selected text') : t('from this point forward')}${made ? t(made === 1 ? ' ({n} section break added)' : ' ({n} section breaks added)', { n: made }) : ''}`;
}

/** Names each section break on the page with its section and its paper, where there is more than one section. */
function refreshSections() {
  if (!page.value) return;
  const ends = sectionEnds();
  ends.forEach((el, n) => {
    const s = doc.sectionSetup(el);
    const paper = PAPERS.find((p) => Math.abs(p.w - Math.min(s.w, s.h)) <= 20 && Math.abs(p.h - Math.max(s.w, s.h)) <= 20);
    const numbering = [s.numFmt && s.numFmt !== 'decimal' ? t('pages {format}', { format: formatLabel(s.numFmt) }) : '', s.numStart != null ? t('from {n}', { n: s.numStart }) : ''].filter(Boolean).join(' ');
    const columns = s.cols > 1 ? ` · ${t('{n} columns', { n: s.cols })}` : '';
    const kind = el.dataset.sectionEnd === 'continuous' ? t('section break (continuous)') : t('section break (next page)');
    const note = `${t('Section {n}', { n: n + 1 })}: ${paper ? `${paper.name} ` : ''}${t(s.w > s.h || s.orient === 'landscape' ? 'landscape' : 'portrait')}${columns}${numbering ? ` · ${numbering}` : ''} · ${kind}`;
    if (el.dataset.sectionNote !== note) el.dataset.sectionNote = note;
  });
  for (const el of page.value.querySelectorAll('[data-section-note]')) if (el.dataset.sectionEnd == null) delete el.dataset.sectionNote;
}

// ------------------------------------------------------------------ a different first page, odd and even pages

/**
 * A different first page and different odd and even pages, as Word's Header & Footer options:
 * the headers and footers they use that the document hasn't got yet are made, empty, and put in
 * the page, where the tabs above and below it choose them.
 */
function applyHeaderFooterFlags(titlePg, evenAndOdd) {
  doc.setHeaderFooterFlags({ titlePg, evenAndOdd });
  // Each one now used that isn't in the page yet (made new, or in the file but unused till now).
  for (const which of ['header', 'footer']) {
    const list = which === 'header' ? doc.headers : doc.footers;
    list.forEach((variant, i) => {
      if (!doc.variantShown(variant) || page.value.querySelector(`.docx-hf-part[data-which="${which}"][data-variant="${i}"]`)) return;
      const band = page.value.querySelector(`:scope > .docx-hf-band.${which}`);
      if (band) band.append(doc.renderVariantHost(which, variant, i));
      else if (which === 'header') page.value.prepend(doc.renderBand(which, [variant]));
      else page.value.append(doc.renderBand(which, [variant]));
    });
  }
  refreshCss(); // the Header and Footer styles may be new
  doc.relabelVariants(page.value);
}

// ------------------------------------------------------------------ Modify Style

// Word's Modify Style, for a paragraph style: its font, size, bold and italic, colour, alignment
// and spacing, changed in the style itself (styles.xml), so every paragraph in it follows.
// Its indents are shown as the Paragraph dialog shows them (Left is where the first line starts
// with a hanging indent), in centimetres or inches.
const modifyDialog = reactive({
  open: false, x: 0, y: 0, id: '', name: '', font: '', size: 11, bold: false, italic: false, auto: true, color: '#000000', align: 'left', before: 0, after: 0, line: '',
  unit: 'cm', left: 0, right: 0, special: 'none', by: 0,
});
const modifyDialogEl = ref(null);
let modifyInitial = null;
let modifyTwips = null; // the indents as they were, exactly, for switching units and for those left alone
/** The preview's first line, indented (or hanging) as the dialog has it, up to a point. */
const modifyPreviewIndent = computed(() => {
  const by = Math.min(4, Math.max(0, Number(modifyDialog.by) || 0) * (modifyDialog.unit === 'in' ? 2.54 : 1));
  if (modifyDialog.special === 'firstLine') return { textIndent: `${by}cm` };
  if (modifyDialog.special === 'hanging') return { textIndent: `-${by}cm`, paddingLeft: `${by}cm` };
  return {};
});
const LINE_CHOICES = [[240, 'Single'], [276, '1.15'], [360, '1.5'], [480, 'Double']];

function openModifyStyle() {
  if (textLocked.value) return explainLock();
  restoreSelection();
  const [block] = selectedParagraphs();
  const id = block?.dataset.style || doc.styles.defaultParagraph;
  if (!id || !doc.styles.styleEl(id)) {
    message.value = block ? 'This paragraph’s style isn’t in the document’s styles: it can’t be changed here' : 'Click in a paragraph in the style to change';
    return;
  }
  const f = doc.styles.styleFormat(id);
  modifyInitial = {
    font: f.font, size: f.size, bold: f.bold, italic: f.italic, auto: !f.color, color: `#${f.color || '000000'}`,
    align: f.align, before: Math.round(f.before / 2) / 10, after: Math.round(f.after / 2) / 10, line: f.line == null ? '' : String(f.line),
  };
  const unit = docUnit();
  modifyTwips = { left: f.left - f.hanging, right: f.right, by: f.hanging || f.firstLine };
  Object.assign(modifyInitial, {
    left: fromTwips(modifyTwips.left, unit), right: fromTwips(f.right, unit),
    special: f.hanging ? 'hanging' : f.firstLine ? 'firstLine' : 'none', by: fromTwips(modifyTwips.by, unit),
  });
  loadChoices();
  Object.assign(modifyDialog, modifyInitial, { unit, open: true, id, name: doc.styleLabel(id), x: Math.max(8, window.innerWidth / 2 - 170), y: 120 });
  nextTick(() => {
    keepInView(modifyDialog, modifyDialogEl.value);
    modifyDialogEl.value?.querySelector('select')?.focus();
  });
}

/** Modify Style's indents in another unit (remembered for the next dialogs). */
function setModifyUnit(unit) {
  saveUnit(unit);
  switchUnit(modifyDialog, modifyInitial, modifyTwips, ['left', 'right', 'by'], unit);
}

function closeModifyStyle() {
  modifyDialog.open = false;
  refocus();
}

/** The fonts Modify Style offers: the document's and Word's usual ones, by name, and the style's own. */
const modifyFonts = computed(() => {
  const names = [...new Set([modifyDialog.font, ...fontChoices.value.map((f) => f.name)].filter(Boolean))];
  return names;
});

function applyModifyStyle() {
  const d = { ...modifyDialog };
  const i = modifyInitial;
  modifyDialog.open = false;
  refocus();
  const num = (v, min, max) => Math.min(max, Math.max(min, Number(v) || 0));
  const change = {};
  if (d.font !== i.font) change.font = d.font;
  if (Number(d.size) !== i.size) change.size = num(d.size, 1, 1638);
  if (d.bold !== i.bold) change.bold = d.bold;
  if (d.italic !== i.italic) change.italic = d.italic;
  if (d.auto !== i.auto || (!d.auto && d.color.toLowerCase() !== i.color.toLowerCase())) change.color = d.auto ? null : d.color.slice(1).toUpperCase();
  if (d.align !== i.align) change.align = d.align;
  if (Number(d.before) !== i.before) change.before = num(d.before, 0, 1584) * 20;
  if (Number(d.after) !== i.after) change.after = num(d.after, 0, 1584) * 20;
  if (d.line !== i.line && d.line) change.line = Number(d.line);
  // Indents, as the Paragraph dialog writes them: with a hanging indent, w:left is Left and the
  // hanging together, so it is written again when either changes; a length left alone keeps its
  // exact value (what is shown is rounded).
  const len = (v) => (Number.isFinite(Number(v)) && v !== '' ? Number(v) : 0);
  const leftChanged = len(d.left) !== i.left;
  const specialChanged = d.special !== i.special || (d.special !== 'none' && len(d.by) !== i.by);
  const by = len(d.by) === i.by && d.special === i.special ? modifyTwips.by : Math.max(0, toTwips(len(d.by), d.unit));
  const hangs = (special) => special === 'hanging';
  if (leftChanged || (specialChanged && (hangs(d.special) || hangs(i.special)))) {
    const left = leftChanged ? toTwips(len(d.left), d.unit) : modifyTwips.left;
    change.left = left + (hangs(d.special) ? by : 0);
  }
  if (len(d.right) !== i.right) change.right = toTwips(len(d.right), d.unit);
  if (specialChanged) {
    if (d.special === 'hanging') change.hanging = by;
    else change.firstLine = d.special === 'firstLine' ? by : 0;
  }
  if (!Object.keys(change).length) return;
  undoStack.flush();
  doc.styles.modifyStyle(d.id, change);
  refreshCss();
  doc.renumber(page.value);
  pageChanged();
  undoStack.commit();
  markDirty();
  onSelectionChange();
  message.value = `Style ${d.name} changed: every paragraph in it follows`;
}

/** The paragraph style menu: a style, or Modify style… for the one the caret is in. */
function onStyleMenu(e) {
  const value = e.target.value;
  if (value === '__modify') {
    e.target.value = state.block;
    openModifyStyle();
    return;
  }
  applyBlockStyle(value);
}

// ------------------------------------------------------------------ watermark

// Word's text watermark (Design > Watermark): WordArt in the headers, behind the text, on every
// page (see doc.newWatermark). The page draws it over each page it estimates, and prints it on
// every page; in the file it is the headers' (a header is made for a document without one).
const watermark = reactive({ text: '', diagonal: true, color: 'silver', font: 'Calibri', size: 0, centers: [] });
let watermarkPrint = null; // the print stylesheet that puts it on every printed page

/** The watermark the document has now, as the Page setup dialog shows it. */
function watermarkNow() {
  const el = page.value?.querySelector('.docx-hf-band.header .watermark');
  const text = el?.dataset.watermark || '';
  return {
    mark: !text ? '' : WATERMARKS.includes(text) ? text : 'custom',
    markText: WATERMARKS.includes(text) ? '' : text,
    markLayout: el && !el.dataset.diagonal ? 'horizontal' : 'diagonal',
  };
}

/**
 * Puts watermark `text` (none: '') in every header, at the start of its first paragraph, in place
 * of the one there; a document without a header gets one (saved only while it is in the page).
 */
function setWatermark(text, diagonal) {
  for (const el of page.value.querySelectorAll('.docx-hf-band.header .watermark')) el.remove();
  let band = page.value.querySelector(':scope > .docx-hf-band.header');
  if (!text) {
    // A header made only for the watermark goes with it.
    const parts = band ? [...band.querySelectorAll('.docx-hf-part')] : [];
    if (parts.length && parts.every((h) => doc.headers.find((v) => v.path === h.dataset.part)?.added && !h.textContent.trim() && !h.querySelector('.raw-inline:not(.raw-hidden), img'))) band.remove();
    return;
  }
  if (!band) {
    band = doc.newHeader();
    refreshCss();
    page.value.prepend(band);
  }
  let n = 0;
  for (const host of band.querySelectorAll('.docx-hf-part')) {
    let p = host.querySelector(BLOCK_SELECTOR);
    if (!p) {
      p = doc.emptyParagraph();
      host.prepend(p);
    }
    p.prepend(doc.newWatermark(text, { diagonal, n: ++n }));
  }
}

/** VML's colours (a name, #RGB or #RRGGBB, maybe with " [n]") as CSS has them. */
const vmlColor = (c) => (c || 'silver').replace(/\s*\[\d+\]$/, '').trim() || 'silver';

/** The watermark shown: one in the middle of each page estimated (`breaks`: client y where pages start). */
function refreshWatermark(breaks = []) {
  const el = page.value?.querySelector('.docx-hf-band.header .watermark');
  const text = el?.dataset.watermark || '';
  watermark.text = text;
  if (!text) {
    if (watermarkPrint) watermarkPrint.textContent = '';
    return;
  }
  watermark.diagonal = el.dataset.diagonal === '1';
  watermark.color = vmlColor(el.dataset.color);
  watermark.font = el.dataset.font || 'Calibri';
  // As Word fits the WordArt: about the width between the margins, less for a short word.
  const letters = Math.max(1, [...text].length);
  const across = (doc.page.width - doc.page.left - doc.page.right) * (watermark.diagonal ? 1.1 : 0.95) * Math.min(1, 0.55 + letters * 0.04);
  watermark.size = Math.round(Math.min(across / (letters * 0.62), across / 2.2));
  const sheet = page.value.closest('.docx-sheet');
  const base = sheet.getBoundingClientRect();
  const zoom = sheetZoom.value || 1;
  const tops = [0, ...breaks.map((y) => (y - base.top) / zoom)];
  const bottom = sheet.offsetHeight;
  const pageHeight = doc.page.height || 1123;
  watermark.centers = tops.map((t, i) => {
    const end = i + 1 < tops.length ? tops[i + 1] : Math.min(bottom, t + pageHeight);
    return Math.round(t + Math.max(0, end - t) / 2);
  });
  // Printing: on every printed page, in its middle, behind the text.
  watermarkPrint ||= Object.assign(document.createElement('style'), { id: 'intact-watermark-print' });
  if (!watermarkPrint.isConnected) document.head.append(watermarkPrint);
  const cssText = text.replace(/[\\"]/g, '\\$&').replace(/[\n\r]/g, ' ');
  watermarkPrint.textContent = `@media print { .docx-sheet .docx-page::after { content: "${cssText}"; position: fixed; left: 50%; top: 50%; z-index: -1;`
    + ` transform: translate(-50%, -50%)${watermark.diagonal ? ' rotate(-45deg)' : ''}; font: 700 ${watermark.size}px '${watermark.font.replace(/'/g, '')}', Calibri, Carlito, sans-serif;`
    + ` color: ${watermark.color}; opacity: .5; white-space: nowrap; pointer-events: none; } }`;
}
onBeforeUnmount(() => watermarkPrint?.remove());

/** Ctrl+Enter: a page break at the caret (tracked while tracking). Not inside form fields. */
function insertPageBreak() {
  const range = selectionInPage();
  if (!range) return;
  if (inField(range.startContainer)) {
    message.value = 'A page break can’t go inside a field';
    return;
  }
  if (refuseInNote(range, 'A page break')) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (!range.collapsed && refuseRemoving([range])) return;
  undoStack.flush();
  const opts = trackedBy(range.startContainer);
  if (!range.collapsed) {
    if (opts) deleteTracked(range, false, opts);
    else deleteSelection();
  }
  const pageBreak = doc.newPageBreak();
  if (opts) insertTracked([pageBreak], opts);
  else insertAtCaret([pageBreak]);
  onInput();
}

// ------------------------------------------------------------------ track changes

let pendingSplit = null; // the author of a paragraph break being inserted while tracking

/** The author and date for a tracked edit made now, or null when this edit isn't tracked. */
function trackedBy(node) {
  return tracking.value && node && !inField(node) ? { author: author.value, date: revisionDate() } : null;
}

/**
 * Typing, deleting, Enter and Shift+Enter while tracking: recorded as tracked changes instead of
 * changing the text (see editor/track.js). True if the event was dealt with here.
 */
function trackedBeforeInput(e) {
  const range = selectionInPage();
  const opts = trackedBy(range?.startContainer);
  if (!opts) {
    // Track Changes off: text typed in (or right after) one of this session's insertions is plain.
    if (e.inputType === 'insertText' && e.data && range?.collapsed && typeUntracked(e.data)) {
      e.preventDefault();
      onInput();
      return true;
    }
    return false;
  }
  const type = e.inputType;
  if (type === 'insertFromDrop' || type === 'deleteByDrag') {
    e.preventDefault();
    message.value = 'Dragging text isn’t tracked – cut and paste it instead while tracking changes';
    return true;
  }
  if (type === 'insertCompositionText') return true; // wrapped as an insertion on compositionend
  const deleting = type.startsWith('delete');
  if (!deleting && !['insertText', 'insertReplacementText', 'insertParagraph', 'insertLineBreak'].includes(type)) return false;
  let removal = null;
  if (!range.collapsed) removal = range.cloneRange();
  else if (type === 'deleteContentBackward' || type === 'deleteContentForward') removal = unitRange(type === 'deleteContentBackward');
  else if (deleting) removal = rangesRemovedBy(e)[0] || null;
  if (deleting && !removal) {
    e.preventDefault(); // at the very start (or end): nothing to delete
    return true;
  }
  if (removal && refuseRemoving([removal])) {
    e.preventDefault();
    return true;
  }
  if (type === 'insertParagraph' && !removal) {
    pendingSplit = opts; // the browser splits the paragraph; onInput marks the new break
    return true;
  }
  if (type === 'insertText' && !removal && !typeTracked(e.data || '', opts)) return true; // extends an insertion
  e.preventDefault();
  if (removal) deleteTracked(removal, /Backward$/.test(type), opts);
  if (type === 'insertText' || type === 'insertReplacementText') {
    const text = e.data ?? e.dataTransfer?.getData('text/plain') ?? '';
    if (removal && text && !typeTracked(text, opts)) {
      document.execCommand('insertText', false, text); // into an insertion; fires input
      return true;
    }
  } else if (type === 'insertParagraph') {
    pendingSplit = opts;
    document.execCommand('insertParagraph'); // fires input, which marks the break
    return true;
  } else if (type === 'insertLineBreak') {
    document.execCommand('insertLineBreak');
    wrapInsertedBreak(opts);
    undoStack.schedule();
    return true;
  }
  onInput();
  return true;
}

/** Text from an input method can't be intercepted: it becomes an insertion once it is complete. */
function onCompositionEnd(e) {
  const range = selectionInPage();
  const opts = trackedBy(range?.startContainer);
  if (!opts) return;
  wrapComposedText(e.data, opts);
  undoStack.schedule();
}

function noteUntrackedFormatting() {
  if (tracking.value) message.value = 'Formatting isn’t tracked – it’s applied directly';
}

function toggleTracking() {
  if (lockMode.value) return explainLock();
  if (tracking.value) {
    tracking.value = false;
    message.value = 'Track Changes is off: your edits change the text directly';
  } else if (named.value) {
    startTracking();
  } else {
    openAuthorDialog(true);
  }
}

function startTracking() {
  tracking.value = true;
  message.value = `Track Changes is on: your edits are marked as ${author.value}’s`;
  page.value.focus();
  restoreSelection();
}

async function openAuthorDialog(start = false) {
  start = start === true; // not a click event
  const anchor = (start ? trackButton.value : document.querySelector('.tracking-status')) || trackButton.value;
  const rect = anchor.getBoundingClientRect();
  Object.assign(authorDialog, {
    open: true,
    start,
    name: named.value ? author.value : '',
    x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)),
    y: start ? rect.bottom + 4 : Math.max(8, rect.top - 140),
  });
  await nextTick();
  authorInput.value?.focus();
  authorInput.value?.select();
}

function closeAuthorDialog() {
  authorDialog.open = false;
  page.value.focus();
  restoreSelection();
}

function confirmAuthor() {
  const name = authorDialog.name.trim();
  if (!name) return;
  author.value = name;
  try { localStorage.setItem(AUTHOR_KEY, name); } catch { /* remembered for this session only */ }
  named.value = true;
  authorDialog.open = false;
  if (authorDialog.start) startTracking();
  else {
    message.value = `Your new tracked changes are marked as ${name}’s`;
    page.value.focus();
    restoreSelection();
  }
}

/** Deletes the selection with execCommand (which fires no beforeinput), keeping comment ranges. */
function deleteSelection() {
  const before = commentMarkers(page.value);
  document.execCommand('delete');
  restoreCommentMarkers(before, page.value);
}

/**
 * Refuses an edit that would remove a tracked change, saying why: true if `ranges` (what the edit
 * would delete) touch one. Edits made with execCommand fire no beforeinput, so call this first.
 */
function refuseRemoving(ranges) {
  const el = trackedChangeAt(ranges, page.value);
  if (!el) return false;
  message.value = refusalMessage(el);
  flash(el);
  return true;
}

/** Flashes what an edit can't change. An animation, not a class, so undo snapshots don't change. */
function flash(el) {
  el.animate?.([
    { outline: '2px solid #b3261e', outlineOffset: '1px', backgroundColor: 'rgb(179 38 30 / 15%)' },
    { outline: '2px solid transparent', outlineOffset: '1px' },
  ], { duration: 1200, easing: 'ease-out' });
}

// ------------------------------------------------------------------ read-only content

// The read-only block or cell last clicked. The editor gives up focus then (the browser would make
// up a caret for the next key otherwise), and keys typed next say why nothing happens.
let lastReadOnly = null;
let pointerDown = null;

/** The read-only element the selection starts or ends in, or null. */
function readOnlySelection() {
  const range = selectionInPage();
  return range && (readOnlyAt(range.startContainer, page.value) || readOnlyAt(range.endContainer, page.value));
}

/** Says in the status bar what read-only `el` is and where to edit it; `refused` flashes it too. */
function explainReadOnly(el, refused = false) {
  message.value = readOnlyMessage(el, page.value);
  if (refused) flash(el);
}

function onPageMousedown(e) {
  selectedChange = null;
  const pic = e.target.closest?.('.picture');
  if (pic && page.value.contains(pic) && !lockMode.value) {
    onPictureMousedown(e, pic);
    return;
  }
  const range = selectionInPage();
  // Typing in the text, then in a header, footer, note or text box, is two steps for Undo.
  const hostOf = (n) => (n?.nodeType === 1 ? n : n?.parentElement)?.closest('[contenteditable="true"]');
  if (range && hostOf(e.target) !== hostOf(range.startContainer)) undoStack.flush();
  pointerDown = { x: e.clientX, y: e.clientY, range: range?.cloneRange() || null };
}

/**
 * A click (not a drag) on read-only content says what it is. The browser leaves the caret where it
 * was, so typing would land there, out of sight: put it beside inline content, on the side clicked,
 * or clear it for a block or cell.
 */
function clickReadOnly(el, e) {
  const down = pointerDown;
  pointerDown = null;
  if (down && Math.hypot(e.clientX - down.x, e.clientY - down.y) > 4) return;
  explainReadOnly(el);
  const now = selectionInPage();
  const moved = now && down?.range && !(now.startContainer === down.range.startContainer && now.startOffset === down.range.startOffset
    && now.endContainer === down.range.endContainer && now.endOffset === down.range.endOffset);
  if (moved && !readOnlyAt(now.startContainer, page.value)) return; // the browser placed the caret itself
  const sel = window.getSelection();
  const beside = caretBeside(el, e.clientX, e.clientY);
  sel.removeAllRanges();
  lastReadOnly = beside ? null : el;
  if (!beside) {
    document.activeElement?.blur();
    state.table = !!tableSelection(); // a read-only cell: still in its table, for Table layout to say why
    return;
  }
  el.parentElement.closest('[contenteditable="true"]')?.focus({ preventScroll: true });
  sel.addRange(beside);
}

/**
 * Typing while the selection is in a header, footer or comment (outside the editable page), or
 * after clicking a read-only block, does nothing; say why. True if the key was explained.
 */
function explainKeyOutsideText(e) {
  if (!isEditingKey(e) || document.activeElement?.matches('input, textarea, select, [contenteditable="true"]')) return false;
  if (lastReadOnly?.isConnected) {
    explainReadOnly(lastReadOnly, true);
    return true;
  }
  const node = window.getSelection().anchorNode;
  const el = node?.nodeType === 1 ? node : node?.parentElement;
  if (el?.closest('.comments') && e.key !== 'Enter') {
    message.value = 'Read-only: comments are shown for reference – reply to or resolve them in Word';
    return true;
  }
  return false;
}

/**
 * End (and Shift+End) as Word has it, also on a line with something that can't be typed into, such
 * as a page number or a footnote's reference: Chrome's End doesn't move past one. There the caret
 * (or the selection's end) goes to the end of the caret's line, found where the paragraph's right
 * edge is level with the caret, and past a page number, a note's reference or a field's hidden
 * marks that end it. True if it was dealt with.
 */
function endOfLine(extend) {
  const sel = window.getSelection();
  if (!sel.rangeCount || !sel.focusNode || !page.value?.contains(sel.focusNode)) return false;
  const block = closestBlock(sel.focusNode);
  // Only a line with a page number or a note's reference (elsewhere the browser's End works: on
  // such a line Chrome's End stops before the number, or the reference), and not right to left.
  if (!block?.querySelector('.page-field, .note-ref') || getComputedStyle(block).direction === 'rtl') return false;
  const caret = document.createRange();
  caret.setStart(sel.focusNode, sel.focusOffset);
  let at = caret.getBoundingClientRect();
  if (!at.height) at = block.getClientRects()[0] || block.getBoundingClientRect();
  const box = block.getBoundingClientRect();
  const x = box.right - parseFloat(getComputedStyle(block).paddingRight || 0) - 1;
  const y = at.top + at.height / 2;
  const pos = document.caretPositionFromPoint?.(x, y);
  let node = pos ? pos.offsetNode : document.caretRangeFromPoint?.(x, y)?.startContainer;
  let offset = pos ? pos.offset : document.caretRangeFromPoint?.(x, y)?.startOffset;
  if (!node || !block.contains(node)) {
    node = block;
    offset = block.childNodes.length;
  } else {
    // Inside something that can't be typed into: just after it, and after hidden marks that follow.
    const locked = (node.nodeType === 1 ? node : node.parentElement).closest('[contenteditable="false"]');
    if (locked && block.contains(locked)) {
      let last = locked;
      while (last.nextSibling?.nodeType === 1 && last.nextSibling.classList.contains('raw-hidden')) last = last.nextSibling;
      node = last.parentNode;
      offset = [...node.childNodes].indexOf(last) + 1;
    }
    // Before a note's reference (or a page number) that ends the line: just after it, as Word's End.
    [node, offset] = pastLockedAtEnd(block, node, offset, at);
  }
  if (extend) sel.extend(node, offset);
  else sel.collapse(node, offset);
  return true;
}

/**
 * Home (and Shift+Home) as Word has it on a line that starts with a page number: Chrome's Home puts
 * the caret inside the number's field there, where typed text would become the field's result. The
 * caret goes to the start of the line, found where the paragraph's left edge is level with the
 * caret, and before the field's hidden marks, so text typed goes before the field. True if it was
 * dealt with; only lines with a page number (elsewhere the browser's Home works), not right to left.
 */
function startOfLine(extend) {
  const sel = window.getSelection();
  if (!sel.rangeCount || !sel.focusNode || !page.value?.contains(sel.focusNode)) return false;
  const block = closestBlock(sel.focusNode);
  if (!block?.querySelector('.page-field') || getComputedStyle(block).direction === 'rtl') return false;
  const caret = document.createRange();
  caret.setStart(sel.focusNode, sel.focusOffset);
  let at = caret.getBoundingClientRect();
  if (!at.height) at = block.getClientRects()[0] || block.getBoundingClientRect();
  const box = block.getBoundingClientRect();
  const x = box.left + parseFloat(getComputedStyle(block).paddingLeft || 0) + 1;
  const y = at.top + at.height / 2;
  const pos = document.caretPositionFromPoint?.(x, y);
  let node = pos ? pos.offsetNode : document.caretRangeFromPoint?.(x, y)?.startContainer;
  let offset = pos ? pos.offset : document.caretRangeFromPoint?.(x, y)?.startOffset;
  if (!node || !block.contains(node)) return false;
  // Inside something that can't be typed into (the number's field): just before it.
  let locked = (node.nodeType === 1 ? node : node.parentElement).closest('[contenteditable="false"]');
  if (locked && !block.contains(locked)) locked = null;
  if (!locked) {
    // Only where the line starts with the page number (after hidden marks): elsewhere the browser's.
    let next = node.nodeType === 3 ? (offset === 0 ? node : null) : node.childNodes[offset];
    while (next && ((next.nodeType === 3 && !next.data.replace(/[\u200b\ufeff]/g, '')) || next.classList?.contains('raw-hidden'))) next = next.nextSibling;
    if (!next?.matches?.('.page-field, .field:has(> .page-field)')) return false;
  }
  if (locked) {
    while (locked.parentElement !== block && locked.parentElement?.closest('[contenteditable="false"]') && block.contains(locked.parentElement.closest('[contenteditable="false"]'))) {
      locked = locked.parentElement.closest('[contenteditable="false"]');
    }
    node = locked.parentNode;
    offset = [...node.childNodes].indexOf(locked);
  }
  // Before hidden marks (a field's start and instruction) just before it, at the start of the line.
  const empty = (n) => n.nodeType === 3 && !n.data.replace(/[\u200b\ufeff]/g, '');
  if (node.nodeType === 3 && offset === 0) {
    offset = [...node.parentNode.childNodes].indexOf(node);
    node = node.parentNode;
  }
  if (node.nodeType === 1) {
    while (offset > 0) {
      const prev = node.childNodes[offset - 1];
      if (!(empty(prev) || (prev.nodeType === 1 && prev.classList.contains('raw-hidden')))) break;
      offset--;
    }
  }
  if (extend) sel.extend(node, offset);
  else sel.collapse(node, offset);
  return true;
}

/**
 * From (`node`, `offset`) at the end of a line of `block`, past what can't be typed into (a note's
 * reference, a page number) and hidden marks that follow it on the same line (`line`: the caret's
 * rectangle). Returns the position after them, or the one given.
 */
function pastLockedAtEnd(block, node, offset, line) {
  const empty = (n) => n.nodeType === 3 && !n.data.replace(/[\u200b\ufeff]/g, '');
  const onLine = (el) => [...el.getClientRects()].some((r) => r.top < line.bottom && r.bottom > line.top);
  // Up to the element holding the position, while it is at the end of a text node or a run.
  let at = node;
  let i = offset;
  while (at !== block && (at.nodeType === 3 ? i >= at.data.length : i >= at.childNodes.length) && block.contains(at.parentNode)) {
    i = [...at.parentNode.childNodes].indexOf(at) + 1;
    at = at.parentNode;
  }
  if (at.nodeType === 3) return [node, offset];
  let next = at.childNodes[i];
  let past = null;
  while (next && (empty(next) || (next.nodeType === 1 && (next.classList.contains('raw-hidden')
    || (next.matches('[contenteditable="false"]:is(.note-ref, .page-field)') && onLine(next)))))) {
    if (next.nodeType === 1 && !next.classList.contains('raw-hidden')) past = next;
    next = next.nextSibling;
  }
  if (!past) return [node, offset];
  // After hidden marks that follow it too, as after a page number.
  let last = past;
  while (last.nextSibling?.nodeType === 1 && last.nextSibling.classList.contains('raw-hidden')) last = last.nextSibling;
  return [last.parentNode, [...last.parentNode.childNodes].indexOf(last) + 1];
}

function onKeydown(e) {
  selectedChange = null;
  if (formKeydown(e)) return;
  if (e.key === 'End' && !e.ctrlKey && !e.metaKey && !e.altKey && endOfLine(e.shiftKey)) {
    e.preventDefault();
    return;
  }
  if (e.key === 'Home' && !e.ctrlKey && !e.metaKey && !e.altKey && startOfLine(e.shiftKey)) {
    e.preventDefault();
    return;
  }
  if (e.key === 'Escape' && picture.cropping) {
    e.preventDefault();
    toggleCropping();
    return;
  }
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  const range = selectionInPage();
  // Select All in the text selects the text, not the notes below it (in a note, the note), as in Word.
  const notes = page.value.querySelector(':scope > :is(.docx-notes, .docx-hf-band)');
  if (mod && key === 'a' && !e.shiftKey && !e.altKey && notes && range && !noteAt(range.startContainer)) {
    e.preventDefault();
    // From the first block to the end of the last one before the notes and footer (after the
    // header): deleting it leaves a paragraph.
    const text = [...page.value.children].filter((c) => !c.matches('.docx-notes, .docx-hf-band'));
    const first = text[0];
    const last = text[text.length - 1];
    if (!last) return;
    const all = document.createRange();
    all.setStart(first, 0);
    all.setEnd(last, last.childNodes.length);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(all);
    return;
  }
  if (isEditingKey(e) || (e.key === 'Tab' && !mod && !e.altKey)) {
    const locked = readOnlySelection();
    if (locked) {
      e.preventDefault();
      explainReadOnly(locked, true);
      return;
    }
  }
  const field = range && controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content');
  if ((e.altKey && !mod && e.key === 'ArrowDown') || (e.key === ' ' && !mod && !e.altKey && range && !range.collapsed)) {
    const near = fieldAtSelection(range, page.value);
    if (near) {
      e.preventDefault();
      activateField(near);
      return;
    }
  }
  if (mod && !e.altKey && ['b', 'i', 'u'].includes(key) && inPlainField()) {
    e.preventDefault();
    return;
  }
  if (e.key === 'F9' && !mod && !e.altKey && !e.shiftKey) {
    e.preventDefault();
    updateAllFields();
    return;
  }
  // The Menu key and Shift+F10 open the right-click menu at the caret; not every browser makes a
  // contextmenu event of them.
  if ((e.key === 'ContextMenu' || (e.key === 'F10' && e.shiftKey)) && !mod && !e.altKey) {
    e.preventDefault();
    openContextMenu();
    return;
  }
  if (e.key === 'F7' && e.shiftKey && !mod && !e.altKey) {
    e.preventDefault();
    openThesaurus();
    return;
  }
  if (e.key === 'F3' && e.shiftKey && !mod && !e.altKey) {
    e.preventDefault();
    changeCaseTo(null);
    return;
  }
  if (e.key === 'Escape' && painter.active) {
    e.preventDefault();
    stopPainter();
    return;
  }
  // Esc stops reading aloud, as Word's Read Aloud.
  if (e.key === 'Escape' && reading.on) {
    e.preventDefault();
    reader?.stop();
    message.value = t('Reading aloud stopped');
    return;
  }
  const shortcut = shortcutFor(e);
  // Ctrl+Shift+V pastes formatting once some was copied (Ctrl+Shift+C); before, it is the browser's
  // paste as plain text.
  if (shortcut === 'pasteFormat' && !painter.copied) return;
  if (shortcut) {
    e.preventDefault();
    runShortcut(shortcut);
    return;
  }
  if (e.key === 'Enter' && field?.tagName === 'SPAN') {
    // A field inside a line can't hold paragraphs; multi-line plain text fields take line breaks.
    e.preventDefault();
    if (e.shiftKey || field.dataset.plain === 'multi') document.execCommand('insertLineBreak');
    return;
  }
  if (mod && key === 's') {
    e.preventDefault();
    requestDownload();
  } else if (mod && !e.altKey && (key === 'z' || key === 'y')) {
    e.preventDefault();
    if (key === 'y' || e.shiftKey) redo();
    else undo();
  } else if (e.key === 'Tab' && !mod && !e.altKey) {
    e.preventDefault();
    onTab(e);
  } else if ((e.key === 'Enter' || e.key === 'Backspace') && !mod && !e.altKey && !e.shiftKey && !e.isComposing) {
    // Word: Enter on an empty list item, or Backspace at the start of one, ends the item's list.
    const range = selectionInPage();
    const block = range?.collapsed && closestBlock(range.startContainer);
    if (!block || !doc.listOf(block)) return;
    if (e.key === 'Enter' ? !isEmptyBlock(block) : !isAtBlockStart(range, block)) return;
    e.preventDefault();
    if (e.key === 'Enter') endListItem(block);
    else changeLists(() => doc.removeList(block));
  }
}

function onPaste(e) {
  e.preventDefault();
  pasteFrom(e.clipboardData);
}

/** Pastes clipboard data (a DataTransfer: a paste event's, or one made from the clipboard read). */
function pasteFrom(clipboardData) {
  const range = selectionInPage();
  const locked = readOnlySelection();
  if (locked) {
    explainReadOnly(locked, true);
    return;
  }
  if (!range || refuseRemoving([range])) return;
  // A picture on its own (a screenshot, an image copied from an app): inserted as a picture.
  // Some browsers list a copied picture only among the clipboard's items, not its files.
  const files = clipboardData.files?.length ? [...clipboardData.files]
    : [...(clipboardData.items || [])].filter((i) => i.kind === 'file').map((i) => i.getAsFile()).filter(Boolean);
  const pictures = files.filter((f) => f.type.startsWith('image/'));
  if (pictures.length && !clipboardData.getData('text/plain').trim()) {
    insertPictures(pictures);
    return;
  }
  undoStack.flush();
  const html = clipboardData.getData('text/html');
  let text = clipboardData.getData('text/plain').replace(/\r\n?/g, '\n');
  const field = controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content')
    || formTextAt(range.startContainer);
  if (field?.tagName === 'SPAN' || field?.dataset.plain) {
    // Fields inside a line, and plain-text fields, take text only; inside a line it stays on one
    // line (or becomes line breaks in multi-line fields).
    if (field.tagName === 'SPAN' && field.dataset.plain !== 'multi') text = text.replace(/\s*\n\s*/g, ' ');
    document.execCommand('insertText', false, text);
    undoStack.commit();
    onInput();
    return;
  }
  let { blocks, skippedImages } = parseClipboard(html, text);
  if (!blocks.length) return;
  // A table goes in the text: where one can't go, its rows are pasted as lines with tabs.
  const noTable = blocks.some((b) => b.table) ? tableRefusal(range) : null;
  if (noTable) blocks = tablesAsText(blocks);
  // Pasted into a link: its text joins that link (a link can't hold another).
  if ((range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('a')) {
    for (const b of paragraphBlocks(blocks)) for (const r of b.runs) r.href = null;
  }
  const tracked = trackedBy(range.startContainer);
  if (!range.collapsed) {
    if (tracked) deleteTracked(range, false, tracked);
    else deleteSelection();
  }
  const tables = pasteMixed(blocks, tracked);
  onInput();
  undoStack.commit();
  const notes = [];
  if (tables.length === 1) notes.push(`Table pasted: ${tables[0].cols} × ${tables[0].rows}${tracked ? ' (tracked)' : ''}`);
  else if (tables.length) notes.push(`${tables.length} tables pasted${tracked ? ' (tracked)' : ''}`);
  if (noTable) notes.push(`The table was pasted as text, with tabs between its cells: ${noTable}`);
  if (skippedImages) {
    notes.push(`${skippedImages} ${skippedImages === 1 ? 'picture was' : 'pictures were'} not pasted with the text – copy a picture on its own and paste it, or use Insert picture`);
  }
  if (notes.length) message.value = notes.join(' · ');
}

/** The paragraph blocks of parsed clipboard blocks, those in tables' cells too. */
function paragraphBlocks(blocks) {
  return blocks.flatMap((b) => (b.table ? b.table.cells.flatMap((c) => c.blocks) : [b]));
}

/** Why a pasted table can't go at `range` (it is pasted as text there), or null if it can. */
function tableRefusal(range) {
  const note = noteAt(range.startContainer);
  if (note) return `a table can’t go in ${storyName(note)} here`;
  if (controlAt(range.startContainer, page.value) || formTextAt(range.startContainer)) return 'a table can’t go in a field';
  const el = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  if (el?.closest('td')) return 'a table can’t go inside a table’s cell here';
  return null;
}

/**
 * Pastes parsed blocks, paragraphs and tables in turn (see pasteBlocks and pasteTable). Returns the
 * tables pasted, as { rows, cols }.
 */
function pasteMixed(blocks, tracked) {
  const tables = [];
  for (let i = 0; i < blocks.length;) {
    if (blocks[i].table) {
      pasteTable(blocks[i].table, tracked);
      tables.push(blocks[i].table);
      i++;
      continue;
    }
    let j = i;
    while (j < blocks.length && !blocks[j].table) j++;
    pasteBlocks(blocks.slice(i, j), tracked);
    i = j;
  }
  if (tables.length) {
    doc.renumber(page.value);
    refreshChanges();
  }
  return tables;
}

/** The rendered cell of `table` that starts at grid row `r`, column `c` (cells span columns). */
function cellAtGrid(table, r, c) {
  let at = 0;
  for (const td of table.rows[r]?.cells || []) {
    if (at === c) return td;
    at += td.colSpan || 1;
    if (at > c) return null;
  }
  return null;
}

/**
 * A table pasted at the caret, as Word pastes one from Excel, Google Sheets or a web page: where
 * Insert table puts one (before the paragraph if the caret is at its start, else the paragraph is
 * split there), in Word's Table Grid style, its columns in proportion to the copied widths, merged
 * cells merged, and each cell's text pasted into it with its basic formatting. While tracking, its
 * rows (and the text in them) are tracked insertions. The caret goes to the paragraph after it.
 */
function pasteTable({ rows, cols, cells, widths }, tracked) {
  const range = selectionInPage();
  if (!range) return null;
  range.collapse(true);
  let block = closestBlock(range.startContainer);
  if (!block) {
    block = doc.emptyParagraph();
    page.value.append(block);
  } else if (!isEmptyBlock(block) && !isAtBlockStart(range, block)) {
    document.execCommand('insertParagraph');
    if (tracked) markSplit(tracked);
    block = closestBlock(selectionInPage().startContainer) || block;
  }
  const merges = cells.filter((c) => c.rowspan > 1 || c.colspan > 1)
    .map((c) => ({ r0: c.r, r1: c.r + c.rowspan - 1, c0: c.c, c1: c.c + c.colspan - 1 }));
  const table = doc.newTable(rows, cols, tracked, { widths, merges });
  block.before(table);
  refreshCss(); // the Table Grid style may be new
  for (const cell of cells) {
    if (!cell.blocks.length) continue;
    const target = cellAtGrid(table, cell.r, cell.c)?.querySelector(BLOCK_SELECTOR);
    if (!target) continue;
    placeCaret(target, true);
    pasteBlocks(cell.blocks, tracked, false);
  }
  placeCaret(block, true);
  return table;
}

/**
 * Inserts parsed blocks with Word's "merge formatting": text takes the formatting where it lands plus
 * its own bold/italic/underline/strike. The first block joins the paragraph at the caret; each
 * further one starts a new paragraph that keeps the destination paragraph's settings, unless it is
 * a list item, which joins the destination list if it is of the same kind or else a new list.
 */
function pasteBlocks(blocks, tracked = null, renumber = true) {
  const startRange = selectionInPage();
  const dest = startRange && closestBlock(startRange.startContainer);
  const destWasEmpty = !!dest && isEmptyBlock(dest);
  const destNum = dest ? { numId: dest.dataset.numId, ilvl: dest.dataset.ilvl } : {};
  const destList = dest ? doc.listOf(dest) : null;
  const destKind = doc.listKind(destList);
  const newLists = {};

  const applyList = (el, block) => {
    if (block.list) {
      const { kind, level } = block.list;
      const numId = destList && destKind === kind ? destList.numId : (newLists[kind] ||= doc.newList(kind));
      doc.setList(el, numId, doc.hasListLevel(numId, level) ? level : 0);
    } else if (destNum.numId != null) {
      el.dataset.numId = destNum.numId;
      el.dataset.ilvl = destNum.ilvl ?? '0';
    } else {
      delete el.dataset.numId;
      delete el.dataset.ilvl;
    }
  };

  blocks.forEach((block, i) => {
    if (i > 0) {
      document.execCommand('insertParagraph');
      if (tracked) markSplit(tracked);
    }
    const here = selectionInPage();
    const el = here && closestBlock(here.startContainer);
    if (el && (i > 0 || destWasEmpty)) applyList(el, block);
    if (tracked && block.runs.length) insertTracked(runNodes(block.runs), tracked);
    else insertAtCaret(runNodes(block.runs));
  });
  if (renumber) doc.renumber(page.value);
}

/** Picture files dropped on the page are inserted where they are dropped; other files aren't. */
function onDrop(e) {
  if (textLocked.value) {
    e.preventDefault();
    explainLock();
    return;
  }
  const files = [...(e.dataTransfer?.files || [])];
  if (!files.length) return;
  e.preventDefault();
  const pictures = files.filter((f) => f.type.startsWith('image/'));
  if (!pictures.length) {
    if (/\.(docx|dotx|docm|dotm)$/i.test(files[0].name)) offerNewTab(files[0]);
    else message.value = 'Only pictures can be dropped into the document';
    return;
  }
  let at = null;
  if (document.caretPositionFromPoint) {
    const pos = document.caretPositionFromPoint(e.clientX, e.clientY);
    if (pos) {
      at = document.createRange();
      at.setStart(pos.offsetNode, pos.offset);
    }
  } else if (document.caretRangeFromPoint) {
    at = document.caretRangeFromPoint(e.clientX, e.clientY);
  }
  insertPictures(pictures, at && page.value.contains(at.startContainer) ? at : null);
}

function onClick(e) {
  if (e.target.closest('.picture')) return; // selected on mousedown
  const formChoice = e.target.closest(FORM_CHOICE);
  if (formChoice && page.value.contains(formChoice)) {
    activateFormField(formChoice);
    return;
  }
  const noteRef = e.target.closest('.note-ref[data-note-ref], .note-mark');
  if (noteRef && goToNote(noteRef)) return;
  const marker = e.target.closest('.comment-ref');
  if (marker) {
    commentsOpen.value = true;
    selectComment(marker.dataset.commentId, false);
    return;
  }
  if (textLocked.value && !(lockMode.value === 'forms' && e.target.closest(OPEN_FIELD))) {
    const a = e.target.closest('a[href]');
    if (a && (e.ctrlKey || e.metaKey) && /^(https?:|mailto:)/i.test(a.getAttribute('href'))) {
      e.preventDefault();
      window.open(a.href, '_blank', 'noopener');
      return;
    }
    if (window.getSelection().isCollapsed) explainLock(); // not after selecting text, to copy it
    return;
  }
  selectBlankFormText(e.target);
  const placeholderCell = e.target.closest('td.sdt-cell.showing-placeholder');
  if (placeholderCell) {
    selectCell(placeholderCell); // a field around a table cell: typing replaces its placeholder
    return;
  }
  const locked = readOnlyAt(e.target, page.value);
  if (locked) {
    clickReadOnly(locked, e);
    return;
  }
  const control = controlAt(e.target, page.value);
  if (control?.classList.contains('showing-placeholder') && e.target.closest('.sdt-content')) {
    selectCell(controlBox(control)); // the browser resets the selection on mouseup
    return;
  }
  if (control && !e.target.closest('.sdt-content')) {
    const kind = doc.control(control).kind;
    if (kind === 'checkbox') toggleCheckbox(control);
    else if (kind !== 'text') openChooser(control);
    return;
  }
  const a = e.target.closest('a[href]');
  if (a && (e.ctrlKey || e.metaKey) && /^(https?:|mailto:)/i.test(a.getAttribute('href'))) {
    e.preventDefault();
    window.open(a.href, '_blank', 'noopener');
  }
}

// ------------------------------------------------------------------ links

const linkButton = ref(null);
const linkEl = ref(null);
const linkInput = ref(null);
const linkDialog = reactive({ open: false, x: 0, y: 0, address: '', text: '', askText: false, editing: false, error: '' });
let linkRange = null; // where the link goes: the selection when the dialog opened
let linkTarget = null; // the link being changed, if the selection was in one

/** The link (<a>) the selection is in, or null. */
function linkAt(range) {
  const el = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
  const a = el?.closest('a[href]');
  return a && page.value.contains(a) ? a : null;
}

/**
 * Ctrl+K, as in Word: a link to a web or e-mail address or a DOI for the selected text (or new text
 * at the caret), or the link the caret is in, to change or remove.
 */
async function openLinkDialog() {
  restoreSelection();
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text, or select some, to add a link';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  const field = controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content');
  if (field?.dataset.plain) {
    message.value = 'A plain-text field can’t hold a link';
    return;
  }
  if (closestBlock(range.startContainer) !== closestBlock(range.endContainer)) {
    message.value = 'Select text within one paragraph to make it a link';
    return;
  }
  linkTarget = linkAt(range);
  linkRange = range.cloneRange();
  const anchor = anchorRect(linkButton.value);
  Object.assign(linkDialog, {
    open: true,
    address: linkTarget ? linkTarget.getAttribute('href') : '',
    text: '',
    askText: !linkTarget && range.collapsed,
    editing: !!linkTarget,
    error: '',
    x: Math.max(8, Math.min(anchor?.left ?? 8, window.innerWidth - 340)),
    y: (anchor?.bottom ?? 60) + 4,
  });
  await nextTick();
  linkInput.value?.focus();
  linkInput.value?.select();
}

function closeLinkDialog() {
  linkDialog.open = false;
  page.value.focus();
  if (linkRange) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(linkRange);
  }
}

function confirmLink() {
  const href = normalizeLink(linkDialog.address);
  if (!href) {
    linkDialog.error = 'Enter a web address (https://… or www.…), an e-mail address or a DOI (10.…)';
    return;
  }
  linkDialog.open = false;
  page.value.focus();
  undoStack.flush();
  const sel = window.getSelection();
  let done;
  if (linkTarget) {
    // A new address for a link: the file's own link keeps its settings, with the new address.
    if (linkTarget.getAttribute('href') !== href) {
      linkTarget.setAttribute('href', href);
      linkTarget.title = `${href} (Ctrl+click to open)`;
      if (linkTarget.dataset.link !== 'new') linkTarget.dataset.link = 'changed';
    }
    done = 'Link updated';
  } else if (linkRange.collapsed) {
    // New text, shown as the link: typed, so tracked while tracking.
    sel.removeAllRanges();
    sel.addRange(linkRange);
    const text = linkDialog.text.trim() || linkDialog.address.trim();
    const a = newLink(href, [document.createTextNode(text)]);
    const tracked = trackedBy(linkRange.startContainer);
    if (tracked) insertTracked([a], tracked);
    else insertAtCaret([a]);
    done = tracked ? 'Link inserted (tracked)' : 'Link inserted';
  } else {
    // The selected text becomes the link (links inside it are taken out first, as in Word).
    const frag = linkRange.extractContents();
    for (const inner of [...frag.querySelectorAll('a')]) inner.replaceWith(...inner.childNodes);
    const a = newLink(href, [...frag.childNodes]);
    linkRange.insertNode(a);
    const after = document.createRange();
    after.setStartAfter(a);
    after.collapse(true);
    sel.removeAllRanges();
    sel.addRange(after);
    done = 'Link added';
  }
  linkRange = null;
  onInput();
  undoStack.commit();
  message.value = done; // after onInput, which clears the status
  if (done === 'Link added') noteUntrackedFormatting();
}

function removeLink() {
  linkDialog.open = false;
  page.value.focus();
  if (!linkTarget?.isConnected) return;
  undoStack.flush();
  const first = linkTarget.firstChild;
  linkTarget.replaceWith(...linkTarget.childNodes);
  if (first) {
    const r = document.createRange();
    r.setStartBefore(first);
    r.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }
  linkRange = null;
  onInput();
  undoStack.commit();
  message.value = 'Link removed – its text stays';
}

// ------------------------------------------------------------------ footnotes and endnotes

/**
 * The note (footnote or endnote, below the text), header, footer or text box that `node` is in, or
 * null: text with its own editing host, apart from the body's.
 */
function noteAt(node) {
  const el = node?.nodeType === 1 ? node : node?.parentElement;
  return el?.closest('.docx-note, .docx-hf-part, .textbox-content') || null;
}

/** "a footnote", "a header", "a text box": where the caret is, for messages. */
function storyName(el) {
  if (el.classList.contains('docx-note')) return `a ${el.dataset.note}`;
  if (el.classList.contains('docx-hf-part')) return `a ${el.dataset.which}`;
  return 'a text box';
}

/** The blocks of the text itself: not in headers, footers, notes or text boxes. */
function bodyBlocks() {
  return [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((b) => !b.closest('.docx-notes, .docx-hf-band, .textbox'));
}

/** "footnote 2", for messages. */
function noteName(note) {
  return `${note.dataset.note} ${note.dataset.n}`;
}

/**
 * Clicking a note's number in the text goes to the note, to edit it; its number in the note goes
 * back to the text, as in Word. False if there is nowhere to go.
 */
function goToNote(ref) {
  const sel = window.getSelection();
  if (ref.classList.contains('note-mark')) {
    const note = ref.closest('.docx-note');
    const target = [...page.value.querySelectorAll(`.note-ref[data-note-ref="${note.dataset.note}"][data-note-id="${CSS.escape(note.dataset.noteId)}"]`)]
      .find((r) => !r.closest('.docx-notes'));
    if (!target) return false;
    page.value.focus({ preventScroll: true });
    const range = document.createRange();
    range.setStartAfter(target);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    target.scrollIntoView({ block: 'center', behavior: 'smooth' });
    message.value = `Back in the text, after ${noteName(note)}`;
    return true;
  }
  const note = page.value.querySelector(`.docx-note[data-note="${ref.dataset.noteRef}"][data-note-id="${CSS.escape(ref.dataset.noteId)}"]`);
  if (!note || note.hidden) return false;
  note.focus({ preventScroll: true });
  const blocks = note.querySelectorAll(BLOCK_SELECTOR);
  placeCaret(blocks[blocks.length - 1] || note);
  note.scrollIntoView({ block: 'center', behavior: 'smooth' });
  const name = noteName(note);
  message.value = `${name[0].toUpperCase()}${name.slice(1)}: type to edit it – click its number to go back to the text`;
  return true;
}

/** Says why something can't go in a note, if the selection is in one: true then. */
function refuseInNote(range, what) {
  const note = range && noteAt(range.startContainer);
  if (!note) return false;
  message.value = `${what} can’t go in ${storyName(note)} here: add it in Word`;
  return true;
}

// ------------------------------------------------------------------ content controls

/** Applies a change to a control as one undoable step. */
function changeControl(el, change) {
  undoStack.flush();
  change();
  clearPlaceholder(el, (id) => doc.isPlaceholderStyle(id));
  undoStack.commit();
  markDirty();
}

function toggleCheckbox(el) {
  const { symbols } = doc.control(el);
  changeControl(el, () => {
    el.dataset.checked = el.dataset.checked === '1' ? '0' : '1';
    setControlText(el, el.dataset.checked === '1' ? symbols.on : symbols.off);
  });
}

function openChooser(el) {
  const info = doc.control(el);
  const box = el.getBoundingClientRect();
  chooserTarget = el;
  Object.assign(chooser, {
    open: true,
    kind: info.kind,
    x: Math.max(8, Math.min(box.left, window.innerWidth - 240)),
    y: box.bottom + 4,
    choices: info.choices,
    current: el.classList.contains('showing-placeholder') ? '' : controlBox(el).textContent,
    date: info.date?.iso || new Date().toISOString().slice(0, 10),
  });
  nextTick(() => (dateInput.value || chooserEl.value?.querySelector('[aria-selected="true"], button'))?.focus());
}

function closeChooser(refocus = false) {
  const el = chooserTarget;
  chooser.open = false;
  chooserTarget = null;
  if (refocus) {
    // Back to the document, with the caret just after the field.
    page.value.focus();
    if (el?.isConnected) {
      const range = document.createRange();
      range.setStartAfter(el);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }
}

/** ↑ / ↓ / Home / End move between the items of a drop-down list. */
function onChooserKeydown(e) {
  if (chooser.kind === 'date' || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
  e.preventDefault();
  const items = [...chooserEl.value.querySelectorAll('button')];
  const i = items.indexOf(document.activeElement);
  const next = { ArrowDown: Math.min(i + 1, items.length - 1), ArrowUp: Math.max(i - 1, 0), Home: 0, End: items.length - 1 }[e.key];
  items[next]?.focus();
}

// A date typed digit by digit is "complete" (and fires change) before the year is finished, so typed
// dates are taken on Enter; a date picked from the calendar is taken at once.
let lastDateKey = 0;
function onDateKeydown(e) {
  lastDateKey = Date.now();
  if (e.key === 'Enter') {
    e.preventDefault();
    chooseDate(e.target.value);
  }
}
function onDateChange(e) {
  if (Date.now() - lastDateKey > 400) chooseDate(e.target.value);
}

/** Alt+↓, or Space on a selected field: tick a check box, or open a list or date picker. */
function activateField(el) {
  if (el.matches(FORM_CHOICE)) activateFormField(el);
  else if (doc.control(el).kind === 'checkbox') toggleCheckbox(el);
  else openChooser(el);
}

const canMarkFields = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

/** Marks the field the caret is touching, so it's clear that Alt+↓ will act on it. */
function markNearField(range) {
  if (!canMarkFields) return;
  const field = fieldAtSelection(range, page.value);
  if (!field) {
    CSS.highlights.delete('field-near');
    return;
  }
  const r = document.createRange();
  r.selectNodeContents(field);
  CSS.highlights.set('field-near', new Highlight(r));
}

function onDocumentMousedown(e) {
  if (chooser.open && !chooserEl.value?.contains(e.target) && !chooserTarget?.contains(e.target)) closeChooser();
  if (palette.open && !paletteEl.value?.contains(e.target) && !paletteButton.value?.contains(e.target)) palette.open = false;
  if (picturePanel.open && !picturePanelEl.value?.contains(e.target) && !layoutOptionsButton.value?.contains(e.target)) picturePanel.open = false;
  if (splitDialog.open && !splitDialogEl.value?.contains(e.target)) splitDialog.open = false;
  if (paraDialog.open && !paraDialogEl.value?.contains(e.target)) paraDialog.open = false;
  if (pageDialog.open && !pageDialogEl.value?.contains(e.target) && !pageSetupButton.value?.contains(e.target)) pageDialog.open = false;
  if (spacingMenu.open && !spacingMenuEl.value?.contains(e.target) && !spacingButton.value?.contains(e.target)) spacingMenu.open = false;
  if (tableMenu.open && !tableMenuEl.value?.contains(e.target) && !layoutButton.value?.contains(e.target)) tableMenu.open = false;
  if (tablePicker.open && !tablePickerEl.value?.contains(e.target) && !tableButton.value?.contains(e.target)) tablePicker.open = false;
  if (cellFormat.open && !cellFormatEl.value?.contains(e.target)) cellFormat.open = false;
  if (dropOffer.file && !dropOfferEl.value?.contains(e.target)) dropOffer.file = null;
  if (saveMenu.open && !saveMenuEl.value?.contains(e.target) && !saveMenuButton.value?.contains(e.target)) saveMenu.open = false;
  if (insertMenu.open && !insertMenuEl.value?.contains(e.target) && !insertButton.value?.contains(e.target)) insertMenu.open = false;
  if (caseMenu.open && !caseMenuEl.value?.contains(e.target) && !caseButton.value?.contains(e.target)) caseMenu.open = false;
  if (pageNumberMenu.open && !pageNumberMenuEl.value?.contains(e.target) && !pageNumberButton.value?.contains(e.target)) pageNumberMenu.open = false;
  if (numDialog.open && !numDialogEl.value?.contains(e.target) && !pageNumberButton.value?.contains(e.target)) numDialog.open = false;
  if (listLibrary.open && !listLibraryEl.value?.contains(e.target) && !listLibraryButton.value?.contains(e.target)) listLibrary.open = false;
  if (sortDialog.open && !sortDialogEl.value?.contains(e.target) && !layoutButton.value?.contains(e.target)) sortDialog.open = false;
  if (formulaDialog.open && !formulaDialogEl.value?.contains(e.target) && !layoutButton.value?.contains(e.target)) formulaDialog.open = false;
  if (ctxMenu.open && !ctxMenuEl.value?.contains(e.target)) ctxMenu.open = false;
  if (shareReady.file && !shareReadyEl.value?.contains(e.target)) shareReady.file = null;
  if (pdfDialog.open && !pdfDialogEl.value?.contains(e.target)) pdfDialog.open = false;
  if (changeMenu.open && !changeMenuEl.value?.contains(e.target) && !acceptMenuButton.value?.contains(e.target) && !rejectMenuButton.value?.contains(e.target)) changeMenu.open = null;
}

function choose(choice) {
  const el = chooserTarget;
  closeChooser(true);
  if (!el?.isConnected) return;
  if (el.matches(FORM_CHOICE)) {
    changeControl(el, () => doc.chooseFormEntry(el, choice.value));
    if (lockMode.value === 'forms' && textLocked.value) el.focus();
    return;
  }
  changeControl(el, () => {
    setControlText(el, choice.text);
    el.dataset.value = choice.value;
  });
}

function chooseDate(iso) {
  const el = chooserTarget;
  if (!iso || !el?.isConnected) return;
  closeChooser(true);
  changeControl(el, () => {
    setControlText(el, doc.formatDate(el, iso));
    el.dataset.fullDate = iso;
  });
}

// ------------------------------------------------------------------ legacy form fields

/** The legacy text form field (`.ff-text`) holding `node`, if any. */
function formTextAt(node) {
  const el = (node?.nodeType === 1 ? node : node?.parentElement)?.closest('.ff-text');
  return el && page.value.contains(el) ? el : null;
}

/** Ticks or clears a legacy check box, or opens a legacy drop-down list, as one undoable step. */
function activateFormField(el) {
  if (textLocked.value && lockMode.value !== 'forms') return explainLock();
  const info = doc.formField(el);
  if (!info) return;
  if (!info.enabled) {
    message.value = `${el.title}`;
    return;
  }
  if (readOnlyAt(el.parentElement, page.value)) {
    message.value = 'This field is in content kept as it is: fill it in in Word';
    return;
  }
  if (info.kind === 'checkbox') {
    changeControl(el, () => doc.toggleFormCheck(el));
    message.value = `${info.name || 'Check box'} ${info.checked ? 'cleared' : 'ticked'}`;
    return;
  }
  const box = el.getBoundingClientRect();
  chooserTarget = el;
  Object.assign(chooser, {
    open: true,
    kind: 'dropdown',
    x: Math.max(8, Math.min(box.left, window.innerWidth - 240)),
    y: box.bottom + 4,
    choices: info.entries.map((text, value) => ({ text, value })),
    current: info.entries[info.selected] ?? '',
  });
  nextTick(() => chooserEl.value?.querySelector('[aria-selected="true"], button')?.focus());
}

/**
 * A text form field still holding Word's blank result (five spaces) is selected when clicked, so
 * typing replaces it, as in Word.
 */
function selectBlankFormText(target) {
  const el = formTextAt(target);
  if (!el || el.textContent.trim() || !window.getSelection().isCollapsed) return;
  selectFormText(el);
}

function selectFormText(el) {
  const range = document.createRange();
  range.selectNodeContents(el);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

/**
 * Typing over, or deleting, the whole text of a legacy text form field: the browser would take the
 * field's element (and, while only filling in forms is allowed, the place to type) with the text, so
 * the text is replaced here, in the field. Deleted, it gets Word's blank (five spaces) back. True if
 * the input was handled. While tracking, the browser's own editing (tracked) does it.
 */
function replaceFormText(e) {
  const deleting = e.inputType.startsWith('delete');
  if (e.inputType !== 'insertText' && e.inputType !== 'insertReplacementText' && !deleting) return false;
  const range = selectionInPage();
  const el = range && formTextAt(range.startContainer);
  if (!el || !el.contains(range.endContainer) || trackedBy(range.startContainer)) return false;
  const all = el.textContent;
  const blank = !all.trim();
  // A caret in a blank field types over the blank; otherwise, only a selection of all its text.
  if (!(blank || (!range.collapsed && range.toString() === all))) return false;
  const text = deleting ? '\u2002'.repeat(5) : (e.data ?? e.dataTransfer?.getData('text/plain') ?? '');
  if (!text) return false;
  e.preventDefault();
  undoStack.flush();
  setControlText(el, text);
  const node = [...el.childNodes].find((n) => n.nodeType === 3) || el.querySelector('*')?.firstChild;
  const sel = window.getSelection();
  const after = document.createRange();
  if (node?.nodeType === 3) {
    if (deleting) after.selectNodeContents(node);
    else after.setStart(node, node.data.length);
  } else after.selectNodeContents(el);
  sel.removeAllRanges();
  sel.addRange(after);
  undoStack.commit();
  onInput();
  return true;
}

/** The fields Tab moves between while only filling in forms is allowed, in the page's order. */
function fillableFields() {
  return [...page.value.querySelectorAll('.ff-text[data-ff-host], .ff-check:not([data-ff-off]), .ff-drop:not([data-ff-off]), .sdt:not(.sdt-locked) > .sdt-content')]
    .filter((el) => !el.closest('.docx-hf-band, [hidden]'));
}

function focusField(el) {
  if (el.matches(FORM_CHOICE)) {
    el.focus();
    el.scrollIntoView({ block: 'nearest' });
    message.value = `${el.title.replace(/ – .*/, '')}: ${el.matches('.ff-check') ? 'Space ticks or clears it' : 'Space or Alt+↓ opens its list'}`;
    return;
  }
  el.focus();
  selectFormText(el);
  el.scrollIntoView({ block: 'nearest' });
}

/**
 * Keys for legacy form fields. While only filling in forms is allowed, Tab and Shift+Tab (and Enter
 * in a text form field) go to the next or previous field, as in a form protected in Word; on a
 * check box or list, Space, Enter or Alt+↓ ticks it or opens it. True if the key was used.
 */
function formKeydown(e) {
  const mod = e.ctrlKey || e.metaKey;
  const choice = document.activeElement?.matches?.(FORM_CHOICE) && page.value.contains(document.activeElement) ? document.activeElement : null;
  if (choice && !mod && (e.key === ' ' || e.key === 'Enter' || (e.altKey && e.key === 'ArrowDown'))) {
    e.preventDefault();
    activateFormField(choice);
    return true;
  }
  if (!(textLocked.value && lockMode.value === 'forms')) return false;
  const range = selectionInPage();
  const inText = range && formTextAt(range.startContainer);
  if ((e.key === 'Tab' || (e.key === 'Enter' && inText)) && !mod && !e.altKey) {
    const fields = fillableFields();
    if (!fields.length) return false;
    e.preventDefault();
    const here = choice || inText || (range && controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content'));
    const i = fields.indexOf(here);
    const step = e.shiftKey ? -1 : 1;
    const next = i < 0 ? (step > 0 ? 0 : fields.length - 1) : (i + step + fields.length) % fields.length;
    focusField(fields[next]);
    return true;
  }
  return false;
}

// ------------------------------------------------------------------ comments

/**
 * Top-level comments in the order they appear in the text, each with its replies, leaving out those
 * whose text was deleted (ids in `removed`).
 */
function commentThreads(removed = new Set()) {
  // As shown: with the text and resolved state kept on their marks in the page (see commentMark).
  const comments = doc.comments.filter((c) => !removed.has(c.id) && (!doc.storyOf(c.id) || doc.storyCommentLive(c))).map((c) => {
    const mark = commentMark(c.id);
    // In a header, footer or note: its changes are kept by the document (see doc.story).
    const place = mark ? null : doc.storyPlace(c.id);
    const text = mark ? mark.dataset.edit : place ? doc.story.edits[c.id] : undefined;
    const done = mark?.dataset.done != null ? mark.dataset.done === '1' : place && doc.story.done[c.id] != null ? doc.story.done[c.id] : c.done;
    return {
      ...c,
      anchored: !!mark || !!place,
      place,
      paragraphs: text != null ? text.split('\n').map((s) => s.trim()).filter(Boolean) : c.paragraphs,
      done,
    };
  });
  const ids = new Set(comments.map((c) => c.id));
  const order = (c) => doc.commentOrder.get(c.id) ?? Infinity;
  return comments
    .filter((c) => !c.parentId || !ids.has(c.parentId))
    .sort((a, b) => order(a) - order(b))
    .map((c) => ({ ...c, n: doc.commentOrder.get(c.id) || null, replies: comments.filter((r) => r.parentId === c.id) }));
}

// As in Word, deleting the text a comment is attached to deletes the comment. The panel follows the
// text, and the status bar says so, since Undo is the way back.
let removedComments = new Set();

/** A comment's mark in the page: it carries the comment's edited text and resolved state. */
function commentMark(id) {
  return page.value?.querySelector(`[data-comment-ref="${CSS.escape(id)}"]`) || null;
}

function refreshComments() {
  if (!doc.comments.length) return;
  doc.renumberComments(page.value); // marks may have been added, removed or brought back by Undo
  const removed = doc.removedComments(page.value);
  const added = [...removed].filter((id) => !removedComments.has(id));
  removedComments = removed;
  threads.value = commentThreads(removed); // edits and resolved state live on the marks, so always
  paintComments();
  if (added.length) message.value = commentRemovalMessage(added);
}

function commentRemovalMessage(ids) {
  const removed = doc.comments.filter((c) => ids.includes(c.id));
  const top = removed.filter((c) => !c.parentId || !ids.includes(c.parentId));
  const what = top.length === 1
    ? `the comment by ${top[0].author}${removed.length > 1 ? ` and ${removed.length - 1 === 1 ? 'its reply' : `its ${removed.length - 1} replies`}` : ''}`
    : `${removed.length} comments`;
  return `Deleted ${what} with the text ${removed.length === 1 ? 'it was' : 'they were'} attached to – Undo (Ctrl+Z) brings ${removed.length === 1 ? 'it' : 'them'} back`;
}

/** The commented text: from the range start marker to the end marker (or the reference mark). */
function commentRange(id) {
  const q = (sel) => page.value.querySelector(`[${sel}="${CSS.escape(id)}"]`);
  const start = q('data-comment-start');
  const end = q('data-comment-end');
  const ref = q('data-comment-id');
  if (!start && !ref) return null;
  const range = document.createRange();
  if (start) range.setStartAfter(start);
  else range.setStartBefore(ref);
  if (end) range.setEndBefore(end);
  else range.setEndAfter(ref);
  return range;
}

const canHighlight = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

function paintComments() {
  if (!canHighlight) return;
  const all = new Highlight();
  const active = new Highlight();
  active.priority = 1;
  for (const t of threads.value) {
    const range = commentRange(t.id);
    if (!range) continue;
    all.add(range);
    if (t.id === activeComment.value) active.add(range);
  }
  if (composer.open && composerRange) active.add(composerRange); // the text a new comment will be on
  CSS.highlights.set('comment', all);
  CSS.highlights.set('comment-active', active);
}

function selectComment(id, reveal = true) {
  activeComment.value = id;
  paintComments();
  if (reveal) {
    const range = commentRange(id);
    if (range) revealRect(range.getBoundingClientRect());
  } else {
    nextTick(() => document.querySelector('.comments .thread.active')?.scrollIntoView({ block: 'nearest' }));
  }
}

/**
 * Previous and Next comment, as Word's Review tab: from the comment shown in the panel, or else from
 * the caret, to the one before or after it in the text (round to the other end), its text selected
 * and its thread shown in the panel. Comments in headers, footers and notes are in the panel only.
 */
function gotoComment(dir) {
  const anchored = threads.value.map((t) => ({ id: t.id, range: commentRange(t.id) })).filter((t) => t.range);
  if (!anchored.length) {
    message.value = threads.value.length ? 'The comments are in headers, footers or notes: they are listed in the panel' : 'There are no comments to go to';
    if (threads.value.length) commentsOpen.value = true;
    return;
  }
  let i = anchored.findIndex((t) => t.id === activeComment.value);
  const caret = selectionInPage() || savedRange;
  if (i >= 0) i = (i + dir + anchored.length) % anchored.length;
  else if (caret) {
    // The first starting after the caret, or the last ending before it (START_TO_END compares a
    // range's end with the caret's start).
    const after = (t) => t.range.compareBoundaryPoints(Range.START_TO_START, caret) > 0;
    const before = (t) => t.range.compareBoundaryPoints(Range.START_TO_END, caret) < 0;
    if (dir > 0) i = anchored.findIndex(after);
    else i = anchored.findLastIndex(before);
    if (i < 0) i = dir > 0 ? 0 : anchored.length - 1;
  } else i = dir > 0 ? 0 : anchored.length - 1;
  const { id, range } = anchored[i];
  commentsOpen.value = true;
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  savedRange = range.cloneRange();
  selectComment(id);
  const t = threads.value.find((x) => x.id === id);
  message.value = `Comment ${i + 1} of ${anchored.length}${t?.author ? `, by ${t.author}` : ''}`;
  nextTick(() => document.querySelector('.comments .thread.active')?.scrollIntoView({ block: 'nearest' }));
}

// ------------------------------------------------------------------ markup view

// How tracked changes show (`markup`, declared with the lock), as Word's Display for Review: 'all'
// (all markup), 'simple' (the text as it will be, with a line in the margin beside each change),
// 'none' (no markup) or 'original' (the text before the changes: insertions hidden, deletions shown
// as ordinary text). Only the page's look changes (see .docx-sheet[data-markup] in the styles); the
// changes and the file stay as they are, and every document opens with all markup, so no change
// goes unseen. The original is shown read-only, as typing into it would edit text that is no
// longer there.
const MARKUP_NAMES = { all: 'All markup', simple: 'Simple markup', none: 'No markup', original: 'Original' };

function setMarkup(value) {
  if (!MARKUP_NAMES[value] || value === markup.value) return;
  const wasOriginal = markup.value === 'original';
  markup.value = value;
  if (wasOriginal || value === 'original') applyLock(textLocked.value);
  const n = changes.total;
  const kept = n ? ` – ${count(n, 'tracked change')} kept in the file` : '';
  message.value = value === 'all'
    ? 'All markup: tracked changes are shown'
    : value === 'original'
      ? `Original: the text before the changes, read-only${kept}`
      : `${MARKUP_NAMES[value]}: the text as it will be once the changes are accepted${kept}`;
  nextTick(() => { pageChanged(); schedulePages(0); });
}

/** Back to all markup, to show a tracked change being gone to, accepted or rejected. */
function showAllMarkup() {
  if (markup.value === 'all') return;
  const wasOriginal = markup.value === 'original';
  markup.value = 'all';
  if (wasOriginal) applyLock(textLocked.value);
  nextTick(() => { pageChanged(); schedulePages(0); });
}

function revealRect(rect) {
  const canvas = page.value.closest('.canvas');
  const box = canvas.getBoundingClientRect();
  if (rect.top < box.top + 8 || rect.bottom > box.bottom - 8) canvas.scrollTop += rect.top - box.top - box.height / 3;
}

// ------------------------------------------------------------------ find and replace

/** Ctrl+F / Ctrl+H anywhere in the editor open Find (and Replace) instead of the browser's own. */
function onWindowKeydown(e) {
  if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.code === 'Slash' || e.key === '/')) {
    e.preventDefault(); // Help, as in Google Docs and many editors
    if (!reviewData.value) helpOpen.value = !helpOpen.value;
    return;
  }
  if (e.key === 'F10' && e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey) {
    e.preventDefault(); // to the toolbar, as in Word and WAI-ARIA's toolbar pattern; Esc comes back
    toolbarNav?.focus();
    return;
  }
  // Typing into a locked document: nothing changes (nothing is editable), so say why.
  if (textLocked.value && !e.defaultPrevented && isEditingKey(e) && !document.activeElement?.matches('input, textarea, select, [contenteditable="true"]')
    && (selectionInPage() || document.activeElement === document.body)) {
    explainLock();
    return;
  }
  if (explainKeyOutsideText(e)) return;
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  if (mod && !e.shiftKey && !e.altKey && key === 'o') {
    e.preventDefault(); // Word's Open: another document, in a new tab (the browser's would leave the app)
    openAnother();
    return;
  }
  if (mod && e.shiftKey && !e.altKey && key === 'e') {
    e.preventDefault(); // Word's Track Changes shortcut, wherever the focus is in the editor
    toggleTracking();
    return;
  }
  if (!mod || e.altKey || e.shiftKey || (key !== 'f' && key !== 'h')) return;
  e.preventDefault();
  openFind(key === 'h');
}

async function openFind(replace) {
  const range = selectionInPage() || savedRange;
  const selected = range && !range.collapsed ? range.toString() : '';
  const text = selected && selected.length <= 200 && !/\n/.test(selected) ? selected : '';
  // Re-pressing Ctrl+F / Ctrl+H while Find is open keeps the current match.
  const from = range && !findOpen.value ? { node: range.startContainer, offset: range.startOffset } : null;
  fillIn.open = false; // both paint the same highlights
  findOpen.value = true;
  await nextTick();
  findBar.value?.open({ replace, text, from });
}

function closeFind() {
  findOpen.value = false;
  page.value.focus();
}

/** What Find searches, in order: the page (its header, the text, notes and footer), then comment texts. */
function searchRoots() {
  return [page.value, ...document.querySelectorAll('.comments .comment .text')];
}

/** Brings a match into view, switching to its header/footer tab or opening Comments if need be. */
async function revealRange(range) {
  const el = range.startContainer.parentElement;
  const variant = el?.closest('.docx-hf-part');
  if (variant && hf[variant.dataset.which] !== Number(variant.dataset.variant)) hf[variant.dataset.which] = Number(variant.dataset.variant);
  if (el?.closest('.comments') && !commentsOpen.value) commentsOpen.value = true;
  await nextTick();
  const scroller = el?.closest('.canvas, .comments');
  if (!scroller) return;
  const rect = range.getBoundingClientRect();
  const box = scroller.getBoundingClientRect();
  if (rect.top < box.top + 8 || rect.bottom > box.bottom - 8) scroller.scrollTop += rect.top - box.top - box.height / 3;
}

/** Closing Find leaves the current match selected in the document, as in Word. */
function onFindClosed(range) {
  closeFind();
  if (range && page.value.contains(range.startContainer)) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

/** Runs a change made outside the browser's editing (such as Replace) as one undoable edit. */
function applyEdit(change) {
  undoStack.flush();
  const touched = change() || [];
  for (const node of touched) {
    const control = controlAt(node, page.value);
    if (control) clearPlaceholder(control, (id) => doc.isPlaceholderStyle(id));
  }
  page.value.normalize();
  undoStack.commit();
  markDirty();
}

// ------------------------------------------------------------------ other documents, in other tabs

// Another document opens in a new tab, beside this one (see tabs.js): the Open button and Ctrl+O
// (Word's Open) open the start page there; a document dropped on the page is offered to open
// there, and handed over in memory. This tab holds its draft's lock, so the start page in another
// tab shows the draft as open here rather than offering to restore a second copy.
const dropOffer = reactive({ file: null });
const dropOfferEl = ref(null);
let releaseDraft = () => {};
onMounted(() => { releaseDraft = holdDraft(draftId); });
onBeforeUnmount(() => releaseDraft());

function openAnother() {
  if (!openStartInNewTab()) message.value = 'The browser didn’t open a new tab: allow pop-ups for this site, or open Intact in a new tab yourself';
}
function offerNewTab(file) {
  dropOffer.file = file;
  nextTick(() => dropOfferEl.value?.querySelector('.confirm')?.focus());
}
function openDroppedInNewTab() {
  const file = dropOffer.file;
  dropOffer.file = null;
  const result = openFileInNewTab(file);
  message.value = {
    opened: `“${file.name}” is opening in a new tab`,
    blocked: 'The browser didn’t open a new tab: allow pop-ups for this site, or open the file from a new tab',
    unsupported: `A new tab is open: choose “${file.name}” there`,
  }[result];
  page.value?.focus();
}

// ------------------------------------------------------------------ save as PDF

// Download's menu: the Word file, or a PDF. The browser makes the PDF, from its print window (there
// is nothing to upload and no PDF library to load): a short dialog says how, and what the PDF will
// show that the sender may not want in it (tracked changes, highlights). Printing sets the headers
// and footers on every page and suggests the file's name (see onBeforePrint).
const saveMenu = reactive({ open: false, x: 0, y: 0 });
const saveMenuEl = ref(null);
const saveMenuButton = ref(null);
const pdfDialog = reactive({ open: false, x: 0, y: 0 });
const pdfDialogEl = ref(null);
const hasHighlight = ref(false);
// Hidden text (w:vanish) isn't printed unless asked, as in Word: a test prints without its answers,
// its answer key with them.
const hasHidden = ref(false);
const printHidden = ref(false);
const pdfBands = ref({}); // what the dialog says about the header and footer

/** The header and footer in the PDF, in words: on every page, or once (and why). */
function describeBands() {
  const modes = printModes(page.value, doc, sheetZoom.value);
  const names = (mode) => ['header', 'footer'].filter((w) => modes[w] === mode).join(' and ');
  const once = names('once');
  return {
    every: names('every'),
    once,
    onceWhere: once === 'header' ? 'at the start' : once === 'footer' ? 'at the end' : 'at the start and the end',
    why: modes.why,
  };
}

function toggleSaveMenu() {
  if (saveMenu.open) return closeSaveMenu();
  const rect = saveMenuButton.value.getBoundingClientRect();
  Object.assign(saveMenu, { open: true, x: Math.max(8, Math.min(rect.right - 220, window.innerWidth - 228)), y: rect.bottom + 4 });
  nextTick(() => {
    keepInView(saveMenu, saveMenuEl.value);
    saveMenuEl.value?.querySelector('button')?.focus();
  });
}
function closeSaveMenu(refocus = false) {
  saveMenu.open = false;
  if (refocus) saveMenuButton.value?.focus();
}
function saveMenuAction(what) {
  closeSaveMenu();
  if (what === 'docx') requestDownload();
  else if (what === 'copy' || what === 'saveAs') requestDownload(what);
  else if (what === 'share') requestShare();
  else if (what === 'merge') openMailMerge();
  else if (what === 'password') openPasswordDialog();
  else if (what === 'web' || what === 'markdown') copyForWeb(what === 'web' ? 'html' : 'markdown');
  else openPdfDialog();
}

// ------------------------------------------------------------------ mail merge

// Mail merge from a CSV (see docx/merge.js and MailMergeDialog): the document as it would be
// downloaded, with each row of a list in its placeholders, as one document or a .zip of one for
// each row. The document itself is left as it is.
const mailMerge = reactive({ open: false, placeholders: [], singleBlocked: null });

function openMailMerge() {
  fillIn.open = false;
  const all = findPlaceholders(page.value);
  const notes = page.value.querySelector('.note-ref[data-note-ref]');
  Object.assign(mailMerge, {
    open: true,
    placeholders: all.map((p) => ({ text: p.text, label: p.label })),
    singleBlocked: notes ? 'it has footnotes or endnotes, which each copy would refer to' : null,
  });
}

function closeMailMerge() {
  mailMerge.open = false;
  refocus();
}

async function runMailMerge({ rows, mapping, mode, nameColumn }) {
  await draftSettled();
  // The document itself, not encrypted: the merged documents are made from it on this device.
  const saved = await withHostsUnlocked(() => doc.save(page.value, { plain: true }));
  const bytes = await saved.arrayBuffer();
  if (mode === 'single') {
    const main = await (await JSZip.loadAsync(bytes)).file(doc.mainPath).async('string');
    const why = singleBlocker(main);
    if (why) throw new Error(`One document can’t be made of this one: ${why}. Choose a separate document for each row.`);
  }
  const parts = { main: doc.mainPath, others: [...doc.headers, ...doc.footers, ...Object.values(doc.noteParts)].map((v) => v.path).filter(Boolean) };
  const result = await mergeDocuments(bytes, {
    rows, placeholders: mailMerge.placeholders, mapping, parts, mode, nameColumn,
    baseName: baseName.value.trim() || 'document', mime: doc.format.mime,
  });
  downloadBlob(result.blob, result.name);
  const kept = result.keptInHeaders.length ? ` – in the header or footer, ${result.keptInHeaders.join(', ')} ${result.keptInHeaders.length === 1 ? 'is' : 'are'} left as written, as every copy shares it` : '';
  const unprotected = passwordProtected.value ? ' – not protected with a password: protect each document before sending it, if it needs one' : '';
  message.value = `Merged ${result.count} ${result.count === 1 ? 'row' : 'rows'}: ${result.name}${kept}${unprotected}`;
}

/** Saves `blob` as a download named `name`. */
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
/** Print (File tab): the browser's print window, as Ctrl+P opens it (see printBands and the print styles). */
function printDocument() {
  closeChooser();
  window.print();
}

function openPdfDialog() {
  updatePages();
  hasHighlight.value = !!page.value.querySelector('mark');
  hasHidden.value = !!page.value.querySelector('[data-hidden]');
  pdfBands.value = describeBands();
  const rect = saveMenuButton.value?.getBoundingClientRect();
  Object.assign(pdfDialog, { open: true, x: Math.max(8, Math.min((rect?.right ?? window.innerWidth) - 360, window.innerWidth - 368)), y: (rect?.bottom ?? 48) + 4 });
  nextTick(() => pdfDialogEl.value?.querySelector('.confirm')?.focus());
}
function closePdfDialog() {
  pdfDialog.open = false;
  page.value.focus();
  restoreSelection();
}
function printToPdf() {
  pdfDialog.open = false;
  // After the dialog has gone from the screen, so the print window opens on the page alone.
  // Hidden text printed only when asked (see the print styles), for this print only.
  document.documentElement.classList.toggle('print-hidden-text', hasHidden.value && printHidden.value);
  window.addEventListener('afterprint', () => document.documentElement.classList.remove('print-hidden-text'), { once: true });
  nextTick(() => requestAnimationFrame(() => window.print()));
}

// ------------------------------------------------------------------ file actions

// "Review my changes" before Download: what will go back to the recipient. It can be skipped; the
// choice is remembered in this browser only, and the toolbar's Review button still shows it.
const SKIP_REVIEW_KEY = 'intact.skipReview';
const reviewData = ref(null);
const reviewFocus = ref('');
const skipReview = ref((() => { try { return localStorage.getItem(SKIP_REVIEW_KEY) === '1'; } catch { return false; } })());

// Save to the file opened, where the browser gave a handle to it (see editor/fileHandles.js): the
// header's Save (Ctrl+S) writes into it, as Word's Save does; Download a copy and Save as… (in
// Download's menu) are still there. Without a handle, Download keeps a copy in Downloads.
const fileHandle = shallowRef(doc.fileHandle || null);
const canSaveAs = canPickSaveFile();
const pendingSave = ref(''); // what the review's button does: '' (Save or Download), 'copy' or 'saveAs'
const saveAction = computed(() => {
  if (pendingShare.value) return 'share';
  if (pendingSave.value === 'saveAs') return 'saveAs';
  return fileHandle.value && pendingSave.value !== 'copy' ? 'save' : 'download';
});

/** Save (Ctrl+S), or `kind` from Download's menu ('copy', 'saveAs'): through the review, unless skipped. */
function requestDownload(kind = '') {
  pendingShare.value = false;
  pendingSave.value = kind;
  if (skipReview.value) {
    const action = saveAction.value;
    pendingSave.value = '';
    save(action);
  } else openReview();
}

// ------------------------------------------------------------------ share

// Share… (Download's menu, where the browser can share files: Safari on iPhone and iPad, among
// others) hands the saved file to the device's share sheet, to send it back by mail or a messaging
// app without finding it in Downloads. The file goes from this page to the app chosen, on the
// device: nothing is uploaded by Intact. It goes through "Review your changes" as Download does.
const DOCX_MIME_TEST = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const canShareFiles = ref(false);
const pendingShare = ref(false);
const shareReady = reactive({ file: null });
const shareReadyEl = ref(null);

function detectShare() {
  try {
    canShareFiles.value = typeof navigator.share === 'function' && typeof navigator.canShare === 'function'
      && navigator.canShare({ files: [new File([''], 'test.docx', { type: doc.format.mime || DOCX_MIME_TEST })] });
  } catch {
    canShareFiles.value = false;
  }
}

function requestShare() {
  pendingSave.value = '';
  pendingShare.value = true;
  if (skipReview.value) {
    pendingShare.value = false;
    save('share');
  } else openReview();
}

/** Opens the share sheet with `file`. True once shared (not when cancelled or refused). */
async function shareFile(file) {
  shareReady.file = null;
  try {
    await navigator.share({ files: [file], title: file.name });
  } catch (e) {
    if (e?.name === 'AbortError') {
      message.value = `${file.name} wasn’t shared – Download keeps a copy on this device`;
      return false;
    }
    if (e?.name === 'NotAllowedError') {
      // Saving took too long for the browser to count the share as asked for by a tap: one more tap.
      shareReady.file = file;
      nextTick(() => shareReadyEl.value?.querySelector('.confirm')?.focus());
      return false;
    }
    console.error(e);
    message.value = `Couldn’t share ${file.name}: ${e?.message || e} – use Download instead`;
    return false;
  }
  dirty.value = false;
  discardDraft({ closing: false }); // sent on: as with Download, the file is the copy now
  message.value = `Shared ${file.name}`;
  return true;
}

// Help (Ctrl+/): the shortcuts, and the review's checks without downloading (see HelpDialog).
const helpOpen = ref(false);
function closeHelp() {
  helpOpen.value = false;
  refocus();
}
function checkFromHelp(part) {
  helpOpen.value = false;
  openReview(part);
}

/** Opens "Review your changes"; `part` ('release', 'accessibility' or 'screening'): opened at that part. */
async function openReview(part = '') {
  reviewFocus.value = typeof part === 'string' ? part : '';
  closeChooser();
  await draftSettled();
  undoStack.flush();
  const review = withHostsUnlocked(() => doc.review(page.value));
  // What happens to the restrictions and signatures in the file (see "restrictions" above).
  // Written as Download would, to know whether anything changed (signatures, properties).
  await withHostsUnlocked(() => doc.write(page.value));
  // Removing personal information, chosen for you where nothing would be lost by it; written again
  // when that changes what Download writes.
  const before = doc.removePersonal;
  choosePersonalDefault(inspection());
  if (doc.removePersonal !== before) await withHostsUnlocked(() => doc.write(page.value));
  const signature = doc.signatures ? { count: doc.signatures.count, removed: doc.changed } : null;
  // The page's elements stay here, out of Vue's reach; the dialog gets what to say about them.
  issues = checkAccessibility(page.value, doc);
  const screening = checkScreening(page.value, baseName.value);
  issues.push(...screening.issues);
  const forDialog = (list) => list.map(({ targets, ...issue }) => ({ ...issue, count: targets.length }));
  reviewData.value = {
    ...review, signature, protection: doc.protection, overridden: unlocked.value, restored: doc.restored, password: passwordReview(),
    startedHere: !!(doc.startedHere || doc.fromTemplate),
    label: doc.properties.label, marking: doc.properties.marking, inspection: inspection(),
    properties: doc.propertiesUpdate(),
    accessibility: forDialog(issues.filter((i) => !i.id.startsWith('ats-'))),
    screening: { cv: screening.cv, issues: forDialog(screening.issues) },
  };
  if (reviewData.value.fields) reviewData.value.fields.updatedHere = fieldsUpdatedAt === page.value.innerHTML;
}

// ------------------------------------------------------------------ accessibility and properties

let issues = []; // the last accessibility and screening checks (see editor/accessibility.js, screening.js)
const shown = new Map(); // issue id -> the target shown last, to show the next one each time

/** Shows the next place an accessibility issue was found, closing the review; pictures open their alt text. */
async function showIssue(id) {
  const issue = issues.find((i) => i.id === id);
  const live = issue?.targets.filter((el) => el.isConnected) || [];
  if (!live.length) return;
  const i = ((shown.get(id) ?? -1) + 1) % live.length;
  shown.set(id, i);
  const el = live[i];
  closeReview();
  if (el.matches('.docx-hf-part *')) {
    const part = el.closest('.docx-hf-part');
    hf[part.dataset.which] = Number(part.dataset.variant);
    await nextTick();
  }
  el.scrollIntoView({ block: 'center' });
  const where = live.length > 1 ? ` (${i + 1} of ${live.length})` : '';
  if (el.matches('.picture') && !lockMode.value) {
    showPicture(el);
    await nextTick();
    if (issue.id === 'alt') {
      if (!picturePanel.open) togglePicturePanel();
      await nextTick();
      picturePanelEl.value?.querySelector('.alt textarea')?.focus();
    }
  } else if (el.matches('table')) {
    const cell = el.querySelector('td');
    if (cell) placeCaret(cell.querySelector(BLOCK_SELECTOR) || cell, true);
  } else if (!el.matches('.picture')) {
    placeCaret(el, true);
  }
  message.value = `${issue.title}${where}`;
}

/** The title written into the document's properties on saving (Before you release › No title). */
async function setTitle(title) {
  if (textLocked.value) return explainLock();
  doc.newTitle = title.trim() || null;
  const core = doc.properties.core.filter((p) => p.label !== 'Title');
  if (doc.newTitle) core.unshift({ label: 'Title', value: doc.newTitle });
  doc.properties.core = core;
  markDirty();
  await openReview();
}

// Recording an edit in the document's properties, as Word does: on unless switched off, which this
// browser remembers. Who edited is the name given for comments and tracked changes.
const PROPERTIES_KEY = 'intact.updateProperties';
doc.updateProperties = (() => { try { return localStorage.getItem(PROPERTIES_KEY) !== '0'; } catch { return true; } })();
watch([author, named], () => { doc.modifiedBy = named.value ? author.value : ''; }, { immediate: true });

async function setUpdateProperties(on) {
  doc.updateProperties = on;
  try { localStorage.setItem(PROPERTIES_KEY, on ? '1' : '0'); } catch { /* this session only */ }
  await openReview();
}

function closeReview() {
  reviewData.value = null;
  pendingShare.value = false;
  pendingSave.value = '';
  backFromReview();
}

/**
 * Back in the text where you were, after the review: the selection as it was before it opened.
 * (Focusing the page alone puts the caret at its start, in a table's cell too.)
 */
function backFromReview() {
  const range = savedRange;
  page.value.focus();
  if (range?.startContainer.isConnected && page.value.contains(range.commonAncestorContainer)) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  } else {
    restoreSelection();
  }
}

function downloadReviewed() {
  reviewData.value = null;
  backFromReview(); // carry on editing where you were
  const action = saveAction.value;
  pendingShare.value = false;
  pendingSave.value = '';
  save(action);
}

function setSkipReview(on) {
  skipReview.value = on;
  try { localStorage.setItem(SKIP_REVIEW_KEY, on ? '1' : '0'); } catch { /* this session only */ }
}

/**
 * Saves the document: 'download' (or 'copy') to Downloads, 'share' to the share sheet, 'save' into
 * the file opened (its handle) and 'saveAs' into a file picked now, which later Saves write to.
 */
async function save(action = 'download') {
  if (saving.value) return;
  saving.value = true;
  try {
    let name = `${baseName.value.trim() || 'document'}.${fileExt.value}`;
    // A file renamed in the header is a new file: Save as…, where the browser can.
    if (action === 'save' && fileHandle.value.name !== name) action = canSaveAs ? 'saveAs' : 'download';
    // First, while the click still counts as asking: the file to write to, and leave to write it.
    let target = null;
    if (action === 'saveAs') {
      target = await pickFileToSave(name, fileExt.value);
      if (!target) {
        message.value = 'Not saved';
        return;
      }
    } else if (action === 'save') {
      target = fileHandle.value;
      if (!(await mayWrite(target))) {
        message.value = `Not saved: the browser didn’t allow writing to ${target.name} – Download a copy (▾) keeps one in Downloads`;
        return;
      }
    }
    await draftSettled();
    const blob = await withHostsUnlocked(() => doc.save(page.value));
    if (target) {
      await writeFile(target, blob);
      fileHandle.value = target;
      doc.fileHandle = target;
      name = target.name;
      baseName.value = name.replace(/\.(docx|dotx|docm|dotm)$/i, '');
    }
    doc.name = name;
    if (action === 'share') {
      saving.value = false;
      await shareFile(new File([blob], name, { type: doc.format.mime || blob.type }));
      return;
    }
    if (!target) downloadBlob(blob, name);
    dirty.value = false;
    discardDraft({ closing: false }); // downloaded or written: the file is the copy now
    const n = doc.commentsRemoved.size;
    message.value = `Saved ${name}${target ? ' on this device' : ''}${n ? ` – ${n} ${n === 1 ? 'comment' : 'comments'} removed with ${n === 1 ? 'its' : 'their'} text` : ''}`
      + (doc.signatures ? (doc.signaturesRemoved ? ' – digital signature removed: it no longer matched the document' : ' – unchanged, so its digital signature is still valid') : '')
      + (passwordProtected.value ? ' – encrypted with its password' : '');
  } catch (e) {
    console.error(e);
    message.value = `Save failed: ${e.message}`;
  } finally {
    saving.value = false;
  }
}

async function close() {
  const ok = await router.push({ name: 'home' });
  if (ok === undefined) setDocument(null); // navigation succeeded
}
</script>

<style scoped>
.editor-shell { height: 100vh; height: 100dvh; display: flex; flex-direction: column; }
/* The file's name gives way (cut off) before the header's buttons do, so a narrow screen never
   scrolls sideways, whatever the fonts. */
.file { display: flex; align-items: center; gap: 2px; min-width: 0; flex: 1; overflow: hidden; }
.filename {
  min-width: 0; max-width: 100%; font: inherit; color: #fff; background: transparent;
  border: 1px solid transparent; border-radius: 4px; padding: 4px 6px;
}
.filename:hover, .filename:focus { border-color: rgb(255 255 255 / 50%); outline: none; background: rgb(255 255 255 / 10%); }
.ext { opacity: .7; }
/* The sensitivity label, as Office shows it beside the name: read-only. */
.label-chip {
  display: inline-flex; align-items: center; gap: 5px; margin-left: 10px; padding: 2px 9px; min-width: 0; max-width: 40ch;
  border: 1px solid rgb(255 255 255 / 55%); border-radius: 999px; font-size: 12.5px; font-weight: 600; cursor: default; white-space: nowrap;
}
.label-chip svg { flex: none; }
.label-chip .text { overflow: hidden; text-overflow: ellipsis; }
.label-chip:focus-visible { outline: 2px solid #fff; outline-offset: 2px; }
.dirty { color: #ffd166; margin-left: 6px; }
.primary { background: #fff; color: var(--brand); border: 0; border-radius: 4px; padding: 6px 14px; font-weight: 600; cursor: pointer; }
.primary:disabled { opacity: .6; }
.open-tab { display: inline-flex; align-items: center; gap: 6px; }
.close-doc { display: inline-flex; align-items: center; gap: 6px; }
.close-doc svg { display: none; }
.chooser.drop-offer { left: 50%; top: 72px; transform: translateX(-50%); width: 340px; max-width: calc(100vw - 16px); max-height: none; font-size: 13px; }
.drop-offer p { margin: 0; line-height: 1.45; }
.password-dialog { display: flex; flex-direction: column; gap: 6px; }
.password-dialog .pw-note { color: var(--muted); }
.password-dialog label { font-weight: 600; margin-top: 4px; }
.password-dialog .pw-error { color: var(--danger); }
.password-dialog .pw-error:empty { display: none; }
/* A phone: Open and Close are their icons, so the file's name has room to be read. */
@media (max-width: 560px) {
  .open-tab, .close-doc { padding: 5px 8px; }
  .open-tab .label, .close-doc .label { display: none; }
  .close-doc svg { display: block; }
  .primary { padding: 6px 10px; }
}
/* On a phone the mark alone is the way to the start page, leaving room for the file and its buttons. */
@media (max-width: 420px) { :deep(.brand span) { display: none; } }
/* A narrow phone: the header's gaps and buttons tighten, so its buttons keep fitting with wider fonts. */
@media (max-width: 360px) {
  :deep(.app-header) { gap: 8px; padding-left: 10px; padding-right: 10px; }
  .primary { padding: 6px 10px; }
  .save-split .primary.arrow { padding: 6px 5px; }
  .ghost { padding: 5px 8px; }
}
.save-split { display: inline-flex; gap: 1px; }
.save-split .primary:first-child { border-radius: 4px 0 0 4px; }
.save-split .primary.arrow { border-radius: 0 4px 4px 0; padding: 6px 7px; font-weight: 400; }
.save-menu { min-width: 220px; }
.insert-menu { min-width: 220px; }
.insert-menu button { display: flex; justify-content: space-between; gap: 16px; }
.insert-menu .hint { color: var(--muted); font-size: 12px; }
.list-library { width: min(460px, calc(100vw - 16px)); }
.list-library button > :first-child, .list-library button { white-space: nowrap; }
.list-library .hint { white-space: normal; text-align: right; }
.formula-dialog input[type="text"] { font-family: ui-monospace, Consolas, monospace; min-width: 14em; }
.formula-dialog .note.error { color: var(--danger); }
.insert-menu .auto-date { margin-top: 4px; border-top: 1px solid var(--border); border-radius: 0; justify-content: flex-start; gap: 6px; flex-wrap: wrap; }
.insert-menu .auto-date .hint { flex-basis: 100%; padding-left: 20px; }
.page-number-menu { min-width: 260px; }
.page-number-menu .menu-sep { margin-top: 4px; border-top: 1px solid var(--border); border-radius: 0; }
.insert-menu .menu-label { margin: 6px 8px 2px; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }
.chooser.share-ready { left: 50%; top: 72px; transform: translateX(-50%); width: 300px; max-width: calc(100vw - 16px); padding: 12px; gap: 10px; font-size: 14px; }
.share-ready p { margin: 0; overflow-wrap: anywhere; }
.share-ready .actions { display: flex; gap: 8px; justify-content: flex-end; }
.share-ready .actions button { border: 1px solid var(--border); border-radius: 4px; padding: 6px 14px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.share-ready .actions .confirm { background: var(--brand); border-color: var(--brand); color: #fff; }
.save-menu button { display: flex; justify-content: space-between; gap: 16px; }
.save-menu .shortcut { color: var(--muted); font-size: 12px; }
.chooser.pdf-dialog { width: 360px; max-width: calc(100vw - 16px); max-height: calc(100vh - 16px); font-size: 13px; }
.pdf-dialog h2 { margin: 0; font-size: 15px; }
.pdf-dialog p { margin: 0; line-height: 1.45; }
.pdf-dialog .pdf-notes { margin: 0; padding-left: 18px; display: flex; flex-direction: column; gap: 4px; line-height: 1.4; }
.pdf-dialog .pdf-notes .warn::marker { color: #b26a00; }
.pdf-dialog .pdf-word { color: var(--muted); }
@media (max-width: 560px) { .ext, :deep(.nav) { display: none; } }
.ghost { background: transparent; color: #fff; border: 1px solid rgb(255 255 255 / 50%); border-radius: 4px; padding: 5px 12px; cursor: pointer; }

.toolbar {
  display: flex; flex-wrap: wrap; align-items: center; gap: 2px; padding: 6px 12px;
  background: var(--surface); border-bottom: 1px solid var(--border);
}
/* The ribbon's tabs, above its tools: one tab's tools at a time (a wider screen than a phone's). */
.ribbon-tabs { display: flex; align-items: flex-end; gap: 2px; padding: 4px 12px 0; background: var(--surface); }
.ribbon-tabs [role="tab"] {
  position: relative; padding: 6px 12px 7px; border: 0; border-bottom: 2px solid transparent; background: none; color: var(--muted);
  font: inherit; font-size: 13px; font-weight: 600; cursor: pointer; border-radius: 4px 4px 0 0;
}
.ribbon-tabs [role="tab"]:hover { background: var(--hover); color: var(--text); }
.ribbon-tabs [role="tab"][aria-selected="true"] { color: var(--accent); border-bottom-color: var(--accent); }
.ribbon-tabs [role="tab"]:focus-visible, .ribbon-mode:focus-visible { outline: 2px solid var(--accent); outline-offset: -2px; }
/* In a table, the Layout tab (which has Table layout) is marked, as Word shows its table tabs. */
.ribbon-tabs [role="tab"].context::after { content: ''; position: absolute; top: 5px; right: 4px; width: 6px; height: 6px; border-radius: 50%; background: var(--accent); }
.ribbon-tabs .tab-badge { margin-left: 6px; padding: 0 6px; border-radius: 9px; background: var(--active); color: var(--accent); font-size: 11px; }
.ribbon-spacer { flex: 1; }
.ribbon-mode { margin-bottom: 4px; padding: 3px 10px; border: 1px solid var(--border) !important; border-radius: 4px; background: var(--surface); color: var(--muted); font: inherit; font-size: 12px !important; cursor: pointer; }
.ribbon-mode:hover { background: var(--hover) !important; color: var(--text); }
.toolbar .ribbon-mode { margin: 0 0 0 auto; height: 28px !important; }
.toolbar.ribbon { min-height: 45px; box-sizing: border-box; }
/* One tab's tools on one row of a laptop's width: the style and font boxes a little narrower. */
.toolbar.ribbon select.style-choice { min-width: 0; width: 118px; }
.toolbar.ribbon select.font-choice { width: 120px; }
.toolbar.ribbon .sep { margin: 0 2px; }
/* The File and View tabs' own tools are the ribbon's: Download's menu (▾) and the status bar (zoom,
   Headings) have them in "All tools" and on a phone. */
.toolbar:not(.ribbon) [data-tab="file"], .toolbar:not(.ribbon) [data-tab="view"] { display: none !important; }
@media (max-width: 720px) { .toolbar [data-tab="file"], .toolbar [data-tab="view"] { display: none !important; } }
.toolbar select.zoom-choice { width: 78px; min-width: 0; }
@media (min-width: 721px) {
  /* A tool shows in the tabs data-tab names (a space-separated list, as ¶ is in Home and View). */
  .toolbar.ribbon[data-active="file"] [data-tab]:not([data-tab~="file"]),
  .toolbar.ribbon[data-active="home"] [data-tab]:not([data-tab~="home"]),
  .toolbar.ribbon[data-active="insert"] [data-tab]:not([data-tab~="insert"]),
  .toolbar.ribbon[data-active="layout"] [data-tab]:not([data-tab~="layout"]),
  .toolbar.ribbon[data-active="references"] [data-tab]:not([data-tab~="references"]),
  .toolbar.ribbon[data-active="review"] [data-tab]:not([data-tab~="review"]),
  .toolbar.ribbon[data-active="view"] [data-tab]:not([data-tab~="view"]) { display: none !important; }
  .toolbar.ribbon .help-toggle { margin-left: auto; }
}
/* A phone: its one row of every tool, scrolling sideways, whatever the choice here. */
@media (max-width: 720px) {
  .ribbon-tabs, .toolbar .ribbon-mode { display: none; }
}
.toolbar button {
  min-width: 30px; height: 32px; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid transparent; border-radius: 4px; background: transparent; color: inherit; font-size: 15px; cursor: pointer;
}
.toolbar button:hover:not(:disabled) { background: var(--hover); }
/* Groups of tools turned off together by a restriction; they don't change the layout. */
.toolbar fieldset.group { display: contents; border: 0; margin: 0; padding: 0; min-width: 0; }
.toolbar select:disabled { opacity: .4; }

/* Word's message bar for restricted and signed documents. */
.restriction-bar {
  --bar: #fff4ce; --bar-line: #e8d48a;
  display: flex; align-items: center; gap: 10px; padding: 6px 12px; font-size: 13px; line-height: 1.4;
  background: var(--bar); border-bottom: 1px solid var(--bar-line); color: var(--text);
}
.restriction-bar.unlocked { --bar: var(--hover); --bar-line: var(--border); }
.restriction-bar svg { flex: none; }
.restriction-bar p { margin: 0; flex: 1; }
.restriction-bar button { flex: none; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; padding: 3px 10px; cursor: pointer; }
.restriction-bar button:hover { background: var(--hover); }
/* The template bar's two buttons go below its text where there isn't room beside it (phones). */
.template-bar { flex-wrap: wrap; row-gap: 6px; }
.template-bar p { flex: 1 1 280px; min-width: 0; }
@media (prefers-color-scheme: dark) { .restriction-bar:not(.unlocked) { --bar: #3a3216; --bar-line: #5c4f20; } }
.toolbar button:disabled { opacity: .4; cursor: default; }
.toolbar button.on { background: var(--active); border-color: var(--border); }
.toolbar select { height: 32px; min-width: 130px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; padding: 0 6px; }
.sep { width: 1px; height: 22px; background: var(--border); margin: 0 4px; }
/* Editing tools, then finding and reviewing: one row of controls, wrapping where it must. */
.tb-group { display: contents; }

.workspace { flex: 1; min-height: 0; display: flex; }
.canvas { flex: 1; min-width: 0; overflow: auto; background: var(--canvas); padding: 24px 16px 48px; }
.docx-sheet {
  position: relative; isolation: isolate; /* pictures behind the text stay on the page */ box-sizing: border-box; margin: 0 auto; max-width: 100%; background: #fff; color: #000;
  box-shadow: 0 1px 3px rgb(0 0 0 / 20%), 0 4px 16px rgb(0 0 0 / 8%);
  display: flex; flex-direction: column;
}
.docx-page {
  flex: 1; box-sizing: border-box; outline: none; min-width: 0;
  white-space: pre-wrap; overflow-wrap: break-word; tab-size: 36pt; /* Word's default stops, half an inch apart; a paragraph's own: see doc.applyTabStops */ caret-color: #000;
  /* What lies past the page's edges (a floating picture placed partly off the page, or one placed
     for a wider page than a phone's) is cut off there, as Word cuts it off at the edge of the paper,
     rather than making the document scroll sideways. Clip, not hidden: nothing becomes scrollable. */
  overflow-x: clip;
}
.hf-tabs {
  position: absolute; z-index: 3; left: 8px; right: 8px; display: flex; flex-wrap: wrap; gap: 4px;
  font: 11px/1.4 system-ui, sans-serif; white-space: normal;
}
.hf-tabs.header { top: 6px; }
.hf-tabs.footer { bottom: 6px; }
.hf-tabs button {
  padding: 1px 8px; border: 1px solid #c9d1de; border-radius: 9px; background: #fff; color: #44516a; cursor: pointer;
}
.hf-tabs button[aria-selected="true"] { background: var(--sheet-accent); border-color: var(--sheet-accent); color: #fff; }
.comments-toggle { gap: 3px; padding: 0 6px; }
.fill-in-toggle { gap: 3px; padding: 0 6px; }
.script { font-size: 14px; line-height: 1; }
.fonts-missing { cursor: help; text-decoration: underline dotted; }
.draft-status { cursor: help; }
.script sup, .script sub { font-size: 9px; }
.split { display: inline-flex; }
.split .arrow { min-width: 16px !important; padding: 0 2px; font-size: 10px !important; }
.changes-group { display: inline-flex; align-items: center; gap: 2px; }
.labelled { gap: 4px; padding: 0 6px !important; font-size: 13px !important; }
/* Icon buttons' text labels, shown where the toolbar has room for them (see editor/toolbar.js fitLabels). */
.toolbar .short-label { display: none; }
.toolbar button.with-label { gap: 4px; padding: 0 6px; }
.toolbar button.with-label .short-label { display: inline; font-size: 13px; white-space: nowrap; }
.change-menu {
  position: fixed; z-index: 10; min-width: 200px; padding: 4px; display: flex; flex-direction: column;
  background: var(--surface); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.change-menu button { text-align: left; padding: 6px 10px; border: 0; border-radius: 4px; background: transparent; color: inherit; font: inherit; cursor: pointer; }
.context-menu { min-width: 230px; }
.context-menu button { display: flex; justify-content: space-between; gap: 16px; }
.context-menu button:disabled { opacity: .45; cursor: default; }
.context-menu .key { color: var(--muted); font-size: 12px; }
.context-menu hr { width: 100%; margin: 4px 0; border: 0; border-top: 1px solid var(--border); }
.context-menu .menu-note { margin: 4px 0 0; padding: 6px 10px 2px; border-top: 1px solid var(--border); font-size: 11px; color: var(--muted); max-width: 230px; }
.change-menu button:hover, .change-menu button:focus-visible { background: var(--hover); }
.spacing-menu button { padding-left: 28px; position: relative; }
.spacing-menu button[aria-checked="true"]::before { content: '✓'; position: absolute; left: 10px; }
.chooser.para-dialog { width: 300px; max-width: calc(100vw - 16px); max-height: calc(100vh - 16px); overflow-y: auto; }
.para-dialog fieldset { display: flex; flex-direction: column; gap: 6px; margin: 0; padding: 6px 10px 10px; border: 1px solid var(--border); border-radius: 4px; }
.para-dialog .tab-stops ul { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 2px; }
.para-dialog .tab-stops li { display: grid; grid-template-columns: 64px 60px 1fr 24px; align-items: center; gap: 6px; font-variant-numeric: tabular-nums; }
.para-dialog .tab-stops li button { width: 24px; height: 22px; padding: 0; border: 0; background: none; color: inherit; border-radius: 4px; cursor: pointer; font-size: 16px; line-height: 1; }
.para-dialog .tab-stops li button:hover { background: var(--hover); }
.para-dialog .tab-stops .leader { color: var(--muted); }
.para-dialog .tab-stops .hint { margin: 0; color: var(--muted); font-size: 12px; }
.para-dialog .tab-add { display: flex; flex-direction: column; gap: 6px; padding-top: 6px; border-top: 1px solid var(--border); }
.para-dialog .tab-buttons { display: flex; gap: 6px; justify-content: flex-end; }
.page-dialog .apply-to { display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 6px; }
.para-dialog .check { justify-content: flex-start !important; gap: 6px !important; }
.para-dialog .check input { width: auto; }
.modify-dialog { width: 340px; }
.modify-dialog h2 { margin: 0 0 8px; font-size: 15px; }
.modify-dialog .note { margin: 4px 0 0; font-size: 12px; color: var(--muted); }
.modify-dialog .checks { display: flex; gap: 16px; }
.modify-dialog .color-pick { display: flex; align-items: center; gap: 8px; }
.modify-dialog .color-pick input[type=color] { width: 40px; height: 26px; padding: 0 2px; }
/* Two lines, so a first-line or hanging indent shows (clamped inside the padding). */
.modify-dialog .preview { margin: 10px 0 4px; padding: 8px 10px; border: 1px solid var(--border); border-radius: 4px; background: #fff; }
.modify-dialog .preview span { display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 2; line-clamp: 2; overflow: hidden; }
.para-dialog legend { font-size: 12px; font-weight: 600; padding: 0 4px; }
.page-dialog .note { margin: 0; font-size: 12px; color: var(--muted); }
.para-dialog label { display: flex; justify-content: space-between; align-items: center; gap: 8px; }
.para-dialog input, .para-dialog select { width: 120px; font: inherit; }
.para-dialog select { padding: 5px 6px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; }
.picture-frame { position: absolute; z-index: 6; border: 1px solid var(--accent); pointer-events: none; box-sizing: border-box; }
.picture-frame .handle { position: absolute; width: 9px; height: 9px; margin: -5px 0 0 -5px; background: #fff; border: 1px solid var(--accent); border-radius: 50%; pointer-events: auto; touch-action: none; }
.picture-frame .nw { left: 0; top: 0; cursor: nwse-resize; }
.picture-frame .n { left: 50%; top: 0; cursor: ns-resize; }
.picture-frame .ne { left: 100%; top: 0; cursor: nesw-resize; }
.picture-frame .e { left: 100%; top: 50%; cursor: ew-resize; }
.picture-frame .se { left: 100%; top: 100%; cursor: nwse-resize; }
.picture-frame .s { left: 50%; top: 100%; cursor: ns-resize; }
.picture-frame .sw { left: 0; top: 100%; cursor: nesw-resize; }
.picture-frame .w { left: 0; top: 50%; cursor: ew-resize; }
.picture-frame .rotate { left: 50%; top: -22px; cursor: grab; }
.picture-frame.cropping { border: 1px dashed #000; }
.picture-frame.cropping .handle { width: 10px; height: 10px; border-radius: 0; background: #000; border-color: #fff; }
.picture-panel details { border: 1px solid var(--border); border-radius: 4px; padding: 4px 8px; }
.picture-panel details[open] { display: flex; flex-direction: column; gap: 6px; padding-bottom: 8px; }
.picture-panel summary { cursor: pointer; font-weight: 600; }
.picture-panel .buttons { display: flex; flex-wrap: wrap; gap: 4px; }
.picture-panel .buttons button[aria-pressed="true"] { background: var(--active); }
.picture-panel input[type=color] { width: 100%; height: 26px; padding: 0 2px; }
.picture-panel fieldset.sub { border-style: dashed; }
.picture-panel .note { margin: 0; color: var(--muted); font-size: 12px; }
.layout-options {
  position: absolute; z-index: 6; width: 26px; height: 26px; padding: 0; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: var(--text); cursor: pointer;
}
.picture-panel {
  position: fixed; z-index: 10; width: 300px; box-sizing: border-box; max-height: calc(100vh - 16px); overflow: auto; padding: 10px; display: flex; flex-direction: column; gap: 8px;
  background: var(--surface); color: var(--text); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%); font-size: 13px;
}
.picture-panel fieldset { margin: 0; padding: 6px 8px; border: 1px solid var(--border); border-radius: 4px; display: flex; flex-direction: column; gap: 4px; }
.picture-panel fieldset:disabled { opacity: .5; }
.picture-panel legend { font-weight: 600; padding: 0 4px; }
.picture-panel .wraps { display: grid; grid-template-columns: 1fr 1fr; }
.picture-panel label { display: flex; flex-direction: column; gap: 2px; }
.picture-panel .wraps label, .picture-panel label.inline { flex-direction: row; align-items: center; gap: 4px; }
.picture-panel input[type=number], .picture-panel select, .picture-panel textarea { font: inherit; padding: 2px 4px; width: 100%; min-width: 0; box-sizing: border-box; }
.picture-panel .size-row { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
.picture-panel button { font: inherit; padding: 3px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; cursor: pointer; align-self: flex-start; }
.picture-panel .actions { display: flex; justify-content: flex-end; }
.change-menu button:disabled { opacity: .45; cursor: default; background: transparent; }
.split-dialog label { display: flex; justify-content: space-between; gap: 8px; align-items: center; }
.split-dialog input { width: 64px; }
.change-menu hr { width: 100%; margin: 4px 0; border: 0; border-top: 1px solid var(--border); }
.table-picker {
  position: fixed; z-index: 10; padding: 8px; display: flex; flex-direction: column; gap: 8px;
  background: var(--surface); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.table-picker .size { margin: 0; font-size: 13px; color: var(--muted); }
.table-picker .grid { display: grid; grid-template-columns: repeat(10, 18px); gap: 3px; }
.table-picker .cell { width: 18px; height: 18px; padding: 0; border: 1px solid var(--border); border-radius: 2px; background: var(--surface); cursor: pointer; }
.table-picker .cell.on { background: color-mix(in srgb, var(--accent) 25%, transparent); border-color: var(--accent); }
.table-picker .cell:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.table-picker .custom { display: flex; gap: 6px; align-items: center; font-size: 13px; border-top: 1px solid var(--border); padding-top: 8px; }
.table-picker .custom label { display: flex; gap: 4px; align-items: center; }
.table-picker .custom input { width: 48px; font: inherit; padding: 2px 4px; }
.table-picker .custom button { font: inherit; padding: 2px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; cursor: pointer; }
.palette {
  position: fixed; z-index: 10; width: 190px; padding: 8px; display: flex; flex-direction: column; gap: 8px;
  background: var(--surface); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.swatches { display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; }
.swatch { height: 24px; border: 1px solid rgb(0 0 0 / 25%); border-radius: 3px; cursor: pointer; padding: 0; }
.swatch.current, .swatch:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.case-icon, .symbol-icon { display: inline-block; min-width: 16px; font-size: 14px; line-height: 16px; text-align: center; }
.case-icon { font-weight: 600; letter-spacing: -.02em; }
.symbol-icon { font-family: 'Times New Roman', Times, serif; font-size: 16px; }
/* The Format Painter's brush, over the page while it is on. */
.docx-page.painting, .docx-page.painting * {
  cursor: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='20' height='20' viewBox='0 0 20 20'%3E%3Cpath d='M3 2v14' stroke='%23000' stroke-width='1.5'/%3E%3Crect x='7' y='3' width='9' height='4' rx='.8' fill='%23fff' stroke='%23000' stroke-width='1.2'/%3E%3Cpath d='M16 5h2v3h-6.5v2.5' fill='none' stroke='%23000' stroke-width='1.2'/%3E%3Crect x='10.2' y='10.5' width='2.6' height='6' rx='.6' fill='%23000'/%3E%3C/svg%3E") 3 2, text;
}
.palette-label { margin: 8px 0 4px; font-size: 12px; color: var(--muted); }
.color-palette { width: 220px; }
.cell-format { width: 280px; max-height: calc(100vh - 16px); overflow-y: auto; }
.border-choices { display: grid; grid-template-columns: 1fr 1fr; gap: 4px; }
.border-choices button { display: flex; align-items: center; gap: 6px; padding: 5px 6px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; font-size: 12px; cursor: pointer; text-align: left; }
.border-choices button:hover, .border-choices button:focus-visible { background: var(--hover); }
.toolbar select.font-choice { min-width: 0; width: 132px; }
.toolbar select.size-choice { min-width: 0; width: 58px; }
.toolbar select.char-style { min-width: 0; width: 124px; }
.toolbar select.markup-choice { min-width: 0; width: 128px; }
.toolbar .comment-nav { display: inline-flex; }
.page-count, .smart-typing, .nav-toggle { background: none; border: 0; padding: 0; font: inherit; color: inherit; cursor: pointer; text-decoration: underline dotted; }
.nav-toggle[aria-pressed="true"] { color: var(--accent); }
/* The Headings pane is for a wider screen (it is hidden on a phone, and so is its button). */
@media (max-width: 700px) { .nav-toggle { display: none; } }
.no-color { border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; padding: 4px; font: inherit; font-size: 13px; cursor: pointer; }
.marks-toggle { font-size: 16px !important; }
.track-toggle, .review-toggle { gap: 6px; padding: 0 8px !important; font-size: 13px !important; }
@media (max-width: 720px) { .track-toggle .label, .review-toggle .label, .labelled .label { display: none; } }
/* Phones: the toolbar is one row that scrolls sideways rather than four rows over the text, faded at
   the edge where more controls are (see editor/toolbar.js); the status bar is one line too. */
@media (max-width: 720px) {
  .toolbar { flex-wrap: nowrap; overflow-x: auto; overscroll-behavior-x: contain; scrollbar-width: none; padding: 4px 8px; }
  .toolbar::-webkit-scrollbar { display: none; }
  .toolbar > *, .toolbar .tb-group > *, .toolbar .group > * { flex: none; }
  .toolbar[data-more="end"] { mask-image: linear-gradient(to right, #000 calc(100% - 32px), transparent); }
  .toolbar[data-more="start"] { mask-image: linear-gradient(to left, #000 calc(100% - 32px), transparent); }
  .toolbar[data-more="both"] { mask-image: linear-gradient(to right, transparent, #000 32px, #000 calc(100% - 32px), transparent); }
  .statusbar { gap: 16px; overflow-x: auto; white-space: nowrap; scrollbar-width: none; }
  .statusbar > * { flex: none; }
  /* What just happened ("Signature inserted", "Shared letter.docx") comes first in the one line,
     rather than out of sight to the right; a long message ends in "…" (its tooltip, and screen
     readers, have all of it), and the counts follow, a swipe away. */
  .statusbar [role="status"] { order: -1; max-width: calc(100vw - 32px); overflow: hidden; text-overflow: ellipsis; color: var(--text); }
  .statusbar [role="status"]:empty { display: none; }
  /* The page already fits the screen's width, and the phone pinches to zoom. */
  .statusbar > span.zoom { display: none; }
  /* The tools a phone is used for first: undo, bold, italic, underline, highlight, lists, check box
     or date, signature, picture and comment; the rest follow, a swipe away. */
  .toolbar .ph { order: -1; }
  .toolbar .sep { display: none; }
  /* A document with tracked changes or comments to go through: Undo and Redo, then the review's
     tools (previous and next change, Accept and Reject, with their names, and the comments), then
     the rest, as approving a document between meetings needs them. */
  .toolbar.review-first .ur { order: -4; }
  .toolbar.review-first .changes-group.rv { order: -3; }
  .toolbar.review-first .rv { order: -2; }
  .toolbar.review-first .changes-group .labelled .label { display: inline; }
}
/* Touch: targets a finger can hit. */
@media (pointer: coarse) {
  .toolbar button { min-width: 40px; height: 40px; }
  .toolbar .split .arrow { min-width: 24px !important; }
  .toolbar select { height: 40px; }
}
/* The word count opens Word count (a length target, readability); over the target, or the headline
   over its limit, it says so in amber (7:1 on the surface in both themes). */
.word-count, .headline-count { background: none; border: 0; padding: 0; font: inherit; color: inherit; cursor: pointer; }
.word-count:hover, .headline-count { text-decoration: underline dotted; }
.word-count.over, .headline-count.over { color: #8a5300; font-weight: 600; }
@media (prefers-color-scheme: dark) { .word-count.over, .headline-count.over { color: #ffcf80; } }
.tracking-status { border: 0; background: none; color: var(--own-change); font: inherit; padding: 0; cursor: pointer; text-decoration: underline dotted; }
.author-dialog { gap: 8px; padding: 12px; min-width: 260px; }
.link-dialog { gap: 6px; padding: 12px; width: 320px; max-width: calc(100vw - 16px); max-height: none; }
.link-dialog .error { margin: 0; color: #b3261e; font-size: 12px; }
.author-dialog label, .link-dialog label { font-size: 13px; }
.author-dialog input, .link-dialog input { font: inherit; padding: 6px 8px; border: 1px solid var(--border); border-radius: 4px; }
.author-dialog .actions, .link-dialog .actions { display: flex; gap: 8px; justify-content: flex-end; }
.author-dialog .actions button, .link-dialog .actions button { border: 1px solid var(--border); border-radius: 4px; padding: 5px 12px; background: var(--surface); color: inherit; cursor: pointer; text-align: center; }
.author-dialog .actions .confirm, .link-dialog .actions .confirm { background: var(--brand); border-color: var(--brand); color: #fff; }
.badge { font-size: 11px; color: var(--muted); }
.badge.zero { opacity: .6; }
.changes-group .badge { min-width: 3ch; margin-left: 4px; padding: 1px 5px; border-radius: 9px; background: var(--hover); text-align: center; font-variant-numeric: tabular-nums; } /* the count changing moves nothing */
.chooser {
  position: fixed; z-index: 10; min-width: 160px; max-width: 280px; max-height: 280px; overflow: auto;
  display: flex; flex-direction: column; padding: 4px; background: var(--surface); color: inherit;
  border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.chooser button { text-align: left; padding: 6px 10px; border: 0; border-radius: 4px; background: transparent; color: inherit; font: inherit; cursor: pointer; }
.chooser button:hover, .chooser button:focus-visible, .chooser button[aria-selected="true"] { background: var(--hover); }
.chooser input { font: inherit; padding: 4px; }
.chooser .empty { margin: 6px 10px; color: var(--muted); }
.statusbar {
  display: flex; gap: 24px; padding: 4px 16px; font-size: 12px; color: var(--muted);
  background: var(--surface); border-top: 1px solid var(--border);
}
.statusbar .zoom { margin-left: auto; display: inline-flex; align-items: center; gap: 2px; flex: none; }
.statusbar .zoom button {
  min-width: 24px; padding: 0 4px; border: 1px solid transparent; border-radius: 4px; background: none;
  color: inherit; font: inherit; line-height: 18px; cursor: pointer;
}
.statusbar .zoom button:hover:not(:disabled) { background: var(--hover); }
.statusbar .zoom button:disabled { opacity: .4; cursor: default; }
.statusbar .zoom .zoom-level { min-width: 44px; font-variant-numeric: tabular-nums; }

@media screen and (max-width: 860px) {
  .canvas { padding: 0; }
  .docx-sheet { min-height: 100% !important; width: 100% !important; box-shadow: none; }
  .docx-page { padding: 24px 16px !important; }
  /* The page's own margins (inline) give way to the phone's. */
  .docx-sheet .docx-page { --hf-header: 12px !important; --hf-top: 0px !important; --hf-footer: 12px !important; --hf-bottom: 0px !important; --hf-left: 16px !important; --hf-right: 16px !important; }
  .docx-sheet:has(.hf-tabs.header) .docx-page { --hf-header: 34px !important; }
  .docx-sheet:has(.hf-tabs.footer) .docx-page { --hf-footer: 34px !important; }
}
</style>

<style>
/* Unscoped: applies to the generated document HTML. */
/* Printing (Ctrl+P) prints the document, not the editor around it: on the document's paper size, in
   its margins (the @page rule is set from the document when it opens), with its page breaks. Headers
   and footers go on every page where they can (set just before printing: editor/printBands.js);
   otherwise they are printed once, at the start and the end, as the page shows them. */
@media print {
  html, body { background: #fff !important; }
  .editor-shell { display: block !important; height: auto !important; }
  .editor-shell > :not(.workspace), .workspace > :not(.canvas), .docx-sheet > :not(.docx-page) { display: none !important; }
  .workspace { display: block !important; }
  .canvas { overflow: visible !important; padding: 0 !important; background: none !important; }
  .docx-sheet { width: auto !important; max-width: none !important; min-height: 0 !important; margin: 0 !important; box-shadow: none !important; zoom: 1 !important; }
  .docx-page { padding: 0 !important; overflow: visible !important; }
  .docx-page, .docx-sheet [contenteditable] { caret-color: transparent !important; }
  .docx-sheet .docx-hf-band.header { padding-top: 0; min-height: 0; }
  .docx-sheet .docx-hf-band.footer { position: static; padding: 12pt 0 0; min-height: 0; }
  .docx-sheet .docx-hf-part { opacity: 1 !important; box-shadow: none !important; }
  .docx-sheet .docx-hf-part::after { display: none !important; }
  .docx-sheet .raw-inline.page-break { border: 0 !important; margin: 0 !important; break-after: page; }
  .docx-sheet [data-page-break-before] { break-before: page; }
  /* Hidden text (w:vanish) isn't printed, as in Word, unless Print hidden text is chosen. */
  .docx-sheet [data-hidden] { display: none !important; }
  .print-hidden-text .docx-sheet [data-hidden] { display: inline !important; opacity: 1 !important; text-decoration: none !important; }
  .docx-sheet [data-section-end="nextPage"], .docx-sheet [data-section-end="oddPage"], .docx-sheet [data-section-end="evenPage"] { break-after: page; }
}
.docx-sheet :is(h1, h2, h3, h4, h5, h6) { font-size: inherit; font-weight: inherit; }
.docx-sheet :is(p, h1, h2, h3, h4, h5, h6) { position: relative; }
.docx-sheet .cs-before { margin-top: 0 !important; }
.docx-sheet .cs-after { margin-bottom: 0 !important; }
.docx-sheet .list::before {
  content: attr(data-marker); position: absolute; transform: translateX(calc(-100% - 0.5em));
  text-indent: 0; white-space: nowrap;
}
/* A list level centred on its own line ("Pasal 1", "BAB I"): its number in the line. */
/* Clicks go through it to the paragraph: Firefox puts no caret where generated text is clicked, and
   an empty "Pasal 1" line is all number. */
.docx-sheet .list[data-marker-inline]::before { position: static; transform: none; margin-right: .3em; pointer-events: none; }
.docx-sheet a { color: #0563c1; }
.docx-sheet img { max-width: 100%; vertical-align: bottom; }
/* Hidden markers (comment ranges, bookmarks, field codes) are empty: shown inline, not display:none,
   because the browser's editing drops unrendered nodes when it moves text, as in merging paragraphs. */
.docx-sheet .raw-hidden { display: none; }
.docx-sheet span.raw-hidden { display: inline; }
/* Legacy form fields (docx/formFields.js): the box or chosen entry is drawn from the marker's
   attributes, so the paragraph's text is unchanged; grey as Word shades form fields. */
.docx-sheet span.raw-hidden.ff-check::before { content: '\2610'; font-family: 'Segoe UI Symbol', 'MS Gothic', 'DejaVu Sans', sans-serif; }
.docx-sheet span.raw-hidden.ff-check[data-checked="1"]::before { content: '\2612'; }
.docx-sheet span.raw-hidden.ff-drop::before { content: attr(data-ff-text); }
.docx-sheet span.raw-hidden.ff-drop::after { content: ' \25BE'; font-size: 0.8em; opacity: 0.6; }
.docx-sheet span.raw-hidden:is(.ff-check, .ff-drop) { background: #e6e6e6; cursor: pointer; border-radius: 2px; }
.docx-sheet span.raw-hidden:is(.ff-check, .ff-drop):hover { background: var(--sheet-tint); }
.docx-sheet span.raw-hidden:is(.ff-check, .ff-drop):focus-visible { outline: 2px solid var(--sheet-accent); outline-offset: 1px; }
.docx-sheet span.raw-hidden:is(.ff-check, .ff-drop)[data-ff-off] { cursor: default; opacity: 0.6; }
.docx-sheet .ff-text { background: #e6e6e6; }
.docx-sheet .ff-text[data-ff-host] { outline: none; }
.docx-sheet .ff-text[data-ff-host]:focus { box-shadow: 0 0 0 2px var(--sheet-accent); }
.docx-sheet .ff-shadow { display: none; }
.docx-sheet .raw-block, .docx-sheet .raw-inline:not(:has(img)) {
  background: #f1f3f6; color: #555; border-radius: 3px; cursor: default; user-select: none;
}
.docx-sheet .raw-block { padding: 4px 8px; margin: 4px 0; border-left: 3px solid #c3c9d4; white-space: pre-wrap; }
.docx-sheet .raw-inline.deleted { text-decoration: line-through; color: #b3261e; background: transparent; }
.docx-sheet .raw-inline.inserted { text-decoration: underline; color: #1e6b34; background: transparent; }
/* Formatting marks (editor/marks.js): a layer over the page that never takes clicks. */
.docx-sheet .marks-layer { position: absolute; inset: 0; pointer-events: none; z-index: 1; overflow: hidden; }
/* Line numbers in the margin, as Word prints them: drawn over the sheet, never part of the page. */
.docx-sheet .line-numbers-layer { position: absolute; inset: 0; pointer-events: none; z-index: 1; overflow: hidden; }
/* The watermark: behind the text (the page comes after it), in the middle of each page. */
/* A section break, named with its section's paper, where the document has several. */
.docx-sheet .docx-page [data-section-note]::after {
  content: attr(data-section-note); display: block; margin-top: 4px; padding-top: 2px; border-top: 1px dashed var(--sheet-line);
  font: 600 10px/14px system-ui, sans-serif; color: var(--sheet-accent); text-indent: 0; text-align: left; letter-spacing: 0;
  user-select: none; pointer-events: none;
}
@media print { .docx-sheet .docx-page [data-section-note]::after { content: none; } }
/* Hidden text (w:vanish): shown faintly (see styles.rPrCss), with Word's dotted line under it. */
.docx-sheet [data-hidden] { text-decoration-line: underline; text-decoration-style: dotted; text-decoration-color: currentColor; }
.docx-sheet .watermark-layer { position: absolute; inset: 0; pointer-events: none; overflow: hidden; z-index: -1; }
.docx-sheet .watermark-mark {
  position: absolute; left: 50%; transform: translate(-50%, -50%); font-weight: 700; line-height: 1; opacity: .5;
  white-space: nowrap; user-select: none;
}
.docx-sheet .watermark-mark.diagonal { transform: translate(-50%, -50%) rotate(-45deg); }
.docx-sheet .page-breaks-layer { position: absolute; inset: 0; pointer-events: none; z-index: 1; overflow: hidden; }
/* Where a page would begin, about (editor/pages.js): a dashed line across the sheet, numbered in the margin. */
.docx-sheet .page-estimate { position: absolute; left: 0; right: 0; height: 0; border-top: 1px dashed var(--sheet-line); }
.docx-sheet .page-estimate span {
  position: absolute; right: 6px; top: -8px; padding: 0 4px; border-radius: 3px; background: #fff;
  font: 600 10px/16px system-ui, sans-serif; color: var(--sheet-accent); white-space: nowrap; user-select: none;
}
.docx-sheet .line-number { position: absolute; font-size: 9pt; color: #6b6b6b; text-align: right; white-space: nowrap; user-select: none; }
.docx-sheet .fmark { position: absolute; color: #4f74b3; font-family: system-ui, sans-serif; white-space: pre; user-select: none; }
.docx-sheet .fmark.space { transform: translateX(-50%); }
/* Highlight, in Word's colours. */
.docx-sheet mark { background: #ffff00; color: inherit; }
.docx-sheet mark[data-color="green"] { background: #00ff00; } .docx-sheet mark[data-color="cyan"] { background: #00ffff; }
.docx-sheet mark[data-color="magenta"] { background: #ff00ff; } .docx-sheet mark[data-color="blue"] { background: #0000ff; }
.docx-sheet mark[data-color="red"] { background: #ff0000; } .docx-sheet mark[data-color="darkBlue"] { background: #000080; }
.docx-sheet mark[data-color="darkCyan"] { background: #008080; } .docx-sheet mark[data-color="darkGreen"] { background: #008000; }
.docx-sheet mark[data-color="darkMagenta"] { background: #800080; } .docx-sheet mark[data-color="darkRed"] { background: #800000; }
.docx-sheet mark[data-color="darkYellow"] { background: #808000; } .docx-sheet mark[data-color="darkGray"] { background: #808080; }
.docx-sheet mark[data-color="lightGray"] { background: #c0c0c0; } .docx-sheet mark[data-color="black"] { background: #000000; }
.docx-sheet mark[data-color="white"] { background: #ffffff; }
/* This session's tracked changes. */
/* The insertion's underline is drawn, not a text-decoration: the browser reads a text-decoration as
   the text being underlined, so the Underline button showed as on and Ctrl+U did nothing in it. */
.docx-sheet .tc-ins {
  color: #6f2dbd;
  background-image: linear-gradient(currentColor, currentColor);
  background-repeat: no-repeat; background-position: 0 100%; background-size: 100% 1px;
}
.docx-sheet .tc-del { color: #6f2dbd; text-decoration: line-through; cursor: default; }
.docx-sheet [data-tc-mark]::after { content: '¶'; color: #6f2dbd; margin-left: 1px; }
.docx-sheet [data-tc-mark="del"]::after { text-decoration: line-through; }
/* Tracked formatting changes (of a run: dashed underline; of a paragraph: a bar in the margin) and
   tracked table rows, as Word marks them. */
.docx-sheet [data-fmt="run"] { border-bottom: 1px dashed #8e44ad; }
.docx-sheet [data-fmt-para] { box-shadow: -10px 0 0 -8px #8e44ad; }
.docx-sheet table.docx-table[data-fmt-table] { outline: 2px dashed #8e5cc7; outline-offset: 2px; }
.docx-sheet tr[data-fmt-row] > td:first-child { box-shadow: inset 3px 0 0 #8e5cc7; }
.docx-sheet table.docx-table td[data-rev-merge] { box-shadow: inset 0 0 0 2px rgb(111 45 189 / 45%); }
.docx-sheet table.docx-table td[data-fmt-cell] { box-shadow: inset 0 0 0 1px #8e5cc7; }
.docx-sheet tr[data-rev-row="ins"] > td { background-color: #e8f5ea; }
.docx-sheet tr[data-rev-row="del"] > td { background-color: #fbe9e7; text-decoration: line-through; }
/* Someone else's tracked paragraph break, in the colours of their insertions and deletions. */
.docx-sheet [data-rev-mark]::after { content: '¶'; margin-left: 1px; }
.docx-sheet [data-rev-mark="ins"]::after { color: #1e6b34; }
.docx-sheet [data-rev-mark="del"]::after { color: #b3261e; text-decoration: line-through; }
/* Simple and no markup (Review > Show markup), as Word's Display for Review: the text as it will be
   once the changes are accepted. Deletions are hidden, insertions look like the text around them and
   formatting changes aren't marked; nothing in the page or the file changes, only how it shows.
   The original is the other way round: insertions hidden, deletions as ordinary text. */
.docx-sheet:is([data-markup="simple"], [data-markup="none"]) :is(.tc-del, .raw-inline.deleted, tr[data-rev-row="del"]),
.docx-sheet[data-markup="original"] :is(.tc-ins, .raw-inline.inserted, tr[data-rev-row="ins"]) { display: none; }
.docx-sheet[data-markup] .tc-ins, .docx-sheet[data-markup] .tc-del { color: inherit; background-image: none; text-decoration: none; }
.docx-sheet[data-markup] .raw-inline.inserted, .docx-sheet[data-markup] .raw-inline.deleted { color: inherit; text-decoration: none; }
.docx-sheet[data-markup] tr[data-rev-row="del"] > td { background-color: transparent; text-decoration: none; }
.docx-sheet[data-markup] [data-tc-mark]::after, .docx-sheet[data-markup] [data-rev-mark]::after { content: none; }
.docx-sheet[data-markup] [data-fmt="run"] { border-bottom: 0; }
.docx-sheet[data-markup] [data-fmt-para], .docx-sheet[data-markup] tr[data-fmt-row] > td:first-child,
.docx-sheet[data-markup] table.docx-table td[data-rev-merge], .docx-sheet[data-markup] table.docx-table td[data-fmt-cell] { box-shadow: none; }
.docx-sheet[data-markup] table.docx-table[data-fmt-table] { outline: 0; }
.docx-sheet[data-markup] tr[data-rev-row="ins"] > td { background-color: transparent; }
/* Simple markup: a red line in the margin beside each paragraph with a change, as Word draws it (a
   2px line 10px out, the white shadow over the red one's inner part). */
.docx-sheet[data-markup="simple"] .docx-page :is(p, h1, h2, h3, h4, h5, h6):is(:has(.tc-ins, .tc-del, .raw-inline.inserted, .raw-inline.deleted, [data-fmt]), [data-tc-mark], [data-rev-mark], [data-fmt-para]) {
  box-shadow: -10px 0 0 0 #fff, -12px 0 0 0 #b3261e;
}
/* No markup hides the comments' highlights too, as Word does; the panel still lists them. */
.docx-sheet:is([data-markup="none"], [data-markup="original"]) ::highlight(comment), .docx-sheet:is([data-markup="none"], [data-markup="original"]) ::highlight(comment-active) { background-color: transparent; }
.docx-sheet .raw-inline.picture { cursor: pointer; }
.docx-sheet .raw-inline.picture img { display: block; }
.docx-sheet .raw-inline.picture::selection, .docx-sheet .raw-inline.picture img::selection { background: transparent; } /* the frame shows it is selected */
.docx-sheet .raw-inline.picture[data-wrap="inline"] { display: inline-block; vertical-align: bottom; }
.docx-sheet .raw-inline.page-break { display: block; height: 0; border-top: 1px dashed #aab; margin: 8px 0; background: none; }
.docx-sheet .raw-inline.note-ref { vertical-align: super; font-size: smaller; background: transparent; color: var(--sheet-accent); }
.docx-sheet .raw-inline.note-ref::before { content: attr(data-n); }
.docx-sheet .raw-inline.note-ref[data-note-ref], .docx-sheet .raw-inline.note-mark { cursor: pointer; }
.docx-sheet .raw-inline.note-ref[data-note-ref]:hover, .docx-sheet .raw-inline.note-mark:hover { text-decoration: underline; }
/* Footnotes and endnotes, below the text (as in Word's draft view): each note is edited in place. */
/* Headers and footers: in the page, in its margins, dimmed until you work in them, as in Word. */
.docx-sheet .docx-hf-band { cursor: default; overflow-x: clip; } /* cut off at the page's edges, as the page is */
.docx-sheet .docx-hf-band.header { padding-top: var(--hf-header); min-height: var(--hf-top); margin-bottom: 6px; }
.docx-sheet .docx-hf-band.footer {
  position: absolute; left: var(--hf-left); right: var(--hf-right); bottom: 0;
  padding-bottom: var(--hf-footer); padding-top: 6px; min-height: var(--hf-bottom);
  display: flex; flex-direction: column; justify-content: flex-end;
}
.docx-sheet .docx-hf-part { outline: none; border-radius: 2px; cursor: text; opacity: .6; position: relative; }
.docx-sheet .docx-hf-part:hover { opacity: .85; box-shadow: 0 0 0 1px var(--sheet-line); }
.docx-sheet .docx-hf-part:focus-within { opacity: 1; box-shadow: 0 0 0 1px var(--sheet-accent); }
.docx-sheet .docx-hf-part:is(:hover, :focus-within)::after {
  content: attr(data-label); position: absolute; left: 0; font: 600 10px/1.6 system-ui, sans-serif;
  padding: 0 6px; border-radius: 3px; background: var(--sheet-tint); color: var(--sheet-accent); pointer-events: none; white-space: nowrap;
}
.docx-sheet .header .docx-hf-part::after { top: 100%; margin-top: 3px; }
.docx-sheet .footer .docx-hf-part::after { bottom: 100%; margin-bottom: 3px; }
/* Text boxes and shapes with text: a frame holding text you edit. */
.docx-sheet .textbox {
  display: inline-flex; flex-direction: column; box-sizing: border-box; vertical-align: top;
  text-indent: 0; white-space: pre-wrap; cursor: default; overflow-wrap: break-word; max-width: 100%;
}
.docx-sheet .textbox[data-wrap="topAndBottom"] { display: flex; }
.docx-sheet .textbox-content { outline: none; cursor: text; min-height: 1em; }
.docx-sheet .textbox:hover { box-shadow: 0 0 0 1px var(--sheet-line); }
.docx-sheet .textbox:focus-within { box-shadow: 0 0 0 2px var(--sheet-accent); }
.docx-sheet .textbox-content :is(p, h1, h2, h3, h4, h5, h6):first-child { margin-top: 0; }
/* A drop cap: its paragraph floats, and the next lines flow around it. */
.docx-sheet .drop-cap { float: left; margin-top: 0 !important; margin-bottom: 0 !important; }
.docx-sheet .drop-cap + * { margin-top: 0; }
.docx-sheet .docx-columns { margin: 0; }
.docx-sheet .docx-columns > :first-child { margin-top: 0; }
.docx-sheet .docx-notes { margin-top: 24px; padding-top: 6px; cursor: default; }
.docx-sheet .docx-notes::before {
  content: attr(data-label); display: block; width: 33%; border-top: 1px solid #9aa0aa; padding-top: 4px; margin-bottom: 2px;
  font: 600 11px/1.4 system-ui, sans-serif; color: #6b7280; letter-spacing: .02em; text-transform: uppercase; user-select: none;
}
.docx-sheet .docx-note { outline: none; border-radius: 2px; cursor: text; }
.docx-sheet .docx-note:focus-within { background: var(--sheet-tint); box-shadow: 0 0 0 2px var(--sheet-tint); }
.docx-sheet .raw-inline.symbol { background: transparent; color: inherit; }
/* Equations, drawn as MathML (docx/omml.js): read-only, so they look like the text around them. */
.docx-sheet .raw-inline.equation { background: transparent; color: inherit; border-radius: 3px; }
.docx-sheet .raw-inline.equation:hover { background: var(--sheet-tint); }
.docx-sheet .raw-inline.equation.display { display: block; margin: 4px 0; }
.docx-sheet .equation math { font-family: 'Cambria Math', 'STIX Two Math', 'Latin Modern Math', math; }
.docx-sheet .raw-inline.comment-ref { display: inline; cursor: pointer; user-select: none; background: transparent; padding: 0; }
.docx-sheet .comment-ref::before {
  content: attr(data-n); display: inline-block; min-width: 1.3em; margin: 0 1px; padding: 0 3px; border-radius: 7px;
  font: 600 10px/1.4 system-ui, sans-serif; text-align: center; vertical-align: super; background: #d0705c; color: #fff;
}
.docx-sheet .field { border-bottom: 1px dotted #9aa6b8; }
/* A page number shows # (each page has its own) and the page count about how many pages there
   are: drawn over the number the file holds, which stays in the page as it is. */
.docx-sheet .field.page-field[data-show]:not([data-show=""]) { position: relative; -webkit-text-fill-color: transparent; }
/* Inline (an inline-block would take Home and End off the line): room made for a longer count. */
.docx-sheet .field.page-field[data-digits="2"] { padding-right: .6em; }
.docx-sheet .field.page-field[data-digits="3"] { padding-right: 1.2em; }
.docx-sheet .field.page-field[data-show]:not([data-show=""])::after { content: attr(data-show); position: absolute; left: 0; top: 0; -webkit-text-fill-color: currentColor; white-space: nowrap; }
/* A page number without a result (see markBarePageField): # as the others show, printed numbered
   where headers and footers print on every page, and blank where they print once. */
.docx-sheet span.raw-hidden.page-field-bare::after { content: "#"; }
@media print {
  .docx-sheet span.raw-hidden.page-field-bare::after { content: none; }
  /* Where the caret is (a table cell's, a content control's outline) isn't part of the page. */
  .docx-sheet table.docx-table td:focus-within, .docx-sheet .sdt:focus-within { box-shadow: none !important; }
  .docx-sheet .sdt:not(.sdt-locked):focus-within { background: none !important; }
  .docx-sheet .field.page-field[data-show] { -webkit-text-fill-color: inherit; padding-right: 0; }
  .docx-sheet .field.page-field[data-show]::after { content: none; }
}
::highlight(comment) { background-color: #fbe3dc; }
/* Read aloud: the paragraph being read, and its word. */
::highlight(read-aloud) { background-color: var(--sheet-tint); }
::highlight(read-aloud-word) { background-color: var(--sheet-line); color: #000; }
::highlight(field-near) { background-color: var(--sheet-tint); }
::highlight(comment-active) { background-color: #f6b9a8; }
.docx-sheet .sdt { border-radius: 2px; transition: background-color .1s, box-shadow .1s; }
.docx-sheet div.sdt { box-shadow: 0 0 0 1px transparent; }
.docx-sheet .sdt:not(.sdt-locked):hover, .docx-sheet .sdt:not(.sdt-locked):focus-within { background: var(--sheet-tint); box-shadow: 0 0 0 1px var(--sheet-line); }
.docx-sheet .sdt-content { outline: none; }
.docx-sheet span.sdt-content:empty { padding: 0 1ch; }
.docx-sheet .showing-placeholder { color: #767676; }
.docx-sheet :is(.sdt-checkbox, .sdt-dropdown, .sdt-date) { cursor: pointer; }
.docx-sheet :is(span.sdt-dropdown, span.sdt-combo, span.sdt-date)::after { content: '▾'; font-size: .8em; color: #5b6b85; margin-left: 2px; cursor: pointer; }
.docx-sheet .sdt-locked { cursor: default; user-select: none; background: #f1f3f6; }
.docx-sheet table.docx-table { border-collapse: collapse; width: 100%; margin: 4px 0; }
.docx-sheet table.docx-table td { border: 1px solid #bfbfbf; padding: 2px 6px; vertical-align: top; outline: none; }
.docx-sheet table.docx-table td:focus-within { box-shadow: inset 0 0 0 2px rgb(14 100 103 / 33%); }
/* A cell merged with the one above: no line between them. "hidden" wins over every other border in
   the collapsing model, the bottom border of the cell above (drawn there, as its own) included. */
.docx-sheet table.docx-table td.vmerge { border-top-style: hidden !important; }
.docx-sheet table.docx-table td[data-rev-cell="ins"] { background-color: rgb(30 107 52 / 8%) !important; box-shadow: inset 0 0 0 1px rgb(30 107 52 / 45%); }
.docx-sheet table.docx-table td[data-rev-cell="del"] { background-color: rgb(179 38 30 / 8%) !important; box-shadow: inset 0 0 0 1px rgb(179 38 30 / 45%); }
.docx-sheet table.docx-table td.vmerge > * { visibility: hidden; }
</style>
