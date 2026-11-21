<template>
  <header class="app-header">
    <RouterLink to="/" class="brand" :aria-label="t('Intact, start page')">
      <!-- The Intact mark (public/logo.svg), reversed for the teal bar: a page with a text cursor. -->
      <svg class="mark" viewBox="0 0 64 64" width="26" height="26" aria-hidden="true">
        <path class="page" d="M18 4H38L54 20V52A8 8 0 0 1 46 60H18A8 8 0 0 1 10 52V12A8 8 0 0 1 18 4Z" />
        <path class="fold" d="M38 4V14A6 6 0 0 0 44 20H54Z" />
        <path class="beam" d="M24 18Q32 18 32 24Q32 18 40 18M32 24V40M24 46Q32 46 32 40Q32 46 40 46" />
      </svg>
      <span>Intact</span>
    </RouterLink>
    <slot />
    <nav class="nav">
      <!-- The interface's language (see i18n.js), on the start page and About. -->
      <select v-if="showLanguage" class="language" :value="lang" :aria-label="t('Language')" :title="t('Language of the interface (remembered in this browser); the documents keep their own')" @change="setLanguage($event.target.value)">
        <option v-for="l in LANGUAGES" :key="l.code" :value="l.code" :lang="l.code">{{ l.name }}</option>
      </select>
      <RouterLink to="/about">{{ t('About') }}</RouterLink>
    </nav>
  </header>
</template>

<script setup>
import { t, lang, setLanguage, LANGUAGES } from '../i18n.js';

defineProps({ showLanguage: { type: Boolean, default: false } });
</script>

<style scoped>
.app-header {
  display: flex; align-items: center; gap: 16px; padding: 8px 16px; min-height: 48px;
  background: var(--brand); color: #fff;
}
/* A phone: tighter, so a document's name has room beside its buttons. */
@media (max-width: 560px) { .app-header { gap: 8px; padding: 8px 12px; } }
.brand { display: flex; align-items: center; gap: 8px; color: inherit; text-decoration: none; font-weight: 600; font-size: 1.05rem; }
.mark { flex: none; }
.mark .page { fill: #fff; }
.mark .fold { fill: rgb(255 255 255 / 60%); }
.mark .beam { fill: none; stroke: var(--brand); stroke-width: 4.5; stroke-linecap: round; stroke-linejoin: round; }
.nav { margin-left: auto; display: flex; align-items: center; gap: 14px; }
.language {
  font: inherit; font-size: .9rem; color: #fff; background: transparent; border: 1px solid rgb(255 255 255 / 55%);
  border-radius: 6px; padding: 3px 6px; cursor: pointer;
}
.language option { color: initial; background: initial; }
.language:focus-visible { outline: 2px solid #fff; outline-offset: 2px; }
.nav a { color: #fff; opacity: .9; text-decoration: none; }
.nav a:hover { opacity: 1; text-decoration: underline; }
</style>
