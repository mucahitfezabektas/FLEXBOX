<script lang="ts">
  import { invoke } from '@tauri-apps/api/core';
  import { moduleManager } from '$lib/state/moduleManager.svelte';
  import type { ThemeMode } from '$lib/utils/theme';

  let selectedTheme = $state<ThemeMode>('light');
  let statusMessage = $state('');
  let saving = $state(false);

  async function clearSavedWorkspaceSession() {
    await moduleManager.clearWorkspaceSession();
    statusMessage = 'Kayitli workspace oturumu temizlendi. Sonraki acilista varsayilan duzen kullanilacak.';
  }

  async function saveTheme() {
    saving = true;
    statusMessage = '';
    moduleManager.setTheme(selectedTheme);

    try {
      await invoke('save_theme_preference', {
        theme: selectedTheme
      });
      statusMessage = 'Tema tercihi Tauri katmanina kaydedildi.';
    } catch {
      statusMessage = 'Tema guncellendi. Tauri shell olmadan calistigi icin kalici kayit atlandi.';
    } finally {
      saving = false;
    }
  }
</script>

<div class="grid h-full gap-4 xl:grid-cols-[0.94fr_1.06fr]">
  <div class="surface-panel-alt rounded-md p-5 shadow-enterprise">
    <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Appearance</p>
    <h2 class="mt-2 text-xl font-semibold tracking-tight text-enterprise-text-primary">Workspace preferences</h2>
    <p class="mt-2 text-sm text-enterprise-text-muted">
      Tema ve oturum davranisi burada yonetilir.
    </p>

    <div class="mt-4 grid gap-3">
      <button
        class="rounded-md border border-enterprise-border-strong bg-white/86 px-5 py-5 text-left shadow-enterprise transition hover:bg-white"
        onclick={() => (selectedTheme = 'light')}
      >
        <p class="font-semibold text-enterprise-text-primary">Enterprise Standard</p>
        <p class="mt-1 text-sm text-enterprise-text-muted">Acik zemin, kompakt pencere yogunlugu ve sade shell dili.</p>
      </button>
    </div>

    <div class="mt-4 flex flex-wrap gap-2">
      <button
        class="enterprise-button-primary rounded-md px-4 py-3 text-sm font-semibold shadow-enterprise transition disabled:cursor-not-allowed disabled:opacity-60"
        disabled={saving}
        onclick={saveTheme}
      >
        {saving ? 'Kaydediliyor...' : 'Temayi Kaydet'}
      </button>

      <button
        class="toolbar-button rounded-md px-4 py-3 text-sm font-medium shadow-enterprise transition"
        onclick={() => moduleManager.resetWorkspaceLayout()}
      >
        Duzeni Sifirla
      </button>

      <button
        class="toolbar-button rounded-md px-4 py-3 text-sm font-medium shadow-enterprise transition"
        onclick={clearSavedWorkspaceSession}
      >
        Oturumu Temizle
      </button>
    </div>

    {#if statusMessage}
      <div class="status-info mt-4 rounded-md p-4 text-sm shadow-enterprise">
        {statusMessage}
      </div>
    {/if}
  </div>

  <div class="flex flex-col gap-4">
    <div class="surface-panel rounded-md p-5 shadow-enterprise">
      <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Workspace Status</p>
      <div class="mt-4 grid gap-3 sm:grid-cols-2">
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Theme</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">{selectedTheme.toUpperCase()}</p>
        </div>
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Autosave</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">ENABLED</p>
        </div>
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Registered Modules</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">{moduleManager.definitions.length}</p>
        </div>
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Open Windows</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">{moduleManager.openModuleCount}</p>
        </div>
      </div>
    </div>

    <div class="surface-panel rounded-md p-5 shadow-enterprise">
      <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Notes</p>
      <div class="mt-4 space-y-3 text-sm text-enterprise-text-secondary">
        <p>Tek aktif tema kurumsal acik varyanttir.</p>
        <p>Calisma alani oturumu yerel olarak saklanir ve uygulama tekrar acildiginda geri yuklenir.</p>
        <p>Klavye kisayollari ve pencere yerlesim komutlari shell seviyesinde aktif kalir.</p>
      </div>
    </div>
  </div>
</div>
