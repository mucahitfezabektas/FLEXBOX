<script lang="ts">
  import { invoke, isTauri } from '@tauri-apps/api/core';

  let confirming = $state(false);
  let shuttingDown = $state(false);
  let statusMessage = $state('');

  const isDesktopRuntime = isTauri();

  async function runShutdown() {
    if (!isDesktopRuntime || shuttingDown) {
      return;
    }

    shuttingDown = true;
    statusMessage = 'FLEXBOX kapatma akisi baslatiliyor...';

    try {
      await invoke('shutdown_application');
    } catch (error) {
      shuttingDown = false;
      statusMessage =
        error instanceof Error ? error.message : 'Shutdown komutu gonderilemedi.';
    }
  }
</script>

<div class="grid h-full gap-4 xl:grid-cols-[0.96fr_1.04fr]">
  <div class="surface-panel-alt rounded-md p-5 shadow-enterprise">
    <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">System Control</p>
    <h2 class="mt-2 text-xl font-semibold tracking-tight text-enterprise-text-primary">Controlled shutdown</h2>
    <p class="mt-2 text-sm text-enterprise-text-muted">
      Shell penceresi, Tauri runtime ve sidecar backend tek akista kapatilir.
    </p>

    <div class="mt-5 rounded-md border border-[rgba(159,49,49,0.2)] bg-[rgba(180,55,55,0.06)] p-4 text-sm text-enterprise-text-secondary shadow-enterprise">
      Bu islem mevcut FLEXBOX oturumunu sonlandirir. Acik modul durumlari kayitliysa bir sonraki acilista geri yuklenir.
    </div>

    <div class="mt-5 flex flex-wrap gap-2">
      {#if !confirming}
        <button
          class="rounded-md border border-[rgba(159,49,49,0.3)] bg-[linear-gradient(180deg,rgba(191,68,68,0.96),rgba(143,39,39,0.96))] px-4 py-3 text-sm font-semibold text-white shadow-enterprise transition hover:translate-y-[-1px]"
          onclick={() => {
            confirming = true;
            statusMessage = 'Kapatma islemini onaylayin.';
          }}
        >
          Shutdown FLEXBOX
        </button>
      {:else}
        <button
          class="rounded-md border border-[rgba(159,49,49,0.3)] bg-[linear-gradient(180deg,rgba(191,68,68,0.96),rgba(143,39,39,0.96))] px-4 py-3 text-sm font-semibold text-white shadow-enterprise transition hover:translate-y-[-1px] disabled:cursor-not-allowed disabled:opacity-60"
          onclick={runShutdown}
          disabled={!isDesktopRuntime || shuttingDown}
        >
          {shuttingDown ? 'Shutting down...' : 'Confirm shutdown'}
        </button>
        <button
          class="toolbar-button rounded-md px-4 py-3 text-sm font-medium shadow-enterprise transition"
          onclick={() => {
            confirming = false;
            statusMessage = 'Shutdown iptal edildi.';
          }}
          disabled={shuttingDown}
        >
          Cancel
        </button>
      {/if}
    </div>

    {#if statusMessage}
      <div class="status-warning mt-4 rounded-md p-4 text-sm shadow-enterprise">
        {statusMessage}
      </div>
    {/if}
  </div>

  <div class="flex flex-col gap-4">
    <div class="surface-panel rounded-md p-5 shadow-enterprise">
      <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Scope</p>
      <div class="mt-4 grid gap-3 sm:grid-cols-2">
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Runtime</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">
            {isDesktopRuntime ? 'TAURI' : 'BROWSER'}
          </p>
        </div>
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Mode</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">CONTROLLED EXIT</p>
        </div>
      </div>
    </div>

    <div class="surface-panel rounded-md p-5 shadow-enterprise">
      <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Behavior</p>
      <div class="mt-4 space-y-3 text-sm text-enterprise-text-secondary">
        <p>Frontend shell kapatilir.</p>
        <p>Backend sidecar process tree sonlandirilir.</p>
        <p>Windows installer lock problemi olusmaması icin kapanis tek komutla yonetilir.</p>
      </div>
    </div>
  </div>
</div>
