<script lang="ts">
  import { onMount } from 'svelte';
  import { ApiError, getApiBaseUrl, getJson, postJson } from '$lib/api/client';

  type LoginResponse = {
    success: boolean;
    message: string;
    token?: string;
  };

  let username = $state('admin');
  let password = $state('uniframe');
  let loading = $state(false);
  let response = $state<LoginResponse | null>(null);
  let errorMessage = $state('');
  let backendStatus = $state<'checking' | 'online' | 'offline'>('checking');
  let backendMessage = $state('Baglanti kontrol ediliyor...');

  async function checkBackendHealth() {
    backendStatus = 'checking';
    backendMessage = 'Baglanti kontrol ediliyor...';

    try {
      const result = await getJson<{ status: string }>('/api/health', 2500);
      backendStatus = result.status === 'ok' ? 'online' : 'offline';
      backendMessage =
        result.status === 'ok'
          ? `FastAPI servisi erisilebilir: ${getApiBaseUrl()}`
          : 'Saglik denetimi beklenmeyen bir durum dondu.';
    } catch (error) {
      backendStatus = 'offline';
      backendMessage = error instanceof Error ? error.message : 'API servisine ulasilamadi.';
    }
  }

  async function login() {
    loading = true;
    response = null;
    errorMessage = '';

    try {
      const data = await postJson<LoginResponse>(
        '/api/auth/login',
        {
          username,
          password
        },
        6000
      );

      backendStatus = 'online';
      backendMessage = `FastAPI servisi erisilebilir: ${getApiBaseUrl()}`;
      response = data;
    } catch (error) {
      if (error instanceof ApiError && error.status === 0) {
        backendStatus = 'offline';
        backendMessage = error.message;
      }

      errorMessage = error instanceof Error ? error.message : 'Beklenmeyen bir hata olustu.';
    } finally {
      loading = false;
    }
  }

  onMount(() => {
    void checkBackendHealth();
  });

  function handleSubmit(event: SubmitEvent) {
    event.preventDefault();
    void login();
  }
</script>

<div class="grid h-full gap-4 xl:grid-cols-[0.92fr_1.08fr]">
  <form class="surface-panel-alt rounded-md p-5 shadow-enterprise" onsubmit={handleSubmit}>
    <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Sign In</p>
    <h2 class="mt-2 text-xl font-semibold tracking-tight text-enterprise-text-primary">Authentication</h2>
    <p class="mt-2 text-sm text-enterprise-text-muted">
      Yerel sidecar uzerinden kimlik dogrulama istegi gonderir.
    </p>

    <div class="mt-4 grid gap-4">
      <label class="grid gap-2">
        <span class="text-sm font-medium text-enterprise-text-secondary">Kullanici adi</span>
        <input class="enterprise-input rounded-md px-4 py-3 text-sm" bind:value={username} placeholder="admin" />
      </label>

      <label class="grid gap-2">
        <span class="text-sm font-medium text-enterprise-text-secondary">Sifre</span>
        <input
          class="enterprise-input rounded-md px-4 py-3 text-sm"
          bind:value={password}
          placeholder="uniframe"
          type="password"
        />
      </label>

      <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 text-data-sm text-enterprise-text-muted shadow-enterprise">
        Varsayilan test hesabi:
        <span class="data-code ml-1 font-medium text-enterprise-text-secondary">admin / uniframe</span>
      </div>

      <div class="flex flex-wrap gap-2">
        <button
          type="submit"
          class="enterprise-button-primary rounded-md px-4 py-3 text-sm font-semibold shadow-enterprise transition disabled:cursor-not-allowed disabled:opacity-60"
          disabled={loading}
        >
          {loading ? 'Giris yapiliyor...' : 'Giris Yap'}
        </button>

        <button
          type="button"
          class="toolbar-button rounded-md px-4 py-3 text-sm font-medium shadow-enterprise transition"
          onclick={checkBackendHealth}
        >
          Baglantiyi Yenile
        </button>
      </div>
    </div>
  </form>

  <div class="flex flex-col gap-4">
    <div class="surface-panel rounded-md p-5 shadow-enterprise">
      <div class="grid gap-3 sm:grid-cols-3">
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">API Status</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">{backendStatus.toUpperCase()}</p>
        </div>
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Endpoint</p>
          <p class="data-code mt-1 truncate text-sm font-semibold text-enterprise-text-primary">{getApiBaseUrl()}</p>
        </div>
        <div class="rounded-md border border-enterprise-border-strong bg-white/76 px-4 py-3 shadow-enterprise">
          <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Mode</p>
          <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">LOCAL</p>
        </div>
      </div>
      <p class="mt-4 text-sm text-enterprise-text-muted">{backendMessage}</p>
    </div>

    <div class="surface-panel rounded-md p-5 shadow-enterprise">
      <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Last Response</p>

      {#if response}
        <div class="status-success mt-4 rounded-md p-4 text-sm shadow-enterprise">
          <p class="font-semibold">Basarili</p>
          <p class="mt-1">{response.message}</p>
          {#if response.token}
            <p class="data-code mt-2 break-all text-data-sm opacity-90">{response.token}</p>
          {/if}
        </div>
      {:else if errorMessage}
        <div class="status-error mt-4 rounded-md p-4 text-sm shadow-enterprise">
          {errorMessage}
        </div>
      {:else}
        <div class="mt-4 rounded-md border border-dashed border-enterprise-border-strong bg-white/68 px-4 py-4 text-sm text-enterprise-text-muted">
          Henuz bir giris istegi gonderilmedi.
        </div>
      {/if}
    </div>
  </div>
</div>
