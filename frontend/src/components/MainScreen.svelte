<script lang="ts">
  import { onMount } from 'svelte';
  import type { ResizeDirection } from '@tauri-apps/api/window';
  import DraggableResizableWindow from '$components/DraggableResizableWindow.svelte';
  import { getApiBaseUrl, getJson } from '$lib/api/client';
  import { moduleRegistry } from '$lib/modules/registry';
  import type { ModuleId } from '$lib/modules/types';
  import { moduleManager } from '$lib/state/moduleManager.svelte';

  type ShellRuntimeStatus = 'checking' | 'online' | 'offline';
  type SystemContext = {
    application: string;
    version?: string;
    environment: string;
    backend: string;
    transport: string;
    modules: string[];
    authMode?: string;
  };

  let workspaceWidth = $state(0);
  let workspaceHeight = $state(0);
  let workspaceElement = $state<HTMLDivElement | null>(null);
  let tauriWindowControls = $state<{
    minimize: () => Promise<void>;
    close: () => Promise<void>;
    toggleMaximize: () => Promise<void>;
    startResizeDragging: (direction: ResizeDirection) => Promise<void>;
    isMaximized: () => Promise<boolean>;
  } | null>(null);
  let shellIsMaximized = $state(false);
  let shellRuntime = $state<'Tauri' | 'Browser'>('Browser');
  let runtimeStatus = $state<ShellRuntimeStatus>('checking');
  let runtimeMessage = $state('Sidecar durumu kontrol ediliyor...');
  let systemContext = $state<SystemContext | null>(null);
  let currentTime = $state('');

  const launcherModules = $derived(moduleManager.definitions);
  const focusedModule = $derived(moduleManager.focusedModule);

  function updateClock() {
    currentTime = new Intl.DateTimeFormat('tr-TR', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    }).format(new Date());
  }

  async function refreshRuntimeContext() {
    runtimeStatus = 'checking';

    try {
      const context = await getJson<SystemContext>('/api/system/context', 3000);
      systemContext = context;
      runtimeStatus = 'online';
      runtimeMessage = `${context.backend} aktif, ulasim ${context.transport}`;
    } catch (error) {
      runtimeStatus = 'offline';
      runtimeMessage = error instanceof Error ? error.message : 'FastAPI sidecar erisilemedi.';
    }
  }

  async function minimizeShellWindow() {
    await tauriWindowControls?.minimize();
  }

  async function closeShellWindow() {
    await tauriWindowControls?.close();
  }

  async function syncShellMaximizeState() {
    if (!tauriWindowControls) {
      shellIsMaximized = false;
      return;
    }

    shellIsMaximized = await tauriWindowControls.isMaximized();
  }

  async function toggleShellMaximize() {
    if (!tauriWindowControls) {
      return;
    }

    await tauriWindowControls.toggleMaximize();
    await syncShellMaximizeState();
  }

  async function startShellResize(direction: ResizeDirection, event: PointerEvent) {
    event.preventDefault();
    event.stopPropagation();
    await tauriWindowControls?.startResizeDragging(direction);
  }

  function openModule(id: ModuleId) {
    moduleManager.openModule(id);
  }

  function openModuleByIndex(index: number) {
    const definition = launcherModules[index];

    if (!definition) {
      return;
    }

    openModule(definition.id);
  }

  function handleToolbarAction(moduleId: ModuleId, actionId: string) {
    switch (actionId) {
      case 'center-window':
        moduleManager.centerModule(moduleId);
        break;
      case 'reset-window':
        moduleManager.resetModuleLayout(moduleId);
        break;
      default:
        break;
    }
  }

  function closeModuleFromTaskbar(event: MouseEvent, id: ModuleId) {
    event.preventDefault();
    event.stopPropagation();
    moduleManager.closeModule(id);
  }

  function toggleModuleMaximizeFromTaskbar(event: MouseEvent, id: ModuleId) {
    event.preventDefault();
    event.stopPropagation();

    const state = moduleManager.moduleStates[id];
    if (state.isMinimized) {
      state.isMinimized = false;
      moduleManager.focusModule(id);
      return;
    }

    moduleManager.toggleMaximize(id);
  }

  function getTabStatusClass(id: ModuleId) {
    const tone = moduleRegistry[id].status?.tone;

    switch (tone) {
      case 'success':
        return 'bg-enterprise-success';
      case 'warning':
        return 'bg-enterprise-warning';
      case 'danger':
        return 'bg-enterprise-danger';
      default:
        return 'bg-enterprise-info';
    }
  }

  function runtimeBadgeClass(status: ShellRuntimeStatus) {
    switch (status) {
      case 'online':
        return 'status-success';
      case 'offline':
        return 'status-danger';
      default:
        return 'status-warning';
    }
  }

  function isTextEntryTarget(target: EventTarget | null) {
    if (!(target instanceof HTMLElement)) {
      return false;
    }

    const tagName = target.tagName;
    return (
      target.isContentEditable ||
      tagName === 'INPUT' ||
      tagName === 'TEXTAREA' ||
      tagName === 'SELECT'
    );
  }

  function handleGlobalKeydown(event: KeyboardEvent) {
    if (isTextEntryTarget(event.target)) {
      return;
    }

    if (event.altKey && !event.shiftKey && !event.ctrlKey && !event.metaKey) {
      const hotkeyIndex = Number.parseInt(event.key, 10) - 1;

      if (Number.isInteger(hotkeyIndex) && hotkeyIndex >= 0 && hotkeyIndex < launcherModules.length) {
        event.preventDefault();
        openModuleByIndex(hotkeyIndex);
        return;
      }

      if (event.key.toLowerCase() === 'c') {
        event.preventDefault();
        moduleManager.cascadeWindows();
        return;
      }

      if (event.key.toLowerCase() === 't') {
        event.preventDefault();
        moduleManager.tileVisibleModules();
      }
    }

    if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 'r') {
      event.preventDefault();
      moduleManager.resetWorkspaceLayout();
    }
  }

  onMount(() => {
    let disposeResizeListener: (() => void) | undefined;
    updateClock();
    const clockTimer = window.setInterval(updateClock, 1000);
    const runtimeTimer = window.setInterval(() => {
      void refreshRuntimeContext();
    }, 15000);
    window.addEventListener('keydown', handleGlobalKeydown);

    void refreshRuntimeContext();

    void (async () => {
      try {
        const { getCurrentWindow } = await import('@tauri-apps/api/window');
        const currentWindow = getCurrentWindow();

        tauriWindowControls = {
          minimize: () => currentWindow.minimize(),
          close: () => currentWindow.close(),
          toggleMaximize: () => currentWindow.toggleMaximize(),
          startResizeDragging: (direction) => currentWindow.startResizeDragging(direction),
          isMaximized: () => currentWindow.isMaximized()
        };
        shellRuntime = 'Tauri';
        await syncShellMaximizeState();
        disposeResizeListener = await currentWindow.onResized(() => {
          void syncShellMaximizeState();
        });
      } catch {
        tauriWindowControls = null;
        shellRuntime = 'Browser';
        shellIsMaximized = false;
      }
    })();

    return () => {
      disposeResizeListener?.();
      window.clearInterval(clockTimer);
      window.clearInterval(runtimeTimer);
      window.removeEventListener('keydown', handleGlobalKeydown);
    };
  });

  $effect(() => {
    if (!workspaceElement) {
      moduleManager.setWorkspaceSize(workspaceWidth, workspaceHeight);
      return;
    }

    const rect = workspaceElement.getBoundingClientRect();
    moduleManager.setWorkspaceFrame(rect.left, rect.top, workspaceWidth, workspaceHeight);
  });
</script>

<div class="relative flex min-h-screen flex-col overflow-hidden bg-enterprise-bg text-enterprise-text-primary">
  <header class="top-bar flex h-12 items-center gap-3 border-b px-3">
    <div
      class="flex h-full min-w-[220px] items-center gap-3 px-1"
      data-tauri-drag-region
      role="presentation"
    >
      <div class="flex h-8 w-8 items-center justify-center rounded-sm border border-black/10 bg-[rgba(17,24,39,0.92)] text-[11px] font-semibold tracking-[0.18em] text-white shadow-enterprise">
        UF
      </div>
      <div class="min-w-0">
        <p class="truncate text-sm font-semibold text-enterprise-text-primary">
          {systemContext?.application ?? 'FLEXBOX'}
        </p>
        <p class="truncate text-data-xs text-enterprise-text-muted">
          Modular desktop shell for integrated subsystems
        </p>
      </div>
    </div>

    <nav class="flex min-w-0 items-center gap-2 overflow-x-auto">
      {#each launcherModules as definition, index (definition.id)}
        <button
          class="flex shrink-0 items-center gap-2 rounded-sm border border-enterprise-border-strong bg-white/80 px-3 py-1.5 text-data-sm font-medium text-enterprise-text-secondary shadow-enterprise transition hover:bg-white hover:text-enterprise-text-primary"
          onclick={() => openModule(definition.id)}
          title={`Alt+${index + 1}`}
        >
          <span class="data-code text-data-xs text-enterprise-text-muted">{definition.icon}</span>
          <span>{definition.shortTitle}</span>
        </button>
      {/each}
    </nav>

    <div
      class="min-w-8 flex-1 self-stretch"
      data-tauri-drag-region
      role="presentation"
    ></div>

    <div class="flex items-center gap-2">
      <span class={`${runtimeBadgeClass(runtimeStatus)} rounded-sm px-2.5 py-1 text-data-xs font-semibold uppercase tracking-[0.14em]`}>
        {runtimeStatus}
      </span>
      <div class="hidden items-center gap-2 rounded-sm border border-enterprise-border-strong bg-white/76 px-3 py-1.5 text-data-sm text-enterprise-text-secondary shadow-enterprise lg:flex">
        <span class="data-code">{shellRuntime}</span>
        <span class="h-3 w-px bg-enterprise-border-strong"></span>
        <span class="data-code">{systemContext?.environment ?? 'DEV'}</span>
        <span class="h-3 w-px bg-enterprise-border-strong"></span>
        <span class="data-code">v{systemContext?.version ?? '0.1.0'}</span>
        <span class="h-3 w-px bg-enterprise-border-strong"></span>
        <span class="data-code">{currentTime}</span>
      </div>
      <button
        class="flex h-7 w-7 items-center justify-center rounded-sm border border-black/10 bg-white/82 text-xs text-enterprise-text-secondary shadow-enterprise transition hover:bg-white hover:text-enterprise-text-primary"
        onclick={minimizeShellWindow}
        aria-label="Minimize application window"
        title="Minimize"
      >
        <span aria-hidden="true">-</span>
      </button>
      <button
        class="flex h-7 w-7 items-center justify-center rounded-sm border border-black/10 bg-white/82 text-[11px] text-enterprise-text-secondary shadow-enterprise transition hover:bg-white hover:text-enterprise-text-primary"
        onclick={toggleShellMaximize}
        aria-label={shellIsMaximized ? 'Restore application window' : 'Maximize application window'}
        title={shellIsMaximized ? 'Restore' : 'Maximize'}
      >
        <span aria-hidden="true">{shellIsMaximized ? '<>' : '[]'}</span>
      </button>
      <button
        class="flex h-7 w-7 items-center justify-center rounded-sm border border-[rgba(255,59,48,0.22)] bg-[rgba(255,59,48,0.10)] text-xs text-[var(--enterprise-accent)] shadow-enterprise transition hover:bg-[rgba(255,59,48,0.18)]"
        onclick={closeShellWindow}
        aria-label="Close application window"
        title="Close"
      >
        <span aria-hidden="true">x</span>
      </button>
    </div>
  </header>

  <main class="workspace-shell relative flex-1 overflow-hidden">
    <div class="absolute inset-0 overflow-hidden" bind:this={workspaceElement} bind:clientWidth={workspaceWidth} bind:clientHeight={workspaceHeight}>
      {#if moduleManager.visibleModules.length === 0}
        <section class="flex h-full items-center justify-center px-8">
          <div class="surface-panel w-full max-w-4xl rounded-md p-6 shadow-window">
            <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Workspace</p>
            <h1 class="mt-2 text-2xl font-semibold tracking-tight text-enterprise-text-primary">
              FLEXBOX
            </h1>
            <p class="mt-3 max-w-2xl text-sm leading-6 text-enterprise-text-secondary">
              Calisma alanini baslatmak icin bir modul acin. Shell; pencere yonetimi, taskbar, oturum geri yukleme
              ve sidecar baglantisini ortak olarak saglar.
            </p>

            <div class="mt-5 grid gap-3 sm:grid-cols-3">
              <div class="rounded-md border border-enterprise-border-strong bg-white/84 px-4 py-3 shadow-enterprise">
                <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Runtime</p>
                <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">{shellRuntime}</p>
              </div>
              <div class="rounded-md border border-enterprise-border-strong bg-white/84 px-4 py-3 shadow-enterprise">
                <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Backend</p>
                <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">{getApiBaseUrl()}</p>
              </div>
              <div class="rounded-md border border-enterprise-border-strong bg-white/84 px-4 py-3 shadow-enterprise">
                <p class="text-data-xs uppercase tracking-[0.14em] text-enterprise-text-muted">Modules</p>
                <p class="data-code mt-1 text-sm font-semibold text-enterprise-text-primary">{moduleManager.definitions.length}</p>
              </div>
            </div>

            <div class="mt-5 flex flex-wrap gap-2">
              {#each launcherModules as definition, index (definition.id)}
                <button
                  class="enterprise-button-primary rounded-sm px-4 py-2 text-sm font-semibold shadow-enterprise"
                  onclick={() => openModule(definition.id)}
                  title={`Alt+${index + 1}`}
                >
                  Open {definition.shortTitle}
                </button>
              {/each}
            </div>
          </div>
        </section>
      {/if}

      {#each moduleManager.visibleModules as definition (definition.id)}
        {@const moduleEntry = moduleRegistry[definition.id]}
        {@const ModuleComponent = moduleEntry.component}
        <DraggableResizableWindow
          moduleId={definition.id}
          title={definition.title}
          windowMark={definition.icon}
          subtitle={definition.subtitle}
          status={moduleEntry.status}
          toolbarActions={moduleEntry.toolbarActions}
          footerMeta={moduleEntry.footerMeta}
          onToolbarAction={(actionId) => handleToolbarAction(definition.id, actionId)}
        >
          <ModuleComponent />
        </DraggableResizableWindow>
      {/each}
    </div>
  </main>

  <footer class="task-bar flex items-center justify-between gap-3 border-t px-3 py-1">
    <div class="enterprise-scrollbar flex min-w-0 flex-1 items-center gap-2 overflow-x-auto py-1">
      {#each moduleManager.openModules as definition (definition.id)}
        {@const state = moduleManager.moduleStates[definition.id]}
        <div
          class={`task-tab-group max-w-[18rem] rounded-sm border shadow-enterprise transition ${
            state.isMinimized
              ? 'border-enterprise-border bg-white/58 text-enterprise-text-muted'
              : moduleManager.isFocused(definition.id)
                ? 'border-enterprise-border-strong bg-white text-enterprise-text-primary shadow-window'
                : 'border-enterprise-border bg-white/74 text-enterprise-text-secondary'
          }`}
        >
          <div
            class={`task-tab relative flex w-full items-center gap-3 px-3 py-1.5 text-left transition ${
              moduleManager.isFocused(definition.id) && !state.isMinimized ? 'task-tab-active' : ''
            }`}
            onclick={() => moduleManager.handleTaskbarClick(definition.id)}
            ondblclick={() => moduleManager.toggleMaximize(definition.id)}
            onkeydown={(event) => {
              if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                moduleManager.handleTaskbarClick(definition.id);
              }
            }}
            title={`${definition.title} - ${definition.subtitle}`}
            role="button"
            tabindex="0"
          >
            <span class="min-w-0 flex-1">
              <span class="flex items-center gap-2">
                <span class="truncate text-sm font-semibold">{definition.shortTitle}</span>
                <span class={`h-2 w-2 rounded-sm ${getTabStatusClass(definition.id)}`}></span>
                {#if state.isMinimized}
                  <span class="rounded-sm border border-enterprise-border bg-enterprise-bg-soft px-1.5 py-0.5 text-[10px] uppercase tracking-[0.12em] text-enterprise-text-muted">
                    Min
                  </span>
                {/if}
              </span>
            </span>

            <span class="flex shrink-0 items-center gap-1">
              <button
                class="task-tab-icon"
                onclick={(event) => toggleModuleMaximizeFromTaskbar(event, definition.id)}
                aria-label={state.isMaximized ? 'Restore module window' : 'Maximize module window'}
                title={state.isMaximized ? 'Restore' : 'Maximize'}
              >
                <span aria-hidden="true">{state.isMaximized ? '<>' : '[]'}</span>
              </button>
              <button
                class="task-tab-icon task-tab-icon-danger"
                onclick={(event) => closeModuleFromTaskbar(event, definition.id)}
                aria-label="Close module window"
                title="Close"
              >
                <span aria-hidden="true">x</span>
              </button>
            </span>
          </div>
        </div>
      {/each}

      {#if moduleManager.openModules.length === 0}
        <p class="px-2 text-sm text-enterprise-text-muted">Acik modul bulunmuyor.</p>
      {/if}
    </div>

    <div class="hidden items-center gap-2 xl:flex">
      <button class="toolbar-button rounded-sm px-3 py-1.5 text-data-sm font-medium shadow-enterprise" onclick={() => moduleManager.cascadeWindows()}>
        Cascade
      </button>
      <button class="toolbar-button rounded-sm px-3 py-1.5 text-data-sm font-medium shadow-enterprise" onclick={() => moduleManager.tileVisibleModules()}>
        Tile
      </button>
      <button class="toolbar-button rounded-sm px-3 py-1.5 text-data-sm font-medium shadow-enterprise" onclick={() => moduleManager.resetWorkspaceLayout()}>
        Reset
      </button>
    </div>

    <div class="hidden items-center gap-3 rounded-sm border border-enterprise-border-strong bg-white/76 px-3 py-1.5 text-data-sm text-enterprise-text-secondary shadow-enterprise lg:flex">
      <span class="data-code">{moduleManager.openModuleCount.toString().padStart(2, '0')} open</span>
      <span class="h-3 w-px bg-enterprise-border-strong"></span>
      <span class="data-code">{moduleManager.minimizedModuleCount.toString().padStart(2, '0')} min</span>
      <span class="h-3 w-px bg-enterprise-border-strong"></span>
      <span class="truncate">{focusedModule?.shortTitle ?? 'Workspace idle'}</span>
      <span class="h-3 w-px bg-enterprise-border-strong"></span>
      <span class="max-w-64 truncate text-enterprise-text-muted" title={runtimeMessage}>{runtimeMessage}</span>
    </div>
  </footer>

  <div class="absolute inset-x-4 top-0 z-40 h-1.5 cursor-n-resize touch-none" onpointerdown={(event) => startShellResize('North', event)} role="presentation"></div>
  <div class="absolute inset-x-4 bottom-0 z-40 h-1.5 cursor-s-resize touch-none" onpointerdown={(event) => startShellResize('South', event)} role="presentation"></div>
  <div class="absolute inset-y-4 left-0 z-40 w-1.5 cursor-w-resize touch-none" onpointerdown={(event) => startShellResize('West', event)} role="presentation"></div>
  <div class="absolute inset-y-4 right-0 z-40 w-1.5 cursor-e-resize touch-none" onpointerdown={(event) => startShellResize('East', event)} role="presentation"></div>
  <div class="absolute left-0 top-0 z-40 h-4 w-4 cursor-nw-resize touch-none" onpointerdown={(event) => startShellResize('NorthWest', event)} role="presentation"></div>
  <div class="absolute right-0 top-0 z-40 h-4 w-4 cursor-ne-resize touch-none" onpointerdown={(event) => startShellResize('NorthEast', event)} role="presentation"></div>
  <div class="absolute bottom-0 left-0 z-40 h-4 w-4 cursor-sw-resize touch-none" onpointerdown={(event) => startShellResize('SouthWest', event)} role="presentation"></div>
  <div class="absolute bottom-0 right-0 z-40 h-4 w-4 cursor-se-resize touch-none" onpointerdown={(event) => startShellResize('SouthEast', event)} role="presentation"></div>
</div>
