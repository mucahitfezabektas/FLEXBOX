<script lang="ts">
  import type { Snippet } from 'svelte';
  import { onDestroy } from 'svelte';
  import { cubicOut } from 'svelte/easing';
  import { moduleManager, type ModuleId, type SnapTarget } from '$lib/state/moduleManager.svelte';
  import type {
    WindowStatus,
    WindowStatusTone,
  } from '$lib/types/windowChrome';
  import type { ModuleTheme, WindowBounds } from '$lib/modules/types';

  type ResizeHandle = 'n' | 's' | 'e' | 'w' | 'nw' | 'ne' | 'sw' | 'se';

  interface DragInteraction {
    type: 'drag';
    originX: number;
    originY: number;
    startX: number;
    startY: number;
    startWidth: number;
    startHeight: number;
    fromDocked?: boolean;
    hasRestored?: boolean;
    restorePointerRatio?: number;
    restoreOffsetY?: number;
  }

  interface ResizeInteraction {
    type: 'resize';
    handle: ResizeHandle;
    originX: number;
    originY: number;
    startX: number;
    startY: number;
    startWidth: number;
    startHeight: number;
  }

  let {
    moduleId,
    title,
    windowMark = '',
    subtitle = '',
    theme = 'default',
    status = null,
    children
  }: {
    moduleId: ModuleId;
    title: string;
    windowMark?: string;
    subtitle?: string;
    theme?: ModuleTheme;
    status?: WindowStatus | null;
    children: Snippet;
  } = $props();

  const windowState = $derived(moduleManager.moduleStates[moduleId]);
  const highestZIndex = $derived(
    Math.max(...moduleManager.openModules.map(({ id }) => moduleManager.moduleStates[id].zIndex), 0)
  );
  const isFocused = $derived(windowState.zIndex === highestZIndex);
  let lastHeaderDoubleClickAt = $state(0);
  let interaction = $state<DragInteraction | ResizeInteraction | null>(null);
  let activeSnapTarget = $state<SnapTarget | null>(null);
  let snapPreview = $state<WindowBounds | null>(null);
  let shouldAnimateFrame = $state(true);
  let frameSyncId = 0;
  let queuedFrameBounds: WindowBounds | null = null;

  function getInitialFrameBounds(): WindowBounds {
    const state = moduleManager.moduleStates[moduleId];

    return {
      x: state.x,
      y: state.y,
      width: state.width,
      height: state.height
    };
  }

  let frameBounds = $state<WindowBounds>(getInitialFrameBounds());

  $effect(() => {
    if (interaction) {
      return;
    }

    shouldAnimateFrame = true;
    frameBounds.x = windowState.x;
    frameBounds.y = windowState.y;
    frameBounds.width = windowState.width;
    frameBounds.height = windowState.height;
  });

  function windowEnter(node: Element) {
    return {
      duration: 180,
      easing: cubicOut,
      css: (t: number) =>
        `opacity:${t};transform:translate3d(${frameBounds.x}px, ${frameBounds.y + (1 - t) * 10}px, 0) scale(${0.985 + t * 0.015});`
    };
  }

  function toneClass(tone: WindowStatusTone) {
    switch (tone) {
      case 'success':
        return 'status-success';
      case 'warning':
        return 'status-warning';
      case 'danger':
        return 'status-danger';
      default:
        return 'status-info';
    }
  }

  function snapLabel(target: SnapTarget) {
    switch (target) {
      case 'maximize':
        return 'Full Screen';
      case 'left':
        return 'Left Split';
      case 'right':
        return 'Right Split';
      case 'top-left':
        return 'Top Left';
      case 'top-right':
        return 'Top Right';
      case 'bottom-left':
        return 'Bottom Left';
      case 'bottom-right':
        return 'Bottom Right';
    }
  }

  function beginGlobalInteraction() {
    shouldAnimateFrame = false;
    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', stopInteraction);
    document.body.style.userSelect = 'none';
    document.body.style.cursor = interaction?.type === 'resize' ? 'nwse-resize' : 'grabbing';
  }

  function stopGlobalInteraction() {
    window.removeEventListener('pointermove', handlePointerMove);
    window.removeEventListener('pointerup', stopInteraction);
    document.body.style.userSelect = '';
    document.body.style.cursor = '';
  }

  function flushQueuedFrameBounds() {
    frameSyncId = 0;

    if (!queuedFrameBounds) {
      return;
    }

    frameBounds.x = queuedFrameBounds.x;
    frameBounds.y = queuedFrameBounds.y;
    frameBounds.width = queuedFrameBounds.width;
    frameBounds.height = queuedFrameBounds.height;
    queuedFrameBounds = null;
  }

  function scheduleFrameBounds(next: WindowBounds) {
    queuedFrameBounds = next;

    if (frameSyncId) {
      return;
    }

    frameSyncId = window.requestAnimationFrame(flushQueuedFrameBounds);
  }

  function applyFrameBoundsImmediately(next: WindowBounds) {
    queuedFrameBounds = null;

    if (frameSyncId) {
      window.cancelAnimationFrame(frameSyncId);
      frameSyncId = 0;
    }

    frameBounds.x = next.x;
    frameBounds.y = next.y;
    frameBounds.width = next.width;
    frameBounds.height = next.height;
  }

  function startDrag(event: PointerEvent) {
    if (event.button !== 0) {
      return;
    }

    if (event.detail > 1) {
      return;
    }

    if ((event.target as HTMLElement).closest('button')) {
      return;
    }

    if (Date.now() - lastHeaderDoubleClickAt < 240) {
      return;
    }

    event.preventDefault();
    moduleManager.focusModule(moduleId);
    const headerElement = event.currentTarget as HTMLElement;

    if (windowState.isMaximized || windowState.snapTarget) {
      const headerRect = headerElement.getBoundingClientRect();
      interaction = {
        type: 'drag',
        originX: event.clientX,
        originY: event.clientY,
        startX: windowState.x,
        startY: windowState.y,
        startWidth: windowState.width,
        startHeight: windowState.height,
        fromDocked: true,
        hasRestored: false,
        restorePointerRatio: Math.min(
          Math.max((event.clientX - headerRect.left) / Math.max(1, headerRect.width), 0.18),
          0.82
        ),
        restoreOffsetY: Math.min(Math.max(headerRect.height * 0.55, 18), 28)
      };
      beginGlobalInteraction();
      return;
    }

    frameBounds.x = windowState.x;
    frameBounds.y = windowState.y;
    frameBounds.width = windowState.width;
    frameBounds.height = windowState.height;
    interaction = {
      type: 'drag',
      originX: event.clientX,
      originY: event.clientY,
      startX: windowState.x,
      startY: windowState.y,
      startWidth: windowState.width,
      startHeight: windowState.height
    };
    beginGlobalInteraction();
  }

  function startResize(handle: ResizeHandle, event: PointerEvent) {
    if (event.button !== 0 || windowState.isMaximized) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    moduleManager.focusModule(moduleId);
    frameBounds.x = windowState.x;
    frameBounds.y = windowState.y;
    frameBounds.width = windowState.width;
    frameBounds.height = windowState.height;
    interaction = {
      type: 'resize',
      handle,
      originX: event.clientX,
      originY: event.clientY,
      startX: windowState.x,
      startY: windowState.y,
      startWidth: windowState.width,
      startHeight: windowState.height
    };
    beginGlobalInteraction();
  }

  function handlePointerMove(event: PointerEvent) {
    if (!interaction) {
      return;
    }

    const deltaX = event.clientX - interaction.originX;
    const deltaY = event.clientY - interaction.originY;

    if (interaction.type === 'drag') {
      if (interaction.fromDocked && !interaction.hasRestored) {
        const travel = Math.hypot(deltaX, deltaY);

        if (travel < 8) {
          return;
        }

        const restored = moduleManager.restoreFromDocked(moduleId);
        const next = moduleManager.getConstrainedBounds(moduleId, {
          x:
            event.clientX -
            moduleManager.workspaceFrame.left -
            restored.width * (interaction.restorePointerRatio ?? 0.5),
          y:
            event.clientY -
            moduleManager.workspaceFrame.top -
            (interaction.restoreOffsetY ?? 24),
          width: restored.width,
          height: restored.height
        });

        moduleManager.updateBounds(moduleId, next);
        applyFrameBoundsImmediately(next);
        interaction.startX = next.x;
        interaction.startY = next.y;
        interaction.startWidth = next.width;
        interaction.startHeight = next.height;
        interaction.originX = event.clientX;
        interaction.originY = event.clientY;
        interaction.hasRestored = true;
      }

      const dragDeltaX = event.clientX - interaction.originX;
      const dragDeltaY = event.clientY - interaction.originY;
      activeSnapTarget = resolveSnapTarget(event.clientX, event.clientY);
      snapPreview = activeSnapTarget ? moduleManager.getSnapBounds(moduleId, activeSnapTarget) : null;

      const next = moduleManager.getConstrainedBounds(moduleId, {
        x: interaction.startX + dragDeltaX,
        y: interaction.startY + dragDeltaY,
        width: frameBounds.width,
        height: frameBounds.height
      });
      scheduleFrameBounds(next);
      return;
    }

    activeSnapTarget = null;
    snapPreview = null;

    let nextX = interaction.startX;
    let nextY = interaction.startY;
    let nextWidth = interaction.startWidth;
    let nextHeight = interaction.startHeight;

    if (interaction.handle.includes('e')) {
      nextWidth = interaction.startWidth + deltaX;
    }

    if (interaction.handle.includes('s')) {
      nextHeight = interaction.startHeight + deltaY;
    }

    if (interaction.handle.includes('w')) {
      nextX = interaction.startX + deltaX;
      nextWidth = interaction.startWidth - deltaX;
    }

    if (interaction.handle.includes('n')) {
      nextY = interaction.startY + deltaY;
      nextHeight = interaction.startHeight - deltaY;
    }

    const next = moduleManager.getConstrainedBounds(moduleId, {
      x: nextX,
      y: nextY,
      width: nextWidth,
      height: nextHeight
    });
    scheduleFrameBounds(next);
  }

  function stopInteraction() {
    if (interaction) {
      if (queuedFrameBounds) {
        applyFrameBoundsImmediately(queuedFrameBounds);
      }

      if (interaction.type === 'drag' && interaction.fromDocked && !interaction.hasRestored) {
        interaction = null;
        activeSnapTarget = null;
        snapPreview = null;
        stopGlobalInteraction();
        return;
      }

      if (interaction.type === 'drag' && activeSnapTarget) {
        moduleManager.applySnapTarget(moduleId, activeSnapTarget, {
          x: interaction.startX,
          y: interaction.startY,
          width: interaction.startWidth,
          height: interaction.startHeight
        });
      } else {
        moduleManager.updateBounds(moduleId, {
          x: frameBounds.x,
          y: frameBounds.y,
          width: frameBounds.width,
          height: frameBounds.height
        });
      }
    }

    interaction = null;
    shouldAnimateFrame = true;
    activeSnapTarget = null;
    snapPreview = null;
    stopGlobalInteraction();
  }

  function resolveSnapTarget(clientX: number, clientY: number): SnapTarget | null {
    const frame = moduleManager.workspaceFrame;

    if (!frame.width || !frame.height) {
      return null;
    }

    const horizontalThreshold = Math.min(72, Math.max(36, frame.width * 0.06));
    const verticalThreshold = Math.min(72, Math.max(36, frame.height * 0.07));
    const nearLeft = clientX <= frame.left + horizontalThreshold;
    const nearRight = clientX >= frame.left + frame.width - horizontalThreshold;
    const nearTop = clientY <= frame.top + verticalThreshold;
    const nearBottom = clientY >= frame.top + frame.height - verticalThreshold;

    if (nearTop && nearLeft) {
      return 'top-left';
    }

    if (nearTop && nearRight) {
      return 'top-right';
    }

    if (nearBottom && nearLeft) {
      return 'bottom-left';
    }

    if (nearBottom && nearRight) {
      return 'bottom-right';
    }

    if (nearTop) {
      return 'maximize';
    }

    if (nearLeft) {
      return 'left';
    }

    if (nearRight) {
      return 'right';
    }

    return null;
  }

  function getEdgeState(bounds: WindowBounds) {
    const workspaceWidth = moduleManager.workspaceSize.width;
    const workspaceHeight = moduleManager.workspaceSize.height;
    const epsilon = 1;
    return {
      touchesLeft: bounds.x <= epsilon,
      touchesTop: bounds.y <= epsilon,
      touchesRight: bounds.x + bounds.width >= workspaceWidth - epsilon,
      touchesBottom: bounds.y + bounds.height >= workspaceHeight - epsilon
    };
  }

  function getEdgeShapeClasses(bounds: WindowBounds) {
    const { touchesLeft, touchesTop, touchesRight, touchesBottom } = getEdgeState(bounds);

    if (touchesLeft && touchesTop && touchesRight && touchesBottom) {
      return 'rounded-none';
    }

    const classes = ['rounded-md'];

    if (touchesLeft && touchesTop) {
      classes.push('rounded-tl-none');
    }

    if (touchesRight && touchesTop) {
      classes.push('rounded-tr-none');
    }

    if (touchesLeft && touchesBottom) {
      classes.push('rounded-bl-none');
    }

    if (touchesRight && touchesBottom) {
      classes.push('rounded-br-none');
    }

    return classes.join(' ');
  }

  function isDocked(bounds: WindowBounds) {
    const { touchesLeft, touchesTop, touchesRight, touchesBottom } = getEdgeState(bounds);
    return touchesLeft || touchesTop || touchesRight || touchesBottom;
  }

  function getDockBorderClasses(bounds: WindowBounds) {
    const { touchesLeft, touchesTop, touchesRight, touchesBottom } = getEdgeState(bounds);
    const classes: string[] = [];

    if (touchesLeft) {
      classes.push('border-l-0');
    }

    if (touchesTop) {
      classes.push('border-t-0');
    }

    if (touchesRight) {
      classes.push('border-r-0');
    }

    if (touchesBottom) {
      classes.push('border-b-0');
    }

    return classes.join(' ');
  }

  function handleHeaderDoubleClick(event: MouseEvent) {
    if ((event.target as HTMLElement).closest('button')) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    lastHeaderDoubleClickAt = Date.now();
    stopInteraction();
    moduleManager.toggleMaximize(moduleId);
  }

  onDestroy(() => {
    if (frameSyncId) {
      window.cancelAnimationFrame(frameSyncId);
    }
    stopGlobalInteraction();
  });
</script>

  {#if snapPreview && interaction?.type === 'drag'}
    <div
      class="snap-preview pointer-events-none absolute rounded-md border border-[rgba(59,130,246,0.42)] bg-[rgba(255,255,255,0.18)] shadow-enterprise-focus"
      style={`transform: translate3d(${snapPreview.x}px, ${snapPreview.y}px, 0); width: ${snapPreview.width}px; height: ${snapPreview.height}px; z-index: ${windowState.zIndex - 1};`}
    >
      {#if activeSnapTarget}
        <div class="snap-preview-label absolute left-3 top-3 rounded-sm border border-[rgba(59,130,246,0.24)] bg-white/84 px-2 py-1 text-data-xs font-semibold uppercase tracking-[0.14em] text-enterprise-text-primary">
          {snapLabel(activeSnapTarget)}
        </div>
      {/if}
    </div>
  {/if}

<div
  in:windowEnter
  class={`window-panel absolute flex flex-col overflow-hidden ${
    theme === 'sheets' ? 'window-theme-sheets' : 'window-theme-default'
  } ${
    shouldAnimateFrame
      ? 'transition-[transform,width,height,opacity,border-color,box-shadow] duration-180 ease-[cubic-bezier(0.22,1,0.36,1)]'
      : 'transition-[border-color,box-shadow] duration-100'
  } ${
    getEdgeShapeClasses(frameBounds)
  } ${
    interaction ? 'will-change-transform' : ''
  } ${
    interaction ? 'drag-active' : ''
  } ${
    isDocked(frameBounds) ? `bg-white backdrop-blur-none ${getDockBorderClasses(frameBounds)}` : ''
  } ${
    windowState.isMaximized
      ? 'border-enterprise-border-strong shadow-none'
      : isDocked(frameBounds)
        ? 'border-enterprise-border-strong shadow-none'
      : interaction?.type === 'drag'
        ? 'border-[rgba(59,130,246,0.28)] shadow-enterprise-focus'
      : isFocused
      ? 'border-[var(--enterprise-window-focus)] shadow-enterprise-focus'
      : 'border-white/55 shadow-window'
  }`}
  style={`transform: translate3d(${frameBounds.x}px, ${frameBounds.y}px, 0); width: ${frameBounds.width}px; height: ${frameBounds.height}px; z-index: ${windowState.zIndex};`}
  onpointerdown={() => moduleManager.focusModule(moduleId)}
  role="dialog"
  aria-label={title}
  tabindex="-1"
>
  <div class="window-theme-accent absolute inset-x-0 top-0 h-[2px]" aria-hidden="true"></div>

  <div class={`window-header select-none flex flex-col ${isDocked(frameBounds) ? 'bg-white' : ''}`} onpointerdown={startDrag} ondblclick={handleHeaderDoubleClick} role="presentation">
    <div class="flex h-11 cursor-grab items-center justify-between gap-3 px-3.5 active:cursor-grabbing">
      <div class="flex min-w-0 items-center gap-3">
        <div class="window-theme-mark flex h-7 w-7 shrink-0 items-center justify-center rounded-sm border text-[10px] font-semibold uppercase tracking-[0.18em] text-white shadow-[inset_0_1px_0_rgba(255,255,255,0.08)]">
          {windowMark || title.slice(0, 2)}
        </div>
        <div class="min-w-0">
          <p class="truncate text-sm font-semibold text-enterprise-text-primary">{title}</p>
          {#if subtitle}
            <p class="truncate text-data-xs text-enterprise-text-muted">{subtitle}</p>
          {/if}
        </div>
      </div>

      <div class="flex items-center gap-2">
        {#if status}
          <span class={`${toneClass(status.tone)} rounded-sm px-2.5 py-1 text-data-xs font-medium`}>
            {status.label}
          </span>
        {/if}
        <button
          class="window-control-button flex h-8 w-10 items-center justify-center rounded-none border-0 bg-transparent text-xs text-enterprise-text-secondary shadow-none transition"
          onpointerdown={(event) => event.stopPropagation()}
          onclick={() => moduleManager.toggleMinimize(moduleId)}
          aria-label="Minimize window"
          title="Minimize"
        >
          <span aria-hidden="true">_</span>
        </button>
        <button
          class="window-control-button flex h-8 w-10 items-center justify-center rounded-none border-0 bg-transparent text-[11px] text-enterprise-text-secondary shadow-none transition"
          onpointerdown={(event) => event.stopPropagation()}
          onclick={() => moduleManager.toggleMaximize(moduleId)}
          aria-label={windowState.isMaximized ? 'Restore window' : 'Maximize window'}
          title={windowState.isMaximized ? 'Restore' : 'Maximize'}
        >
          <span aria-hidden="true">{windowState.isMaximized ? '<>' : '[]'}</span>
        </button>
        <button
          class="window-control-button window-control-button-danger flex h-8 w-12 items-center justify-center rounded-none border-0 bg-transparent text-xs shadow-none transition"
          onpointerdown={(event) => event.stopPropagation()}
          onclick={() => moduleManager.closeModule(moduleId)}
          aria-label="Close window"
          title="Close"
        >
          <span aria-hidden="true">x</span>
        </button>
      </div>
    </div>

  </div>

  <div class={`enterprise-scrollbar flex-1 overflow-auto p-4 ${isDocked(frameBounds) ? 'bg-white' : 'bg-white/72 backdrop-blur-[6px]'}`}>
    {@render children()}
  </div>

  {#if !windowState.isMaximized}
    <div class="absolute inset-x-3 top-0 z-20 h-2 cursor-n-resize touch-none" onpointerdown={(event) => startResize('n', event)} role="presentation"></div>
    <div class="absolute inset-x-3 bottom-0 z-20 h-2 cursor-s-resize touch-none" onpointerdown={(event) => startResize('s', event)} role="presentation"></div>
    <div class="absolute inset-y-3 right-0 z-20 w-2 cursor-e-resize touch-none" onpointerdown={(event) => startResize('e', event)} role="presentation"></div>
    <div class="absolute inset-y-3 left-0 z-20 w-2 cursor-w-resize touch-none" onpointerdown={(event) => startResize('w', event)} role="presentation"></div>
    <div class="absolute left-0 top-0 z-30 h-3.5 w-3.5 cursor-nw-resize touch-none" onpointerdown={(event) => startResize('nw', event)} role="presentation"></div>
    <div class="absolute right-0 top-0 z-30 h-3.5 w-3.5 cursor-ne-resize touch-none" onpointerdown={(event) => startResize('ne', event)} role="presentation"></div>
    <div class="absolute bottom-0 left-0 z-30 h-3.5 w-3.5 cursor-sw-resize touch-none" onpointerdown={(event) => startResize('sw', event)} role="presentation"></div>
    <div class="absolute bottom-0 right-0 z-30 h-3.5 w-3.5 cursor-se-resize touch-none" onpointerdown={(event) => startResize('se', event)} role="presentation"></div>
    <div class="absolute bottom-1.5 right-1.5 z-30 flex h-5 w-5 cursor-se-resize touch-none items-end justify-end rounded-sm border border-black/10 bg-white/92 pr-0.5 pb-0.5 text-[10px] leading-none text-enterprise-text-muted shadow-enterprise" onpointerdown={(event) => startResize('se', event)} role="presentation">
      //
    </div>
  {/if}
</div>

<style>
  .window-theme-default {
    --module-window-accent-strong: rgba(59, 130, 246, 0.88);
    --module-window-accent-soft: rgba(59, 130, 246, 0.08);
    --module-window-mark-bg: rgba(17, 24, 39, 0.96);
    --module-window-mark-border: rgba(0, 0, 0, 0.1);
  }

  .window-theme-sheets {
    --module-window-accent-strong: rgba(111, 116, 88, 0.96);
    --module-window-accent-soft: rgba(111, 116, 88, 0.12);
    --module-window-mark-bg: linear-gradient(180deg, #73795b 0%, #565c44 100%);
    --module-window-mark-border: rgba(79, 86, 62, 0.24);
  }

  .window-theme-accent {
    background: linear-gradient(
      90deg,
      var(--module-window-accent-strong),
      var(--module-window-accent-soft)
    );
  }

  .window-theme-mark {
    border-color: var(--module-window-mark-border);
    background: var(--module-window-mark-bg);
  }

  .window-theme-sheets .window-header {
    background:
      linear-gradient(180deg, rgba(251, 250, 244, 0.98) 0%, rgba(240, 237, 224, 0.98) 100%);
  }

</style>
