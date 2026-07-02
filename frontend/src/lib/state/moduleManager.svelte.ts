import { applyTheme, type ThemeMode } from '$lib/utils/theme';
import { moduleDefinitions } from '$lib/modules/definitions';
import type { ModuleDefinition, ModuleId, WindowBounds } from '$lib/modules/types';
import {
  clearWorkspaceSession,
  loadWorkspaceSession,
  saveWorkspaceSession,
  type WorkspaceSessionSnapshot
} from '$lib/state/workspaceSession';

export type { ModuleDefinition, ModuleId, WindowBounds } from '$lib/modules/types';
export type SnapTarget =
  | 'maximize'
  | 'left'
  | 'right'
  | 'top-left'
  | 'top-right'
  | 'bottom-left'
  | 'bottom-right';

export interface ModuleWindowState {
  id: ModuleId;
  isOpen: boolean;
  isMinimized: boolean;
  x: number;
  y: number;
  width: number;
  height: number;
  minWidth: number;
  minHeight: number;
  zIndex: number;
  hasOpened: boolean;
  isMaximized: boolean;
  snapTarget: SnapTarget | null;
  restoreBounds: WindowBounds | null;
}

function createInitialState(definition: ModuleDefinition): ModuleWindowState {
  return {
    id: definition.id,
    isOpen: false,
    isMinimized: false,
    x: definition.defaultBounds.x,
    y: definition.defaultBounds.y,
    width: definition.defaultBounds.width,
    height: definition.defaultBounds.height,
    minWidth: definition.defaultBounds.minWidth,
    minHeight: definition.defaultBounds.minHeight,
    zIndex: 1,
    hasOpened: false,
    isMaximized: false,
    snapTarget: null,
    restoreBounds: null
  };
}

class ModuleManager {
  private persistTimer: number | null = null;
  private hydratingSession = false;
  registeredModules = $state(moduleDefinitions);
  moduleStates = $state<Record<ModuleId, ModuleWindowState>>({
    auth: createInitialState(moduleDefinitions.auth),
    customize: createInitialState(moduleDefinitions.customize),
    spreadsheet: createInitialState(moduleDefinitions.spreadsheet),
    'system-shutdown': createInitialState(moduleDefinitions['system-shutdown']),
    empty: createInitialState(moduleDefinitions.empty)
  });
  workspaceSize = $state({
    width: 0,
    height: 0
  });
  workspaceFrame = $state({
    left: 0,
    top: 0,
    width: 0,
    height: 0
  });
  theme = $state<ThemeMode>('light');
  _zCounter = $state(20);
  _openOffset = $state(0);

  get definitions() {
    return Object.values(this.registeredModules);
  }

  get openModules() {
    return this.definitions.filter(({ id }) => this.moduleStates[id].isOpen);
  }

  get visibleModules() {
    return this.openModules.filter(({ id }) => !this.moduleStates[id].isMinimized);
  }

  get openModuleCount() {
    return this.openModules.length;
  }

  get minimizedModuleCount() {
    return this.openModules.filter(({ id }) => this.moduleStates[id].isMinimized).length;
  }

  get focusedModule() {
    const openModules = this.openModules;

    if (openModules.length === 0) {
      return null;
    }

    return openModules.reduce((current, next) =>
      this.moduleStates[next.id].zIndex > this.moduleStates[current.id].zIndex ? next : current
    );
  }

  setWorkspaceSize(width: number, height: number) {
    this.workspaceSize.width = width;
    this.workspaceSize.height = height;
    this.workspaceFrame.width = width;
    this.workspaceFrame.height = height;

    for (const { id } of this.definitions) {
      if (this.moduleStates[id].isMaximized) {
        this.applyMaximizedBounds(id);
      } else if (this.moduleStates[id].snapTarget) {
        const target = this.moduleStates[id].snapTarget;
        if (target) {
          this.applySnapTarget(id, target, this.moduleStates[id].restoreBounds ?? undefined, false);
        }
      } else {
        this.updateBounds(id, {});
      }
    }

    this.schedulePersist();
  }

  setWorkspaceFrame(left: number, top: number, width: number, height: number) {
    if (
      this.workspaceFrame.left === left &&
      this.workspaceFrame.top === top &&
      this.workspaceFrame.width === width &&
      this.workspaceFrame.height === height
    ) {
      return;
    }

    this.workspaceFrame.left = left;
    this.workspaceFrame.top = top;
    this.setWorkspaceSize(width, height);
  }

  setTheme(theme: ThemeMode) {
    this.theme = theme;
    applyTheme(theme);
    this.schedulePersist();
  }

  openModule(id: ModuleId) {
    const state = this.moduleStates[id];

    state.isOpen = true;
    state.isMinimized = false;

    if (!state.hasOpened) {
      const definition = this.registeredModules[id];
      const offset = this._openOffset * 28;

      state.x = definition.defaultBounds.x + offset;
      state.y = definition.defaultBounds.y + offset;
      state.width = definition.defaultBounds.width;
      state.height = definition.defaultBounds.height;
      state.hasOpened = true;
      this._openOffset = (this._openOffset + 1) % 6;
    }

    this.focusModule(id);
    this.updateBounds(id, {});
  }

  closeModule(id: ModuleId) {
    const state = this.moduleStates[id];
    state.isOpen = false;
    state.isMinimized = false;
    state.isMaximized = false;
    state.snapTarget = null;
    state.restoreBounds = null;
    this.schedulePersist();
  }

  toggleMinimize(id: ModuleId) {
    const state = this.moduleStates[id];

    if (!state.isOpen) {
      this.openModule(id);
      return;
    }

    state.isMinimized = !state.isMinimized;

    if (!state.isMinimized) {
      this.focusModule(id);
    }

    this.schedulePersist();
  }

  handleTaskbarClick(id: ModuleId) {
    const state = this.moduleStates[id];

    if (!state.isOpen) {
      this.openModule(id);
      return;
    }

    if (state.isMinimized) {
      state.isMinimized = false;
      this.focusModule(id);
      this.schedulePersist();
      return;
    }

    if (!this.isFocused(id)) {
      this.focusModule(id);
      return;
    }

    state.isMinimized = true;
    this.schedulePersist();
  }

  focusModule(id: ModuleId) {
    this._zCounter += 1;
    this.moduleStates[id].zIndex = this._zCounter;
    this.schedulePersist();
  }

  isFocused(id: ModuleId) {
    const openIds = this.openModules.map(({ id: moduleId }) => moduleId);

    if (openIds.length === 0) {
      return false;
    }

    const highestZIndex = Math.max(...openIds.map((moduleId) => this.moduleStates[moduleId].zIndex));
    return this.moduleStates[id].zIndex === highestZIndex;
  }

  toggleMaximize(id: ModuleId) {
    const state = this.moduleStates[id];

    if (state.isMaximized && state.restoreBounds) {
      state.isMaximized = false;
      const restoreBounds = state.restoreBounds;
      state.restoreBounds = null;
      this.updateBounds(id, restoreBounds);
      this.focusModule(id);
      this.schedulePersist();
      return;
    }

    state.restoreBounds = {
      x: state.x,
      y: state.y,
      width: state.width,
      height: state.height
    };
    state.isMaximized = true;
    this.applyMaximizedBounds(id);
    this.focusModule(id);
    this.schedulePersist();
  }

  restoreFromDocked(id: ModuleId): WindowBounds {
    const state = this.moduleStates[id];
    const restoreBounds =
      state.restoreBounds ?? {
        x: state.x,
        y: state.y,
        width: Math.min(state.width, this.workspaceSize.width || state.width),
        height: Math.min(state.height, this.workspaceSize.height || state.height)
      };

    state.isMaximized = false;
    state.snapTarget = null;
    state.restoreBounds = null;

    const next = this.normalizeBounds(
      {
        ...state,
        ...restoreBounds
      },
      16
    );

    state.x = next.x;
    state.y = next.y;
    state.width = next.width;
    state.height = next.height;

    this.schedulePersist();

    return {
      x: next.x,
      y: next.y,
      width: next.width,
      height: next.height
    };
  }

  resetModuleLayout(id: ModuleId) {
    const definition = this.registeredModules[id];
    const state = this.moduleStates[id];

    state.x = definition.defaultBounds.x;
    state.y = definition.defaultBounds.y;
    state.width = definition.defaultBounds.width;
    state.height = definition.defaultBounds.height;
    state.isMinimized = false;
    state.isMaximized = false;
    state.snapTarget = null;
    state.restoreBounds = null;

    if (state.isOpen) {
      this.focusModule(id);
      this.updateBounds(id, {});
    }

    this.schedulePersist();
  }

  resetWorkspaceLayout() {
    for (const { id } of this.definitions) {
      this.resetModuleLayout(id);
    }

    this.schedulePersist();
  }

  updatePosition(id: ModuleId, x: number, y: number) {
    this.updateBounds(id, { x, y });
  }

  updateSize(id: ModuleId, width: number, height: number) {
    this.updateBounds(id, { width, height });
  }

  centerModule(id: ModuleId) {
    const state = this.moduleStates[id];
    const width = Math.min(state.width, Math.max(320, this.workspaceSize.width - 32));
    const height = Math.min(state.height, Math.max(240, this.workspaceSize.height - 32));
    const x = Math.max(16, Math.floor((this.workspaceSize.width - width) / 2));
    const y = Math.max(16, Math.floor((this.workspaceSize.height - height) / 2));

    state.isMinimized = false;
    state.isMaximized = false;
    state.snapTarget = null;
    state.restoreBounds = null;
    this.updateBounds(id, { x, y, width, height });
    this.focusModule(id);
    this.schedulePersist();
  }

  cascadeWindows() {
    const visibleModules = this.visibleModules;
    const offset = 28;

    visibleModules.forEach(({ id }, index) => {
      const definition = this.registeredModules[id];
      const x = 24 + index * offset;
      const y = 24 + index * offset;
      const width = Math.min(definition.defaultBounds.width, Math.max(480, this.workspaceSize.width - 120));
      const height = Math.min(definition.defaultBounds.height, Math.max(360, this.workspaceSize.height - 120));

      const state = this.moduleStates[id];
      state.isMaximized = false;
      state.isMinimized = false;
      state.snapTarget = null;
      state.restoreBounds = null;
      this.updateBounds(id, { x, y, width, height });
      this.focusModule(id);
    });

    this.schedulePersist();
  }

  tileVisibleModules() {
    const visibleModules = this.visibleModules;
    const count = visibleModules.length;

    if (count === 0 || !this.workspaceSize.width || !this.workspaceSize.height) {
      return;
    }

    const gap = 12;
    const columns = Math.ceil(Math.sqrt(count));
    const rows = Math.ceil(count / columns);
    const width = Math.floor((this.workspaceSize.width - gap * (columns - 1)) / columns);
    const height = Math.floor((this.workspaceSize.height - gap * (rows - 1)) / rows);

    visibleModules.forEach(({ id }, index) => {
      const column = index % columns;
      const row = Math.floor(index / columns);
      const x = column * (width + gap);
      const y = row * (height + gap);

      const state = this.moduleStates[id];
      state.isMaximized = false;
      state.isMinimized = false;
      state.snapTarget = null;
      state.restoreBounds = null;
      this.updateBounds(id, { x, y, width, height });
    });

    this.focusModule(visibleModules[visibleModules.length - 1].id);
    this.schedulePersist();
  }

  getConstrainedBounds(id: ModuleId, patch: Partial<WindowBounds>): WindowBounds {
    const state = this.moduleStates[id];
    const next = this.normalizeBounds(
      {
        ...state,
        ...patch
      },
      16
    );

    return {
      x: next.x,
      y: next.y,
      width: next.width,
      height: next.height
    };
  }

  getSnapBounds(id: ModuleId, target: SnapTarget): WindowBounds {
    const state = this.moduleStates[id];
    const workspaceWidth = this.workspaceSize.width;
    const workspaceHeight = this.workspaceSize.height;

    if (!workspaceWidth || !workspaceHeight) {
      return {
        x: state.x,
        y: state.y,
        width: state.width,
        height: state.height
      };
    }

    const halfWidth = Math.floor(workspaceWidth / 2);
    const halfHeight = Math.floor(workspaceHeight / 2);
    let nextBounds: WindowBounds;

    switch (target) {
      case 'maximize':
        nextBounds = { x: 0, y: 0, width: workspaceWidth, height: workspaceHeight };
        break;
      case 'left':
        nextBounds = { x: 0, y: 0, width: halfWidth, height: workspaceHeight };
        break;
      case 'right':
        nextBounds = {
          x: workspaceWidth - halfWidth,
          y: 0,
          width: halfWidth,
          height: workspaceHeight
        };
        break;
      case 'top-left':
        nextBounds = { x: 0, y: 0, width: halfWidth, height: halfHeight };
        break;
      case 'top-right':
        nextBounds = {
          x: workspaceWidth - halfWidth,
          y: 0,
          width: halfWidth,
          height: halfHeight
        };
        break;
      case 'bottom-left':
        nextBounds = {
          x: 0,
          y: workspaceHeight - halfHeight,
          width: halfWidth,
          height: halfHeight
        };
        break;
      case 'bottom-right':
        nextBounds = {
          x: workspaceWidth - halfWidth,
          y: workspaceHeight - halfHeight,
          width: halfWidth,
          height: halfHeight
        };
        break;
    }

    const next = this.normalizeBounds(
      {
        ...state,
        minWidth: 0,
        minHeight: 0,
        ...nextBounds
      },
      0
    );

    return {
      x: next.x,
      y: next.y,
      width: next.width,
      height: next.height
    };
  }

  applySnapTarget(id: ModuleId, target: SnapTarget, restoreBounds?: WindowBounds, shouldFocus = true) {
    const state = this.moduleStates[id];
    const next = this.getSnapBounds(id, target);
    const nextRestoreBounds =
      restoreBounds ??
      state.restoreBounds ?? {
        x: state.x,
        y: state.y,
        width: state.width,
        height: state.height
      };

    if (target === 'maximize') {
      state.restoreBounds = nextRestoreBounds;
      state.isMaximized = true;
      state.snapTarget = 'maximize';
    } else {
      state.isMaximized = false;
      state.snapTarget = target;
      state.restoreBounds = nextRestoreBounds;
    }

    state.x = next.x;
    state.y = next.y;
    state.width = next.width;
    state.height = next.height;

    if (shouldFocus) {
      this.focusModule(id);
    }

    this.schedulePersist();
  }

  restoreFromMaximized(id: ModuleId): WindowBounds {
    return this.restoreFromDocked(id);
  }

  updateBounds(id: ModuleId, patch: Partial<ModuleWindowState>) {
    const state = this.moduleStates[id];

    if (state.isMaximized && this.hasBoundsPatch(patch)) {
      state.isMaximized = false;
      state.restoreBounds = null;
    }

    if (this.hasBoundsPatch(patch)) {
      state.snapTarget = null;
    }

    const next = this.normalizeBounds(
      {
        ...state,
        ...patch
      },
      16
    );

    state.x = next.x;
    state.y = next.y;
    state.width = next.width;
    state.height = next.height;
    this.schedulePersist();
  }

  async restoreWorkspaceSession() {
    const snapshot = await loadWorkspaceSession();

    if (!snapshot) {
      return null;
    }

    this.hydratingSession = true;

    try {
      this.theme = snapshot.theme === 'light' ? snapshot.theme : 'light';
      applyTheme(this.theme);
      this._zCounter = Math.max(20, snapshot.zCounter || 20);
      this._openOffset = Math.max(0, Math.min(5, snapshot.openOffset || 0));

      for (const { id } of this.definitions) {
        const persisted = snapshot.modules[id];
        if (!persisted) {
          continue;
        }

        const state = this.moduleStates[id];
        state.isOpen = Boolean(persisted.isOpen);
        state.isMinimized = Boolean(persisted.isMinimized);
        state.x = this.safeNumber(persisted.x, state.x);
        state.y = this.safeNumber(persisted.y, state.y);
        state.width = this.safeNumber(persisted.width, state.width);
        state.height = this.safeNumber(persisted.height, state.height);
        state.minWidth = this.safeNumber(persisted.minWidth, state.minWidth);
        state.minHeight = this.safeNumber(persisted.minHeight, state.minHeight);
        state.zIndex = this.safeNumber(persisted.zIndex, state.zIndex);
        state.hasOpened = Boolean(persisted.hasOpened);
        state.isMaximized = Boolean(persisted.isMaximized);
        state.snapTarget = persisted.snapTarget ?? null;
        state.restoreBounds = persisted.restoreBounds
          ? {
              x: this.safeNumber(persisted.restoreBounds.x, state.x),
              y: this.safeNumber(persisted.restoreBounds.y, state.y),
              width: this.safeNumber(persisted.restoreBounds.width, state.width),
              height: this.safeNumber(persisted.restoreBounds.height, state.height)
            }
          : null;
      }
    } finally {
      this.hydratingSession = false;
    }

    return snapshot.theme;
  }

  async clearWorkspaceSession() {
    await clearWorkspaceSession();
  }

  private applyMaximizedBounds(id: ModuleId) {
    const state = this.moduleStates[id];
    const workspaceWidth = this.workspaceSize.width;
    const workspaceHeight = this.workspaceSize.height;

    if (!workspaceWidth || !workspaceHeight) {
      return;
    }

    const next = this.normalizeBounds(
      {
        ...state,
        x: 0,
        y: 0,
        width: workspaceWidth,
        height: workspaceHeight
      },
      0
    );

    state.x = next.x;
    state.y = next.y;
    state.width = next.width;
    state.height = next.height;
  }

  private hasBoundsPatch(patch: Partial<ModuleWindowState>) {
    return patch.x !== undefined || patch.y !== undefined || patch.width !== undefined || patch.height !== undefined;
  }

  private safeNumber(value: unknown, fallback: number) {
    return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
  }

  private schedulePersist() {
    if (this.hydratingSession || typeof window === 'undefined') {
      return;
    }

    if (this.persistTimer) {
      window.clearTimeout(this.persistTimer);
    }

    this.persistTimer = window.setTimeout(() => {
      this.persistTimer = null;
      void this.persistWorkspaceSession();
    }, 180);
  }

  private async persistWorkspaceSession() {
    if (typeof window === 'undefined') {
      return;
    }

    const modules = {} as WorkspaceSessionSnapshot['modules'];

    for (const { id } of this.definitions) {
      const state = this.moduleStates[id];
      modules[id] = {
        id,
        isOpen: state.isOpen,
        isMinimized: state.isMinimized,
        x: state.x,
        y: state.y,
        width: state.width,
        height: state.height,
        minWidth: state.minWidth,
        minHeight: state.minHeight,
        zIndex: state.zIndex,
        hasOpened: state.hasOpened,
        isMaximized: state.isMaximized,
        snapTarget: state.snapTarget,
        restoreBounds: state.restoreBounds
      };
    }

    await saveWorkspaceSession({
      version: 1,
      theme: this.theme,
      zCounter: this._zCounter,
      openOffset: this._openOffset,
      modules
    });
  }

  private normalizeBounds(
    state: Pick<ModuleWindowState, 'x' | 'y' | 'width' | 'height' | 'minWidth' | 'minHeight'>,
    padding: number
  ) {
    const workspaceWidth = this.workspaceSize.width;
    const workspaceHeight = this.workspaceSize.height;

    if (!workspaceWidth || !workspaceHeight) {
      return state;
    }

    const paddedWidth = Math.max(1, workspaceWidth - padding * 2);
    const paddedHeight = Math.max(1, workspaceHeight - padding * 2);
    const minWidth = Math.min(state.minWidth, workspaceWidth);
    const minHeight = Math.min(state.minHeight, workspaceHeight);
    const maxWidth = Math.max(minWidth, paddedWidth);
    const maxHeight = Math.max(minHeight, paddedHeight);
    const width = Math.min(Math.max(state.width, minWidth), maxWidth);
    const height = Math.min(Math.max(state.height, minHeight), maxHeight);

    const maxX = Math.max(padding, workspaceWidth - width - padding);
    const maxY = Math.max(padding, workspaceHeight - height - padding);
    const x = Math.min(Math.max(state.x, padding), maxX);
    const y = Math.min(Math.max(state.y, padding), maxY);

    return {
      ...state,
      x,
      y,
      width,
      height
    };
  }
}

export const moduleManager = new ModuleManager();
