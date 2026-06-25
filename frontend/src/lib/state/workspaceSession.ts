import { invoke } from '@tauri-apps/api/core';
import type { ModuleId, WindowBounds } from '$lib/modules/types';
import type { ThemeMode } from '$lib/utils/theme';

export type PersistedSnapTarget =
  | 'maximize'
  | 'left'
  | 'right'
  | 'top-left'
  | 'top-right'
  | 'bottom-left'
  | 'bottom-right';

export interface PersistedModuleWindowState extends WindowBounds {
  id: ModuleId;
  isOpen: boolean;
  isMinimized: boolean;
  minWidth: number;
  minHeight: number;
  zIndex: number;
  hasOpened: boolean;
  isMaximized: boolean;
  snapTarget: PersistedSnapTarget | null;
  restoreBounds: WindowBounds | null;
}

export interface WorkspaceSessionSnapshot {
  version: 1;
  theme: ThemeMode;
  zCounter: number;
  openOffset: number;
  modules: Record<ModuleId, PersistedModuleWindowState>;
}

const STORAGE_KEY = 'uniframe.workspace.session.v1';

function loadWorkspaceSessionFromBrowser() {
  if (typeof localStorage === 'undefined') {
    return null;
  }

  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return null;
    }

    const parsed = JSON.parse(raw) as WorkspaceSessionSnapshot;
    if (parsed?.version !== 1 || !parsed.modules) {
      return null;
    }

    return parsed;
  } catch {
    return null;
  }
}

function saveWorkspaceSessionToBrowser(snapshot: WorkspaceSessionSnapshot) {
  if (typeof localStorage === 'undefined') {
    return;
  }

  localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
}

function clearWorkspaceSessionFromBrowser() {
  if (typeof localStorage === 'undefined') {
    return;
  }

  localStorage.removeItem(STORAGE_KEY);
}

export async function loadWorkspaceSession() {
  try {
    const raw = await invoke<string | null>('load_workspace_session');
    if (!raw) {
      return null;
    }

    const parsed = JSON.parse(raw) as WorkspaceSessionSnapshot;
    if (parsed?.version !== 1 || !parsed.modules) {
      return null;
    }

    return parsed;
  } catch {
    return loadWorkspaceSessionFromBrowser();
  }
}

export async function saveWorkspaceSession(snapshot: WorkspaceSessionSnapshot) {
  try {
    await invoke('save_workspace_session', {
      sessionJson: JSON.stringify(snapshot)
    });
    saveWorkspaceSessionToBrowser(snapshot);
  } catch {
    saveWorkspaceSessionToBrowser(snapshot);
  }
}

export async function clearWorkspaceSession() {
  try {
    await invoke('clear_workspace_session');
  } catch {}

  clearWorkspaceSessionFromBrowser();
}
