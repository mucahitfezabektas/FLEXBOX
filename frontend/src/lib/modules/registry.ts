import type { Component } from 'svelte';
import { moduleCatalogById } from '$lib/modules/catalog';
import type { ModuleId } from '$lib/modules/types';
import type { WindowFooterMeta, WindowStatus, WindowToolbarAction } from '$lib/types/windowChrome';

export interface ModuleRegistryEntry {
  definition: (typeof moduleCatalogById)[ModuleId];
  component: Component;
  status: WindowStatus | null;
  toolbarActions: WindowToolbarAction[];
  footerMeta: WindowFooterMeta[];
}

export const moduleRegistry: Record<ModuleId, ModuleRegistryEntry> = Object.fromEntries(
  Object.entries(moduleCatalogById).map(([id, manifest]) => [
    id,
    {
      definition: manifest,
      component: manifest.component,
      status: manifest.status,
      toolbarActions: manifest.toolbarActions,
      footerMeta: manifest.footerMeta
    }
  ])
) as Record<ModuleId, ModuleRegistryEntry>;
