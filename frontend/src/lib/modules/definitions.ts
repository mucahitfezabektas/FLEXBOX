import { moduleCatalogById } from '$lib/modules/catalog';
import type { ModuleDefinition, ModuleId } from '$lib/modules/types';

export const moduleDefinitions: Record<ModuleId, ModuleDefinition> = Object.fromEntries(
  Object.entries(moduleCatalogById).map(([id, manifest]) => {
    const { component: _component, status: _status, toolbarActions: _toolbarActions, footerMeta: _footerMeta, ...definition } =
      manifest;
    return [id, definition];
  })
) as Record<ModuleId, ModuleDefinition>;
