import type { Component } from 'svelte';
import type { ModuleDefinition, ModuleId } from '$lib/modules/types';
import type { WindowFooterMeta, WindowStatus, WindowToolbarAction } from '$lib/types/windowChrome';
import AuthModule from '$modules/AuthModule.svelte';
import CustomizeModule from '$modules/CustomizeModule.svelte';
import SpreadsheetModule from '$modules/SpreadsheetModule.svelte';
import EmptyModule from '$modules/EmptyModule.svelte';

export interface ModuleManifest extends ModuleDefinition {
  component: Component;
  status: WindowStatus | null;
  toolbarActions: WindowToolbarAction[];
  footerMeta: WindowFooterMeta[];
}

export const moduleCatalog = [
  {
    id: 'auth',
    title: 'Authentication Console',
    shortTitle: 'Auth',
    subtitle: 'Identity and access verification',
    description: 'FastAPI sidecar uzerinden yerel kimlik dogrulama ve servis erisimi.',
    group: 'Security',
    icon: 'AU',
    keywords: ['login', 'security', 'access', 'identity'],
    defaultBounds: {
      x: 40,
      y: 36,
      width: 860,
      height: 620,
      minWidth: 720,
      minHeight: 480
    },
    component: AuthModule,
    status: { label: 'SECURE', tone: 'success' },
    toolbarActions: [
      { id: 'center-window', label: 'Center' },
      { id: 'reset-window', label: 'Reset' }
    ],
    footerMeta: [
      { label: 'Area', value: 'IAM' },
      { label: 'Mode', value: 'LOCAL' }
    ]
  },
  {
    id: 'customize',
    title: 'Workspace Settings',
    shortTitle: 'Settings',
    subtitle: 'Policy and appearance controls',
    description: 'Tema, pencere duzeni ve framework davranislari icin ayar paneli.',
    group: 'Configuration',
    icon: 'WS',
    keywords: ['settings', 'theme', 'workspace', 'policy'],
    defaultBounds: {
      x: 110,
      y: 84,
      width: 760,
      height: 560,
      minWidth: 640,
      minHeight: 440
    },
    component: CustomizeModule,
    status: { label: 'POLICY', tone: 'info' },
    toolbarActions: [
      { id: 'center-window', label: 'Center' },
      { id: 'reset-window', label: 'Reset' }
    ],
    footerMeta: [
      { label: 'Area', value: 'UX' },
      { label: 'Mode', value: 'LOCAL' }
    ]
  },
  {
    id: 'spreadsheet',
    title: 'Spreadsheet Workbench',
    shortTitle: 'Sheets',
    subtitle: 'High-density Excel virtualization',
    description: 'Rust tarafinda tutulan dev Excel tablolarini sanal viewport ile kaydirsiz gösterim.',
    group: 'Data',
    theme: 'sheets',
    icon: 'SS',
    keywords: ['excel', 'spreadsheet', 'xlsx', 'virtualization', 'data-grid'],
    defaultBounds: {
      x: 148,
      y: 104,
      width: 1180,
      height: 760,
      minWidth: 920,
      minHeight: 560
    },
    component: SpreadsheetModule,
    status: { label: 'XL', tone: 'info' },
    toolbarActions: [
      { id: 'center-window', label: 'Center' },
      { id: 'reset-window', label: 'Reset' }
    ],
    footerMeta: [
      { label: 'Area', value: 'DATA' },
      { label: 'Mode', value: 'LOCAL' }
    ]
  },
  {
    id: 'empty',
    title: 'Framework Blueprint',
    shortTitle: 'Blueprint',
    subtitle: 'Subsystem extension template',
    description: 'Yeni moduller, entegrasyonlar ve ortak pencere sozlesmeleri icin referans alan.',
    group: 'Framework',
    icon: 'FW',
    keywords: ['template', 'framework', 'extension', 'module'],
    defaultBounds: {
      x: 180,
      y: 120,
      width: 940,
      height: 600,
      minWidth: 760,
      minHeight: 460
    },
    component: EmptyModule,
    status: { label: 'SDK', tone: 'warning' },
    toolbarActions: [
      { id: 'center-window', label: 'Center' },
      { id: 'reset-window', label: 'Reset' }
    ],
    footerMeta: [
      { label: 'Area', value: 'FRAMEWORK' },
      { label: 'Mode', value: 'TEMPLATE' }
    ]
  }
] as const satisfies readonly ModuleManifest[];

export const moduleCatalogById: Record<ModuleId, ModuleManifest> = moduleCatalog.reduce(
  (registry, entry) => {
    registry[entry.id] = entry;
    return registry;
  },
  {} as Record<ModuleId, ModuleManifest>
);
