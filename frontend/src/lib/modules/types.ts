export type ModuleId = 'auth' | 'customize' | 'spreadsheet' | 'system-shutdown' | 'empty';
export type ModuleGroup = 'Security' | 'Configuration' | 'Data' | 'System' | 'Framework';
export type ModuleTheme = 'default' | 'sheets';

export interface WindowBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ModuleDefinition {
  id: ModuleId;
  title: string;
  shortTitle: string;
  subtitle: string;
  description: string;
  group: ModuleGroup;
  theme?: ModuleTheme;
  icon: string;
  keywords: string[];
  defaultBounds: WindowBounds & {
    minWidth: number;
    minHeight: number;
  };
}
