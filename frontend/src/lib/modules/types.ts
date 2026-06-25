export type ModuleId = 'auth' | 'customize' | 'spreadsheet' | 'empty';
export type ModuleGroup = 'Security' | 'Configuration' | 'Data' | 'Framework';

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
  icon: string;
  keywords: string[];
  defaultBounds: WindowBounds & {
    minWidth: number;
    minHeight: number;
  };
}
