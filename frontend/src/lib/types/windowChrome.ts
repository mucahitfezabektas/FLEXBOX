export type WindowStatusTone = 'info' | 'success' | 'warning' | 'danger';

export interface WindowStatus {
  label: string;
  tone: WindowStatusTone;
}

export interface WindowToolbarAction {
  id: string;
  label: string;
  kind?: 'default' | 'primary' | 'danger';
}

export interface WindowFooterMeta {
  label: string;
  value: string;
}
