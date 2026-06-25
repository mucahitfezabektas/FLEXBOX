export type ThemeMode = 'light';

export function applyTheme(theme: ThemeMode) {
  if (typeof document === 'undefined') {
    return;
  }

  document.documentElement.classList.remove('dark');
  document.documentElement.dataset.theme = theme;
}
