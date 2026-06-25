<script lang="ts">
  import { onMount } from 'svelte';
  import { invoke } from '@tauri-apps/api/core';
  import MainScreen from '$components/MainScreen.svelte';
  import { moduleManager } from '$lib/state/moduleManager.svelte';
  import type { ThemeMode } from '$lib/utils/theme';

  const DEFAULT_THEME: ThemeMode = 'light';

  onMount(async () => {
    const restoredTheme = await moduleManager.restoreWorkspaceSession();
    moduleManager.setTheme(restoredTheme ?? DEFAULT_THEME);

    try {
      const savedTheme = await invoke<ThemeMode | null>('load_theme_preference');
      if (savedTheme) {
        moduleManager.setTheme(savedTheme);
      }
    } catch {
      // Browser-only dev mode can run without a Tauri shell.
    }
  });
</script>

<MainScreen />
