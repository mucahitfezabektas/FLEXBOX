import { defineConfig } from 'vite';
import { svelte } from '@sveltejs/vite-plugin-svelte';

export default defineConfig({
  plugins: [svelte()],
  root: 'frontend',
  resolve: {
    alias: {
      $components: '/src/components',
      $lib: '/src/lib',
      $modules: '/src/modules'
    }
  },
  server: {
    host: '0.0.0.0',
    port: 1420,
    strictPort: true
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true
  }
});
