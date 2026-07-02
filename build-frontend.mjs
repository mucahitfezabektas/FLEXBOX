import { fileURLToPath, URL } from 'node:url';
import { build } from 'vite';
import { svelte } from '@sveltejs/vite-plugin-svelte';

const frontendRoot = fileURLToPath(new URL('./frontend', import.meta.url));
const frontendDist = fileURLToPath(new URL('./frontend/dist', import.meta.url));

await build({
  configFile: false,
  root: frontendRoot,
  plugins: [svelte()],
  resolve: {
    alias: {
      $components: fileURLToPath(new URL('./frontend/src/components', import.meta.url)),
      $lib: fileURLToPath(new URL('./frontend/src/lib', import.meta.url)),
      $modules: fileURLToPath(new URL('./frontend/src/modules', import.meta.url))
    }
  },
  server: {
    host: '0.0.0.0',
    port: 1420,
    strictPort: true
  },
  build: {
    outDir: frontendDist,
    emptyOutDir: true
  }
});
