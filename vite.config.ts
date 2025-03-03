import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    strictPort: true,
    https: {
      key: fs.readFileSync('.cert/key.pem'),
      cert: fs.readFileSync('.cert/cert.pem'),
    },
    headers: {
      'Access-Control-Allow-Origin': 'https://api.pexels.com'
    }
  },
  build: {
    outDir: 'dist',
    sourcemap: true
  }
});