import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: '/iso27001-skos-taxonomy/',
  build: {
    outDir: '../docs',
    emptyOutDir: true,
  },
})
