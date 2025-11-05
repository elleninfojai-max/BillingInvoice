import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import path from 'path'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    open: true,
    fs: {
      // Allow serving files from one level up to the project root
      allow: ['..']
    }
  },
  optimizeDeps: {
    include: ['mammoth'],
    exclude: ['pdfjs-dist'], // Exclude pdfjs-dist from optimization
    esbuildOptions: {
      resolveExtensions: ['.js', '.jsx', '.ts', '.tsx', '.mjs']
    }
  },
  resolve: {
    dedupe: ['pdfjs-dist', 'mammoth']
  },
  build: {
    commonjsOptions: {
      include: [/pdfjs-dist/, /mammoth/],
      transformMixedEsModules: true
    },
    rollupOptions: {
      external: ['pdfjs-dist'] // Mark as external to skip bundling
    }
  },
  ssr: {
    noExternal: ['pdfjs-dist', 'mammoth'] // Include in SSR build
  }
})

