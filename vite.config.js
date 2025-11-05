import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import path from 'path'

export default defineConfig({
  plugins: [react({
    jsxRuntime: 'automatic'
  })],
  server: {
    port: 3000,
    open: true,
    fs: {
      // Allow serving files from one level up to the project root
      allow: ['..']
    }
  },
  optimizeDeps: {
    include: ['mammoth', 'framer-motion', 'react', 'react-dom'],
    exclude: ['pdfjs-dist'], // Exclude pdfjs-dist from optimization
    esbuildOptions: {
      resolveExtensions: ['.js', '.jsx', '.ts', '.tsx', '.mjs']
    }
  },
  resolve: {
    dedupe: ['react', 'react-dom', 'pdfjs-dist', 'mammoth'],
    alias: {
      'react': path.resolve(__dirname, './node_modules/react'),
      'react-dom': path.resolve(__dirname, './node_modules/react-dom')
    }
  },
  build: {
    commonjsOptions: {
      include: [/pdfjs-dist/, /mammoth/, /node_modules/],
      transformMixedEsModules: true
    },
    rollupOptions: {
      output: {
        manualChunks: {
          'react-vendor': ['react', 'react-dom'],
          'motion': ['framer-motion']
        }
      }
    }
  }
})

