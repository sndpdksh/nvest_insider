import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: '/nvest-insider/',
  server: {
    port: 5173,
    strictPort: true,
  },
  build: {
    chunkSizeWarningLimit: 1000,
    rollupOptions: {
      output: {
        manualChunks: {
          'pdf-lib': ['pdfjs-dist'],
          'azure-msal': ['@azure/msal-browser', '@azure/msal-react'],
          'graph-client': ['@microsoft/microsoft-graph-client'],
          'react-vendor': ['react', 'react-dom'],
        },
      },
    },
  },
})
