import { defineConfig } from 'vite'

// host:true expone el servidor en tu red local para abrirlo desde el móvil.
// base:'./' hace que el build funcione también abriendo index.html directamente.
export default defineConfig({
  base: './',
  server: { host: true, port: 5173 },
  preview: { host: true, port: 4173 },
})
