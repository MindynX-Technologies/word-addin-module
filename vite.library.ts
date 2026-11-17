import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    lib: { entry: 'src/library.ts', formats: ['es'], fileName: 'index', cssFileName: 'style' },
    rollupOptions: {
      external: [/^react($|\/)/, /^react-dom($|\/)/, /^@zip\.js\//, 'lucide-react'],
    },
  },
});
