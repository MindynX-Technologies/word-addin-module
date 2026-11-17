import { defineConfig } from 'vite';

export default defineConfig({
  server: { proxy: { '/word-converter': {
    target: 'http://127.0.0.1:4175',
    rewrite: path => path.replace(/^\/word-converter/, ''),
  } } },
});
