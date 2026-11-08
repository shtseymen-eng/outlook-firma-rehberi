import {defineConfig} from 'vite';
import {resolve} from 'node:path';
export default defineConfig({base:'/outlook-firma-rehberi/',build:{outDir:'dist',emptyOutDir:true,rollupOptions:{input:{main:resolve(import.meta.dirname,'index.html'),auth:resolve(import.meta.dirname,'auth.html')},output:{entryFileNames:'[name]-[hash].js',chunkFileNames:'chunk-[hash].js',assetFileNames:'[name]-[hash][extname]'}}},server:{host:'127.0.0.1',port:8767}});
