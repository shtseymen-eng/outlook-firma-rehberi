import {defineConfig} from 'vite';
import {resolve} from 'node:path';
export default defineConfig({base:process.env.LOCAL_HELPER==='1'?'/':'/outlook-firma-rehberi/',build:{outDir:'dist',emptyOutDir:true,rollupOptions:{input:{main:resolve(import.meta.dirname,'index.html'),auth:resolve(import.meta.dirname,'auth.html'),downloads:resolve(import.meta.dirname,'downloads.html')},output:{entryFileNames:'[name]-[hash].js',chunkFileNames:'chunk-[hash].js',assetFileNames:'[name]-[hash][extname]'}}},server:{host:'127.0.0.1',port:8767}});
