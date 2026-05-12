import UnoCSS from 'unocss/vite';
import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import { fileURLToPath, URL } from 'node:url';
import { copyFile } from 'wpsjs/vite_plugins';
import AutoImport from 'unplugin-auto-import/vite';
import Components from 'unplugin-vue-components/vite';
import { NaiveUiResolver } from 'unplugin-vue-components/resolvers';

export default defineConfig({
    base: './',
    plugins: [
        copyFile({
            src: 'manifest.xml',
            dest: 'manifest.xml'
        }),
        vue(),
        AutoImport({
            imports: ['vue']
        }),
        Components({
            resolvers: [NaiveUiResolver()]
        }),
        UnoCSS({ configFile: './uno.config.js' })
    ],
    resolve: {
        alias: {
            '@': fileURLToPath(new URL('./src', import.meta.url))
        }
    },
    server: {
        host: '0.0.0.0'
    }
});
