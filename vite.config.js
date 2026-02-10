import { defineConfig } from 'vite';

export default defineConfig({
    // Use relative asset paths so built files work in Electron via file://
    base: './'
});
