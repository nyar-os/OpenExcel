import { defineConfig } from '@vmz/vmz';

/**
 * Studio shell served by the WAE WebView host.
 * Profile `static` → assembly `web-static`; the host loads `dist`.
 */
export default defineConfig({
    delivery: {
        default: 'static',
        profiles: {
            static: {
                host: 'browser',
                assembly: 'web-static',
            },
        },
    },
});
