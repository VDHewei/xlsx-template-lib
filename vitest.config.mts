import {defineConfig} from 'vitest/config'
import dotenv from 'dotenv';

dotenv.config({debug: false});

export default defineConfig({
    test: {
        //setupFiles: ['./vitest.setup.ts'],
        unstubEnvs: true,
        include: ['src/**/*.test.ts', 'src/*.test.ts', 'test/**/*.test.ts'],
        tags: [
            {
                name: "backend",
                description: "Tests written for backend.",
                timeout: 100000,
                retry: 3,
            },
            {
                name: "xlsx",
                description: "xlsx test group",
                timeout: 500000,
                retry: 3,
            },
            {
                name: "compile",
                description: "compile xlsx test group",
                timeout: 500000,
            }
        ],
    },
    resolve: {
        conditions: ['import']
    },
    define: {
        'import.meta.vitest': 'undefined',
        'import.meta.env': 'process.env',
    },
})