import {defineConfig} from 'vitest/config'

export default defineConfig({
    test: {
        include: ['src/**/*.test.ts', 'src/*.test.ts', 'test/**/*.test.ts'],
        tags: [
            {
                name: "backend",
                description: "Tests written for backend.",
                timeout: 100000,
            },
        ],
    },
    resolve: {
        conditions: ['import']
    },
    define: {
        'import.meta.vitest': 'undefined',
    },
})