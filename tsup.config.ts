import { defineConfig } from 'tsup'

export default defineConfig({
    // 入口文件
    entry: ['src/index.ts'],
    // 同时输出 CommonJS 和 ES Module 格式
    format: ['cjs', 'esm'],
    // 生成类型声明文件
    dts: {
        resolve: true,
        // 可以在这里传递编译器选项吗？
        // rollup-plugin-dts 接受 compilerOptions
        compilerOptions: {
            ignoreDeprecations: "6.0"
        }
    },
    // 开启代码压缩
    minify: false, // 库代码建议不压缩，方便用户调试，如有需要可改为 true
    // 生成 sourcemap
    sourcemap: true,
    // 清理旧的 dist 文件
    clean: true,
    // 外部依赖不打入包内（非常重要！）
    external: ['@types/elementtree','image-size','jszip'],
})