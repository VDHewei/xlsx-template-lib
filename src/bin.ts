#!/usr/bin/env node
import chalk from 'chalk';
import {Command} from 'commander';
import * as fs from 'node:fs/promises';
import {existsSync} from 'node:fs';
import * as path from 'node:path';
import {
    XlsxRender,
    BufferType,
    compileAll,
    compileRuleSheetName,
    AutoOptions,
} from './index';
import {
    generateOutputFilename,
    resolveFilePath,
    parseRenderData,
    checkSheetAndPlaceholders,
    addRuleToSheet,
    parseRulesFromFile,
    addMultipleRulesToSheet,
} from './bin-helpers';

async function main() {
    // Load package.json version
    let version = '1.0.0';
    try {
        // Try multiple possible paths for package.json
        const possiblePaths = [
            path.join(process.cwd(), 'package.json'),
            path.join(process.cwd(), '..', 'package.json'),
        ];

        for (const packagePath of possiblePaths) {
            if (existsSync(packagePath)) {
                const packageJson = JSON.parse(await fs.readFile(packagePath, 'utf-8'));
                version = packageJson.version;
                break;
            }
        }
    } catch (e) {
        // Fallback to default version
    }

    // Load .env if exists
    const envPath = path.join(process.cwd(), '.env');
    if (existsSync(envPath)) {
        try {
            const dotenv = (await import('dotenv')).default;
            dotenv.config({quiet:true,debug: false, path: envPath});
        } catch (e) {
            // dotenv is optional
        }
    }

    const program = new Command();
    program
        .name('xlsx-cli')
        .version(version);

// 编译存在规则的xlsx 输出 编译后带占位符 的 xlsx
//   为指定 --sheet-name , 默认取 第一个 sheet 名称
//   则自动寻找 export_metadata.config 配置文件，并进行编译替换到指定的sheet 文件中, 不存在则不进行编译
//    --remove 启用时,编译后的 xlsx 文件,移除 export_metadata.config 配置文件(可以调用ExprResolver.removeUnExportSheets)
//   以上都成功 则 输出 渲染后的 xlsx 文件 , 有指定 --save 参数 则输出到指定目录，否则输出到当前目录
program.command('compile')
    .argument('<string>', "xlsx file path")
    .option('-s,--save <string>', "save compiled xlsx file to user dir")
    .option('-n,--sheet-name <string>', "compile xlsx sheet name when xlsx has multiple sheets")
    .option('-r,--remove', 'remove configure rules sheet', false)
    .action(async (xlsxFile: string, options: { [key: string]: any }) => {
        try {
            console.log(chalk.green('📄 Compiling Excel file...'));

            // Resolve file path
            const filePath = await resolveFilePath(xlsxFile);
            console.log(chalk.gray(`Loading file: ${filePath}`));

            // Read file buffer
            const buffer = await fs.readFile(filePath);

            // Determine sheet name
            const xlsx = await XlsxRender.create(buffer);
            const sheets = xlsx.getSheets();
            const sheetName = options.sheetName || sheets[0].name;
            console.log(chalk.gray(`Target sheet: ${sheetName}`));

            // Check if rule sheet exists
            const ruleSheetName = options.sheetName || compileRuleSheetName;

            // Compile file
            console.log(chalk.gray('Compiling rules...'));
            const compiledBuffer = await compileAll(buffer, {
                sheetName: ruleSheetName,
                remove: options.remove || false,
            } as unknown as AutoOptions);
            console.log(chalk.green('✓ Compilation completed'));

            // Determine output path
            const outputFile = options.save || path.join(process.cwd(), generateOutputFilename(xlsxFile));
            console.log(chalk.gray(`Saving to: ${outputFile}`));

            // Write output file
            await fs.writeFile(outputFile, compiledBuffer);

            console.log(chalk.green('✓ Excel file compiled successfully!'));
            console.log(chalk.green(`📁 Output: ${outputFile}`));
        } catch (error) {
            console.error(chalk.red('✗ Compilation failed:'));
            console.error(chalk.red(error instanceof Error ? error.message : String(error)));
            process.exit(1);
        }
    });

// 渲染占位符 xlsx template 文件
// 为指定 --sheet-name , 默认取 第一个 sheet 名称
//  检查 xlsx 对应 sheet 是否存在，检查 对应sheet 是否存在 变量占位符
//  以上检查满足后 渲染 xlsx 文件，不满足 提示 无效的 xlsx 文件或者 sheet不存在
//  渲染的 --data 参数支持 本地文件 json 文件/ json 字符串 / 远程 json 文件
// --compile 参数开启， 则自动寻找 export_metadata.config 配置文件，并进行编译替换到指定的sheet 文件中, 不存在则不进行编译
// 无编译 逻辑，则直接 解析 data 进行 xlsx sheet 渲染
// 以上有任意异常，都输出异常提示 并终止 业务逻辑
// 以上都成功 则 输出 渲染后的 xlsx 文件 , 有指定 --save 参数 则输出到指定目录，否则输出到当前目录
program.command("render")
    .argument('<string>', "xlsx file path")
    .option('-c,--compile', "auto compile flag", false)
    .option('-n,--sheet-name <string>', "render xlsx sheet name when xlsx has multiple sheets")
    .option('-s,--save <string>', "save render xlsx file to user dir")
    .option('-d,--data <string>', "render xlsx file data from")
    .action(async (xlsxFile: string, options: { [key: string]: any }) => {
        try {
            console.log(chalk.green('📄 Rendering Excel template...'));

            // Resolve file path
            const filePath = await resolveFilePath(xlsxFile);
            console.log(chalk.gray(`Loading file: ${filePath}`));

            // Read file buffer
            let buffer = await fs.readFile(filePath);

            // Determine sheet name
            let xlsx = await XlsxRender.create(buffer);
            const sheets = xlsx.getSheets();
            const sheetName = options.sheetName || sheets[0].name;
            console.log(chalk.gray(`Target sheet: ${sheetName}`));

            // Check sheet exists and has placeholders
            checkSheetAndPlaceholders(xlsx, sheetName);
            console.log(chalk.gray('Sheet validation passed'));

            // Parse render data
            const renderData = await parseRenderData(options.data);
            if (Object.keys(renderData).length > 0) {
                console.log(chalk.gray(`Render data loaded with ${Object.keys(renderData).length} keys`));
            }

            // Compile if needed
            if (options.compile) {
                console.log(chalk.gray('Auto-compiling rules...'));
                const ruleSheetName = options.sheetName || compileRuleSheetName;

                const compiledResult = await compileAll(buffer, {
                    sheetName: ruleSheetName,
                    remove: false,  // Don't remove rule sheet during render
                } as unknown as AutoOptions);
                buffer = Buffer.from(compiledResult);
                xlsx = await XlsxRender.create(buffer);
                console.log(chalk.green('✓ Auto-compilation completed'));
            }

            // Render sheet
            console.log(chalk.gray('Rendering template...'));
            await xlsx.render(renderData, sheetName);

            // Generate output
            const outputBuffer = await xlsx.generate({
                type: BufferType.NodeBuffer,
                compression: "DEFLATE",
                compressionOptions: {
                    level: 9
                }
            });

            // Determine output path
            const outputFile = options.save || path.join(process.cwd(), generateOutputFilename(xlsxFile));
            console.log(chalk.gray(`Saving to: ${outputFile}`));

            // Write output file
            await fs.writeFile(outputFile, outputBuffer);

            console.log(chalk.green('✓ Excel template rendered successfully!'));
            console.log(chalk.green(`📁 Output: ${outputFile}`));
        } catch (error) {
            console.error(chalk.red('✗ Rendering failed:'));
            console.error(chalk.red(error instanceof Error ? error.message : String(error)));
            process.exit(1);
        }
    });

// 添加规则配置
// xlsx 不存存在 export_metadata.config sheet 则添加
// 检查 对应 配置 是否存在 , 如果不存在则添加
// 对应 类型 type 规则添加 超过 4列就添加一行，到新行配置记录
// 配置行 都是以 类型 type 值 开头的行 <cell,alias,rowCell,mergeCell>
// 样式要求： type 值为 加粗 居中，配置等式 上下 居中，cell 宽 自适应
program.command("rules")
    .argument('<string>', "xlsx compile rules setting")
    .option('-t,--type <string>', "xlsx compile rule type <cell,alias,rowCell,mergeCell> (optional when using -f)")
    .option('-r,--rule <string>', "xlsx compile rule expr (can be specified multiple times)")
    .option('-f,--file <string>', "read rules from file (format: <type> ruleExpr per line)")
    .option('-s,--save <string>', "save compiled xlsx file to user dir")
    .action(async (xlsxFile: string, options: { [key: string]: any }) => {
        try {
            console.log(chalk.green('📝 Adding rule configuration...'));

            const validTypes = ['cell', 'alias', 'rowCell', 'mergeCell'];
            let rules: { type: string; rule: string }[] = [];

            // Mode 1: Read from file
            if (options.file) {
                console.log(chalk.gray(`Reading rules from file: ${options.file}`));
                rules = await parseRulesFromFile(options.file);
                console.log(chalk.green(`✓ Loaded ${rules.length} rules from file`));
            }
            // Mode 2: Read from command line
            else if (options.rule) {
                // Normalize to array if single rule
                const ruleArray = Array.isArray(options.rule) ? options.rule : [options.rule];

                // Validate type if specified
                if (options.type) {
                    if (!validTypes.includes(options.type)) {
                        console.error(chalk.red(`Invalid rule type: ${options.type}. Must be one of: ${validTypes.join(', ')}`));
                        process.exit(1);
                    }
                    // Add all rules with same type
                    for (const rule of ruleArray) {
                        rules.push({ type: options.type, rule });
                    }
                } else {
                    console.error(chalk.red('Error: -t,--type is required when using -r,--rule'));
                    process.exit(1);
                }
                console.log(chalk.green(`✓ Loaded ${rules.length} rules from command line`));
            } else {
                console.error(chalk.red('Error: Either -f,--file or -r,--rule must be specified'));
                process.exit(1);
            }

            if (rules.length === 0) {
                console.error(chalk.red('Error: No rules to add'));
                process.exit(1);
            }

            // Resolve file path
            const filePath = await resolveFilePath(xlsxFile);
            console.log(chalk.gray(`Loading file: ${filePath}`));

            // Read file buffer
            const buffer = await fs.readFile(filePath);

            // Add all rules to export_metadata.config sheet
            const updatedBuffer = await addMultipleRulesToSheet(buffer, rules);

            // Determine output path
            const outputFile = options.save || path.join(process.cwd(), generateOutputFilename(xlsxFile));
            console.log(chalk.gray(`Saving to: ${outputFile}`));

            // Write output file
            await fs.writeFile(outputFile, updatedBuffer);

            console.log(chalk.green('✓ All rules added successfully!'));
            console.log(chalk.green(`📁 Output: ${outputFile}`));
        } catch (error) {
            console.error(chalk.red('✗ Failed to add rule configuration:'));
            console.error(chalk.red(error instanceof Error ? error.message : String(error)));
            process.exit(1);
        }
    });

    program.parse(process.argv);
}

main().catch(error => {
    console.error(chalk.red('✗ Fatal error:'));
    console.error(chalk.red(error instanceof Error ? error.message : String(error)));
    process.exit(1);
});
