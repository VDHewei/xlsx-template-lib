#!/usr/bin/env node
import chalk from 'chalk';
import {Command} from 'commander';
import * as engine from './index';

// 全局规则: 1. 异常/成功 提示 都使用 chalk 输出
//            2. Command version 最好 使用 package.json 中的 版本号
//            3. 默认文件输出名 都使用 文件名 + _毫秒时间戳.xlsx
//            4. 任务执行前可以动态 加载 .env
//            5. 必须 tsc 编译无错误无 warn 业代码才算 完成
//            6. 功能必须 windows, linux ,macos 平台 都支持
//            7. 必须 使用 单元测试 覆盖 80% 以上 代码
//            8. 必须在 readme 中 描述 如何 使用 命令行工具
const program = new Command();
program
    .name('xlsx-cli')
    .version('1.0.0');

// 编译存在规则的xlsx 输出 编译后带占位符 的 xlsx
program.command('compile')
    .argument('<string>', "xlsx file path")
    .option('-s,--save <string>', "save compiled xlsx file to user dir")
    .option('-n,--sheet-name <string>', "compile xlsx sheet name when xlsx has multiple sheets ")
    .option('-r,--remove', 'remove configure rules sheet', false)
    .action((cmd: Command, xlsxFile: string, options: { [key: string]: any }) => {
       // @TODO
        // 为指定 --sheet-name , 默认取 第一个 sheet 名称
       //  则自动寻找 export_metadata.config 配置文件，并进行编译替换到指定的sheet 文件中, 不存在则不进行编译
        //  --remove 启用时,编译后的 xlsx 文件,移除 export_metadata.config 配置文件(可以调用ExprResolver.removeUnExportSheets)
        // 以上都成功 则 输出 渲染后的 xlsx 文件 , 有指定 --save 参数 则输出到指定目录，否则输出到当前目录
    });

// 渲染占位符 xlsx template 文件
program.command("render")
    .argument('<string>', "xlsx file path")
    .option('-c,--compile', "auto compile flag", false)
    .option('-n,--sheet-name <string>', "render xlsx sheet name when xlsx has multiple sheets ")
    .option('-s,--save <string>', "save render xlsx file to user dir")
    .option('-d,--data <string>', "render xlsx file data from")
    .action((cmd: Command, xlsxFile: string, options: { [key: string]: any }) => {
       // @TODO
        // 为指定 --sheet-name , 默认取 第一个 sheet 名称
        //  检查 xlsx 对应 sheet 是否存在，检查 对应sheet 是否存在 变量占位符
        //  以上检查满足后 渲染 xlsx 文件，不满足 提示 无效的 xlsx 文件或者 sheet不存在
        //  渲染的 --data 参数支持 本地文件 json 文件/ json 字符串 / 远程 json 文件
        // --compile 参数开启， 则自动寻找 export_metadata.config 配置文件，并进行编译替换到指定的sheet 文件中, 不存在则不进行编译
        // 无编译 逻辑，则直接 解析 data 进行 xlsx sheet 渲染
        // 以上有任意异常，都输出异常提示 并终止 业务逻辑
        // 以上都成功 则 输出 渲染后的 xlsx 文件 , 有指定 --save 参数 则输出到指定目录，否则输出到当前目录
    });

// 渲染占位符 xlsx template 文件
program.command("rules")
    .argument('<string>', "xlsx compile rules setting")
    .requiredOption('-t,--type <string>', "xlsx compile rule type <cell,alias,rowCell,mergeCell>")
    .requiredOption('-r,--rule <string>', "xlsx compile rule expr")
    .action((cmd: Command, xlsxFile: string, options: { [key: string]: any }) => {
        // @TODO
        // xlsx 不存存在 export_metadata.config sheet 则添加
        // 检查 对应 配置 是否存在 , 如果不存在则添加
        // 对应 类型 type 规则添加 超过 4列就添加一行，到新行配置记录
        // 配置行 都是以 类型 type 值 开头的行 <cell,alias,rowCell,mergeCell>
        // 样式要求： type 值为 加粗 居中，配置等式 上下 居中，cell 宽 自适应
    });


program.parse(process.argv);