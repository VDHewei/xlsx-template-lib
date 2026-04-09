#!/usr/bin/env node
import chalk from 'chalk';
import {Command} from 'commander';
import * as engine from './index';

const program = new Command();
program
    .name('xlsx-cli')
    .version('1.0.0');

program.command('compile')
    .argument('<string>', "xlsx file path")
    .option('-s,--save <string>', "save compiled xlsx file to user dir")
    .option('-r,--remove', 'remove configure rules sheet', false)
    .action((cmd: Command, xlsxFile: string, options: { [key: string]: any }) => {
        const resolver = engine.ExprResolver;
        console.log(chalk.green('xlsxFile %s'),xlsxFile);
        console.log(chalk.red('options %s'), options);
    });

program.command("render")
    .argument('<string>', "xlsx file path")
    .option('-c,--compile', "auto compile flag", false)
    .option('-s,--save <string>', "save render xlsx file to user dir")
    .action((cmd: Command, xlsxFile: string, options: { [key: string]: any }) => {
        console.log(chalk.green('xlsxFile %s'),xlsxFile);
        console.log(chalk.red('options %s'), options);
    });

program.parse(process.argv);