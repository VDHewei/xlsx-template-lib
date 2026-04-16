import * as path from 'node:path';
import * as url from 'node:url';
import {existsSync} from 'node:fs';
import * as fs from 'node:fs/promises';
import {XlsxRender} from './biz';
import {loadWorkbook, parseWorkSheetRules, RuleToken} from './helper';
import exceljs from 'exceljs';
import chalk from 'chalk';

/**
 * Helper function to generate output filename
 * Format: <basename>_<timestamp>.xlsx
 */
export function generateOutputFilename(inputFile: string): string {
    const basename = path.basename(inputFile, path.extname(inputFile));
    const timestamp = Date.now();
    return `${basename}_${timestamp}.xlsx`;
}

/**
 * Helper function to resolve file path
 * Tries absolute/relative to cwd, then script directory
 */
export async function resolveFilePath(filePath: string): Promise<string> {
    // Try as absolute or relative to cwd
    const resolvedPath = path.resolve(filePath);
    if (existsSync(resolvedPath)) {
        return resolvedPath;
    }
    // Try relative to script directory
    const scriptDir = path.dirname(url.fileURLToPath(import.meta.url));
    const relativePath = path.resolve(scriptDir, filePath);
    if (existsSync(relativePath)) {
        return relativePath;
    }
    throw new Error(`File not found: ${filePath}`);
}

/**
 * Helper function to parse render data
 * Supports: JSON string, file path, or URL
 */
export async function parseRenderData(dataOption: string | undefined): Promise<Record<string, any>> {
    if (!dataOption) {
        return {};
    }

    try {
        // Try as JSON string first
        return JSON.parse(dataOption);
    } catch (e) {
        // Try as file path
        try {
            const filePath = await resolveFilePath(dataOption);
            const fileContent = await fs.readFile(filePath, 'utf-8');
            return JSON.parse(fileContent);
        } catch (e2) {
            // Try as remote URL
            if (dataOption.startsWith('http://') || dataOption.startsWith('https://')) {
                let fetch;
                // Try to import fetch from native fetch (Node.js 18+) or node-fetch
                try {
                    fetch = globalThis.fetch;
                } catch (e3) {
                    try {
                        // @ts-expect-error - node-fetch is optional dependency
                        const nodeFetch = await import('node-fetch');
                        fetch = nodeFetch.default;
                    } catch (e4) {
                        throw new Error('Remote URLs require Node.js 18+ or node-fetch package');
                    }
                }
                const response = await fetch(dataOption);
                return response.json();
            }
            throw new Error(`Failed to parse render data from: ${dataOption}`);
        }
    }
}

/**
 * Helper function to check sheet exists
 * Throws error if sheet not found
 */
export function checkSheetAndPlaceholders(xlsx: XlsxRender, sheetName: string): void {
    const sheets = xlsx.getSheets();
    const sheet = sheets.find(s => s.name === sheetName);

    if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found in Excel file`);
    }
}

/**
 * Helper function to add rule to export_metadata.config sheet
 * Creates sheet if not exists, adds rule with proper styling
 */
export async function addRuleToSheet(
    xlsxBuffer: Buffer,
    ruleType: string,
    ruleExpr: string,
    sheetName: string = 'export_metadata.config'
): Promise<Buffer> {
    const workbook = await loadWorkbook(xlsxBuffer);
    let worksheet = workbook.getWorksheet(sheetName);

    // Create sheet if not exists
    if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
        console.log(chalk.gray(`Created new sheet: ${sheetName}`));
    }

    // Find the next available row for the given rule type
    let startRow = 1;
    let currentRow = 1;
    let columnCount = 0;

    // Scan existing rules to find where to add new rule
    worksheet.eachRow((row, rowNumber) => {
        const cell = row.getCell(1);
        const cellValue = cell.value?.toString().trim();

        // Skip empty rows
        if (!cellValue) {
            return;
        }

        // Check if this is our rule type row
        if (cellValue.toLowerCase() === ruleType.toLowerCase()) {
            // Check how many rules are in this row
            let col = 2;
            while (col <= 4) {
                const ruleCell = row.getCell(col);
                if (ruleCell.value) {
                    columnCount++;
                }
                col++;
            }

            // If we have less than 4 rules, add to this row
            if (columnCount < 4) {
                currentRow = rowNumber;
                startRow = rowNumber;
            } else {
                // Need to create a new row
                currentRow = rowNumber + 1;
            }
        }
    });

    // If no existing rule type found, start from row 1
    if (startRow === 1 && currentRow === 1 && worksheet.rowCount === 0) {
        currentRow = 1;
    }

    // Ensure we have the target row
    let targetRow = worksheet.getRow(currentRow);
    if (!targetRow.getCell(1).value) {
        // New row - set the type header
        const typeCell = targetRow.getCell(1);
        typeCell.value = ruleType;
        typeCell.font = { bold: true };
        typeCell.alignment = { horizontal: 'center', vertical: 'middle' };
    }

    // Determine column to add rule (skip column 1 which contains type)
    let ruleCol = 2;
    while (ruleCol <= 4) {
        const existingCell = targetRow.getCell(ruleCol);
        if (!existingCell.value) {
            break;
        }
        ruleCol++;
    }

    // Add rule expression
    if (ruleCol <= 4) {
        const ruleCell = targetRow.getCell(ruleCol);
        ruleCell.value = ruleExpr;
        ruleCell.alignment = { vertical: 'middle' };

        // Auto-fit column width
        const column = worksheet.getColumn(ruleCol);
        column.width = Math.max(column.width || 10, ruleExpr.length + 2);

        console.log(chalk.gray(`Added rule ${ruleType} at row ${currentRow}, column ${ruleCol}`));
    } else {
        throw new Error(`Cannot add more than 4 rules for type: ${ruleType}`);
    }

    // Write buffer
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
}

/**
 * Helper function to parse rules from file
 * File format: <type> ruleExpr
 * Lines starting with # are treated as comments
 * @param filePath - Path to rules file
 * @returns Array of {type, rule} objects
 */
export async function parseRulesFromFile(filePath: string): Promise<{ type: string; rule: string }[]> {
    const resolvedPath = await resolveFilePath(filePath);
    const fileContent = await fs.readFile(resolvedPath, 'utf-8');
    const lines = fileContent.split('\n');
    const rules: { type: string; rule: string }[] = [];
    const validTypes = ['cell', 'alias', 'rowCell', 'mergeCell'];
    const validRulesSet = new Set<string>();

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // Skip empty lines and comments
        if (!line || line.startsWith('#')) {
            continue;
        }

        // Parse line: <type> ruleExpr
        const spaceIndex = line.indexOf(' ');
        if (spaceIndex === -1) {
            console.log(chalk.yellow(`⚠ Line ${i + 1}: Invalid format. Expected "<type> ruleExpr"`));
            continue;
        }

        const type = line.substring(0, spaceIndex).trim();
        const items = line.substring(spaceIndex + 1).trim().split(' ');

        if (!type) {
            console.log(chalk.yellow(`⚠ Line ${i + 1}: Invalid format. Expected "<type> ruleExpr"`));
            continue;
        }
        if(!items || items.length === 0){
            continue
        }

        // Validate rule type
        if (!validTypes.includes(type)) {
            console.log(chalk.yellow(`⚠ Line ${i + 1}: Invalid rule type "${type}". Must be one of: ${validTypes.join(', ')}`));
            continue;
        }
        for(const rule of items) {
            let str = rule.trim();
            let key = `${type}:${str}`;
            if(validRulesSet.has(key)){
                console.log(chalk.yellow(`⚠ Line ${i + 1}: Duplicate rule "${str}"`));
                continue;
            }
            validRulesSet.add(key);
            rules.push({ type:type, rule:str});
        }
    }

    if (rules.length === 0) {
        throw new Error('No valid rules found in file');
    }

    return rules;
}

/**
 * Helper function to add multiple rules to sheet
 * @param xlsxBuffer - Excel file buffer
 * @param rules - Array of {type, rule} objects
 * @returns Updated Excel buffer
 */
export async function addMultipleRulesToSheet(
    xlsxBuffer: Buffer,
    rules: { type: string; rule: string }[]
): Promise<Buffer> {
    let buffer = xlsxBuffer;

    for (const { type, rule } of rules) {
        console.log(chalk.gray(`Adding ${type} rule: ${rule}`));
        buffer = await addRuleToSheet(buffer, type, rule);
    }

    return buffer;
}
