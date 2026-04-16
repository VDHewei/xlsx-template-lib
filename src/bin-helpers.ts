import * as path from 'node:path';
import {existsSync} from 'node:fs';
import * as fs from 'node:fs/promises';
import {XlsxRender} from './biz';
import {columnNumberToLetter, loadWorkbook, toCellValue} from './helper';
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
    const scriptDir = path.dirname(process.cwd());
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
                let fetch: ((input: string | URL | Request, init?: RequestInit) => Promise<Response>) | ((arg0: string) => any);
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
    workbook: exceljs.Workbook,
    ruleType: string,
    ruleExpr: string,
    sheetName: string = 'export_metadata.config'
): Promise<exceljs.Workbook> {
    let worksheet = workbook.getWorksheet(sheetName);

    // Create sheet if not exists
    if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
        console.log(chalk.gray(`Created new sheet: ${sheetName}`));
    }

    // Find the next available row for the given rule type
    let startRow = 1;
    let currentRow: number = NaN;


    // Scan existing rules to find where to add new rule
    if (worksheet.rowCount > 0) {
        let stop = false;
        worksheet.eachRow((row, rowNumber) => {
            let columnCount = 0;
            const cell = row.getCell(1);
            const cellValue = cell.value;
            // Skip empty rows
            if (!cellValue || stop) {
                return;
            }

            // Check if this is our rule type row
            const value: string = toCellValue(cellValue).trim();
            if (value.toLowerCase() === ruleType.toLowerCase()) {
                // console.log(`worksheet:  ${cellValue}`)
                // Check how many rules are in this row
                for (let col = 2; col < 5; col++) {
                    const ruleCell = row.getCell(col);
                    if (ruleCell.value) {
                        columnCount++;
                    } else {
                        break;
                    }
                }

                // If we have less than 4 rules, add to this row
                if (columnCount < 3) {
                    currentRow = rowNumber;
                    startRow = rowNumber;
                } else {
                    // Need to create a new row
                    currentRow = rowNumber + 1;
                }
                const cell = worksheet.getRow(currentRow).getCell(1);
                // need stop
                if (cell === undefined ||
                    cell.value === null ||
                    cell.value === undefined) {
                    stop = true;
                    return;
                }
                if (toCellValue(cell.value).toLowerCase() !== ruleType.toLowerCase()) {
                    worksheet.insertRow(currentRow, []);
                    currentRow = currentRow + 1;
                    stop = true;
                    return;
                }
                if (toCellValue(cell.value).toLowerCase() === ruleType.toLowerCase() && columnCount < 3) {
                    stop = true;
                    return;
                }
            }
        });
    }
    // empty sheets
    if (isNaN(currentRow)) {
        currentRow = worksheet.rowCount + 1;
        worksheet.insertRow(currentRow, [])
    }

    // If no existing rule type found, start from row 1
    if (startRow === 1 && currentRow === 1 && worksheet.rowCount === 0) {
        currentRow = 1;
    }

    // Ensure we have the target row
    let targetRow = worksheet.getRow(currentRow);
    let lastCell = targetRow.getCell(4);
    if (lastCell !== undefined && lastCell.value !== undefined && lastCell.value !== null && lastCell.value !== "") {
        currentRow = worksheet.rowCount + 1;
        targetRow = worksheet.insertRow(currentRow, [])
    }
    if (!targetRow.getCell(1).value) {
        // New row - set the type header
        const typeCell = targetRow.getCell(1);
        typeCell.value = ruleType;
        typeCell.font = {bold: true};
        typeCell.alignment = {horizontal: 'center', vertical: 'middle'};
    }

    // Determine column to add rule (skip column 1 which contains type)
    let ruleCol = 2;
    while (ruleCol <= 4) {
        const existingCell = targetRow.getCell(ruleCol);
        if (existingCell === undefined ||
            (existingCell.value === undefined || existingCell.value === null || existingCell.value === "")) {
            break;
        }
        ruleCol++;
    }
    const ruleCell = targetRow.getCell(ruleCol);
    ruleCell.value = ruleExpr;
    ruleCell.alignment = {vertical: 'middle', horizontal: 'center'};

    // Auto-fit column width
    const column = worksheet.getColumn(ruleCol);
    column.width = Math.max(column.width || 10, ruleExpr.length + 2);
    if(currentRow === 1 && ruleCol === 2){
        const column = worksheet.getColumn(1);
        column.width = Math.max(column.width || 10, ruleExpr.length + 2);
    }
    // console.log(chalk.gray(`Added rule ${ruleType} at row ${currentRow}, column ${ruleCol}`));
    return workbook;
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

        let items: string[] = [];
        const type: string = line.substring(0, spaceIndex).trim();
        const values: string = line.substring(spaceIndex + 1).trim();
        if (values.indexOf(' ') >= 0) {
            items = values.split(' ');
        } else if (values.indexOf('\t') >= 0) {
            items = values.split('\t');
        } else {
            items = [values];
        }
        items = items.filter(item => item.trim() !== "" && item.indexOf("=") > 0)
        if (!type) {
            console.log(chalk.yellow(`⚠ Line ${i + 1}: Invalid format. Expected "<type> ruleExpr"`));
            continue;
        }
        if (!items || items.length === 0) {
            continue
        }

        // Validate rule type
        if (!validTypes.includes(type)) {
            console.log(chalk.yellow(`⚠ Line ${i + 1}: Invalid rule type "${type}". Must be one of: ${validTypes.join(', ')}`));
            continue;
        }
        for (const rule of items) {
            let str = rule.trim();
            if (str === undefined || str === null
                || str === "" || str === '\t') {
                continue;
            }
            let key = `${type}:${str}`;
            if (validRulesSet.has(key)) {
                console.log(chalk.yellow(`⚠ Line ${i + 1}: Duplicate rule "${str}"`));
                continue;
            }
            validRulesSet.add(key);
            rules.push({type: type, rule: str});
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
    let workbook = await loadWorkbook(xlsxBuffer);

    for (const {type, rule} of rules) {
        workbook = await addRuleToSheet(workbook, type, rule);
    }
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
}
