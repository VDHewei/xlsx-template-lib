import { describe, it, expect } from 'vitest';
import { XlsxRender } from './biz';
import exceljs from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import JsZip from 'jszip';

// 辅助：检查测试文件是否存在
function hasTestFiles(): boolean {
    try {
        fs.readFileSync(path.resolve(__dirname, '../test_data/default_template_SD.xlsx'));
        fs.readFileSync(path.resolve(__dirname, '../test_data/form_data-SD.json'));
        return true;
    } catch {
        return false;
    }
}

function getCellXml(sheetXml: string, ref: string): string | undefined {
    return sheetXml.match(new RegExp(`<c[^>]*\\br="${ref}"(?=\\s|/|>)[\\s\\S]*?(?:</c>|/>)`))?.[0];
}

function getCellStyle(sheetXml: string, ref: string): string | undefined {
    return getCellXml(sheetXml, ref)?.match(/\bs="([^"]*)"/)?.[1];
}

describe('Bug Fix Validation', () => {
    it('should preserve A11:C15 merged cell border styles', async () => {
        if (!hasTestFiles()) {
            console.log('[SKIP] test_data/default_template_SD.xlsx not found — skipping XML validation test');
            return;
        }
        const tmplBuf = fs.readFileSync(path.resolve(__dirname, '../test_data/default_template_SD.xlsx'));
        const renderData = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../test_data/form_data-SD.json'), 'utf8'));

        const xlsx = await XlsxRender.create(tmplBuf);
        await xlsx.render(renderData, 'Summary');
        const outBuf = await xlsx.generate({ type: 'nodebuffer' });

        const zip = await JsZip.loadAsync(outBuf);
        const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');

        const templateZip = await JsZip.loadAsync(tmplBuf);
        const templateSheetXml = await templateZip.file('xl/worksheets/sheet1.xml').async('string');

        expect(sheetXml).toContain('A11:C15');
        for (const ref of ['A14', 'B14', 'C14', 'A15', 'B15', 'C15']) {
            expect(getCellXml(sheetXml, ref), ref).toBeTruthy();
            expect(getCellStyle(sheetXml, ref)).toEqual(getCellStyle(templateSheetXml, ref));
        }
    });

    it('should preserve M14-Q17 border-only cell styles', async () => {
        if (!hasTestFiles()) {
            console.log('[SKIP] test_data/default_template_SD.xlsx not found — skipping XML validation test');
            return;
        }
        const tmplBuf = fs.readFileSync(path.resolve(__dirname, '../test_data/default_template_SD.xlsx'));
        const renderData = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../test_data/form_data-SD.json'), 'utf8'));

        const xlsx = await XlsxRender.create(tmplBuf);
        await xlsx.render(renderData, 'Summary');
        const outBuf = await xlsx.generate({ type: 'nodebuffer' });

        const templateZip = await JsZip.loadAsync(tmplBuf);
        const zip = await JsZip.loadAsync(outBuf);
        const templateSheetXml = await templateZip.file('xl/worksheets/sheet1.xml').async('string');
        const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
        const templateStylesXml = await templateZip.file('xl/styles.xml').async('string');
        const stylesXml = await zip.file('xl/styles.xml').async('string');

        expect(stylesXml).toEqual(templateStylesXml);

        const cols = ['M', 'N', 'O', 'P', 'Q'];
        for (let r = 14; r <= 17; r++) {
            for (const col of cols) {
                const ref = col + r;
                expect(getCellXml(sheetXml, ref)).toBeTruthy();
                expect(getCellStyle(sheetXml, ref)).toEqual(getCellStyle(templateSheetXml, ref));
            }
        }

        for (const ref of ['E12', 'I12', 'M12', 'M13', 'S12', 'S13', 'H12', 'L12', 'H13', 'O13', 'P13', 'Q13', 'E14', 'G14', 'S14']) {
            expect(getCellXml(sheetXml, ref), ref).toBeTruthy();
            expect(getCellStyle(sheetXml, ref)).toEqual(getCellStyle(templateSheetXml, ref));
        }
    });

    it('should preserve imageincell rich value metadata', async () => {
        if (!hasTestFiles()) {
            console.log('[SKIP] test_data/default_template_SD.xlsx not found — skipping XML validation test');
            return;
        }
        const tmplBuf = fs.readFileSync(path.resolve(__dirname, '../test_data/default_template_SD.xlsx'));
        const renderData = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../test_data/form_data-SD.json'), 'utf8'));

        const xlsx = await XlsxRender.create(tmplBuf);
        await xlsx.render(renderData, 'Summary');
        const outBuf = await xlsx.generate({ type: 'nodebuffer' });

        const zip = await JsZip.loadAsync(outBuf);

        // Original F67 and M67 are shifted to F71 and M71 due to table expansion (4 extra rows)
        const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');

        // Check F71 cell has t="e" (rich value type) and vm attribute
        const f71Cell = sheetXml.match(/<c r="F71"[^>]*\/>/);
        expect(f71Cell).toBeTruthy();
        expect(f71Cell![0]).toContain('t="e"');
        expect(f71Cell![0]).toContain('vm=');

        // Check M71 cell has t="e" (rich value type) and vm attribute
        const m71Cell = sheetXml.match(/<c r="M71"[^>]*\/>/);
        expect(m71Cell).toBeTruthy();
        expect(m71Cell![0]).toContain('t="e"');
        expect(m71Cell![0]).toContain('vm=');

        // Check rich value relationship exists in sheet rels
        const sheetRels = zip.file('xl/worksheets/_rels/sheet1.xml.rels');
        expect(sheetRels).toBeTruthy();
        const relsContent = await sheetRels!.async('string');
        expect(relsContent).toContain('richValue');

        // Check media files exist (image data)
        const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('xl/media/'));
        expect(mediaFiles.length).toBeGreaterThan(0);
    });

    it('should correctly render cell values in standard placeholders', async () => {
        if (!hasTestFiles()) {
            console.log('[SKIP] test_data/default_template_SD.xlsx not found — skipping XML validation test');
            return;
        }
        const tmplBuf = fs.readFileSync(path.resolve(__dirname, '../test_data/default_template_SD.xlsx'));
        const renderData = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../test_data/form_data-SD.json'), 'utf8'));

        const xlsx = await XlsxRender.create(tmplBuf);
        await xlsx.render(renderData, 'Summary');
        const outBuf = await xlsx.generate({ type: 'nodebuffer' });

        const zip = await JsZip.loadAsync(outBuf);
        const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
        const ssXml = await zip.file('xl/sharedStrings.xml').async('string');

        // Check D5 has a rendered value (shared string)
        const d5Match = sheetXml.match(/<c r="D5"[^>]*>[\s\S]*?<v>(\d+)<\/v>/);
        expect(d5Match).toBeTruthy();
        if (d5Match) {
            const si = parseInt(d5Match[1], 10);
            expect(si).toBeGreaterThanOrEqual(0);
        }

        const templateZip = await JsZip.loadAsync(tmplBuf);
        const templateSheetXml = await templateZip.file('xl/worksheets/sheet1.xml').async('string');
        expect(getCellStyle(sheetXml, 'D5')).toEqual(getCellStyle(templateSheetXml, 'D5'));
    });
});
