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

describe('Bug Fix Validation', () => {
    it('should preserve A11:C15 merged cell without phantom A14/A15 cells', async () => {
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

        // Check that A14 and A15 do NOT exist as standalone cell elements
        expect(sheetXml).not.toMatch(/<c r="A14"[ >]/);
        expect(sheetXml).not.toMatch(/<c r="A15"[ >]/);

        // Verify merge range A11:C15 still exists
        expect(sheetXml).toContain('A11:C15');
    });

    it('should preserve M14-Q17 border-only cells (shifted to M18-Q22)', async () => {
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

        // Original rows 14-17 are pushed down by 4 rows (5 formStatusHistories items expand table by 4 rows)
        // So cells at M14-Q17 should now be at M18-Q22
        const cols = ['M', 'N', 'O', 'P', 'Q'];
        for (let r = 18; r <= 21; r++) {
            for (const col of cols) {
                const ref = col + r;
                const cellRegex = new RegExp(`<c r="${ref}"`);
                expect(sheetXml).toMatch(cellRegex);
            }
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
    });
});
