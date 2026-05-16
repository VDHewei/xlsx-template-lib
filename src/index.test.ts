import * as fs from "node:fs/promises";
import {BufferType, generateXlsxTemplate, Workbook} from './core'
import {assertType, describe, expect, expectTypeOf, it, Mock, vi} from 'vitest'
import {
    AddCommand,
    Argument,
    compileRuleSheetName,
    CmdFunction,
    generateCommandsXlsxTemplate,
    generateCommandsXlsxTemplateWithCompile,
    getCommands,
} from './extends'
import {
    compileWorkSheet,
    DefaultPlaceholderCellValue,
    exceljs,
    loadWorkbook,
    parseWorkSheetRules,
    PlaceholderCellValue, RuleMapOptions,
    RuleResult,
    RuleToken,
    scanCellSetPlaceholder
} from './helper';

import {
    ZipXlsxTemplateApp,
} from './biz';

async function createMockBuffer(options: {
    targetValue?: string | null;
    merged?: boolean;
    leftValues?: (string | null)[];
}): Promise<Buffer> {
    const wb = new exceljs.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    const colNum = 2;
    const rowNum = 2;
    if (options.merged) {
        const mergeRows = options.leftValues?.length || 1;
        ws.mergeCells(rowNum, colNum, rowNum + mergeRows - 1, colNum);
        if (options.leftValues) {
            options.leftValues.forEach((val, idx) => {
                if (val !== null && val !== undefined) {
                    ws.getCell(rowNum + idx, colNum - 1).value = val;
                }
            });
        }
    }
    if (options.targetValue) {
        ws.getCell(rowNum, colNum).value = options.targetValue;
    }
    return Buffer.from(await wb.xlsx.writeBuffer());
}

function getPlaceholder(): {
    placeholder: PlaceholderCellValue,
    spyToString: Mock<() => string>,
    spyMerge: Mock<(values: string[]) => string>
} {
    const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
    const spyToString = vi.spyOn(placeholder, 'toString');
    const spyMerge = vi.spyOn(placeholder, 'mergeCell');
    return {
        placeholder,
        spyToString,
        spyMerge,
    }
}

function testEnv(key: symbol, value: string, extKey?: string): boolean {
    const k = key.toString();
    let x = k.substring(7, k.length - 1);
    if (extKey !== undefined && extKey !== "") {
        x = `${x}.${extKey}`
    }
    return process.env[x] === value;
}

const XlsxTest = Symbol(`VITE_SAVE_XLSX_TEST`);
const BackendTest = Symbol(`VITE_SAVA_BACKEND_TEST`);
const CompileTest = Symbol(`VITE_SAVE_COMPILE_XLSX_TEST`);

describe('generateXlsxTemplate', {tags: ["backend"]}, () => {
    it('should generate a template', async () => {
        // 创建内联模板（无需外部测试文件）
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '${name}';
        ws.getCell('B1').value = '${age}';
        const xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        const data = { name: 'test', age: '25' };
        const buffer = await generateXlsxTemplate(xlsx, data, { type: BufferType.NodeBuffer });
        if (testEnv(BackendTest, 'true')) {
            await fs.writeFile(`./test_data/test_${new Date().valueOf()}.xlsx`, buffer);
        }
        expect(buffer).toBeInstanceOf(Buffer);
        // 验证数据填充正确
        const w = await loadWorkbook(buffer);
        const sheet = w.getWorksheet('Sheet1');
        expect(sheet.getCell('A1').value).equal('test');
        expect(sheet.getCell('B1').value).equal('25');
    });

    it('should generate a template with data', async () => {
        // 创建内联模板和嵌套数据
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '${user.name}';
        ws.getCell('B1').value = '${user.age}';
        const xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        const values = { user: { name: 'Alice', age: '30' } };
        const buffer = await generateXlsxTemplate(xlsx, values, { type: BufferType.NodeBuffer });
        if (testEnv(BackendTest, 'true')) {
            await fs.writeFile(`./test_data/test_${new Date().valueOf()}_data.xlsx`, buffer);
        }
        expect(buffer).toBeInstanceOf(Buffer);
        const w = await loadWorkbook(buffer);
        const sheet = w.getWorksheet('Sheet1');
        expect(sheet.getCell('A1').value).equal('Alice');
        expect(sheet.getCell('B1').value).equal('30');
    });

    it('should generate a table data', async () => {
        // 创建内联模板：多列表格数据
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Summary');
        ws.getCell('B4').value = '${name}';
        ws.getCell('B5').value = '${birthDate}';
        ws.getCell('B6').value = '${shortDate}';
        ws.getCell('D5').value = '${weather}';
        ws.getCell('D6').value = '${weather}';
        ws.getCell('D7').value = '${count}';
        ws.getCell('A21').value = '${label21}';
        ws.getCell('A27').value = '${label27}';
        // 表格行：${table:items.title} 和 ${table:items.num}
        ws.getCell('E13').value = '${table:items.title}';
        ws.getCell('G13').value = '${table:items.num}';
        const xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        const values = {
            name: 'VARATEST1',
            birthDate: '1992-05-09',
            shortDate: '05-09',
            weather: 'Cloudy',
            count: '1',
            label21: 'Instruction',
            label27: 'Comments',
            items: [
                { title: 'Amah', num: '1' },
                { title: 'Amah (Seconded to ARUP)', num: '2' },
                { title: 'Assistant Construction Manager', num: '3' },
            ],
        };
        const buffer = await generateXlsxTemplate(xlsx, values, { type: BufferType.NodeBuffer });
        expect(buffer).toBeInstanceOf(Buffer);
        const w = await loadWorkbook(buffer);
        expect(w).toBeInstanceOf(exceljs.Workbook);
        const sheet = w.getWorksheet('Summary');
        expect(sheet.getRow(4).getCell('B').value).equal('VARATEST1');
        expect(sheet.getRow(5).getCell('B').value).equal('1992-05-09');
        expect(sheet.getRow(6).getCell('B').value).equal('05-09');

        expect(sheet.getRow(5).getCell('D').value).equal('Cloudy');
        expect(sheet.getRow(6).getCell('D').value).equal('Cloudy');
        expect(sheet.getRow(7).getCell('D').value).equal('1');

        expect(sheet.getRow(21).getCell('A').value).equal('Instruction');
        expect(sheet.getRow(27).getCell('A').value).equal('Comments');

        expect(sheet.getRow(13).getCell('E').value).equal('Amah');
        expect(sheet.getRow(13).getCell('G').value).equal('1');
        expect(sheet.getRow(14).getCell('E').value).equal('Amah (Seconded to ARUP)');
        expect(sheet.getRow(14).getCell('G').value).equal('2');
        expect(sheet.getRow(15).getCell('E').value).equal('Assistant Construction Manager');
        expect(sheet.getRow(15).getCell('G').value).equal('3');
    });

    it('should generate with image in cell', async () => {
        // 1x1 透明 PNG 的 base64 编码（最小有效图片）
        const base64Image = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
        const values = {
            formStatusHistories: [
                { actionSignatureBase64: base64Image },
            ],
        };
        // 创建带 imageincell 占位符的模板
        const JsZip = (await import('jszip')).default;
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '${imageincell:formStatusHistories.0.actionSignatureBase64}';
        const templateBuffer = Buffer.from(await wb.xlsx.writeBuffer());
        const buffer = await generateXlsxTemplate(templateBuffer, values, { type: BufferType.NodeBuffer });
        expect(buffer).toBeInstanceOf(Buffer);
        // 验证图片已嵌入输出档案
        const zip = await JsZip.loadAsync(buffer);
        const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('xl/media/') && !zip.files[f].dir);
        expect(mediaFiles.length).toBeGreaterThan(0);
        const w = await loadWorkbook(buffer);
        expect(w).toBeInstanceOf(exceljs.Workbook);
        const sheet = w.getWorksheet('Sheet1');
        const cellValue = sheet.getCell('A1').value;
        // 图片通过绘图层嵌入，单元格值不应是 base64 字符串
        expect(cellValue).not.equal(base64Image);
    });

    it('should fill existing rows with table data when useExistingRows is true', async () => {
        // Create a template with a table placeholder and pre-formatted empty rows below
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '${table:users.name}';
        // Pre-format 3 empty rows below the template row
        ws.getCell('A2').value = '';
        ws.getCell('A3').value = '';
        ws.getCell('A4').value = '';
        const templateBuffer = Buffer.from(await wb.xlsx.writeBuffer());

        const values = {
            users: [
                {name: 'Alice'},
                {name: 'Bob'},
                {name: 'Charlie'},
            ]
        };

        const buffer = await generateXlsxTemplate(templateBuffer, values, {
            type: BufferType.NodeBuffer,
            useExistingRows: true,
        });
        expect(buffer).toBeInstanceOf(Buffer);

        const w = await loadWorkbook(buffer);
        const sheet = w.getWorksheet('Sheet1');
        // Template row filled with first item
        expect(sheet.getCell('A1').value).equal('Alice');
        // Pre-formatted rows filled with subsequent items
        expect(sheet.getCell('A2').value).equal('Bob');
        expect(sheet.getCell('A3').value).equal('Charlie');
        // Row 4 should be empty (was pre-formatted but data only has 3 items)
        expect(sheet.getCell('A4').value === null || sheet.getCell('A4').value === '').toBeTruthy();
        // No extra rows should be created (total should be 4, the original row count)
        expect(sheet.rowCount).equal(4);
    })

    it('should create new rows when table data exceeds pre-formatted rows', async () => {
        // Template has A1 placeholder and 2 pre-formatted rows
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '${table:users.name}';
        ws.getCell('A2').value = '';
        ws.getCell('A3').value = '';
        const templateBuffer = Buffer.from(await wb.xlsx.writeBuffer());

        const values = {
            users: [
                {name: 'Alice'},
                {name: 'Bob'},
                {name: 'Charlie'},
                {name: 'Diana'},
                {name: 'Eve'},
            ]
        };

        const buffer = await generateXlsxTemplate(templateBuffer, values, {
            type: BufferType.NodeBuffer,
            useExistingRows: true,
        });
        expect(buffer).toBeInstanceOf(Buffer);

        const w = await loadWorkbook(buffer);
        const sheet = w.getWorksheet('Sheet1');
        // Template row + 2 pre-formatted rows + 2 new rows = 5
        expect(sheet.rowCount).equal(5);
        expect(sheet.getCell('A1').value).equal('Alice');
        expect(sheet.getCell('A2').value).equal('Bob');
        expect(sheet.getCell('A3').value).equal('Charlie');
        expect(sheet.getCell('A4').value).equal('Diana');
        expect(sheet.getCell('A5').value).equal('Eve');
    })
})

describe('generateCommandsXlsxTemplate', {tags: ["backend"]}, () => {
    it('should generate a template', async () => {
        // 创建内联模板
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '${name}';
        ws.getCell('B1').value = '${age}';
        const xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        const data = { name: 'test', age: '25' };
        const buffer = await generateCommandsXlsxTemplate(xlsx, data, { type: BufferType.NodeBuffer });
        if (testEnv(BackendTest, 'true')) {
            await fs.writeFile(`./test_data/test_cmd_${new Date().valueOf()}.xlsx`, buffer);
        }
        expect(buffer).toBeInstanceOf(Buffer);
    });

    it("get commands", () => {
        const cmds = getCommands();
        assertType<number>(cmds.size);
        expect(cmds.size).not.equal(0, "empty builtin command")
        for (let [key, cmd] of cmds.entries()) {
            expectTypeOf(key).toEqualTypeOf<string>();
            expectTypeOf(cmd).toEqualTypeOf<CmdFunction>();
        }
    })

    it("add commands", () => {
        const cmds = getCommands();
        let size = cmds.size;
        assertType<number>(size);
        expect(size).not.equal(0, "empty builtin command")
        for (let [key, cmd] of cmds.entries()) {
            expectTypeOf(key).toEqualTypeOf<string>();
            expectTypeOf(cmd).toEqualTypeOf<CmdFunction>();
        }
        AddCommand("test", (values: Object | Record<string, any>, argument: Argument): any | undefined => {
            return "test";
        });
        AddCommand("hello", (values: Object | Record<string, any>, argument: Argument): any | undefined => {
            return "hello";
        });
        expect(cmds.size).equal(size + 2, "add command size not matched")
        expect(cmds.has("test")).equal(true, "check test command failed")
        expect(cmds.has("hello")).equal(true, "check hello command failed")
        assertType<CmdFunction>(cmds.get("test"));
        assertType<CmdFunction>(cmds.get("hello"));
    })

    it('should command generate a template with data', async () => {
        // 创建内联模板（使用简单占位符，避免复杂嵌套路径）
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = '${name}';
        ws.getCell('B1').value = '${age}';
        const xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        const values = { name: 'Bob', age: '28' };
        const buffer = await generateCommandsXlsxTemplate(xlsx, values, { type: BufferType.NodeBuffer });
        if (testEnv(BackendTest, 'true')) {
            await fs.writeFile(`./test_data/test_cmd_${new Date().valueOf()}_data.xlsx`, buffer);
        }
        expect(buffer).toBeInstanceOf(Buffer);
        const w = await loadWorkbook(buffer);
        const sheet = w.getWorksheet('Sheet1');
        expect(sheet.getCell('A1').value).equal('Bob');
        expect(sheet.getCell('B1').value).equal('28');
    });
})

describe('scanCellSetPlaceholder', {tags: ["backend", "xlsx"]}, () => {

    it('未合并且空单元格时调用toString', async () => {
        const buffer = await createMockBuffer({targetValue: null});
        const {placeholder, spyToString, spyMerge} = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, {Row: 'B', Column: 2, Sheet: "Sheet1"}, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_1_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).toHaveBeenCalledOnce();
        expect(spyMerge).not.toHaveBeenCalled();

    });

    it('未合并且非空单元格时不调用任何方法', async () => {
        const buffer = await createMockBuffer({targetValue: 'Existing Data'});
        const {placeholder, spyToString, spyMerge} = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, {Row: 'B', Column: 2, Sheet: "Sheet1"}, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_2_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).not.toHaveBeenCalled();
        expect(spyMerge).not.toHaveBeenCalled();
    });

    it('合并单元格且左侧全空时调用toString', async () => {
        const buffer = await createMockBuffer({merged: true, leftValues: [null, null, null]});
        const {placeholder, spyToString, spyMerge} = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, {Row: 'B', Column: 2, Sheet: "Sheet1"}, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_3_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).toHaveBeenCalledOnce();
        expect(spyMerge).not.toHaveBeenCalled();
    });

    it('合并单元格且左侧非全空时调用 mergeCell 并传入过滤后的数组', async () => {
        const buffer = await createMockBuffer({merged: true, leftValues: ['Val1', null, 'Val2', '']});
        const {placeholder, spyToString, spyMerge} = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, {Row: 'B', Column: 2, Sheet: "Sheet1"}, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_4_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).not.toHaveBeenCalled();
        expect(spyMerge).toHaveBeenCalledWith(['Val1', 'Val2']);
    });

    it('支持 base64 字符串入参', async () => {
        const buffer = await createMockBuffer({targetValue: null});
        const base64Str = buffer.toString('base64');
        const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
        const spyToString = vi.spyOn(placeholder, 'toString');
        const res = await scanCellSetPlaceholder(base64Str, {Row: 'B', Column: 2, Sheet: "Sheet1"}, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_5_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).toHaveBeenCalledOnce();
    });

});

describe('compileWorkSheet', {tags: ["compile"]}, () => {
    it('parse-rules-only', async () => {
        // 创建包含规则配置的测试工作表
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('export_metadata.config');
        ws.getCell('A1').value = 'A:C: = x := y';
        ws.getCell('A2').value = 'M:G: = z := w';
        ws.getCell('A3').value = 'RC:N: = a := b';
        ws.getCell('A4').value = 'MC:O: = c := d';
        const xlsxBuf = Buffer.from(await wb.xlsx.writeBuffer());
        const workbook = await loadWorkbook(xlsxBuf);
        const sheetName = 'export_metadata.config';
        const res = parseWorkSheetRules(workbook.getWorksheet(sheetName));
        expectTypeOf<RuleResult>(res);
        expect(res.rules.size).not.equal(0, '输出结果异常');
    });

    it('compile-only', async () => {
        const sheetName = 'export_metadata.config';
        // 创建内联 xlsx 文件
        const wb = new exceljs.Workbook();
        wb.addWorksheet(sheetName);
        const xlsxBuf = Buffer.from(await wb.xlsx.writeBuffer());
        const res = await compileWorkSheet(xlsxBuf, sheetName);
        expectTypeOf<exceljs.Xlsx | Error[]>(res);
        assertType<exceljs.Xlsx>(res as exceljs.Xlsx);
        if (testEnv(CompileTest, 'true')) {
            const sv = res as exceljs.Xlsx;
            await sv.writeFile(`./test_data/test_compile_${new Date().valueOf()}.xlsx`);
        }
    });

    it('withData', async () => {
        // 创建内联模板和 compile 工作簿（编译测试只需验证流程不报错）
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('export_metadata.config');
        ws.getCell('A1').value = 'A:C: = name := Name';
        const xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        const values = { name: 'test' };
        const compileOptions = new RuleMapOptions();
        compileOptions.sheetName = compileRuleSheetName;
        const save = testEnv(CompileTest, 'true', 'WITH_DATA');
        compileOptions.save = save;
        compileOptions.saveFile = './test_data/withData_';
        compileOptions.skipRemoveUnExportSheet = true;
        const bf = await generateCommandsXlsxTemplateWithCompile(xlsx, values, compileOptions, { type: BufferType.NodeBuffer });
        if (save) {
            await fs.writeFile(`./test_data/test_compile_${new Date().valueOf()}_data.xlsx`, bf);
        }
        expect(bf).toBeInstanceOf(Buffer);
    });
});


describe('compileZip', {tags: ["compile"]}, () => {
    it('zipCompile', async () => {
        // 创建内联 zip 包含一个 xlsx 文件
        const AdmZip = (await import('adm-zip')).default;
        const wb = new exceljs.Workbook();
        wb.addWorksheet('Sheet1');
        const xlsxBuf = Buffer.from(await wb.xlsx.writeBuffer());
        const admZip = new AdmZip();
        admZip.addFile('template.xlsx', xlsxBuf);
        const fd = admZip.toBuffer();
        const processedBuffer = await ZipXlsxTemplateApp.compileTo(Buffer.from(fd), {});
        expectTypeOf<Buffer>(processedBuffer);
        expect(processedBuffer.length).not.equal(0, '输出结果异常');
        if (testEnv(XlsxTest, 'true', 'ZIP_COMPILE')) {
            await fs.writeFile(`./test_data/test_zip_3_${new Date().valueOf()}.zip`, processedBuffer);
        }
    });
});