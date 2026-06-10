import * as fs from "node:fs/promises";
import { constants } from "node:fs";
import {BufferType, generateXlsxTemplate, valueDotGet, Workbook} from './core'
import { assertType, describe, expect, expectTypeOf, it, Mock, vi } from 'vitest'
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
    ExprResolver,
    loadWorkbook,
    parseWorkSheetRules,
    PlaceholderCellValue, RuleMapOptions,
    RuleResult,
    RuleToken,
    scanCellSetPlaceholder
} from './helper';

import {
    formStatusImage,
    ZipXlsxTemplateApp,
} from './biz';

async function fileExists(path: string): Promise<boolean> {
    try {
        await fs.access(path, constants.F_OK);
        return true;
    } catch {
        return false;
    }
}

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

describe('generateXlsxTemplate', { tags: ["backend"] }, () => {
    AddCommand("formStatusImage",formStatusImage);
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
        // 判断 test_data 文件夹是否存在，不存在则创建
        let wb: exceljs.Workbook;
        let xlsx: Buffer;
        await fs.mkdir('./test_data', { recursive: true });
        //判断 test_data/default_template_SD.xlsx 是否存在，存在则加载 exceljs.Workbook 对象
        const templatePath = './test_data/default_template_SD.xlsx';
        const templateDataPath = './test_data/form_data-SD.json';
        const newFile = `./test_data/default_template_${new Date().valueOf()}_SD.xlsx`;
        if (await fileExists(templatePath)) {
            const templateBuffer = await fs.readFile(templatePath);
            xlsx = Buffer.from(templateBuffer);
            wb = await loadWorkbook(templateBuffer);
        } else {
            // 不存在则创建一个新的模板文件
            wb = new exceljs.Workbook();
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

            xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        }

        let values: Record<string, any> = {};
        if (await fileExists(templateDataPath)) {
            const templateData = await fs.readFile(templateDataPath, 'utf-8');
            values = JSON.parse(templateData);
        } else {
            values = {
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
        }

        const buffer = await generateXlsxTemplate(xlsx, values, { type: BufferType.NodeBuffer });
        await fs.writeFile(newFile, buffer);
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

        //expect(sheet.getRow(13).getCell('E').value).equal('Amah.1');
        expect(sheet.getRow(13).getCell('E').value).equal('Amah');
        expect(sheet.getRow(13).getCell('G').value).equal('1');
      //  expect(sheet.getRow(14).getCell('E').value).equal('Amah (Seconded to ARUP).2');
        expect(sheet.getRow(14).getCell('E').value).equal('Amah (Seconded to ARUP)');
        expect(sheet.getRow(14).getCell('G').value).equal('2');
      //  expect(sheet.getRow(15).getCell('E').value).equal('Assistant Construction Manager.3');
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
                { name: 'Alice' },
                { name: 'Bob' },
                { name: 'Charlie' },
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
                { name: 'Alice' },
                { name: 'Bob' },
                { name: 'Charlie' },
                { name: 'Diana' },
                { name: 'Eve' },
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

describe('generateCommandsXlsxTemplate', { tags: ["backend"] }, () => {
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

describe('scanCellSetPlaceholder', { tags: ["backend", "xlsx"] }, () => {

    it('未合并且空单元格时调用toString', async () => {
        const buffer = await createMockBuffer({ targetValue: null });
        const { placeholder, spyToString, spyMerge } = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, { Row: 'B', Column: 2, Sheet: "Sheet1" }, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_1_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).toHaveBeenCalledOnce();
        expect(spyMerge).not.toHaveBeenCalled();

    });

    it('未合并且非空单元格时不调用任何方法', async () => {
        const buffer = await createMockBuffer({ targetValue: 'Existing Data' });
        const { placeholder, spyToString, spyMerge } = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, { Row: 'B', Column: 2, Sheet: "Sheet1" }, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_2_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).not.toHaveBeenCalled();
        expect(spyMerge).not.toHaveBeenCalled();
    });

    it('合并单元格且左侧全空时调用toString', async () => {
        const buffer = await createMockBuffer({ merged: true, leftValues: [null, null, null] });
        const { placeholder, spyToString, spyMerge } = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, { Row: 'B', Column: 2, Sheet: "Sheet1" }, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_3_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).toHaveBeenCalledOnce();
        expect(spyMerge).not.toHaveBeenCalled();
    });

    it('合并单元格且左侧非全空时调用 mergeCell 并传入过滤后的数组', async () => {
        const buffer = await createMockBuffer({ merged: true, leftValues: ['Val1', null, 'Val2', ''] });
        const { placeholder, spyToString, spyMerge } = getPlaceholder();
        const res = await scanCellSetPlaceholder(buffer, { Row: 'B', Column: 2, Sheet: "Sheet1" }, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_4_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).not.toHaveBeenCalled();
        expect(spyMerge).toHaveBeenCalledWith(['Val1', 'Val2']);
    });

    it('支持 base64 字符串入参', async () => {
        const buffer = await createMockBuffer({ targetValue: null });
        const base64Str = buffer.toString('base64');
        const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
        const spyToString = vi.spyOn(placeholder, 'toString');
        const res = await scanCellSetPlaceholder(base64Str, { Row: 'B', Column: 2, Sheet: "Sheet1" }, placeholder);
        expectTypeOf<ArrayBuffer>(res);
        expect(res.byteLength).not.equal(0, "输出结果异常")
        if (testEnv(XlsxTest, "true")) {
            await fs.writeFile(`./test_data/test_scanCell_5_${new Date().valueOf()}.xlsx`, res as any)
        }
        expect(spyToString).toHaveBeenCalledOnce();
    });

});

describe('compileWorkSheet', { tags: ["compile"] }, () => {
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


describe('compileZip', { tags: ["compile"] }, () => {
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

describe('date format placeholders', { tags: ["backend"] }, () => {
    // 统一测试用数据
    const ISO_DATE = '1992-05-09T00:00:00.000Z';
    const UNIX_TS = 705369600; // 1992-05-09 UTC (Math.floor(new Date('1992-05-09T00:00:00Z').getTime() / 1000))
    const DATE_OBJ = new Date('1992-05-09T00:00:00.000Z');

    async function buildAndFill(
        placeholders: Record<string, string>,
        values: Record<string, any>
    ): Promise<exceljs.Worksheet> {
        const wb = new exceljs.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        let col = 1;
        for (const [addr, ph] of Object.entries(placeholders)) {
            ws.getCell(addr).value = ph;
            col++;
        }
        const xlsx = Buffer.from(await wb.xlsx.writeBuffer());
        const buffer = await generateXlsxTemplate(xlsx, values, { type: BufferType.NodeBuffer });
        expect(buffer).toBeInstanceOf(Buffer);
        const w = await loadWorkbook(buffer);
        return w.getWorksheet('Sheet1');
    }

    it('${user.birthday:date} — 普通对象路径，date 格式化 (ISO string)', async () => {
        const sheet = await buildAndFill(
            { A1: '${user.birthday:date}' },
            { user: { birthday: ISO_DATE } }
        );
        expect(sheet.getCell('A1').value).equal('1992-05-09');
    });

    it('${user.birthday:date} — Date 对象输入', async () => {
        const sheet = await buildAndFill(
            { A1: '${user.birthday:date}' },
            { user: { birthday: DATE_OBJ } }
        );
        expect(sheet.getCell('A1').value).equal('1992-05-09');
    });

    it('${user.birthday:date} — Unix 时间戳(数字)输入', async () => {
        const sheet = await buildAndFill(
            { A1: '${user.birthday:date}' },
            { user: { birthday: UNIX_TS } }
        );
        expect(sheet.getCell('A1').value).equal('1992-05-09');
    });

    it('${user.birthday:date} — Unix 时间戳(字符串)输入', async () => {
        const sheet = await buildAndFill(
            { A1: '${user.birthday:date}' },
            { user: { birthday: String(UNIX_TS) } }
        );
        expect(sheet.getCell('A1').value).equal('1992-05-09');
    });

    it('${user.birthday:day} — 普通对象路径，day 格式化 (ISO string)', async () => {
        const sheet = await buildAndFill(
            { A1: '${user.birthday:day}' },
            { user: { birthday: ISO_DATE } }
        );
        expect(sheet.getCell('A1').value).equal('05-09');
    });

    it('${user.birthday:day} — Date 对象输入', async () => {
        const sheet = await buildAndFill(
            { A1: '${user.birthday:day}' },
            { user: { birthday: DATE_OBJ } }
        );
        expect(sheet.getCell('A1').value).equal('05-09');
    });

    it('${user.birthday:day} — Unix 时间戳(数字)输入', async () => {
        const sheet = await buildAndFill(
            { A1: '${user.birthday:day}' },
            { user: { birthday: UNIX_TS } }
        );
        expect(sheet.getCell('A1').value).equal('05-09');
    });

    it('${table:users.birthday:date} — 表格数组 date 格式化', async () => {
        const sheet = await buildAndFill(
            { A1: '${table:users.birthday:date}' },
            {
                users: [
                    { birthday: '1992-05-09' },
                    { birthday: '2000-12-25' },
                ]
            }
        );
        expect(sheet.getCell('A1').value).equal('1992-05-09');
        expect(sheet.getCell('A2').value).equal('2000-12-25');
    });

    it('${table:users.birthday:day} — 表格数组 day 格式化', async () => {
        const sheet = await buildAndFill(
            { A1: '${table:users.birthday:day}' },
            {
                users: [
                    { birthday: '1992-05-09' },
                    { birthday: '2000-12-25' },
                ]
            }
        );
        expect(sheet.getCell('A1').value).equal('05-09');
        expect(sheet.getCell('A2').value).equal('12-25');
    });

    it('${table:users.birthday:date} — 表格中混合 Unix 时间戳', async () => {
        const sheet = await buildAndFill(
            { A1: '${table:users.birthday:date}' },
            {
                users: [
                    { birthday: UNIX_TS },
                    { birthday: '2000-12-25T00:00:00.000Z' },
                ]
            }
        );
        expect(sheet.getCell('A1').value).equal('1992-05-09');
        expect(sheet.getCell('A2').value).equal('2000-12-25');
    });
});

describe('compile subType alias expansion', { tags: ["compile"] }, () => {
    // Builds a workbook with a rule sheet and an empty target sheet, runs compile, returns the target worksheet.
    async function compileAndGetSheet(
        rules: Array<[type: string, value: string]>,
        targetSheetName = 'Sheet1'
    ): Promise<exceljs.Worksheet> {
        const wb = new exceljs.Workbook();
        const configSheet = wb.addWorksheet('export_metadata.config');
        const dataSheet = wb.addWorksheet(targetSheetName);

        rules.forEach(([type, value], idx) => {
            configSheet.getCell(`A${idx + 1}`).value = type;
            configSheet.getCell(`B${idx + 1}`).value = value;
        });
        // Ensure all referenced cells exist (findCell returns undefined for non-existent cells)
        dataSheet.getCell('A1').value = '';
        dataSheet.getCell('B1').value = '';
        dataSheet.getCell('C1').value = '';

        const xlsxBuf = Buffer.from(await wb.xlsx.writeBuffer());
        const compileOptions = new RuleMapOptions();
        compileOptions.compileSheets = [targetSheetName];
        const result = await ExprResolver.compile(xlsxBuf, 'export_metadata.config', compileOptions);
        expect(result.errs, `compile errors: ${result.errs?.map(e => e.message).join(', ')}`).toBeUndefined();
        return result.workbook.getWorksheet(targetSheetName);
    }

    it('${@#.@MY:date} — dual alias with :date subType preserved', async () => {
        const sheet = await compileAndGetSheet([
            ['alias', '#=exportData.LRR'],
            ['alias', 'MY=mothOrYear'],
            ['cell', 'A:1=${@#.@MY:date}'],
        ]);
        expect(sheet.getCell('A1').value).equal('${exportData.LRR.mothOrYear:date}');
    });

    it('${@#.@MY:day} — dual alias with :day subType preserved', async () => {
        const sheet = await compileAndGetSheet([
            ['alias', '#=exportData.LRR'],
            ['alias', 'MY=mothOrYear'],
            ['cell', 'A:1=${@#.@MY:day}'],
        ]);
        expect(sheet.getCell('A1').value).equal('${exportData.LRR.mothOrYear:day}');
    });

    it('${@LRR.mothOrYear:date} — single alias with :date subType preserved', async () => {
        const sheet = await compileAndGetSheet([
            ['alias', 'LRR=exportData.LRR'],
            ['cell', 'A:1=${@LRR.mothOrYear:date}'],
        ]);
        expect(sheet.getCell('A1').value).equal('${exportData.LRR.mothOrYear:date}');
    });

    it('${@T} — alias without subType still works', async () => {
        const sheet = await compileAndGetSheet([
            ['alias', 'T=user.name'],
            ['cell', 'A:1=${@T}'],
        ]);
        expect(sheet.getCell('A1').value).equal('${user.name}');
    });

    it('${@#.@MY:date} — compile then fill produces formatted date', async () => {
        // Full round-trip: compile aliases, then fill with real data to verify :date still formats.
        // Uses a 2-level data path to match the existing date-format test pattern.
        const wb = new exceljs.Workbook();
        const configSheet = wb.addWorksheet('export_metadata.config');
        const dataSheet = wb.addWorksheet('Sheet1');

        configSheet.getCell('A1').value = 'alias';
        configSheet.getCell('B1').value = '#=person';
        configSheet.getCell('A2').value = 'alias';
        configSheet.getCell('B2').value = 'BD=birthday';
        configSheet.getCell('A3').value = 'cell';
        configSheet.getCell('B3').value = 'A:1=${@#.@BD:date}';
        dataSheet.getCell('A1').value = '';

        const xlsxBuf = Buffer.from(await wb.xlsx.writeBuffer());
        const compileOptions = new RuleMapOptions();
        compileOptions.compileSheets = ['Sheet1'];
        const result = await ExprResolver.compile(xlsxBuf, 'export_metadata.config', compileOptions);
        expect(result.errs).toBeUndefined();

        // Verify compile produced the correct placeholder text
        expect(result.workbook.getWorksheet('Sheet1').getCell('A1').value)
            .equal('${person.birthday:date}');

        // Remove config sheet so generateXlsxTemplate only sees the data sheet
        const compiledWb = ExprResolver.removeUnExportSheets(result.workbook, compileOptions);
        const compiledBuf = await ExprResolver.toBuffer(compiledWb);

        const filledBuf = await generateXlsxTemplate(compiledBuf, {
            person: { birthday: '1992-05-09T00:00:00.000Z' }
        }, { type: BufferType.NodeBuffer });

        const filledWb = await loadWorkbook(filledBuf as Buffer);
        const filledSheet = filledWb.getWorksheet('Sheet1');
        expect(filledSheet.getCell('A1').value).equal('1992-05-09');
    });
});