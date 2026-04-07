import * as fs from "node:fs/promises";
import {BufferType, generateXlsxTemplate} from './core'
import {assertType, describe, expect, expectTypeOf, it, Mock, vi} from 'vitest'
import {AddCommand, Argument, CmdFunction, generateCommandsXlsxTemplate, getCommands} from './extends'
import {
    compileWorkSheet,
    DefaultPlaceholderCellValue,
    exceljs,
    loadWorkbook,
    parseWorkSheetRules,
    PlaceholderCellValue,
    RuleResult,
    RuleToken,
    scanCellSetPlaceholder
} from './helper';


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

function testEnv(key: symbol, value: string): boolean {
    const k = key.toString();
    const x = k.substring(7, k.length - 1);
    return process.env[x] === value;
}

const XlsxTest = Symbol(`VITE_SAVE_XLSX_TEST`);
const BackendTest = Symbol(`VITE_SAVA_BACKEND_TEST`);
const CompileTest = Symbol(`VITE_SAVE_COMPILE_XLSX_TEST`);

describe('generateXlsxTemplate', {tags: ["backend"]}, () => {
    it('should generate a template', async () => {
        const columns = [
            {header: 'Age', key: 'age'},
            {header: 'Name', key: 'name'},
        ]
        const xlsx = await fs.readFile("./test_data/test.xlsx");
        const data = {columns, "name": "test"};
        const buffer = await generateXlsxTemplate(xlsx, data, {type: BufferType.NodeBuffer});
        if (testEnv(BackendTest, "true")) {
            await fs.writeFile(`./test_data/test_${new Date().valueOf()}.xlsx`, buffer)
        }
        expect(buffer).toBeInstanceOf(Buffer)
    })

    it('should generate a template with data', async () => {
        const data = await fs.readFile("./test_data/data.json");
        const values = JSON.parse(data.toString('utf-8'));
        const xlsx = await fs.readFile("./test_data/test_data.xlsx");
        values["__alias"] = new Map<string, string>([
            ["#", "exportData.LRR.table"],
            ["T", "template"],
        ]);
        const buffer = await generateXlsxTemplate(xlsx, values, {type: BufferType.NodeBuffer});
        if (testEnv(BackendTest, "true")) {
            await fs.writeFile(`./test_data/test_${new Date().valueOf()}_data.xlsx`, buffer);
        }
        expect(buffer).toBeInstanceOf(Buffer)
    })
})

describe('generateCommandsXlsxTemplate', {tags: ["backend"]}, () => {
    it('should generate a template', async () => {
        const columns = [
            {header: 'Age', key: 'age'},
            {header: 'Name', key: 'name'},
        ]
        const xlsx = await fs.readFile("./test_data/test.xlsx")
        const data = {columns, "name": "test"};
        const buffer = await generateCommandsXlsxTemplate(xlsx, data, {type: BufferType.NodeBuffer})
        if (testEnv(BackendTest, "true")) {
            await fs.writeFile(`./test_data/test_cmd_${new Date().valueOf()}.xlsx`, buffer)
        }
        expect(buffer).toBeInstanceOf(Buffer)
    })

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
        const data = await fs.readFile("./test_data/data.json");
        const values = JSON.parse(data.toString('utf-8'));
        values["__alias"] = new Map<string, string>([
            ["#", "exportData.LRR.table"],
            ["T", "template"],
        ]);
        const xlsx = await fs.readFile("./test_data/test_data.xlsx");
        const buffer = await generateCommandsXlsxTemplate(xlsx, values, {type: BufferType.NodeBuffer});
        if (testEnv(BackendTest, "true")) {
            await fs.writeFile(`./test_data/test_cmd_${new Date().valueOf()}_data.xlsx`, buffer);
        }
        expect(buffer).toBeInstanceOf(Buffer)
    })
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
        const sheetName = `export_metadata.config`;
        const workbook = await loadWorkbook("./test_data/test_data.xlsx");
        const res = parseWorkSheetRules(workbook.getWorksheet(sheetName));
        expectTypeOf<RuleResult>(res);
        expect(res.rules.size).not.equal(0, "输出结果异常")
        expect(res.rules.has(RuleToken.AliasToken)).equal(true, "解析alias规则失败")
        expect(res.rules.has(RuleToken.CellToken)).equal(true, "解析cell规则失败")
        expect(res.rules.has(RuleToken.RowCellToken)).equal(true, "解析rowCell规则失败")
        expect(res.rules.get(RuleToken.RowCellToken).length).not.equal(0, "解析rowCell规则失败")
        expect(res.rules.has(RuleToken.MergeCellToken)).equal(true, "解析mergeCell规则失败")
    });

    it('compile-only', async () => {
        const sheetName = `export_metadata.config`;
        const workbook = "./test_data/test_data.xlsx";
        const res = await compileWorkSheet(workbook, sheetName);
        expectTypeOf<exceljs.Xlsx | Error[]>(res);
        assertType<exceljs.Xlsx>(res as exceljs.Xlsx);
        if (testEnv(CompileTest, "true")) {
            const sv = res as exceljs.Xlsx;
            await sv.writeFile(`./test_data/test_compile_${new Date().valueOf()}.xlsx`);
        }
    });
});