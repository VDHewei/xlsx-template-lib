import {describe, it, vi,expect,expectTypeOf,assertType} from 'vitest'
import {BufferType, generateXlsxTemplate} from './core'
import {exceljs,DefaultPlaceholderCellValue,scanCellSetPlaceholder} from './helper';
import {generateCommandsXlsxTemplate, getCommands, CmdFunction, AddCommand, Argument} from './extends'
import * as fs from "node:fs/promises";



describe('generateXlsxTemplate', {tags: ["backend"]}, () => {
    it('should generate a template', async () => {
        const columns = [
            {header: 'Age', key: 'age'},
            {header: 'Name', key: 'name'},
        ]
        const xlsx = await fs.readFile("./test_data/test.xlsx");
        const data = {columns, "name": "test"};
        const buffer = await generateXlsxTemplate(xlsx, data, {type: BufferType.NodeBuffer});
        await fs.writeFile(`./test_data/test_${new Date().valueOf()}.xlsx`, buffer)
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
        await fs.writeFile(`./test_data/test_${new Date().valueOf()}_data.xlsx`, buffer);
        expect(buffer).toBeInstanceOf(Buffer)
        //expect(sheet!.rowCount).toBe(2)
        //expect(sheet!.getRow(2).getCell(1).value).toBe('John')
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
        await fs.writeFile(`./test_data/test_cmd_${new Date().valueOf()}.xlsx`, buffer)
        expect(buffer).toBeInstanceOf(Buffer)
    })

    it("get commands", () => {
        const cmds = getCommands();
        assertType<number>(cmds.size);
        expect(cmds.size).not.equal(0,"empty builtin command")
        for (let [key, cmd] of cmds.entries()) {
            expectTypeOf(key).toEqualTypeOf<string>();
            expectTypeOf(cmd).toEqualTypeOf<CmdFunction>();
        }
    })

    it("add commands", () => {
        const cmds = getCommands();
        let size = cmds.size;
        assertType<number>(size);
        expect(size).not.equal(0,"empty builtin command")
        for (let [key, cmd] of cmds.entries()) {
            expectTypeOf(key).toEqualTypeOf<string>();
            expectTypeOf(cmd).toEqualTypeOf<CmdFunction>();
        }
        AddCommand("test",(values:Object|Record<string, any>,argument: Argument): any|undefined=>{
                return "test";
        });
        AddCommand("hello",(values:Object|Record<string, any>,argument: Argument): any|undefined=>{
            return "hello";
        });
        expect(cmds.size).equal(size+2,"add command size not matched")
        expect(cmds.has("test")).equal(true,"check test command failed")
        expect(cmds.has("hello")).equal(true,"check hello command failed")
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
        await fs.writeFile(`./test_data/test_cmd_${new Date().valueOf()}_data.xlsx`, buffer);
        expect(buffer).toBeInstanceOf(Buffer)
    })
})

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

describe('scanCellSetPlaceholder', () => {
    it('未合并且空单元格时调用 toString', async () => {
        const buffer = await createMockBuffer({ targetValue: null });
        const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
        const spyToString = vi.spyOn(placeholder, 'toString');
        const spyMerge = vi.spyOn(placeholder, 'mergeCell');
        const res = await scanCellSetPlaceholder(buffer, { X: 'B', Y: 2,Sheet:"Sheet1" }, placeholder);
        expect(res).toBe(true);
        expect(spyToString).toHaveBeenCalledOnce();
        expect(spyMerge).not.toHaveBeenCalled();
    });

    it('未合并且非空单元格时不调用任何方法', async () => {
        const buffer = await createMockBuffer({ targetValue: 'Existing Data' });
        const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
        const spyToString = vi.spyOn(placeholder, 'toString');
        const spyMerge = vi.spyOn(placeholder, 'mergeCell');
        const res = await scanCellSetPlaceholder(buffer, { X: 'B', Y: 2 ,Sheet:"Sheet1" }, placeholder);
        expect(res).toBe(true);
        expect(spyToString).not.toHaveBeenCalled();
        expect(spyMerge).not.toHaveBeenCalled();
    });
    it('合并单元格且左侧全空时调用 toString', async () => {
        const buffer = await createMockBuffer({ merged: true, leftValues: [null, null, null] });
        const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
        const spyToString = vi.spyOn(placeholder, 'toString');
        const spyMerge = vi.spyOn(placeholder, 'mergeCell');
        const res = await scanCellSetPlaceholder(buffer, { X: 'B', Y: 2,Sheet:"Sheet1"  }, placeholder);
        expect(res).toBe(true);
        expect(spyToString).toHaveBeenCalledOnce();
        expect(spyMerge).not.toHaveBeenCalled();
    });
    it('合并单元格且左侧非全空时调用 mergeCell 并传入过滤后的数组', async () => {
        const buffer = await createMockBuffer({ merged: true, leftValues: ['Val1', null, 'Val2', ''] });
        const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
        const spyToString = vi.spyOn(placeholder, 'toString');
        const spyMerge = vi.spyOn(placeholder, 'mergeCell');
        const res = await scanCellSetPlaceholder(buffer, { X: 'B', Y: 2,Sheet:"Sheet1" }, placeholder);
        expect(res).toBe(true);
        expect(spyToString).not.toHaveBeenCalled();
        expect(spyMerge).toHaveBeenCalledWith(['Val1', 'Val2']);
    });

    it('支持 base64 字符串入参', async () => {
        const buffer = await createMockBuffer({ targetValue: null });
        const base64Str = buffer.toString('base64');
        const placeholder = new DefaultPlaceholderCellValue('{{P}}', 'M: ?');
        const spyToString = vi.spyOn(placeholder, 'toString');
        const res = await scanCellSetPlaceholder(base64Str, { X: 'B', Y: 2,Sheet:"Sheet1" }, placeholder);
        expect(res).toBe(true);
        expect(spyToString).toHaveBeenCalledOnce();
    });
});