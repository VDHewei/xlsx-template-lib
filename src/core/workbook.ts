import * as path from 'path';
import * as fs from 'fs';
import * as etree from 'elementtree';
import JsZip from "jszip";
import {imageSize as sizeOf} from 'image-size';
import {isArray, parseInt, toString} from "lodash";
import exceljs from "exceljs";

// 从新模块导入
import {
    Placeholder, Ref, Range, SheetInfo, SheetInfoMust, DrawingInfo, TableInfo, RelsInfo,
    FullOptions, OutputByType,
    CustomReplacer, CustomPlaceholderExtractor, BeforeReplaceHook, AfterReplaceHook,
    CustomFormatter, QueryFunction, CellPlaceholder, CellData, PlaceholderCacheResult
} from './types';
import {
    isImageValue,
    toArrayBuffer,
    toDate,
    updateBooleanCell,
    updateFormulaCell,
    updateHyperlinkCell,
    updateImageCell,
    updateRichTextCell
} from './xml-utils';
import {valueDotGet, defaultFormatters, resolveFullDataPath} from './formatters';
import {defaultExtractPlaceholders} from './placeholders';

// ==================== Workbook 类 ====================
/**
 * XLSX 模板工作簿类
 * 负责加载、解析和替换 XLSX 模板中的占位符
 */
class Workbook {
    // ==================== 属性定义 ====================
    /** 用户配置选项 */
    option: FullOptions;
    /** JSZip 归档实例，管理 XLSX 内部所有文件 */
    archive: JsZip;
    /** 共享字符串列表 */
    sharedStrings: string[] = [];
    /** 共享字符串查询映射（字符串 -> 索引） */
    sharedStringsLookup: Record<string, number> = {};
    /** 共享字符串文件路径 */
    sharedStringsPath: string = "";
    /** 所有工作表信息列表 */
    sheets: SheetInfo[] | SheetInfoMust[] = [];
    /** 当前处理的工作表信息 */
    sheet: SheetInfo | SheetInfoMust | null = null;
    /** 工作簿 XML 根元素 */
    workbook: exceljs.Workbook | null = null;

    /** 内容类型 XML 根元素 */
    contentTypes: any = null;
    /** 工作簿文件前缀路径 */
    prefix: string | null = null;

    /** 取值结果缓存 **/
    _cache: Map<string, any | undefined> = new Map<string, any | undefined>();

    // RichData 相关属性
    private richDataIsInit: boolean = false;
    private _relsrichValueRel: any = null;
    private rdrichvalue: any = null;
    private richValueRel: any = null;
    private metadata: any = null;

    /**
     * 单元格值映射表（按工作表名称分组）
     * 外层 key 为工作表名称，内层 key 为单元格引用（如 "E13"），value 为要写入的值
     * 在 XML 结构性修改完成后通过 exceljs 批量写入，保留单元格样式和边框
     * 通过按工作表分组，支持在 substituteAll() 结束时一次性处理所有工作表
     */
    private _cellValueMap: Map<string, Map<string, any>> = new Map();

    // ==================== 构造函数 ====================
    /**
     * 创建工作簿实例
     * @param option - 可选配置项
     */
    constructor(option?: FullOptions) {
        this.option = {
            moveImages: false,
            substituteAllTableRow: false,
            moveSameLineImages: false,
            imageRatio: 100,
            pushDownPageBreakOnTableSubstitution: false,
            useExistingRows: false,
            imageRootPath: null,
            handleImageError: null,
            ...option
        };
    }

    // ==================== parse 静态构造函数 ====================
    /**
     * 从 XLSX buffer 解析模板，返回 Workbook 实例
     * @param data - XLSX 文件的 Buffer
     * @param option - 可选配置项
     * @returns Workbook 实例
     */
    static async parse(data: Buffer | string, option?: FullOptions): Promise<Workbook> {
        const w = new Workbook(option);
        await w.loadTemplate(data);
        return w;
    }

    // ==================== 扩展方法 ====================
    /**
     * 添加自定义替换器
     * @param replacer - 替换函数
     * @returns this（支持链式调用）
     */
    addReplacer(replacer: CustomReplacer): this {
        if (!this.option.replacers) {
            this.option.replacers = [];
        }
        this.option.replacers.push(replacer);
        return this;
    }

    /**
     * 添加自定义格式化器
     * @param formatter - 格式化函数
     * @returns this（支持链式调用）
     */
    addFormatter(formatter: CustomFormatter): this {
        if (!this.option.formatters) {
            this.option.formatters = [];
        }
        this.option.formatters.push(formatter);
        return this;
    }

    /**
     * 设置替换前钩子
     * @param hook - 钩子函数
     * @returns this（支持链式调用）
     */
    setBeforeReplaceHook(hook: BeforeReplaceHook): this {
        this.option.beforeReplace = hook;
        return this;
    }

    /**
     * 设置替换后钩子
     * @param hook - 钩子函数
     * @returns this（支持链式调用）
     */
    setAfterReplaceHook(hook: AfterReplaceHook): this {
        this.option.afterReplace = hook;
        return this;
    }

    /**
     * 设置自定义占位符提取器
     * @param extractor - 提取函数
     * @returns this（支持链式调用）
     */
    setPlaceholderExtractor(extractor: CustomPlaceholderExtractor): this {
        this.option.customPlaceholderExtractor = extractor;
        return this;
    }

    /**
     * 设置自定义模板正则表达式
     * @param regex - 正则表达式
     * @param enableDefaultParsing - 是否同时启用默认解析，默认 false
     * @returns this（支持链式调用）
     */
    setPlaceholderRegex(regex: RegExp, enableDefaultParsing: boolean = false): this {
        this.option.customPlaceholderRegex = regex;
        this.option.enableDefaultParsing = enableDefaultParsing;
        return this;
    }

    /**
     * 设置数值查询器
     * @param h QueryFunction - 设置数值查询器
     * @returns this（支持链式调用）
     */
    setQueryFunctionHandler(h: QueryFunction): this {
        this.option.customQueryFunction = h
        return this;
    }

    // ==================== 扩展点执行方法 ====================
    /**
     * 执行自定义替换器链
     * @param cell - 单元格元素
     * @param stringValue - 字符串值
     * @param placeholder - 占位符信息
     * @param substitution - 替换值
     * @returns 替换结果或 undefined
     */
    private executeReplacers(
        cell: any,
        stringValue: string,
        placeholder: Placeholder,
        substitution: any
    ): string | undefined {
        // 首先执行 customReplacer（向后兼容）
        if (this.option.customReplacer) {
            const result = this.option.customReplacer(cell, stringValue, placeholder, substitution);
            if (result !== undefined && result !== null) {
                return result;
            }
        }
        // 执行 replacers 列表
        if (this.option.replacers && this.option.replacers.length > 0) {
            for (const replacer of this.option.replacers) {
                const result = replacer(cell, stringValue, placeholder, substitution);
                if (result !== undefined && result !== null) {
                    return result;
                }
            }
        }
        return undefined;
    }

    /**
     * 执行格式化器链
     * @param value - 原始值
     * @param placeholder - 占位符信息
     * @param key - 可选键名
     * @returns 格式化后的字符串
     */
    private executeFormatters(value: any, placeholder: Placeholder, key?: string): string {
        // 自定义格式化器
        if (this.option.formatters && this.option.formatters.length > 0) {
            for (const formatter of this.option.formatters) {
                const result = formatter(value, placeholder, key);
                if (result !== undefined && result !== null) {
                    return result;
                }
            }
        }
        // 默认格式化器
        for (const formatter of defaultFormatters) {
            const result = formatter(value, placeholder, key);
            if (result !== undefined && result !== null) {
                return result;
            }
        }
        return "";
    }

    /**
     * 执行替换前钩子
     * @param stringValue - 原始字符串值
     * @param substitutions - 替换数据
     * @returns 处理后的字符串
     */
    private executeBeforeReplaceHook(
        stringValue: string,
        substitutions: Record<string, any>
    ): string {
        if (this.option.beforeReplace) {
            const result = this.option.beforeReplace(stringValue, substitutions);
            if (result !== undefined && result !== null) {
                return result;
            }
        }
        return stringValue;
    }

    /**
     * 执行替换后钩子
     * @param resultString - 替换结果字符串
     * @param stringValue - 原始字符串
     * @param substitutions - 替换数据
     * @returns 最终字符串
     */
    private executeAfterReplaceHook(
        resultString: string,
        stringValue: string,
        substitutions: Record<string, any>
    ): string {
        if (this.option.afterReplace) {
            return this.option.afterReplace(resultString, stringValue, substitutions);
        }
        return resultString;
    }

    /**
     * 从字节数组加载 .xlsx 文件
     * @param data - XLSX 文件的 Buffer 或二进制字符串
     */
    async loadTemplate(data: Buffer | string): Promise<void> {
        if (Buffer.isBuffer(data)) {
            data = data.toString('binary');
        }
        const w = new exceljs.Workbook();
        this.workbook = await w.xlsx.load(data as any);
    }

    /**
     * 使用给定的替换数据对所有工作表进行插值
     * @param substitutions - 替换数据对象，key 为占位符名
     */
    async substituteAll(substitutions: Record<string, any>): Promise<void> {
        const sheets = this.loadSheets(this.workbook);
        for (let sheet of sheets) {
            await this.substitute(sheet.id, substitutions);
        }
    }

    // ==================== substitute 方法 ====================
    /**
     * 使用给定的替换数据对指定工作表进行插值
     * @param sheetName - 工作表名称或编号
     * @param substitutions - 替换数据对象
     */
    async substitute(sheetName: string | number, substitutions: Record<string, any>): Promise<void> {
        const sheet = await this.loadSheet(sheetName);
        this.sheet = sheet;
        // 识别 sheet name 是否带占位符
        const sheetNamePlaceholder = this.extractPlaceholders(sheet.name);
        // 更新表格标题-中的占位符
        sheet.name = this.substituteString(sheet.name, sheetNamePlaceholder, substitutions);
        // 扫描所有单元格 Cell  记录带占位符 Cell[]
        const cellsWithPlaceholders: CellPlaceholder[] = await this.scanSheetPlaceholder(sheet.root);
        // 更新所有单元格 Cell 中的 占位符 为实际内容
        await this.updateCells(cellsWithPlaceholders, substitutions, sheet.root);
    }

    async scanSheetPlaceholder(sheet?: exceljs.Worksheet): Promise<CellPlaceholder[]> {
        const list: CellPlaceholder[] = [];
        const cached = new Map<string, boolean>();
        const worksheet = !sheet ? this.sheet.root : sheet;
        for (const row of worksheet.getRows(1, sheet.rowCount)) {
            const columns = row.cellCount;
            for (let i = 1; i < columns; i++) {
                let cell = row.getCell(i);
                let addr = cell.isMerged ? cell.master.address : cell.address;
                if (cached.has(addr)) {
                    continue;
                }
                let text = !cell.text ? cell.value : cell.text;
                cached.set(addr, true);
                let pls: Placeholder[] | undefined = this.extractPlaceholders(text as string);
                if (!pls || pls.length <= 0) {
                    continue;
                }
                const [_row, _column] = addr.split(":");
                const cellValue: CellPlaceholder = {
                    Row: _row,
                    Value: text,
                    Placeholders: pls,
                    Sheet: worksheet.name,
                    Column: parseInt(_column, 10),
                };
                list.push(cellValue);
            }
        }
        return list;
    }

    private async updateCells(cells: CellPlaceholder[], substitutions: Record<string, any>, worksheet?: exceljs.Worksheet): Promise<void> {
        const sheet = !worksheet ? this.sheet.root : worksheet;
        const sheetName = sheet.name;
        for (const cell of cells) {
            let values: CellData[] = await this.resolverCellValue(cell, substitutions);
            if (!cell.Sheet || cell.Sheet === sheetName) {
                await this.setCells(sheet, values);
            } else {
                let work = this.workbook.getWorksheet(cell.Sheet)
                if (work === undefined) {
                    continue;
                }
                await this.setCells(work, values);
            }
        }
        return;
    }

    async resolverCellValue(cell: CellPlaceholder, substitutions: Record<string, any>): Promise<CellData[]> {
        let items: CellData[] = [];
        let loopIndex = this.getLoopIndex(cell.Placeholders);
        if (loopIndex >= 0) {
            // ${table:xxx.xxx}|${table:xxx.xx2}
            // ${prefix}.${table:xxx.name}
            // ${table:xxx.age}
            items = await this.generateCells(cell, substitutions, loopIndex);
        } else {
            // ${user.name}
            // ${image:xxx.png}
            // ${xxx.logo}
            let result: CellData = {
                Row: cell.Row,
                Column: cell.Column,
                Value: await this.resolverPlaceholders(cell, substitutions),
            };
            items.push(result);
        }
        return items;
    }

    async setCells(workSheet: exceljs.Worksheet, values: CellData[]): Promise<void> {
        for (const cell of values) {
            const cellRef = workSheet.getCell(cell.Row, cell.Column);
            await this.safeUpdate(cellRef, cell.Value,workSheet);
        }
    }

    public async safeUpdate(cellRef: exceljs.Cell, value: any,workSheet:exceljs.Worksheet): Promise<exceljs.CellValue> {
        const style = cellRef.style ? JSON.parse(JSON.stringify(cellRef.style)) : {};
        switch (cellRef.type) {
            case exceljs.ValueType.Null:
                if(isImageValue(value)){
                    await  updateImageCell(cellRef,value,workSheet)
                    return cellRef.value
                }
                cellRef.value = value;
                break;
            case exceljs.ValueType.Date:
                cellRef.value = toDate(value);
                break;
            case exceljs.ValueType.Formula:
                cellRef.value = updateFormulaCell(cellRef.value, value);
                break;
            case exceljs.ValueType.RichText:
                cellRef.value = updateRichTextCell(cellRef.value, value);
                break;
            case exceljs.ValueType.Boolean:
                cellRef.value = updateBooleanCell(cellRef.value, value);
                break;
            case exceljs.ValueType.Hyperlink:
                cellRef.value = updateHyperlinkCell(cellRef.value, value);
                break;
            case exceljs.ValueType.Merge:
            case exceljs.ValueType.String:
                cellRef.value = toString(value);
                break;
            case exceljs.ValueType.SharedString:
                cellRef.value = toString(value);
                break;
            default:
                break;
        }
        cellRef.style = style;
        return cellRef.value;
    }

    private async resolverPlaceholders(cell: CellPlaceholder, substitutions: Record<string, any>): Promise<exceljs.CellValue> {
        let value = cell.Value;
        for (const p of cell.Placeholders) {
            let substitution = this.getSubstitution(p, substitutions);
            value = await this.replaceCellValue(value, p.placeholder, substitution);
        }
        return value;
    }

    private async generateCells(cell: CellPlaceholder, substitutions: Record<string, any>, loopIndex: number): Promise<CellData[]> {

        return [];
    }

    private async replaceCellValue(value: exceljs.CellValue, placeholder: string, newValue: any | undefined): Promise<exceljs.CellValue> {

        return value;
    }

    substituteString(value: string, placeholders: Placeholder[], substitutions: Record<string, any>): string {
        // 循环 placeholders 数组，替换 value 中的占位符
        for (const placeholder of placeholders) {
            const substitution = this.getSubstitution(placeholder, substitutions);
            if (substitution !== undefined) {
                value = value.replace(placeholder.placeholder, substitution);
            }
        }
        return value;
    }

    getSubstitution(placeholder: Placeholder, substitutions: Record<string, any>): any | undefined {
        const {value, exists} = this.getPlaceholderValueByCache(placeholder);
        if (exists) {
            return value;
        }
        const res = this.valueGet(substitutions, placeholder);
        if (placeholder.placeholder !== "") {
            this.setPlaceholderValueByCache(placeholder, res)
        }
        return res;
    }

    private getPlaceholderValueByCache(p: Placeholder): PlaceholderCacheResult {
        if (!this._cache.has(p.placeholder)) {
            return {
                value: undefined,
                exists: false,
            }
        }
        return {
            exists: true,
            value: this._cache.get(p.placeholder),
        }
    }

    private setPlaceholderValueByCache(p: Placeholder, value: any | undefined): void {
        this._cache.set(p.placeholder, value);
    }

    getCache(): Map<string, any | undefined> {
        return this._cache;
    }

    private getLoopIndex(values: Placeholder[]): number {
        for (const [i, v] of values.entries()) {
            if (v.type === "table" || v.subType === "table") {
                return i;
            }
        }
        return -1;
    }

    public destory(): void {
        this._cache = null;
        this.workbook = null;
    }

    /**
     * 处理表格展开时产生的新行（图片移动、跨表单元格复制、排序和推挤）
     * 从 substitute() 中拆分出的子方法
     * @param row - 模板行元素
     * @param rowContext - 行上下文信息
     * @param rows - 所有行的数组
     * @param totalRowsInserted - 已插入的累计行数
     * @param namedTables - 命名表格列表
     * @param rels - 关系文件信息
     * @param sheet - 当前工作表信息
     * @returns 更新后的总插入行数
     */
    private async _processNewTableRows(
        row: any,
        rowContext: { newTableRows: any[]; cellsSubstituteTable: any[]; drawing: DrawingInfo | null },
        rows: any[],
        totalRowsInserted: number,
        namedTables: TableInfo[],
        rels: RelsInfo | null,
        sheet: SheetInfoMust
    ): Promise<number> {
        let updatedInserted = totalRowsInserted;
        // 处理图片移动
        if (this.option["moveImages"] && rels) {
            if (rowContext.drawing == null) {
                rowContext.drawing = await this.loadDrawing(sheet.root, sheet.filename, rels.root);
            }
            if (rowContext.drawing != null) {
                this.moveAllImages(rowContext.drawing, row.attrib.r, rowContext.newTableRows.length);
            }
        }
        // 复制非表格单元格到新行
        const cellsOverTable = row.findall("c").filter(
            (cell: any) => !rowContext.cellsSubstituteTable.includes(cell)
        );
        rowContext.newTableRows.forEach((newRow: any) => {
            if (this.option && this.option.substituteAllTableRow) {
                cellsOverTable.forEach((cellOverTable: any) => {
                    const newCell = this.cloneElement(cellOverTable);
                    newCell.attrib.r = this.joinRef({
                        row: newRow.attrib.r,
                        col: this.splitRef(newCell.attrib.r).col
                    });
                    newRow.append(newCell);
                });
                const newSortRow = newRow.findall("c").sort((a: any, b: any) => {
                    const colA = this.splitRef(a.attrib.r).col;
                    const colB = this.splitRef(b.attrib.r).col;
                    return this.charToNum(colA) - this.charToNum(colB);
                });
                this.replaceChildren(newRow, newSortRow);
            }
            rows.push(newRow);
            ++updatedInserted;
        });
        // 推挤后续行和合并单元格
        this.pushDown(this.workbook, sheet.root, namedTables, parseInt(row.attrib.r, 10), rowContext.newTableRows.length);
        return updatedInserted;
    }

    /**
     * 处理 useExistingRows 模式下溢出的表格行，插入到最后一个被消费的行之后
     * 从 substitute() 中拆分出的子方法
     * @param overflowTableRows - 溢出的表格行数组
     * @param rows - 所有行的数组
     * @param totalRowsInserted - 已插入的累计行数
     * @param namedTables - 命名表格列表
     * @param sheet - 当前工作表信息
     * @param consumedRowNumbers - 已消费的行号集合
     * @param currentRow - 当前处理的行号
     * @returns 更新后的总插入行数
     */
    private _insertOverflowRows(
        overflowTableRows: any[],
        rows: any[],
        totalRowsInserted: number,
        namedTables: TableInfo[],
        sheet: SheetInfoMust,
        consumedRowNumbers: Set<number>,
        currentRow: number | null
    ): number {
        const overflowCount = overflowTableRows.length;
        let insertAfterRow = 0;
        if (consumedRowNumbers.size > 0) {
            consumedRowNumbers.forEach(n => {
                if (n > insertAfterRow) insertAfterRow = n;
            });
        } else if (currentRow) {
            insertAfterRow = currentRow;
        }
        const sheetCellMap = this._cellValueMap.get(sheet.name);
        // 更新溢出行及其单元格的引用编号，同时同步 _cellValueMap 中的键
        for (let i = 0; i < overflowTableRows.length; i++) {
            const overflowRow = overflowTableRows[i];
            const overflowRowNum = insertAfterRow + 1 + i;
            overflowRow.attrib.r = `${overflowRowNum}`;
            overflowRow.findall("c").forEach((c: any) => {
                const oldRef = c.attrib.r;
                const newRef = this.joinRef({
                    row: overflowRowNum,
                    col: this.splitRef(oldRef).col
                });
                c.attrib.r = newRef;
                // 同步更新 _cellValueMap 中的键（溢出行在 _createNewTableRow
                // 中使用临时行号创建，需要更新为正确的最终引用）
                if (sheetCellMap && sheetCellMap.has(oldRef)) {
                    const val = sheetCellMap.get(oldRef);
                    sheetCellMap.delete(oldRef);
                    sheetCellMap.set(newRef, val);
                }
            });
        }
        // 在正确位置插入溢出行
        let insertIdx = rows.findIndex(r => parseInt(r.attrib.r, 10) === insertAfterRow);
        if (insertIdx < 0) {
            insertIdx = rows.length - 1;
        }
        rows.splice(insertIdx + 1, 0, ...overflowTableRows);
        // 更新后续行的编号，同时同步 _cellValueMap 中的键
        for (let i = insertIdx + 1 + overflowCount; i < rows.length; i++) {
            const r = rows[i];
            const oldNum = parseInt(r.attrib.r, 10);
            const newNum = oldNum + overflowCount;
            r.attrib.r = `${newNum}`;
            r.findall("c").forEach((c: any) => {
                const oldRef = c.attrib.r;
                const newRef = this.joinRef({
                    row: newNum,
                    col: this.splitRef(oldRef).col
                });
                c.attrib.r = newRef;
                // 同步更新 _cellValueMap（后续行号被推挤后需要更新为新的引用）
                if (sheetCellMap && sheetCellMap.has(oldRef)) {
                    const val = sheetCellMap.get(oldRef);
                    sheetCellMap.delete(oldRef);
                    sheetCellMap.set(newRef, val);
                }
            });
        }
        this.pushDown(this.workbook, sheet.root, namedTables, insertAfterRow, overflowCount);
        return totalRowsInserted + overflowCount;
    }

    /**
     * 处理单个单元格 - 提取占位符并执行替换
     * 从 substitute() 中拆出的子方法
     * @param cell - 单元格元素
     * @param row - 所在行元素
     * @param rowContext - 行上下文信息
     * @param substitutions - 替换数据
     * @param namedTables - 命名表格列表
     * @param rels - 关系文件信息
     * @param originalRowsByNum - 原始行号映射（useExistingRows 模式）
     * @param consumedRowNumbers - 已消费的行号集合
     * @param overflowTableRows - 溢出的表格行
     * @returns 返回插入的列数
     */
    private async _processSingleCell(
        cell: any,
        row: any,
        rowContext: {
            cells: any[];
            cellsInserted: number;
            newTableRows: any[];
            cellsSubstituteTable: any[];
            currentRow: number;
            drawing: DrawingInfo | null
        },
        substitutions: Record<string, any>,
        namedTables: TableInfo[],
        rels: RelsInfo | null,
        originalRowsByNum?: Map<number, any>,
        consumedRowNumbers?: Set<number>,
        overflowTableRows?: any[]
    ): Promise<{ cellsInserted: number }> {
        let appendCell = true;
        let cellsInserted = rowContext.cellsInserted;
        let drawing = rowContext.drawing;
        cell.attrib.r = this.getCurrentCell(cell, rowContext.currentRow, cellsInserted);
        // 如果是字符串列，查找共享字符串
        if (cell.attrib.t === "s") {
            const cellValue = cell.find("v");
            const stringIndex = parseInt(cellValue.text.toString(), 10);
            let strValue = this.sharedStrings[stringIndex];
            if (strValue === undefined) {
                return {cellsInserted};
            }
            // 执行替换前钩子
            strValue = this.executeBeforeReplaceHook(strValue, substitutions);
            // 遍历占位符
            for (let placeholder of this.extractPlaceholders(strValue)) {
                let newCellsInserted = 0;
                let substitution = this.valueGet(substitutions, placeholder);
                // 尝试执行自定义替换器
                const customResult = this.executeReplacers(cell, strValue, placeholder, substitution);
                if (customResult !== undefined) {
                    strValue = customResult;
                    rowContext.cells.push(cell);
                    return {cellsInserted: 0};
                }
                if (placeholder.full && placeholder.type === "table") {
                    // Reconstruct the correct data path from the raw placeholder
                    // The regex splits at the first dot, so placeholder.name/key are incomplete
                    const fullPath = resolveFullDataPath(placeholder);
                    const lastDotIdx = fullPath.lastIndexOf('.');
                    const tableName = lastDotIdx >= 0 ? fullPath.substring(0, lastDotIdx) : fullPath;
                    const tableKey = lastDotIdx >= 0 ? fullPath.substring(lastDotIdx + 1) : '';
                    const tableData = valueDotGet(substitutions, tableName, [], '');
                    if (substitution instanceof Array || isArray(tableData)) {
                        if (placeholder.subType === 'image' && drawing == null) {
                            if (rels) {
                                drawing = await this.loadDrawing(this.sheet!.root, this.sheet!.filename, rels.root);
                            } else {
                                console.log("Need to implement initRels. Or init this with Excel");
                            }
                        }
                        rowContext.cellsSubstituteTable.push(cell);
                        const correctedPlaceholder = {
                            ...placeholder,
                            name: tableName,
                            key: tableKey
                        };
                        const useExisting = this.option.useExistingRows;
                        //newCellsInserted = await this._handleTableSubstitution(
                        //    row, rowContext.newTableRows, rowContext.cells, cell, namedTables,
                        //    tableData, correctedPlaceholder, drawing,
                        //    useExisting ? originalRowsByNum : undefined,
                        //    useExisting ? consumedRowNumbers : undefined,
                        //    useExisting ? overflowTableRows : undefined
                        //);
                        if (newCellsInserted !== 0 || tableData.length) {
                            if (tableData.length === 1) {
                                appendCell = true;
                            }
                            if (tableData[0] && tableData[0][tableKey] instanceof Array) {
                                appendCell = false;
                            }
                        }
                        if (newCellsInserted !== 0) {
                            cellsInserted += newCellsInserted;
                            this.pushRight(this.workbook, this.sheet!.root, cell.attrib.r, newCellsInserted);
                        }
                    }
                }
                if (placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                    appendCell = false;
                    newCellsInserted = this.substituteArray(rowContext.cells, cell, substitution);
                    if (newCellsInserted !== 0) {
                        cellsInserted += newCellsInserted;
                        this.pushRight(this.workbook, this.sheet!.root, cell.attrib.r, newCellsInserted);
                    }
                }
                if (placeholder.type === "image" && placeholder.full) {
                    if (rels != null) {
                        if (drawing == null) {
                            drawing = await this.loadDrawing(this.sheet!.root, this.sheet!.filename, rels.root);
                        }
                        this.substituteImage(cell, strValue, placeholder, substitution, drawing);
                    } else {
                        console.log("Need to implement initRels. Or init this with Excel");
                    }
                }
                if (placeholder.type === "imageincell" && placeholder.full) {
                    // Reconstruct full data path from raw placeholder string
                    // since regex only captures the first dot-separated segment
                    const fullPath = resolveFullDataPath(placeholder);
                    const imageValue = substitution !== undefined && substitution !== null && typeof substitution === 'string' && substitution.length > 0
                        ? substitution
                        : valueDotGet(substitutions, fullPath, '', placeholder.type);
                    await this.substituteImageInCell(cell, imageValue);
                    // 将 richValue 关系添加到工作表关系文件（sheet rels），
                    // 使 Excel 能够通过 rdrichvalue.xml 找到嵌入图片数据
                    if (rels) {
                        const maxId = this.findMaxId(rels.root, 'Relationship', 'Id', /rId(\d*)/);
                        const rel = etree.SubElement(rels.root, 'Relationship');
                        rel.set('Id', 'rId' + maxId);
                        rel.set('Type', 'http://schemas.microsoft.com/office/2017/06/relationships/richValue');
                        rel.set('Target', '../richData/rdrichvalue.xml');
                    }
                } else if (placeholder.type === "table" && placeholder.full) {
                    // Table substitution already handled above by substituteTable()
                } else {
                    if (placeholder.key) {
                        substitution = this.valueGet(substitutions, placeholder, true);
                    }
                    strValue = this.substituteScalar(cell, strValue, placeholder, substitution);
                }
            }
            // 执行替换后钩子
            strValue = this.executeAfterReplaceHook(strValue, strValue, substitutions);
        }
        if (appendCell) {
            rowContext.cells.push(cell);
        }
        rowContext.drawing = drawing;
        return {cellsInserted};
    }


    /**
     * 处理表格数据的第一行：直接填充模板行中的单元格
     * @param cells - 单元格数组
     * @param cell - 当前单元格元素
     * @param value - 要填充的值
     * @param placeholder - 占位符信息
     * @param drawing - 绘图信息
     * @returns 插入的列数
     */
    private async _fillFirstTableRow(
        cells: any[],
        cell: any,
        value: any,
        placeholder: Placeholder,
        drawing: DrawingInfo | null
    ): Promise<number> {
        if (value instanceof Array) {
            return this.substituteArray(cells, cell, value);
        } else if (placeholder.subType === 'image' && value !== "") {
            this.substituteImage(cell, placeholder.placeholder, placeholder, value, drawing);
        } else if (placeholder.subType === "imageincell" && value !== "") {
            await this.substituteImageInCell(cell, value);
        } else {
            const customResult = this.executeReplacers(cell, '', placeholder, value);
            this.recordCellValue(cell, customResult !== undefined ? customResult : value);
        }
        return 0;
    }


    /**
     * 创建新的表格行并填充单元格值
     */
    private async _createNewTableRow(
        row: any,
        cell: any,
        idx: number,
        value: any,
        placeholder: Placeholder,
        drawing: DrawingInfo | null,
        parentTables: TableInfo[],
        templateDataRows: any[],
        colRef: string,
        currentRowNum: number,
        overflowTableRows?: any[],
        newTableRows?: any[]
    ): Promise<void> {
        let newRow: any;
        let newCell: any;
        const newCells: any[] = [];
        let newCellsInsertedOnNewRow = 0;

        // 获取或创建新行容器
        if (overflowTableRows) {
            newRow = this.cloneElement(row, false);
            // 使用高位临时行号（10000+i），避免与已消费行的行号冲突。
            // _insertOverflowRows 后续会将这些临时行号重编号为正确值，
            // 并同步更新 _cellValueMap 中的键。
            newRow.attrib.r = `${10000 + overflowTableRows.length}`;
            overflowTableRows.push(newRow);
        } else if (newTableRows) {
            if ((idx - 1) < newTableRows.length) {
                newRow = newTableRows[idx - 1];
            } else {
                newRow = this.cloneElement(row, false);
                newRow.attrib.r = this.getCurrentRow(row, newTableRows.length + 1);
                newTableRows.push(newRow);
            }
        }

        newCell = this.cloneElement(cell);
        // 应用数据行样式
        //const dataStyle = this._getDataRowStyle(colRef, templateDataRows, currentRowNum + idx);
        //if (dataStyle !== undefined) newCell.attrib.s = dataStyle;
        newCell.attrib.r = this.joinRef({
            row: newRow.attrib.r,
            col: this.splitRef(newCell.attrib.r).col
        });

        if (value instanceof Array) {
            newCellsInsertedOnNewRow = this.substituteArray(newCells, newCell, value);
            newCells.forEach((nc: any) => newRow.append(nc));
            if (newCellsInsertedOnNewRow) this.updateRowSpan(newRow, newCellsInsertedOnNewRow);
        } else if (placeholder.subType === 'image' && value !== '') {
            this.substituteImage(newCell, placeholder.placeholder, placeholder, value, drawing);
        } else if (placeholder.subType === "imageincell" && value !== "") {
            await this.substituteImageInCell(newCell, value);
            newRow.append(newCell);
        } else {
            await this._setCellValue(newCell, value, placeholder, drawing);
            newRow.append(newCell);
        }
        // this._handleMergeCellForNewRow(newCell, newRow, colRef, templateDataRows, currentRowNum, idx);
        // 扩展命名表范围
        this._expandTableRange(parentTables, newCell.attrib.r);
    }

    /**
     * 设置单元格的值（自定义替换器优先，否则直接记录值）
     * 注意：数组类型的值由调用方单独处理（需要自行管理行追加）
     */
    private async _setCellValue(cell: any, value: any, placeholder: Placeholder, drawing: DrawingInfo | null): Promise<void> {
        if (placeholder.subType === 'image' && value !== '') {
            this.substituteImage(cell, placeholder.placeholder, placeholder, value, drawing);
        } else if (placeholder.subType === 'imageincell' && value !== '') {
            await this.substituteImageInCell(cell, value);
        } else {
            const customResult = this.executeReplacers(cell, '', placeholder, value);
            this.recordCellValue(cell, customResult !== undefined ? customResult : value);
        }
    }

    /**
     * 扩展命名表的范围，将新行纳入表格区域
     * @param parentTables - 父命名表列表
     * @param cellRef - 新单元格的引用
     */
    private _expandTableRange(parentTables: TableInfo[], cellRef: string): void {
        parentTables.forEach((namedTable) => {
            const tableRoot = namedTable.root;
            const autoFilter = tableRoot.find("autoFilter");
            const range = this.splitRange(tableRoot.attrib.ref);
            if (!this.isWithin(cellRef, range.start, range.end)) {
                range.end = this.nextRow(range.end);
                tableRoot.attrib.ref = this.joinRange(range);
                if (autoFilter !== null) {
                    autoFilter.attrib.ref = tableRoot.attrib.ref;
                }
            }
        });
    }

    /**
     * 生成新的二进制 .xlsx 文件
     * @param options - JSZip 生成选项
     * @returns 生成的输出数据
     */
    async generate<T extends JsZip.OutputType>(options?: JsZip.JSZipGeneratorOptions<T>): Promise<OutputByType[T]> {
        return await this.archive.generateAsync(options);
    }

    /**
     * 查询占位符对应的数据值
     * @param substitutions - 数据对象
     * @param p - 占位符
     * @param full - 是否使用完整 key
     * @returns 查询到的值
     */
    public valueGet(substitutions: object | Record<string, any>, p: Placeholder, full?: boolean): any {
        if (this.option.customQueryFunction === undefined) {
            if ((full !== undefined && typeof full === "boolean" && full && p.key) || (p.full && p.key)) {
                return valueDotGet(substitutions, p.name + '.' + p.key, p.default || (p.type === 'table' ? [] : ''), p.type);
            }
            return valueDotGet(substitutions, p.name, p.default || (p.type === 'table' ? [] : ''), p.type)
        }
        if (full !== undefined && typeof full === "boolean" && full &&
            p.key && !p.name.endsWith(`.${p.key}`)) {
            p.name = p.name + '.' + p.key
        }
        return this.option.customQueryFunction(substitutions, p)
    }

    // ==================== 单元格值处理（exceljs 集成）====================

    /**
     * 通过 XML 级操作设置公式值
     * 公式必须用 XML 处理，因为 exceljs 在加载时会对公式进行求值，可能导致计算错误
     * @param cell - XML cell 元素
     * @param substitution - 以 "=" 开头的公式字符串
     * @returns 公式文本
     */
    private _xmlInsertFormula(cell: any, substitution: any): string {
        const cellValue = cell.find("v");
        const formula = etree.Element("f");
        formula.text = substitution.substring(1);
        cell.insert(1, formula);
        delete cell.attrib.t;
        if (cellValue) {
            cellValue.text = '';
        }
        return substitution.substring(1);
    }

    /**
     * 记录单元格值到映射表，稍后通过 exceljs 批量写入（保留样式和边框）
     * 如果是公式（以 "=" 开头），则转为 XML 级处理
     * @param cell - XML 单元格元素
     * @param substitution - 要写入的值
     * @returns 字符串化后的值
     */

    /**
     * 记录单元格值到映射表，稍后通过 exceljs 批量写入（保留样式和边框）
     * 如果是公式（以 "=" 开头），则转为 XML 级处理
     * 同时清除 t 属性，让 exceljs 自行推断单元格类型
     * @param cell - XML 单元格元素
     * @param substitution - 要写入的值
     * @returns 字符串化后的值
     */
    /**
     * 记录单元格值，稍后通过 exceljs 批量写入（保留样式和边框）
     * 按工作表名称分组存储，支持 substituteAll() 结尾一次性处理所有工作表
     * 如果是公式（以 "=" 开头），则转为 XML 级处理
     * @param cell - XML 单元格元素
     * @param substitution - 要写入的值
     * @returns 字符串化后的值
     */
    private recordCellValue(cell: any, substitution: any): string {
        const stringify = this.stringify(substitution);
        if (typeof substitution === 'string' && substitution[0] === '=') {
            // 公式单元格：继续保持 XML 级处理
            return this._xmlInsertFormula(cell, substitution);
        }
        const cellRef = cell.attrib.r;
        // 按当前工作表名称分组存储
        const sheetName = this.sheet!.name;
        if (!this._cellValueMap.has(sheetName)) {
            this._cellValueMap.set(sheetName, new Map());
        }
        this._cellValueMap.get(sheetName)!.set(cellRef, substitution);
        // 清除 t 属性（如 "s" 共享字符串类型标记），让 exceljs 自行推断类型
        delete cell.attrib.t;
        return stringify;
    }


    /**
     * 将任意类型的值转换为字符串
     * 支持扩展：使用自定义格式化器
     * @param value - 任意类型的值
     * @param placeholder - 可选的占位符信息
     * @param key - 可选的键名
     * @returns 字符串表示
     */
    stringify(value: any, placeholder?: Placeholder, key?: string): string {
        // 如果提供了占位符信息，使用格式化器链
        if (placeholder) {
            return this.executeFormatters(value, placeholder, key);
        }
        // 默认行为（向后兼容）
        if (value instanceof Date) {
            return Number((value.getTime() / (1000 * 60 * 60 * 24)) + 25569).toString();
        } else if (typeof (value) === "number" || typeof (value) === "boolean") {
            return Number(value).toString();
        } else if (typeof (value) === "string") {
            return String(value).toString();
        }
        return "";
    }

    // ==================== 辅助方法 ====================

    /**
     * 将共享字符串列表写回到 archive 中的 sharedStrings.xml 文件
     * 遍历所有字符串，重建 XML 结构（si -> t 元素树）
     */
    private async writeSharedStrings(): Promise<void> {
        const content = await this.archive.file(this.sharedStringsPath).async("string");
        const root = etree.parse(content).getroot();
        const children = root.getchildren();
        root.delSlice(0, children.length);
        this.sharedStrings.forEach((string) => {
            const si = etree.Element("si");
            const t = etree.Element("t");
            t.text = string;
            si.append(t);
            root.append(si);
        });
        root.attrib.count = `${this.sharedStrings.length}`;
        root.attrib.uniqueCount = `${this.sharedStrings.length}`;
        this.archive.file(this.sharedStringsPath, etree.tostring(root, {encoding: 'utf-8'}));
    }

    /**
     * 添加新字符串到共享字符串列表，返回其索引
     * @param s - 要添加的字符串
     * @returns 新字符串的索引位置
     */
    private addSharedString(s: string): number {
        const idx = this.sharedStrings.length;
        this.sharedStrings.push(s);
        this.sharedStringsLookup[s] = idx;
        return idx;
    }

    /**
     * 获取字符串在共享字符串列表中的索引，不存在则自动添加
     * @param s - 要查询的字符串
     * @returns 字符串索引
     */
    private stringIndex(s: string): number {
        let idx = this.sharedStringsLookup[s];
        if (idx === undefined) {
            idx = this.addSharedString(s);
        }
        return idx;
    }

    /**
     * 替换共享字符串列表中的字符串，若旧字符串不存在则新增
     * @param oldString - 要替换的旧字符串
     * @param newString - 新字符串
     * @returns 字符串在列表中的索引
     */
    private replaceString(oldString: string, newString: string): number {
        let idx = this.sharedStringsLookup[oldString];
        if (idx === undefined) {
            idx = this.addSharedString(newString);
        } else {
            this.sharedStrings[idx] = newString;
            delete this.sharedStringsLookup[oldString];
            this.sharedStringsLookup[newString] = idx;
        }
        return idx;
    }

    /**
     * 从工作簿 XML 中加载所有工作表信息列表
     * @param workbook - 工作簿
     * @returns 工作表信息数组
     */
    private loadSheets(workbook: exceljs.Workbook): SheetInfo[] {
        const sheets: SheetInfo[] = [];
        for (const sheet of workbook.worksheets) {
            sheets.push({
                root: sheet,
                id: sheet.id,
                name: sheet.name,
            });
        }
        return sheets;
    }

    /**
     * 加载指定工作表的内容 XML，返回包含根元素的完整信息
     * @param sheet - 工作表名称或编号
     * @returns 包含 XML 根元素的工作表信息
     */
    async loadSheet(sheet: string | number): Promise<SheetInfoMust> {
        const s = this.workbook.getWorksheet(sheet)
        if (!s) {
            throw new Error(`${sheet} not exists`);
        }
        return {
            filename: s.name,
            name: s.name,
            id: s.id,
            root: s,
        };
    }

    /**
     * 加载工作表的关系文件（_rels/sheetN.xml.rels），不存在则初始化新的关系文件
     * @param sheetFilename - 工作表 XML 文件路径
     * @returns 关系文件信息
     */
    async loadSheetRels(sheetFilename: string): Promise<RelsInfo> {
        const sheetDirectory = path.dirname(sheetFilename);
        const sheetName = path.basename(sheetFilename);
        const relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/');
        const relsFile = this.archive.file(relsFilename);
        if (relsFile === null) {
            return this.initSheetRels(sheetFilename);
        }
        const content = await relsFile.async("string");
        return {
            filename: relsFilename,
            root: etree.parse(content).getroot()
        };
    }

    /**
     * 初始化工作表的关系文件（_rels）空结构
     * @param sheetFilename - 工作表文件名
     * @returns 新建的关系文件信息
     */
    private initSheetRels(sheetFilename: string): RelsInfo {
        const sheetDirectory = path.dirname(sheetFilename);
        const sheetName = path.basename(sheetFilename);
        const relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/');
        const ElementTree = etree.ElementTree;
        const root = etree.Element('Relationships');
        root.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        const relsEtree = new ElementTree(root);
        return {
            filename: relsFilename,
            root: relsEtree.getroot()
        };
    }

    /**
     * 加载工作表的绘制信息（drawing.xml），用于图片位置管理
     * @param sheet - 工作表 XML 根元素
     * @param sheetFilename - 工作表文件名
     * @param rels - 关系根元素
     * @returns 绘制信息对象
     */
    private async loadDrawing(sheet: any, sheetFilename: string, rels: any): Promise<DrawingInfo> {
        // const sheetName = path.basename(sheetFilename);
        const sheetDirectory = path.dirname(sheetFilename);
        const drawing: DrawingInfo = {filename: '', root: null};
        const drawingPart = sheet.find("drawing");
        if (drawingPart === null) {
            return this.initDrawing(sheet, rels);
        }
        const relationshipId = drawingPart.attrib['r:id'];
        const target = rels.find(`Relationship[@Id='${relationshipId}']`).attrib.Target;
        const drawingFilename = path.join(sheetDirectory, target).replace(/\\/g, '/');
        const drawContent = await this.archive.file(drawingFilename).async("string");
        const drawingTree = etree.parse(drawContent);
        drawing.filename = drawingFilename;
        drawing.root = drawingTree.getroot();
        drawing.relFilename = path.dirname(drawingFilename) + '/_rels/' + path.basename(drawingFilename) + '.rels';
        const relFile = this.archive.file(drawing.relFilename);
        if (relFile === null) {
            drawing.relRoot = etree.Element('Relationships');
            drawing.relRoot.set('xmlns', "http://schemas.openxmlformats.org/package/2006/relationships");
        } else {
            const relContent = await relFile.async("string");
            drawing.relRoot = etree.parse(relContent).getroot();
        }
        return drawing;
    }

    /**
     * 注册新的 ContentType 到 [Content_Types].xml
     * @param partName - 资源路径（如 /xl/worksheets/sheet1.xml）
     * @param contentType - 内容类型字符串
     */
    private addContentType(partName: string, contentType: string): void {
        etree.SubElement(this.contentTypes, 'Override', {'ContentType': contentType, 'PartName': partName});
    }

    /**
     * 初始化绘制信息结构，用于图片占位
     * 创建新的 drawing.xml 和对应关系
     * @param sheet - 工作表根元素
     * @param rels - 关系根元素
     * @returns 新建的绘制信息
     */
    private initDrawing(sheet: any, rels: any): DrawingInfo {
        const maxId = this.findMaxId(rels, 'Relationship', 'Id', /rId(\d*)/);
        const rel = etree.SubElement(rels, 'Relationship');
        sheet.insert(sheet._children.length, etree.Element('drawing', {'r:id': 'rId' + maxId}));
        rel.set('Id', 'rId' + maxId);
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing');
        const drawing: DrawingInfo = {} as DrawingInfo;
        const drawingFilename = 'drawing' + this.findMaxFileId(/xl\/drawings\/drawing\d*\.xml/, /drawing(\d*)\.xml/) + '.xml';
        rel.set('Target', '../drawings/' + drawingFilename);
        drawing.root = etree.Element('xdr:wsDr');
        drawing.root.set('xmlns:xdr', "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        drawing.root.set('xmlns:a', "http://schemas.openxmlformats.org/drawingml/2006/main");
        drawing.filename = 'xl/drawings/' + drawingFilename;
        drawing.relFilename = 'xl/drawings/_rels/' + drawingFilename + '.rels';
        drawing.relRoot = etree.Element('Relationships');
        drawing.relRoot.set('xmlns', "http://schemas.openxmlformats.org/package/2006/relationships");
        this.addContentType('/' + drawing.filename, 'application/vnd.openxmlformats-officedocument.drawing+xml');
        return drawing;
    }

    /**
     * 将绘制信息写回到 archive 中
     * @param drawing - 绘制信息（可为 null）
     */
    private writeDrawing(drawing: DrawingInfo | null): void {
        if (drawing !== null) {
            this.archive.file(drawing.filename, etree.tostring(drawing.root, {encoding: "utf-8"}));
            this.archive.file(drawing.relFilename, etree.tostring(drawing.relRoot, {encoding: "utf-8"}));
        }
    }

    /**
     * 移动绘制中所有图片的位置（表格展开时下移）
     * @param drawing - 绘制信息
     * @param fromRow - 起始行号字符串
     * @param nbRow - 要移动的行数
     */
    private moveAllImages(drawing: DrawingInfo, fromRow: string, nbRow: number): void {
        drawing.root.getchildren().forEach((drawElement: any) => {
            if (drawElement.tag == "xdr:twoCellAnchor") {
                this._moveTwoCellAnchor(drawElement, fromRow, nbRow);
            }
        });
    }

    private _moveTwoCellAnchor(drawingElement: any, fromRow: string, nbRow: number): void {
        const _moveImage = (drawingElement: any, fromRow: string, nbRow: number | string) => {
            let num: number;
            if (typeof nbRow === "string") {
                num = Number.parseInt(nbRow, 10);
            } else {
                num = nbRow;
            }
            drawingElement.find('xdr:from').find('xdr:row').text = Number.parseInt(drawingElement.find('xdr:from').find('xdr:row').text, 10) + num;
            drawingElement.find('xdr:to').find('xdr:row').text = Number.parseInt(drawingElement.find('xdr:to').find('xdr:row').text, 10) + num;
        };
        if (this.option["moveSameLineImages"]) {
            if (parseInt(drawingElement.find('xdr:from').find('xdr:row').text) + 1 >= parseInt(fromRow)) {
                _moveImage(drawingElement, fromRow, nbRow);
            }
        } else {
            if (parseInt(drawingElement.find('xdr:from').find('xdr:row').text) + 1 > parseInt(fromRow)) {
                _moveImage(drawingElement, fromRow, nbRow);
            }
        }
    }

    /**
     * 加载工作表中所有命名表（tableParts），返回完整的 XML 树
     * @param sheet - 工作表 XML 根元素
     * @param sheetFilename - 工作表文件名
     * @returns 命名表信息数组
     */
    async loadTables(sheet: etree.Element, sheetFilename: string): Promise<TableInfo[]> {
        const sheetDirectory = path.dirname(sheetFilename);
        const sheetName = path.basename(sheetFilename);
        const relsFilename = sheetDirectory + "/" + '_rels' + "/" + sheetName + '.rels';
        const relsFile = this.archive.file(relsFilename);
        const tables: TableInfo[] = [];
        if (relsFile === null) {
            return tables;
        }
        const relsContent = await relsFile.async("string");
        const rels = etree.parse(relsContent).getroot();
        for (let tablePart of sheet.findall("tableParts/tablePart")) {
            const relationshipId = tablePart.attrib['r:id'];
            const target = rels.find(`Relationship[@Id='${relationshipId}']`).attrib.Target;
            const tableFilename = target.replace('..', this.prefix!);
            const content = await this.archive.file(tableFilename).async("string");
            const tableTree = etree.parse(content);
            tables.push({
                filename: tableFilename,
                root: tableTree.getroot()
            });
        }
        return tables;
    }

    /**
     * 将命名表 XML 写回到 archive 中
     * @param tables - 命名表信息数组
     */
    private writeTables(tables: TableInfo[]): void {
        tables.forEach((namedTable) => {
            this.archive.file(namedTable.filename, etree.tostring(namedTable.root, {encoding: 'utf-8'}));
        });
    }

    /**
     * 替换工作表中所有超链接目标中的占位符
     * @param rels - 关系文件信息
     * @param substitutions - 替换数据
     */
    private substituteHyperlinks(rels: RelsInfo | null, substitutions: Record<string, any>): void {
        if (rels === null) {
            return;
        }
        const relationships = rels.root._children;
        relationships.forEach((relationship: any) => {
            if (relationship.attrib.Type === HYPERLINK_RELATIONSHIP) {
                let target = relationship.attrib.Target;
                target = decodeURI(decodeURI(target));
                this.extractPlaceholders(target).forEach((placeholder) => {
                    const substitution = substitutions[placeholder.name];
                    if (substitution === undefined) {
                        return;
                    }
                    target = target.replace(placeholder.placeholder, this.stringify(substitution));
                    relationship.attrib.Target = encodeURI(target);
                });
            }
        });
    }

    /**
     * 替换命名表列标题中的占位符，支持展开数组为多列
     * @param tables - 命名表列表
     * @param substitutions - 替换数据
     */
    private substituteTableColumnHeaders(tables: TableInfo[], substitutions: Record<string, any>): void {
        tables.forEach((table) => {
            const root = table.root;
            const columns = root.find("tableColumns");
            const autoFilter = root.find("autoFilter");
            const tableRange = this.splitRange(root.attrib.ref);
            let idx = 0;
            let inserted = 0;
            const newColumns: any[] = [];
            columns.findall("tableColumn").forEach((col: any) => {
                ++idx;
                col.attrib.id = Number(idx).toString();
                newColumns.push(col);
                let name = col.attrib.name;
                this.extractPlaceholders(name).forEach((placeholder) => {
                    const substitution = substitutions[placeholder.name];
                    if (substitution === undefined) {
                        return;
                    }
                    if (placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                        substitution.forEach((element: any, i: number) => {
                            let newCol = col;
                            if (i > 0) {
                                newCol = this.cloneElement(newCol);
                                newCol.attrib.id = Number(++idx).toString();
                                newColumns.push(newCol);
                                ++inserted;
                                tableRange.end = this.nextCol(tableRange.end);
                            }
                            newCol.attrib.name = this.stringify(element);
                        });
                    } else {
                        name = name.replace(placeholder.placeholder, this.stringify(substitution));
                        col.attrib.name = name;
                    }
                });
            });
            this.replaceChildren(columns, newColumns);
            if (inserted > 0) {
                columns.attrib.count = Number(idx).toString();
                root.attrib.ref = this.joinRange(tableRange);
                if (autoFilter !== null) {
                    autoFilter.attrib.ref = this.joinRange(tableRange);
                }
            }
            const tableRoot = table.root;
            const tableRange2 = this.splitRange(tableRoot.attrib.ref);
            const tableStart = this.splitRef(tableRange2.start);
            const tableEnd = this.splitRef(tableRange2.end);
            if (tableRoot.attrib.totalsRowCount) {
                const autoFilter2 = tableRoot.find("autoFilter");
                if (autoFilter2 !== null) {
                    autoFilter2.attrib.ref = this.joinRange({
                        start: this.joinRef(tableStart),
                        end: this.joinRef(tableEnd),
                    });
                }
                ++tableEnd.row;
                tableRoot.attrib.ref = this.joinRange({
                    start: this.joinRef(tableStart),
                    end: this.joinRef(tableEnd),
                });
            }
        });
    }

    /**
     * 提取字符串中可能存在的占位符标记
     * 支持扩展：自定义正则表达式和自定义提取器
     * @param inputString - 输入字符串
     * @returns 占位符数组
     */
    extractPlaceholders(inputString: string): Placeholder[] {
        // 如果提供了自定义占位符提取器，使用它
        if (this.option.customPlaceholderExtractor) {
            return this.option.customPlaceholderExtractor(inputString, this.option);
        }
        return defaultExtractPlaceholders(inputString, this.option);
    }

    private splitRef(ref: string): Ref {
        const match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)?(\$)?([0-9]+)/);
        return {
            table: match && match[1] || null,
            colAbsolute: Boolean(match && match[2]),
            col: match && match[3] || "",
            rowAbsolute: Boolean(match && match[4]),
            row: parseInt(match && match[5], 10)
        };
    }

    private joinRef(ref: Ref): string {
        return (ref.table ? ref.table + "!" : "") +
            (ref.colAbsolute ? "$" : "") +
            ref.col.toUpperCase() +
            (ref.rowAbsolute ? "$" : "") +
            Number(ref.row).toString();
    }

    private nextCol(ref: string): string {
        ref = ref.toUpperCase();
        return ref.replace(/[A-Z]+/, (match) => {
            return this.numToChar(this.charToNum(match) + 1);
        });
    }

    private nextRow(ref: string): string {
        ref = ref.toUpperCase();
        return ref.replace(/[0-9]+/, (match) => {
            return (parseInt(match, 10) + 1).toString();
        });
    }

    private charToNum(str: string | number): number {
        let num = 0;
        if (typeof str === "string") {
            for (let idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration) {
                const thisChar = str.charCodeAt(idx) - 64;
                const multiplier = Math.pow(26, iteration);
                num += multiplier * thisChar;
            }
        } else {
            num = str as number;
        }
        return num;
    }

    private numToChar(num: number): string {
        let str = "";
        for (let i = 0; num > 0; ++i) {
            let remainder = num % 26;
            let charCode = remainder + 64;
            num = (num - remainder) / 26;
            if (remainder === 0) {
                charCode = 90;
                --num;
            }
            str = String.fromCharCode(charCode) + str;
        }
        return str;
    }

    private generateUUID(): string {
        const hexDigits = '0123456789ABCDEF';
        let uuid = '{';
        for (let i = 0; i < 36; i++) {
            if (i === 8 || i === 13 || i === 18 || i === 23) {
                uuid += '-';
            } else {
                uuid += hexDigits[Math.floor(Math.random() * 16)];
            }
        }
        uuid += '}';
        return uuid;
    }

    private isRange(ref: string): boolean {
        return ref.indexOf(':') !== -1;
    }

    private isWithin(ref: string, startRef: string, endRef: string): boolean {
        const start = this.splitRef(startRef);
        const end = this.splitRef(endRef);
        const target = this.splitRef(ref);
        start.col = `${this.charToNum(start.col)}`;
        end.col = `${this.charToNum(end.col)}`;
        target.col = `${this.charToNum(target.col)}`;
        return (start.row <= target.row && target.row <= end.row &&
            start.col <= target.col && target.col <= end.col);
    }

    /**
     * 执行单个值的替换
     * 支持扩展：调用自定义替换器
     */
    private substituteScalar(cell: any, string: string, placeholder: Placeholder, substitution: any): string {
        // 尝试执行自定义替换器
        const customResult = this.executeReplacers(cell, string, placeholder, substitution);
        if (customResult !== undefined) {
            if (placeholder.full) {
                return this.recordCellValue(cell, customResult);
            } else {
                cell.attrib.t = "s";
                return this.recordCellValue(cell, customResult);
            }
        }
        // 默认行为
        if (placeholder.full) {
            return this.recordCellValue(cell, substitution);
        } else {
            const newString = string.replace(placeholder.placeholder, this.stringify(substitution, placeholder));
            cell.attrib.t = "s";
            return this.recordCellValue(cell, newString);
        }
    }

    private substituteArray(cells: any[], cell: any, substitution: any[]): number {
        let newCellsInserted = -1;
        let currentCell = cell.attrib.r;
        substitution.forEach((element) => {
            ++newCellsInserted;
            if (newCellsInserted > 0) {
                currentCell = this.nextCol(currentCell);
            }
            const newCell = this.cloneElement(cell);
            this.recordCellValue(newCell, element);
            newCell.attrib.r = currentCell;
            cells.push(newCell);
        });
        return newCellsInserted;
    }

    /**
     * 初始化 RichData（富数据/本地图片）所需的所有 XML 结构和关系
     * 首次调用时解析预设的 XML 模板并加载 archive 中已存在的文件
     */
    private async initRichData(): Promise<void> {
        if (!this.richDataIsInit) {
            // 从 filePaths 数组获取所有文件路径
            const filePaths = [
                RICH_DATA_RELS_FILE,
                RICH_DATA_RV_FILE,
                RICH_DATA_STRUCTURE_FILE,
                RICH_DATA_TYPES_FILE,
                RICH_DATA_VALUE_REL_FILE,
                RICH_DATA_METADATA_FILE
            ];
            // 从 xmlTemplates 数组获取所有 XML 模板字符串
            const xmlTemplates = [
                RICH_DATA_xml_RELS,
                RICH_DATA_xml_RV,
                RICH_DATA_xml_STRUCTURE,
                RICH_DATA_xml_TYPES,
                RICH_DATA_xml_VALUE_REL,
                RICH_DATA_xml_METADATA
            ];
            // 初始化 XML 根元素
            this._relsrichValueRel = etree.parse(xmlTemplates[0]).getroot();
            this.rdrichvalue = etree.parse(xmlTemplates[1]).getroot();
            this.rdrichvaluestructure = etree.parse(xmlTemplates[2]).getroot();
            this.rdRichValueTypes = etree.parse(xmlTemplates[3]).getroot();
            this.richValueRel = etree.parse(xmlTemplates[4]).getroot();
            this.metadata = etree.parse(xmlTemplates[5]).getroot();
            // 加载 archive 中已存在的文件（覆盖默认模板）
            for (let i = 0; i < filePaths.length; i++) {
                if (this.archive.file(filePaths[i])) {
                    const content = await this.archive.file(filePaths[i]).async("string");
                    const parsedRoot = etree.parse(content).getroot();
                    switch (i) {
                        case 0:
                            this._relsrichValueRel = parsedRoot;
                            break;
                        case 1:
                            this.rdrichvalue = parsedRoot;
                            break;
                        case 2:
                            this.rdrichvaluestructure = parsedRoot;
                            break;
                        case 3:
                            this.rdRichValueTypes = parsedRoot;
                            break;
                        case 4:
                            this.richValueRel = parsedRoot;
                            break;
                        case 5:
                            this.metadata = parsedRoot;
                            break;
                    }
                }
            }
            this.richDataIsInit = true;
        }
    }

    private writeRichDataAlreadyExist(element: any, elementSearchName: string, attributeName: string, attributeValue: string): boolean {
        for (const e of element.findall(elementSearchName)) {
            if (e.attrib[attributeName] == attributeValue) {
                return true;
            }
        }
        return false;
    }


    // 待优化
    private async substituteImageInCell(cell: any, substitution: any): Promise<boolean> {
        if (substitution == null || substitution == "") {
            this.recordCellValue(cell, "");
            return true;
        }
        await this.initRichData();
        const maxFildId = this.findMaxFileId(/xl\/media\/image\d*\..*/, /image(\d*)\./);
        const fileExtension = "jpg";
        try {
            substitution = this.imageToBuffer(substitution);
        } catch (error) {
            if (this.option && this.option.handleImageError && typeof this.option.handleImageError === "function") {
                this.option.handleImageError(substitution, error as Error);
            } else {
                throw error;
            }
        }
        this.archive.file('xl/media/image' + maxFildId + '.' + fileExtension, toArrayBuffer(substitution), {
            binary: true,
            base64: false
        });
        const maxIdRichData = this.findMaxId(this._relsrichValueRel, 'Relationship', 'Id', /rId(\d*)/);
        const _rel = etree.SubElement(this._relsrichValueRel, 'Relationship');
        _rel.set('Id', 'rId' + maxIdRichData);
        _rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
        _rel.set('Target', '../media/image' + maxFildId + '.' + fileExtension);
        const currentCountedRichValue = this.rdrichvalue.get('count');
        this.rdrichvalue.set('count', parseInt(currentCountedRichValue) + 1);
        const rv = etree.SubElement(this.rdrichvalue, 'rv');
        rv.set('s', "0");
        const firstV = etree.SubElement(rv, 'v');
        const secondV = etree.SubElement(rv, 'v');
        firstV.text = currentCountedRichValue;
        secondV.text = "5";
        const rel = etree.SubElement(this.richValueRel, 'rel');
        rel.set("r:id", 'rId' + maxIdRichData);
        const futureMetadata = this.metadata.findall('futureMetadata').find((fm: any) => {
            return fm.attrib.name === 'XLRICHVALUE';
        });
        const futureMetadataCount = futureMetadata.get('count');
        futureMetadata.set('count', parseInt(futureMetadataCount) + 1);
        const bk = etree.SubElement(futureMetadata, 'bk');
        const extLst = etree.SubElement(bk, 'extLst');
        const ext = etree.SubElement(extLst, 'ext');
        ext.set("uri", "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}");
        const xlrd_rvb = etree.SubElement(ext, 'xlrd:rvb');
        xlrd_rvb.set("i", futureMetadataCount);
        const valueMetadataCount = this.metadata.find('valueMetadata').get('count');
        this.metadata.find('valueMetadata').set('count', parseInt(valueMetadataCount) + 1);
        const bk_VM = etree.SubElement(this.metadata.find('valueMetadata'), 'bk');
        const rc = etree.SubElement(bk_VM, 'rc');
        const XLRICHVALUEMetaDataTypeIndex = this.metadata.find('metadataTypes').findall('metadataType').findIndex((el: any) => {
            return el.attrib.name === "XLRICHVALUE";
        });
        rc.set("t", "" + (XLRICHVALUEMetaDataTypeIndex + 1));
        rc.set("v", valueMetadataCount);
        cell.set("t", "e");
        cell.set("vm", parseInt(currentCountedRichValue) + 1);
        // 清除旧的 <v> 元素（如共享字符串索引），t="e" 单元格不需要
        const existingV = cell.find('v');
        if (existingV) {
            cell.remove(existingV);
        }
        return true;
    }

    private substituteImage(cell: any, string: string, placeholder: Placeholder, substitution: any, drawing: DrawingInfo | null): boolean {
        this.substituteScalar(cell, string, placeholder, '');
        if (substitution == null || substitution == "") {
            return true;
        }
        const maxId = this.findMaxId(drawing!.relRoot, 'Relationship', 'Id', /rId(\d*)/);
        const maxFildId = this.findMaxFileId(/xl\/media\/image\d*.jpg/, /image(\d*)\.jpg/);
        // 创建图片文件关联关系
        const rel = etree.SubElement(drawing!.relRoot, 'Relationship');
        rel.set('Id', 'rId' + maxId);
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
        rel.set('Target', '../media/image' + maxFildId + '.jpg');
        // 处理图片源（路径/base64/buffer 统一转为 Buffer）
        try {
            substitution = this.imageToBuffer(substitution);
        } catch (error) {
            if (this.option && this.option.handleImageError && typeof this.option.handleImageError === "function") {
                this.option.handleImageError(substitution, error as Error);
            } else {
                throw error;
            }
        }
        // 将图片数据写入 archive
        this.archive.file('xl/media/image' + maxFildId + '.jpg', toArrayBuffer(substitution), {
            binary: true,
            base64: false
        });
        // 计算图片在绘制中的尺寸
        const {imageWidth, imageHeight} = this._calculateImageDimensions(substitution, cell);
        // 构建图片绘制 XML
        this._buildImageDrawingXML(drawing!.root, cell, maxId, imageWidth, imageHeight);
        return true;
    }

    /**
     * 计算图片的绘制尺寸（考虑合并单元格和 imageRatio 选项）
     * @param imageBuffer - 图片数据 Buffer
     * @param cell - 单元格元素
     * @returns 计算后的图片宽高（EMU 单位）
     */
    private _calculateImageDimensions(imageBuffer: Buffer, cell: any): { imageWidth: number; imageHeight: number } {
        const dimension = sizeOf(imageBuffer);
        let imageWidth = this.pixelsToEMUs(dimension.width);
        let imageHeight = this.pixelsToEMUs(dimension.height);
        let imageInMergeCell = false;
        // 检测合并单元格内的图片，按合并区域缩放
        /* for (let mergeCell of this.sheet!.root.findall("mergeCells/mergeCell")) {
             if (this.cellInMergeCells(cell, mergeCell)) {
                 const mergeCellWidth = this.getWidthMergeCell(mergeCell, this.sheet! as SheetInfoMust);
                 const mergeCellHeight = this.getHeightMergeCell(mergeCell, this.sheet! as SheetInfoMust);
                 const mergeWidthEmus = this.columnWidthToEMUs(mergeCellWidth);
                 const mergeHeightEmus = this.rowHeightToEMUs(mergeCellHeight);
                 const widthRate = imageWidth / mergeWidthEmus;
                 const heightRate = imageHeight / mergeHeightEmus;
                 if (widthRate > heightRate) {
                     imageWidth = Math.floor(imageWidth / widthRate);
                     imageHeight = Math.floor(imageHeight / widthRate);
                 } else {
                     imageWidth = Math.floor(imageWidth / heightRate);
                     imageHeight = Math.floor(imageHeight / heightRate);
                 }
                 imageInMergeCell = true;
             }
         }*/
        // 非合并单元格：应用 imageRatio 缩放比例
        if (!imageInMergeCell) {
            let ratio = this.option?.imageRatio || 100;
            if (ratio <= 0) ratio = 100;
            imageWidth = Math.floor(imageWidth * ratio / 100);
            imageHeight = Math.floor(imageHeight * ratio / 100);
        }
        return {imageWidth, imageHeight};
    }

    /**
     * 构建图片占位符的绘制 XML（DrawingML）
     * 创建 oneCellAnchor、from、ext、pic、spPr 等完整绘制结构
     * @param drawingRoot - 绘制根元素
     * @param cell - 单元格元素
     * @param maxId - 最大关系 ID
     * @param imageWidth - 图片宽度（EMU）
     * @param imageHeight - 图片高度（EMU）
     */
    private _buildImageDrawingXML(drawingRoot: any, cell: any, maxId: number, imageWidth: number, imageHeight: number): void {
        const cellRef = this.splitRef(cell.attrib.r);
        const colIndex = (this.charToNum(cellRef.col) - 1).toString();
        const rowIndex = (cellRef.row - 1).toString();

        const imagePart = etree.SubElement(drawingRoot, 'xdr:oneCellAnchor');
        // from 定位
        const fromPart = etree.SubElement(imagePart, 'xdr:from');
        const fromCol = etree.SubElement(fromPart, 'xdr:col');
        fromCol.text = colIndex;
        etree.SubElement(fromPart, 'xdr:colOff').text = '0';
        const fromRow = etree.SubElement(fromPart, 'xdr:row');
        fromRow.text = rowIndex;
        etree.SubElement(fromPart, 'xdr:rowOff').text = '0';
        // ext 尺寸
        etree.SubElement(imagePart, 'xdr:ext', {cx: `${imageWidth}`, cy: `${imageHeight}`});
        // pic 节点
        const picNode = etree.SubElement(imagePart, 'xdr:pic');
        const nvPicPr = etree.SubElement(picNode, 'xdr:nvPicPr');
        etree.SubElement(nvPicPr, 'xdr:cNvPr', {id: `${maxId}`, name: 'image_' + maxId, descr: ''});
        const cNvPicPr = etree.SubElement(nvPicPr, 'xdr:cNvPicPr');
        etree.SubElement(cNvPicPr, 'a:picLocks', {noChangeAspect: '1'});
        // blipFill
        const blipFill = etree.SubElement(picNode, 'xdr:blipFill');
        etree.SubElement(blipFill, 'a:blip', {
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "r:embed": "rId" + maxId
        });
        const stretch = etree.SubElement(blipFill, 'a:stretch');
        etree.SubElement(stretch, 'a:fillRect');
        // spPr 形状属性
        const spPr = etree.SubElement(picNode, 'xdr:spPr');
        const xfrm = etree.SubElement(spPr, 'a:xfrm');
        etree.SubElement(xfrm, 'a:off', {x: "0", y: "0"});
        etree.SubElement(xfrm, 'a:ext', {cx: `${imageWidth}`, cy: `${imageHeight}`});
        etree.SubElement(spPr, 'a:prstGeom', {'prst': 'rect'});
        etree.SubElement(spPr, 'a:avLst');
        // clientData
        etree.SubElement(imagePart, 'xdr:clientData');
    }

    private cloneElement(element: any, deep?: boolean): any {
        const newElement = etree.Element(element.tag, element.attrib);
        newElement.text = element.text;
        newElement.tail = element.tail;
        if (deep !== false) {
            element.getchildren().forEach((child: any) => {
                newElement.append(this.cloneElement(child, deep));
            });
        }
        return newElement;
    }

    private replaceChildren(parent: any, children: any[]): void {
        parent.delSlice(0, parent.len());
        children.forEach((child) => {
            parent.append(child);
        });
    }

    private getCurrentRow(row: any, rowsInserted: number): number {
        return parseInt(row.attrib.r, 10) + rowsInserted;
    }

    private getCurrentCell(cell: any, currentRow: number, cellsInserted: number): string {
        const colRef = this.splitRef(cell.attrib.r).col;
        const colNum = this.charToNum(colRef);
        return this.joinRef({
            row: currentRow,
            col: this.numToChar(colNum + cellsInserted)
        });
    }

    private updateRowSpan(row: any, cellsInserted: number): void {
        if (cellsInserted !== 0 && row.attrib.spans) {
            const rowSpan = row.attrib.spans.split(':').map((f: string) => parseInt(f, 10));
            rowSpan[1] += cellsInserted;
            row.attrib.spans = rowSpan.join(":");
        }
    }

    private splitRange(range: string): Range {
        const split = range.split(":");
        return {
            start: split[0],
            end: split[1]
        };
    }

    private joinRange(range: Range): string {
        return range.start + ":" + range.end;
    }

    private pushRight(workbook: any, sheet: any, currentCell: string, numCols: number): void {
        const cellRef = this.splitRef(currentCell);
        const currentRow = cellRef.row;
        const currentCol = this.charToNum(cellRef.col);
        sheet.findall("mergeCells/mergeCell").forEach((mergeCell: any) => {
            const mergeRange = this.splitRange(mergeCell.attrib.ref);
            const mergeStart = this.splitRef(mergeRange.start);
            const mergeStartCol = this.charToNum(mergeStart.col);
            const mergeEnd = this.splitRef(mergeRange.end);
            const mergeEndCol = this.charToNum(mergeEnd.col);
            if (mergeStart.row === currentRow && currentCol < mergeStartCol) {
                mergeStart.col = this.numToChar(mergeStartCol + numCols);
                mergeEnd.col = this.numToChar(mergeEndCol + numCols);
                mergeCell.attrib.ref = this.joinRange({
                    start: this.joinRef(mergeStart),
                    end: this.joinRef(mergeEnd),
                });
            }
        });
        workbook.findall("definedNames/definedName").forEach((name: any) => {
            const ref = name.text;
            if (this.isRange(ref)) {
                const namedRange = this.splitRange(ref);
                const namedStart = this.splitRef(namedRange.start);
                const namedStartCol = this.charToNum(namedStart.col);
                const namedEnd = this.splitRef(namedRange.end);
                const namedEndCol = this.charToNum(namedEnd.col);
                if (namedStart.row === currentRow && currentCol < namedStartCol) {
                    namedStart.col = this.numToChar(namedStartCol + numCols);
                    namedEnd.col = this.numToChar(namedEndCol + numCols);
                    name.text = this.joinRange({
                        start: this.joinRef(namedStart),
                        end: this.joinRef(namedEnd),
                    });
                }
            } else {
                const namedRef = this.splitRef(ref);
                const namedCol = this.charToNum(namedRef.col);
                if (namedRef.row === currentRow && currentCol < namedCol) {
                    namedRef.col = this.numToChar(namedCol + numCols);
                    name.text = this.joinRef(namedRef);
                }
            }
        });
        sheet.findall("hyperlinks/hyperlink").forEach((hyperlink: any) => {
            const ref = this.splitRef(hyperlink.attrib.ref);
            const colNumber = this.charToNum(ref.col);
            if (colNumber > currentCol) {
                ref.col = this.numToChar(colNumber + numCols);
                hyperlink.attrib.ref = this.joinRef(ref);
            }
        });
    }

    private pushDown(workbook: any, sheet: any, tables: TableInfo[], currentRow: number, numRows: number): void {
        const mergeCells = sheet.find("mergeCells");
        sheet.findall("mergeCells/mergeCell").forEach((mergeCell: any) => {
            const mergeRange = this.splitRange(mergeCell.attrib.ref);
            const mergeStart = this.splitRef(mergeRange.start);
            const mergeEnd = this.splitRef(mergeRange.end);
            if (mergeStart.row > currentRow) {
                mergeStart.row += numRows;
                mergeEnd.row += numRows;
                mergeCell.attrib.ref = this.joinRange({
                    start: this.joinRef(mergeStart),
                    end: this.joinRef(mergeEnd),
                });
            } else if (mergeStart.row == currentRow) {
                for (let i = 1; i <= numRows; i++) {
                    const newMergeCell = this.cloneElement(mergeCell);
                    mergeStart.row += 1;
                    mergeEnd.row += 1;
                    newMergeCell.attrib.ref = this.joinRange({
                        start: this.joinRef(mergeStart),
                        end: this.joinRef(mergeEnd)
                    });
                    mergeCells.attrib.count = parseInt(mergeCells.attrib.count, 10) + 1;
                    mergeCells._children.push(newMergeCell);
                }
            }
        });
        tables.forEach((table) => {
            const tableRoot = table.root;
            const tableRange = this.splitRange(tableRoot.attrib.ref);
            const tableStart = this.splitRef(tableRange.start);
            const tableEnd = this.splitRef(tableRange.end);
            if (tableStart.row > currentRow) {
                tableStart.row += numRows;
                tableEnd.row += numRows;
                tableRoot.attrib.ref = this.joinRange({
                    start: this.joinRef(tableStart),
                    end: this.joinRef(tableEnd),
                });
                const autoFilter = tableRoot.find("autoFilter");
                if (autoFilter !== null) {
                    autoFilter.attrib.ref = tableRoot.attrib.ref;
                }
            }
        });
        workbook.findall("definedNames/definedName").forEach((name: any) => {
            const ref = name.text;
            if (this.isRange(ref)) {
                const namedRange = this.splitRange(ref);
                const namedStart = this.splitRef(namedRange.start);
                const namedEnd = this.splitRef(namedRange.end);
                if (namedStart) {
                    if (namedStart.row > currentRow) {
                        namedStart.row += numRows;
                        namedEnd.row += numRows;
                        name.text = this.joinRange({
                            start: this.joinRef(namedStart),
                            end: this.joinRef(namedEnd),
                        });
                    }
                }
                if (this.option && this.option.pushDownPageBreakOnTableSubstitution) {
                    if (this.sheet!.name == name.text.split("!")[0].replace(/'/gi, "") && namedEnd) {
                        if (namedEnd.row > currentRow) {
                            namedEnd.row += numRows;
                            name.text = this.joinRange({
                                start: this.joinRef(namedStart),
                                end: this.joinRef(namedEnd),
                            });
                        }
                    }
                }
            } else {
                const namedRef = this.splitRef(ref);
                if (namedRef.row > currentRow) {
                    namedRef.row += numRows;
                    name.text = this.joinRef(namedRef);
                }
            }
        });
        sheet.findall("hyperlinks/hyperlink").forEach((hyperlink: any) => {
            const ref = this.splitRef(hyperlink.attrib.ref);
            if (ref.row > currentRow) {
                ref.row += numRows;
                hyperlink.attrib.ref = this.joinRef(ref);
            }
        });
    }


    private getNbRowOfMergeCell(mergeCell: any): number {
        const mergeRange = this.splitRange(mergeCell.attrib.ref);
        const mergeStartRow = this.splitRef(mergeRange.start).row;
        const mergeEndRow = this.splitRef(mergeRange.end).row;
        return mergeEndRow - mergeStartRow + 1;
    }

    private pixelsToEMUs(pixels: number): number {
        return Math.round(pixels * 914400 / 96);
    }

    private columnWidthToEMUs(width: number): number {
        return this.pixelsToEMUs(width * 7.625579987895905);
    }

    private rowHeightToEMUs(height: number): number {
        return Math.round(height / 72 * 914400);
    }

    private findMaxFileId(fileNameRegex: RegExp, idRegex: RegExp): number {
        const files = this.archive.file(fileNameRegex);
        const maxId = files.reduce((p: number, c: any) => {
            const num = parseInt(idRegex.exec(c.name)![1]);
            if (p == null) {
                return num;
            }
            return p > num ? p : num;
        }, 0);
        return maxId + 1;
    }

    private cellInMergeCells(cell: any, mergeCell: any): boolean {
        const cellCol = this.charToNum(this.splitRef(cell.attrib.r).col);
        const cellRow = this.splitRef(cell.attrib.r).row;
        const mergeRange = this.splitRange(mergeCell.attrib.ref);
        const mergeStartCol = this.charToNum(this.splitRef(mergeRange.start).col);
        const mergeEndCol = this.charToNum(this.splitRef(mergeRange.end).col);
        const mergeStartRow = this.splitRef(mergeRange.start).row;
        const mergeEndRow = this.splitRef(mergeRange.end).row;
        if (cellCol >= mergeStartCol && cellCol <= mergeEndCol) {
            if (cellRow >= mergeStartRow && cellRow <= mergeEndRow) {
                return true;
            }
        }
        return false;
    }

    private imageToBuffer(imageObj: any): Buffer {
        function checkImage(buffer: Buffer): Buffer {
            try {
                sizeOf(buffer);
                return buffer;
            } catch (error) {
                throw new TypeError('imageObj cannot be parse as a buffer image');
            }
        }

        if (!imageObj) {
            throw new TypeError('imageObj cannot be null');
        }
        if (imageObj instanceof Buffer) {
            return checkImage(imageObj);
        }
        if (typeof (imageObj) === 'string' || imageObj instanceof String) {
            imageObj = imageObj.toString();
            const imagePath = this.option && this.option.imageRootPath
                ? this.option.imageRootPath + "/" + imageObj
                : imageObj;
            if (fs.existsSync(imagePath)) {
                return checkImage(Buffer.from(fs.readFileSync(imagePath, {encoding: 'base64'}), 'base64'));
            }
            try {
                return checkImage(Buffer.from(imageObj, 'base64'));
            } catch (error) {
                throw new TypeError('imageObj cannot be parse as a buffer');
            }
        }
        throw new TypeError("imageObj type is not supported : " + typeof (imageObj));
    }

    private findMaxId(element: any, tag: string, attr: string, idRegex: RegExp): number {
        let maxId = 0;
        element.findall(tag).forEach((el: any) => {
            const match = idRegex.exec(el.attrib[attr]);
            if (match == null) {
                throw new Error("Can not find the id!");
            }
            const cid = parseInt(match[1]);
            if (cid > maxId) {
                maxId = cid;
            }
        });
        return ++maxId;
    }
}

export {Workbook};
