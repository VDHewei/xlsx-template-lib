import * as path from 'path';
import * as fs from 'fs';
import * as etree from 'elementtree';
import JsZip from "jszip";
import {imageSize as sizeOf} from 'image-size';
import {clone, isArray, parseInt, toString} from "lodash";
import exceljs from "exceljs";

// 从新模块导入
import {
    Placeholder, Ref, Range, SheetInfo, SheetInfoMust, DrawingInfo, TableInfo, RelsInfo,
    FullOptions, OutputByType,
    CustomReplacer, CustomPlaceholderExtractor, BeforeReplaceHook, AfterReplaceHook,
    CustomFormatter, QueryFunction, CellPlaceholder, CellData, PlaceholderCacheResult,
    ImageValue,
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

    private parseCellAddress(addr: string): { colLetter: string; rowNum: number } | null {
        const match = addr.match(/^([A-Z]+)(\d+)$/i);
        if (!match) return null;
        return { colLetter: match[1], rowNum: parseInt(match[2], 10) };
    }

    async scanSheetPlaceholder(sheet?: exceljs.Worksheet): Promise<CellPlaceholder[]> {
        const list: CellPlaceholder[] = [];
        const cached = new Map<string, boolean>();
        const worksheet = !sheet ? this.sheet.root : sheet;
        for (const row of worksheet.getRows(1, worksheet.rowCount)) {
            const columns = row.cellCount;
            for (let i = 1; i <= columns; i++) {
                let cell = row.getCell(i);
                // 合并单元格，只处理第一个单元格
                let addr = cell.isMerged ? cell.master.address : cell.address;
                // 确保 addr 为字符串
                addr = String(addr);
                // 判断是否已经处理过的单元格
                if (cached.has(addr)) {
                    continue;
                }
                // 获取单元格值
                let text = !cell.text && cell.text !== '' ? cell.value : cell.text;
                // 记录单元格地址已经被处理过
                cached.set(addr, true);
                // 提取占位符
                let pls: Placeholder[] | undefined = this.extractPlaceholders(String(text ?? ''));
                if (!pls || pls.length <= 0) {
                    continue;
                }
                // 解析地址为行号和列号
                const parsed = this.parseCellAddress(addr);
                if (!parsed) {
                    continue;
                }
                // 创建 CellPlaceholder 对象（Row 存行号，Column 存列号）
                const cellValue: CellPlaceholder = {
                    Row: String(parsed.rowNum),
                    Value: text as exceljs.CellValue,
                    Placeholders: pls,
                    Sheet: worksheet.name,
                    Column: parsed.colLetter.charCodeAt(0) - 'A'.charCodeAt(0) + 1,
                };
                // 合并单元格，获取/推算最后一个单元格地址
                if (cell.isMerged) {
                    cellValue.Merge = true;
                    // 获取合并单元格的最后一个单元格
                    const lastCell = (cell.master as any).lastCell;
                    if (lastCell) {
                        const lastAddr = String(lastCell.address);
                        const lastParsed = this.parseCellAddress(lastAddr);
                        if (lastParsed) {
                            cellValue.LastRow = String(lastParsed.rowNum);
                            cellValue.LastColumn = lastParsed.colLetter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
                        }
                    }
                }
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
            // 表格/数组 生成需要处理的单元格(行列)
            items = await this.generateCells(cell, substitutions, loopIndex);
        } else {
            // 检测 image / imageincell 占位符 — 需要特殊处理
            const imgPlaceholder = cell.Placeholders.find(
                p => p.type === 'imageincell' || p.type === 'image'
            );
            if (imgPlaceholder && imgPlaceholder.full) {
                const substitution = this.getSubstitution(imgPlaceholder, substitutions);
                if (substitution !== undefined && substitution !== null && substitution !== '') {
                    // 构建 ImageValue 供 safeUpdate / updateImageCell 使用
                    const imgValue: ImageValue = {
                        imageType: 'base64',
                        buffer: typeof substitution === 'string'
                            ? Buffer.from(substitution, 'base64')
                            : (substitution instanceof Buffer ? substitution : undefined),
                        path: typeof substitution === 'string' ? substitution : undefined,
                    };
                    items.push({
                        Row: cell.Row,
                        Column: cell.Column,
                        Value: imgValue as any,
                    });
                } else {
                    // 无图片数据：清空单元格
                    items.push({
                        Row: cell.Row,
                        Column: cell.Column,
                        Value: '',
                    });
                }
            } else {
                // ${user.name}
                // ${xxx.logo}
                // 单个单元格值处理解析
                let result: CellData = {
                    Row: cell.Row,
                    Column: cell.Column,
                    Value: await this.resolverPlaceholders(cell, substitutions),
                };
                items.push(result);
            }
        }
        return items;
    }

    async setCells(workSheet: exceljs.Worksheet, values: CellData[]): Promise<void> {
        for (const cell of values) {
            const cellRef = workSheet.getCell(cell.Row, cell.Column);
            await this.safeUpdate(cellRef, cell.Value,workSheet,this.workbook);
        }
    }

    public async safeUpdate(cellRef: exceljs.Cell, value: any,workSheet:exceljs.Worksheet,w:exceljs.Workbook): Promise<exceljs.CellValue> {
        // 无论单元各类型，先检查是否图片值
        if (isImageValue(value)) {
            await updateImageCell(cellRef, value, workSheet, w);
            return cellRef.value;
        }

        const style = cellRef.style ? JSON.parse(JSON.stringify(cellRef.style)) : {};
        switch (cellRef.type) {
            case exceljs.ValueType.Null:
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
            case exceljs.ValueType.SharedString:
                cellRef.value = toString(value);
                break;
            default:
                cellRef.value = value;
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

    /**
     * 解析表格/数组占位符，将数组数据展开为 CellData 列表
     * @param cell - 包含 table 类型占位符的 CellPlaceholder
     * @param substitutions - 数据源对象
     * @returns Map<行偏移, CellData[]> 每个键对应一行数据
     */
    private parseSubstitutionsArray(cell: CellPlaceholder, substitutions: Record<string, any>): Map<number, CellData[]> {
        const result = new Map<number, CellData[]>();
        const tablePlaceholders = cell.Placeholders.filter(p => p.type === 'table');
        if (tablePlaceholders.length === 0) return result;

        // 获取第一个 table 占位符对应的数组数据
        const first = tablePlaceholders[0];
        const arrayData = this.getSubstitution(first, substitutions);
        const dataArray: any[] = Array.isArray(arrayData) ? arrayData : [];
        if (dataArray.length === 0) return result;

        // 为数组的每个元素生成一行 CellData
        for (let i = 0; i < dataArray.length; i++) {
            const items: CellData[] = [];
            for (const p of tablePlaceholders) {
                const item = dataArray[i];
                let value: any;
                if (p.key && item !== null && item !== undefined) {
                    // 如果 item 是对象，按 key 取值；否则（简单类型）直接使用 item
                    value = typeof item === 'object' && !Array.isArray(item)
                        ? (item as any)[p.key]
                        : item;
                } else {
                    value = item;
                }
                value = value !== undefined && value !== null ? value : '';
                items.push({
                    Row: cell.Row,
                    Column: cell.Column,
                    Value: value,
                });
            }
            result.set(i, items);
        }
        return result;
    }

    private async generateCells(cell: CellPlaceholder, substitutions: Record<string, any>, loopIndex: number): Promise<CellData[]> {
        const loopArray = this.parseSubstitutionsArray(cell, substitutions);
        if (!loopArray || loopArray.size <= 0) {
            return [];
        }

        const result: CellData[] = [];
        const baseRow = typeof cell.Row === 'number' ? cell.Row : parseInt(String(cell.Row), 10);
        if (isNaN(baseRow)) return result;

        for (const [index, values] of loopArray.entries()) {
            // 取该行第一个 CellData 的值作为当前行的 table 数据
            const cellData = values.length > 0 ? values[0] : null;
            const rowData = cellData ? cellData.Value : '';

            // 将原始占位符文本中的 table 和非 table 占位符都替换为实际数据
            let textValue: any = cell.Value;
            for (const p of cell.Placeholders) {
                if (p.type === 'table') {
                    textValue = await this.replaceCellValue(textValue, p.placeholder, rowData);
                } else {
                    const substitution = this.getSubstitution(p, substitutions);
                    textValue = await this.replaceCellValue(textValue, p.placeholder, substitution);
                }
            }
            result.push({
                Row: baseRow + index,
                Column: cell.Column,
                Value: textValue,
            });
        }
        return result;
    }

    private async replaceCellValue(value: exceljs.CellValue, placeholder: string, newValue: any | undefined): Promise<exceljs.CellValue> {
        if (newValue === undefined || newValue === null) return value;
        if (typeof value === 'string') {
            if (value === placeholder) {
                return newValue;
            }
            return value.replace(placeholder, String(newValue));
        }
        return newValue;
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
     * 生成新的二进制 .xlsx 文件
     * @param options - JSZip 生成选项
     * @returns 生成的输出数据
     */
    async generate<T extends JsZip.OutputType>(options?: JsZip.JSZipGeneratorOptions<T>): Promise<OutputByType[T]> {
        // ExcelJS path: use workbook.writeBuffer (preferred, preserves styles)
        if (this.workbook) {
            const buffer = await this.workbook.xlsx.writeBuffer();
            return buffer as OutputByType[T];
        }
        // Fallback: JSZip path (XML-level manipulation)
        if (this.archive) {
            return await this.archive.generateAsync(options);
        }
        throw new Error('No workbook or archive loaded');
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
