import JsZip from "jszip";
import {imageSize as sizeOf} from 'image-size';
import { parseInt, toString} from "lodash";
import exceljs from "exceljs";

// 从新模块导入
import {
    Placeholder, SheetInfo, SheetInfoMust,
    FullOptions, OutputByType,
    CustomReplacer, CustomPlaceholderExtractor, BeforeReplaceHook, AfterReplaceHook,
    CustomFormatter, QueryFunction, CellPlaceholder, CellData, PlaceholderCacheResult,
    ImageValue,
} from './types';
import {
    isImageValue,
    toDate,
    updateBooleanCell,
    updateFormulaCell,
    updateHyperlinkCell,
    updateImageCell,
    updateRichTextCell
} from './xml-utils';
import {valueDotGet, defaultFormatters, resolveFullDataPath} from './formatters';
import {defaultExtractPlaceholders} from './placeholders';
import {commandExtendQuery} from "../extends";

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
    /** 所有工作表信息列表 */
    sheets: SheetInfo[] | SheetInfoMust[] = [];
    /** 当前处理的工作表信息 */
    sheet: SheetInfo | SheetInfoMust | null = null;
    /** 工作簿 XML 根元素 */
    workbook: exceljs.Workbook | null = null;
    /** 取值结果缓存 **/
    _cache: Map<string, any | undefined> = new Map<string, any | undefined>();
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
        return value;
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

    /** 列字母转数字: A→1, Z→26, AA→27, AB→28 … */
    private colLetterToNumber(letters: string): number {
        let num = 0;
        for (let i = 0; i < letters.length; i++) {
            num = num * 26 + (letters.charCodeAt(i) - 64);
        }
        return num;
    }

    private parseCellAddress(addr: string): { colLetter: string; colNumber: number; rowNum: number } | null {
        const match = addr.match(/^([A-Z]+)(\d+)$/i);
        if (!match) return null;
        return {
            colLetter: match[1],
            colNumber: this.colLetterToNumber(match[1]),
            rowNum: parseInt(match[2], 10),
        };
    }

    async scanSheetPlaceholder(sheet?: exceljs.Worksheet): Promise<CellPlaceholder[]> {
        const list: CellPlaceholder[] = [];
        const cached = new Map<string, boolean>();
        const worksheet = !sheet ? this.sheet.root : sheet;
        const maxCol = worksheet.columnCount || 0;
        for (const row of worksheet.getRows(1, worksheet.rowCount)) {
            const columns = Math.max(row.cellCount, maxCol);
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
                    Column: parsed.colNumber,
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
                            cellValue.LastColumn = lastParsed.colNumber;
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
            items = await this.generateCells(cell, substitutions);
        } else {
            // 检测 image / imageincell / fn:image 占位符 — 需要特殊处理
            const imgPlaceholder = cell.Placeholders.find(
                p => p.type === 'imageincell' || p.type === 'image' || (p.type === 'fn' && (p.subType === 'image' || p.subType === 'imageincell'))
            );
            if (imgPlaceholder) {
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
                        Value: null,
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

        // 获取每个 table 占位符各自的数组数据（独立解析，长度可能不同）
        const arrays: { data: any[]; placeholder: Placeholder }[] = [];
        for (const p of tablePlaceholders) {
            const arrayData = this.getSubstitution(p, substitutions);
            const arr: any[] = Array.isArray(arrayData) ? arrayData : [];
            arrays.push({ data: arr, placeholder: p });
        }

        // 取所有数组中最大长度作为行数
        const maxLen = Math.max(...arrays.map(a => a.data.length), 0);
        if (maxLen === 0) return result;

        // 为每行生成 CellData（每个 table 占位符独立取对应行的值）
        for (let i = 0; i < maxLen; i++) {
            const items: CellData[] = [];
            for (const entry of arrays) {
                const p = entry.placeholder;
                const item = i < entry.data.length ? entry.data[i] : undefined;
                let value: any;
                if (p.key && item !== null && item !== undefined) {
                    value = typeof item === 'object' && !Array.isArray(item)
                        ? (item as any)[p.key]
                        : item;
                } else {
                    value = item;
                }
                value = value !== undefined && value !== null ? value : '';
                // Apply subType formatter (e.g. :date, :day) on each row value
                if (p.subType && p.subType !== '' && p.subType !== 'image' && value !== '') {
                    value = this.executeFormatters(value, p, p.subType);
                }
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

    private async generateCells(cell: CellPlaceholder, substitutions: Record<string, any>): Promise<CellData[]> {
        const loopArray = this.parseSubstitutionsArray(cell, substitutions);
        if (!loopArray || loopArray.size <= 0) {
            return [];
        }

        const result: CellData[] = [];
        const baseRow = typeof cell.Row === 'number' ? cell.Row : parseInt(String(cell.Row), 10);
        if (isNaN(baseRow)) return result;

        for (const [index, values] of loopArray.entries()) {
            // 将原始占位符文本中的 table 和非 table 占位符都替换为实际数据
            // 每个 table 占位符使用 values 中对应位置的值
            let textValue: any = cell.Value;
            let tableIdx = 0;
            for (const p of cell.Placeholders) {
                if (p.type === 'table') {
                    const rowData = tableIdx < values.length ? values[tableIdx].Value : '';
                    textValue = await this.replaceCellValue(textValue, p.placeholder, rowData);
                    tableIdx++;
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
        if (typeof value === 'string') {
            if (newValue === undefined || newValue === null) {
                return value.replace(placeholder, '');
            }
            if (value === placeholder) {
                return newValue;
            }
            return value.replace(placeholder, String(newValue));
        }
        if (newValue === undefined || newValue === null) return value;
        return newValue;
    }

    substituteString(value: string, placeholders: Placeholder[], substitutions: Record<string, any>): string {
        // 循环 placeholders 数组，替换 value 中的占位符
        for (const placeholder of placeholders) {
            const substitution = this.getSubstitution(placeholder, substitutions);
            value = value.replace(placeholder.placeholder, substitution !== undefined ? String(substitution) : '');
        }
        return value;
    }

    getSubstitution(placeholder: Placeholder, substitutions: Record<string, any>): any | undefined {
        const {value, exists} = this.getPlaceholderValueByCache(placeholder);
        if (exists) {
            return value;
        }
        let res = this.valueGet(substitutions, placeholder);
        if(placeholder.subType!=="" && placeholder.subType!=="image" 
            && res !== undefined && res !== null){
            res = this.executeFormatters(res, placeholder, placeholder.subType);
        }
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
        if(p.type === "fn"){
            return this.parseFunc(substitutions,p)
        }else {
            if (p.key) {
                return valueDotGet(substitutions, p.name + '.' + p.key, p.default || (p.type === 'table' ? [] : ''), p.type);
            }
            return valueDotGet(substitutions, p.name, p.default || (p.type === 'table' ? [] : ''), p.type)
        }
    }

    // 解析函数
    private parseFunc(substitutions: object | Record<string, any>, p: Placeholder): any {
        if(!this.option.customQueryFunction){
          return commandExtendQuery(substitutions,p)
        }
        return this.option.customQueryFunction(substitutions,p)
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
     * 从工作簿 XML 中加载所有工作表信息列表
     * @param workbook - 工作簿
     * @returns 工作表信息数组
     */
    protected loadSheets(workbook: exceljs.Workbook): SheetInfo[] {
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




}

export {Workbook};
