import * as path from 'path';
import * as fs from 'fs';
import * as etree from 'elementtree';
import {Element} from 'elementtree';
import JsZip from "jszip";
import * as console from "node:console";
import {imageSize as sizeOf} from 'image-size';

// ==================== 常量定义 ====================
const DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CALC_CHAIN_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
const SHARED_STRINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const HYPERLINK_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

// ==================== 类型定义 ====================
/** 占位符信息 */
interface Placeholder {
    /** 完整占位符字符串，如 ${name} 或 ${table:items.key} */
    placeholder: string;
    /** 类型：normal, table, image, imageincell,fn 或自定义类型 */
    type: string;
    /** 变量名 */
    name: string;
    /** 对象属性键（table 类型时使用） */
    key?: string;
    /** 子类型（如 image） */
    subType?: string;
    /** 是否是字符串的全部内容 */
    full: boolean;

    /** 自定义属性，用于扩展 */
    [customKey: string]: any;
}

/** 单元格引用 */
interface Ref {
    table?: string | null;
    colAbsolute?: boolean;
    col: string;
    rowAbsolute?: boolean;
    row: number;
}

/** 范围 */
interface Range {
    start: string;
    end: string;
}

/** 工作表信息 */
interface SheetInfo {
    id: number;
    name: string;
    filename: string;
    root?: etree.Element;
}

interface SheetInfoMust {
    id: number;
    name: string;
    filename: string;
    root: etree.Element;
}

/** 绘图信息 */
interface DrawingInfo {
    filename: string;
    root: any;
    relFilename?: string;
    relRoot?: any;
}

/** 表格信息 */
interface TableInfo {
    filename: string;
    root: any;
}

/** 关系文件信息 */
interface RelsInfo {
    filename: string;
    root: any;
}

/** 基础工作簿选项 */
interface WorkbookOptions {
    moveImages?: boolean;
    substituteAllTableRow?: boolean;
    moveSameLineImages?: boolean;
    imageRatio?: number;
    pushDownPageBreakOnTableSubstitution?: boolean;
    imageRootPath?: string | null;
    handleImageError?: ((imageObj: any, error: Error) => void) | null;
}

/** 生成选项 */
type OutputByType = {
    readonly base64: string;
    readonly string: string;
    readonly text: string;
    readonly binarystring: string;
    readonly array: readonly number[];
    readonly uint8array: Uint8Array;
    readonly arraybuffer: ArrayBuffer;
    readonly blob: Blob;
    readonly nodebuffer: Buffer;
};

// ==================== 扩展类型定义 ====================
/**
 * 自定义替换函数类型
 * @param cell - XML 单元格元素
 * @param stringValue - 当前字符串值
 * @param placeholder - 解析出的占位符信息
 * @param substitution - 替换值
 * @returns 返回替换后的字符串，或 undefined/null 表示不处理（交给下一个替换器）
 */
type CustomReplacer = (
    cell: any,
    stringValue: string,
    placeholder: Placeholder,
    substitution: any
) => string | undefined | null;
/**
 * 自定义占位符提取函数类型
 * @param inputString - 输入字符串
 * @returns 解析出的占位符数组
 */
type CustomPlaceholderExtractor = (
    inputString: string,
    options: ExtensionOptions,
) => Placeholder[];

// 输出Buffer 类型
enum BufferType {
    Base64 = "base64",
    String = "string",
    Text = "text",
    Blob = "blob",
    Array = "array",
    NodeBuffer = "nodebuffer",
    Uint8array = "uint8array",
    Arraybuffer = "arraybuffer",
    BinaryString = "binarystring",
}

/**
 * 替换前钩子函数类型
 * @param stringValue - 原始字符串
 * @param substitutions - 替换数据对象
 * @returns 返回修改后的字符串或 undefined（不修改）
 */
type BeforeReplaceHook = (
    stringValue: string,
    substitutions: Record<string, any>
) => string | undefined | null;

/**
 * 替换后钩子函数类型
 * @param resultString - 替换后的字符串
 * @param stringValue - 原始字符串
 * @param substitutions - 替换数据对象
 * @returns 返回最终字符串
 */
type AfterReplaceHook = (
    resultString: string,
    stringValue: string,
    substitutions: Record<string, any>
) => string;

/**
 * 自定义值格式化函数类型
 * @param value - 原始值
 * @param placeholder - 占位符信息
 * @param key - 可选的键名
 * @returns 格式化后的字符串
 */
type CustomFormatter = (
    value: any,
    placeholder: Placeholder,
    key?: string
) => string | undefined | null;

/**
 * 自定义数据查询函数类型
 * @param obj - 数据对象
 * @param p - 占位符
 * @returns 返回对应键key值或者undefined
 */
type QueryFunction = (
    obj: Object | Record<string, any>,
    p: Placeholder,
) => any | undefined;

/** 扩展配置接口 */
interface ExtensionOptions {
    /**
     * 自定义替换函数，在默认替换逻辑之前调用
     * 返回非空值则使用返回值，否则走默认逻辑
     */
    customReplacer?: CustomReplacer;
    /**
     * 自定义占位符提取函数
     * 如果提供，将完全替代默认的 extractPlaceholders
     */
    customPlaceholderExtractor?: CustomPlaceholderExtractor;
    /**
     * 额外的替换函数列表，按顺序尝试
     * 每个函数返回非空值则停止后续处理
     */
    replacers?: CustomReplacer[];
    /**
     * 替换前钩子，在所有替换开始前调用
     */
    beforeReplace?: BeforeReplaceHook;
    /**
     * 替换后钩子，在所有替换完成后调用
     */
    afterReplace?: AfterReplaceHook;
    /**
     * 自定义值格式化函数列表
     * 按顺序尝试，返回非空值则使用
     */
    formatters?: CustomFormatter[];
    /**
     * 自定义模板正则表达式
     * 用于替代默认的占位符匹配模式
     */
    customPlaceholderRegex?: RegExp;
    /**
     * 是否启用默认占位符解析（当使用 customPlaceholderRegex 时）
     * 默认为 true
     */
    enableDefaultParsing?: boolean;

    /**
     * 设置自定义数据查询处理器
     * 默认为 undefined
     */
    customQueryFunction?: QueryFunction;
}

/** 完整选项类型 */
type FullOptions = WorkbookOptions & ExtensionOptions;

// ==================== 辅助函数 ====================
/*function _get_simple(obj: any, keys: string): any {
    if (keys.indexOf("[") >= 0) {
        const specification = keys.split(/[[[\]]/);
        const property = specification[0];
        const index = specification[1];
        return obj[property][index];
    }
    return obj[keys];
}*/

/*function valueDotGet<T extends OutputValue = OutputValue>(obj: any|object|Record<string, any>, keys: string, defaultValue?: ValuerType[T]): ValuerType[T] {
    const arr = keys.split('.');
    try {
        while (arr.length) {
            obj = _get_simple(obj, arr.shift()!);
        }
    } catch (ex) {
        obj = undefined;
    }
    return obj === undefined ? defaultValue : obj;
}*/

function _getSimple(obj: any, key: string): any {
    if (key.includes("[")) {
        // 修正正则：匹配 [ 和 ] 并进行拆分
        // 例如：'list[0]' -> ['list', '0', '']
        const parts = key.split(/[\[\]]/);
        const property = parts[0];
        const index = parts[1];
        if (property && index !== undefined) {
            return obj?.[property]?.[index];
        }
    }
    return obj?.[key];
}

type PathImpl<T, Key extends string> =
    T extends object
        ? Key extends `${infer K}.${infer Rest}`
            ? K extends keyof T
                ? PathImpl<T[K], Rest>
                : never
            : Key extends keyof T
                ? T[Key]
                : never
        : any;

type PathType<T, Key extends string> = string extends Key ? any : PathImpl<T, Key>;

/**
 * 基于路径从对象中获取值
 * 模拟 lodash 的 get 方法
 */
function valueDotGet<T extends Record<string, any> & object, P extends string>(
    obj: T,
    path: P,
    defaultValue?: PathType<T, P>
): PathType<T, P> {
    if (!path || !obj) return defaultValue as PathType<T, P>;
    const keys = path.split('.');
    let current: any = obj;
    for (const key of keys) {
        if (current === null || current === undefined) return defaultValue as PathType<T, P>;
        current = _getSimple(current, key);
    }
    return current === undefined ? defaultValue as PathType<T, P> : current;
}

/**
 * 基于路径从对象中获取值,默认方法
 * 模拟 lodash 的 get 方法
 */
function defaultValueDotGet<T extends Record<string, any> & object>(obj: T, p: Placeholder): PathType<T, string> {
    return valueDotGet(obj, p.name, p.default || '');
}

// ==================== 内置格式化器 ====================
/** 日期格式化器 */
const dateFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if (value instanceof Date) {
        // Excel 中日期是从 1900/01/01 开始的天数
        return Number((value.getTime() / (1000 * 60 * 60 * 24)) + 25569).toString();
    }
    return undefined;
};

/** 数字格式化器 */
const numberFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if (typeof value === "number") {
        return value.toString();
    }
    return undefined;
};

/** 布尔格式化器 */
const booleanFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if (typeof value === "boolean") {
        return Number(value).toString();
    }
    return undefined;
};

/** 字符串格式化器（默认） */
const stringFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if (typeof value === "string") {
        return value.toString();
    }
    return undefined;
};

const defaultRe = /\${(?:([^{}:]+?):)?([^{}:]+?)(?:\.([^{}:]+?))?(?::([^{}:]+?))??}/g;

/**默认占位符提取器**/
const defaultExtractPlaceholders: CustomPlaceholderExtractor = (inputString: string, options: ExtensionOptions): Placeholder[] => {
    const matches: Placeholder[] = [];
    // 默认正则表达式
    // 使用自定义正则表达式（如果提供）
    const re = options.customPlaceholderRegex || defaultRe;
    // 如果启用了默认解析且使用了自定义正则，先执行默认解析
    if (options.enableDefaultParsing && options.customPlaceholderRegex) {
        let match: RegExpExecArray | null;
        while ((match = defaultRe.exec(inputString)) !== null) {
            matches.push({
                placeholder: match[0],
                type: match[1] || 'normal',
                name: match[2],
                key: match[3],
                subType: match[4],
                full: match[0].length === inputString.length
            });
        }
    }
    // 执行当前正则匹配
    let match: RegExpExecArray | null;
    // 重置 lastIndex（如果正则不是全局的）
    re.lastIndex = 0;
    while ((match = re.exec(inputString)) !== null) {
        // 如果已经启用了默认解析，检查是否重复
        if (options.enableDefaultParsing && options.customPlaceholderRegex) {
            const isDuplicate = matches.some(m => m.placeholder === match![0]);
            if (isDuplicate) continue;
        }
        matches.push({
            placeholder: match[0],
            type: match[1] || 'normal',
            name: match[2],
            key: match[3],
            subType: match[4],
            full: match[0].length === inputString.length
        });
    }
    return matches;
}

/** 默认格式化器列表 */
const defaultFormatters: CustomFormatter[] = [
    dateFormatter,
    numberFormatter,
    booleanFormatter,
    stringFormatter
];

const pattern = new RegExp('^(https?:\\/\\/)?' +
    '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' +
    '((\\d{1,3}\\.){3}\\d{1,3}))' +
    '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' +
    '(\\?[;&a-z\\d%_.~+=-]*)?' +
    '(\\#[-a-z\\d_]*)?$', 'i');

const isUrl = function (str: string): boolean {
    return !!pattern.test(str);
}

const toArrayBuffer = function (buffer: Buffer): ArrayBuffer {
    const ab = new ArrayBuffer(buffer.length);
    const view = new Uint8Array(ab);
    for (let i = 0; i < buffer.length; ++i) {
        view[i] = buffer[i];
    }
    return ab;
}

// ==================== Workbook 类 ====================
class Workbook {
    // ==================== 属性定义 ====================
    option: FullOptions;
    archive: JsZip;
    sharedStrings: string[] = [];
    sharedStringsLookup: Record<string, number> = {};
    sharedStringsPath: string = "";
    sheets: SheetInfo[] | SheetInfoMust[] = [];
    sheet: SheetInfo | SheetInfoMust | null = null;
    workbook: Element | null = null;
    workbookPath: string | null = null;
    contentTypes: any = null;
    prefix: string | null = null;
    workbookRels: Element | null = null;
    calChainRel: any = null;
    calcChainPath: string = "";

    // RichData 相关属性
    private richDataIsInit: boolean = false;
    private _relsrichValueRel: any = null;
    private rdrichvalue: any = null;
    private rdrichvaluestructure: any = null;
    private rdRichValueTypes: any = null;
    private richValueRel: any = null;
    private metadata: any = null;

    // ==================== 构造函数 ====================
    constructor(option?: FullOptions) {
        this.option = {
            moveImages: false,
            substituteAllTableRow: false,
            moveSameLineImages: false,
            imageRatio: 100,
            pushDownPageBreakOnTableSubstitution: false,
            imageRootPath: null,
            handleImageError: null,
            ...option
        };
    }

    // ==================== parse 构造函数 ====================
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

    // ==================== 核心方法 ====================
    /**
     * 删除工作表
     */
    async deleteSheet(sheetName: string | number): Promise<this> {
        const sheet = await this.loadSheet(sheetName);
        const sh = this.workbook.find(`sheets/sheet[@sheetId='${sheet.id}']`);
        const sheets = this.workbook.findall("sheets/sheet");
        const sheetIndex = sheets.indexOf(sh);
        // 移除关联的 definedNames
        const definedNamesParent = this.workbook.find("definedNames");
        if (definedNamesParent) {
            const toRemove: any[] = [];
            this.workbook.findall("definedNames/definedName").forEach((def: any) => {
                if (def.attrib.localSheetId !== undefined) {
                    const localId = parseInt(def.attrib.localSheetId);
                    if (localId === sheetIndex) {
                        toRemove.push(def);
                    } else if (localId > sheetIndex) {
                        def.attrib.localSheetId = (localId - 1).toString();
                    }
                }
            });
            toRemove.forEach((def: any) => {
                definedNamesParent.remove(def);
            });
        }
        this.workbook.find("sheets").remove(sh);
        const rel = this.workbookRels.find(`Relationship[@Id='${sh.attrib['r:id']}']`);
        this.workbookRels.remove(rel);
        this._rebuild();
        return this;
    }

    /**
     * 复制工作表
     */
    async copySheet(sheetName: string | number, copyName?: string, binary: boolean = true): Promise<this> {
        // 警告用户 binary 模式已禁用
        if (binary === false && !process.env.JEST_WORKER_ID) {
            console.warn('Warning: copySheet() called with binary=false. UTF-8 characters may be corrupted.');
        }
        const sheet = await this.loadSheet(sheetName);
        const newSheetIndex = (this.workbook.findall("sheets/sheet").length + 1).toString();
        const fileName = 'worksheets/sheet' + newSheetIndex + '.xml';
        const arcName = this.prefix + '/' + fileName;
        // 以二进制模式复制工作表文件以保留 UTF-8 编码
        const sourceSheetFile = this.archive.file(sheet.filename);
        let sheetContent = await sourceSheetFile.async(`nodebuffer`);
        this.archive.file(arcName, sheetContent);
        this.archive.files[arcName].options.compression = binary ? "STORE" : "DEFLATE";

        // 为新工作表添加内容类型
        const sheetContentType = etree.SubElement(this.contentTypes, 'Override');
        sheetContentType.attrib.PartName = '/' + arcName;
        sheetContentType.attrib.ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
        // 在工作簿中复制工作表名称
        const newSheet = etree.SubElement(this.workbook.find('sheets'), 'sheet');
        const finalSheetName = copyName || 'Sheet' + newSheetIndex;
        newSheet.attrib.name = finalSheetName;
        newSheet.attrib.sheetId = newSheetIndex;
        newSheet.attrib['r:id'] = 'rId' + newSheetIndex;
        // 复制 definedName（如果有）
        this.workbook.findall('definedNames/definedName').forEach((element: any) => {
            if (element.text && element.text.split("!").length && element.text.split("!")[0] == sheetName) {
                const newDefinedName = etree.SubElement(this.workbook.find('definedNames'), 'definedName', element.attrib);
                newDefinedName.text = finalSheetName + "!" + element.text.split("!")[1];
                const index = Number.parseInt(newSheetIndex, 10) - 1;
                newDefinedName.attrib.localSheetId = `${index}`;
            }
        });
        const newRel = etree.SubElement(this.workbookRels, 'Relationship');
        newRel.attrib.Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
        newRel.attrib.Target = fileName;
        // 复制工作表关系及其目标文件
        const sourceRels = await this.loadSheetRels(sheet.filename);
        const relFileName = 'worksheets/_rels/sheet' + newSheetIndex + '.xml.rels';
        const relArcName = this.prefix + '/' + relFileName;
        const newRelsRoot = this.cloneElement(sourceRels.root, true);
        // 为注释生成新的 UUID
        const newCommentUuid = this.generateUUID();
        // 处理每个关系以使用唯一名称复制目标文件
        sourceRels.root.findall('Relationship').forEach((rel: any, index: number) => {
            const relType = rel.attrib.Type;
            const target = rel.attrib.Target;
            const needsFileCopy = [
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
                'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
            ];
            if (needsFileCopy.indexOf(relType) !== -1) {
                const sheetDirectory = path.dirname(sheet.filename);
                const sourceFilePath = path.join(sheetDirectory, target).replace(/\\/g, '/');
                const sourceFile = this.archive.file(sourceFilePath);
                if (sourceFile) {
                    const fileExtension = path.extname(target);
                    const fileBaseName = path.basename(target, fileExtension);
                    const fileDir = path.dirname(target);
                    const baseNameWithoutNumber = fileBaseName.replace(/\d+$/, '');
                    const newFileName = baseNameWithoutNumber + newSheetIndex + fileExtension;
                    const newTarget = path.join(fileDir, newFileName).replace(/\\/g, '/');
                    const newFilePath = path.join(sheetDirectory, newTarget).replace(/\\/g, '/');
                    const content = sourceFile.async(`string`);
                    content.then((binaryContent) => {
                        // 应用特定于文件的转换
                        if (relType === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing') {
                            binaryContent = binaryContent.replace(/data="\d+"/, 'data="' + newSheetIndex + '"');
                        } else if (relType === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments') {
                            const uuidWithoutBraces = newCommentUuid.replace(/[{}]/g, '');
                            binaryContent = binaryContent.replace(/(<author>tc=\{)[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}(\}<\/author>)/gi, '$1' + uuidWithoutBraces + '$2');
                            binaryContent = binaryContent.replace(/(xr:uid="\{)[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}(\}\")/gi, '$1' + uuidWithoutBraces + '$2');
                            const commentsContentType = etree.SubElement(this.contentTypes, 'Override');
                            commentsContentType.attrib.PartName = '/' + newFilePath;
                            commentsContentType.attrib.ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
                        } else if (relType === 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment') {
                            const uuidWithoutBraces = newCommentUuid.replace(/[{}]/g, '');
                            binaryContent = binaryContent.replace(/(\sid="\{)[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}(\}")/gi, '$1' + uuidWithoutBraces + '$2');
                            const threadedCommentContentType = etree.SubElement(this.contentTypes, 'Override');
                            threadedCommentContentType.attrib.PartName = '/' + newFilePath;
                            threadedCommentContentType.attrib.ContentType = 'application/vnd.ms-excel.threadedcomments+xml';
                        }
                        this.archive.file(newFilePath, binaryContent);
                        this.archive.files[newFilePath].options.compression = binary ? "STORE" : "DEFLATE";
                        const newRelInRels = newRelsRoot.findall('Relationship')[index];
                        if (newRelInRels) {
                            newRelInRels.attrib.Target = newTarget;
                        }
                    }).catch(err => {
                        console.log(err)
                    })
                }
            }
        });
        this.archive.file(relArcName, etree.tostring(newRelsRoot, {encoding: 'utf-8'}));
        this.archive.files[relArcName].options.compression = binary ? "STORE" : "DEFLATE";
        this.archive.file('[Content_Types].xml', etree.tostring(this.contentTypes, {encoding: 'utf-8'}));
        this._rebuild();
        return this;
    }

    /**
     * 部分重建（复制/删除工作表后）
     */
    private _rebuild(): void {
        const order = ['worksheet', 'theme', 'styles', 'sharedStrings'];
        this.workbookRels.findall("*")
            .sort((rel1: any, rel2: any) => {
                const index1 = order.indexOf(path.basename(rel1.attrib.Type));
                const index2 = order.indexOf(path.basename(rel2.attrib.Type));
                if (index1 < 0 && index2 >= 0) return 1;
                if (index1 >= 0 && index2 < 0) return -1;
                if (index1 < 0 && index2 < 0) return 0;
                if ((index1 + index2) === 0) {
                    if (rel1.attrib.Id && rel2.attrib.Id) {
                        return rel1.attrib.Id.substring(3) - rel2.attrib.Id.substring(3);
                    }
                    return rel1._id - rel2._id;
                }
                return index1 - index2;
            })
            .forEach((item: any, index: number) => {
                item.attrib.Id = 'rId' + (index + 1);
            });
        this.workbook.findall("sheets/sheet").forEach((item: any, index: number) => {
            item.attrib['r:id'] = 'rId' + (index + 1);
            item.attrib.sheetId = (index + 1).toString();
        });
        this.archive.file(
            this.prefix + '/' + '_rels' + '/' + path.basename(this.workbookPath!) + '.rels',
            etree.tostring(this.workbookRels, {encoding: 'utf-8'})
        );
        this.archive.file(this.workbookPath!, etree.tostring(this.workbook, {encoding: 'utf-8'}));
        this.sheets = this.loadSheets(this.prefix!, this.workbook, this.workbookRels);
    }

    /**
     * 从字节数组加载 .xlsx 文件
     */
    async loadTemplate(data: Buffer | string): Promise<void> {
        if (Buffer.isBuffer(data)) {
            data = data.toString('binary');
        }
        this.archive = await JsZip.loadAsync(data, {base64: false, checkCRC32: true});
        // 加载关系
        const text = await this.archive.file("_rels/.rels").async("string");
        const rels = etree.parse(text).getroot();
        const workbookPath = rels.find(`Relationship[@Type='${DOCUMENT_RELATIONSHIP}']`).attrib.Target;
        this.workbookPath = workbookPath;
        this.prefix = path.dirname(workbookPath);
        const workbookText = await this.archive.file(workbookPath).async(`string`);
        this.workbook = etree.parse(workbookText).getroot();
        const refText = await this.archive.file(this.prefix + "/" + '_rels' + "/" + path.basename(workbookPath) + '.rels').async(`string`);
        this.workbookRels = etree.parse(refText).getroot();
        this.sheets = this.loadSheets(this.prefix, this.workbook, this.workbookRels);
        this.calChainRel = this.workbookRels.find(`Relationship[@Type='${CALC_CHAIN_RELATIONSHIP}']`);
        if (this.calChainRel) {
            this.calcChainPath = this.prefix + "/" + this.calChainRel.attrib.Target;
        }
        this.sharedStringsPath = this.prefix + "/" + this.workbookRels.find(`Relationship[@Type='${SHARED_STRINGS_RELATIONSHIP}']`).attrib.Target;
        this.sharedStrings = [];
        this.sharedStringsLookup = {};
        const sharedText = await this.archive.file(this.sharedStringsPath).async(`string`);
        etree.parse(sharedText).getroot().findall('si').forEach((si: any) => {
            const t = {text: ''};
            si.findall('t').forEach((tmp: any) => {
                t.text += tmp.text;
            });
            si.findall('r/t').forEach((tmp: any) => {
                t.text += tmp.text;
            });
            this.sharedStrings.push(t.text);
            this.sharedStringsLookup[t.text] = this.sharedStrings.length - 1;
        });
        const contentTypeText = await this.archive.file('[Content_Types].xml').async(`string`);
        this.contentTypes = etree.parse(contentTypeText).getroot();
        const jpgType = this.contentTypes.find('Default[@Extension="jpg"]');
        if (jpgType === null) {
            etree.SubElement(this.contentTypes, 'Default', {'ContentType': 'image/png', 'Extension': 'jpg'});
        }
    }

    /**
     * 使用给定的替换数据对所有工作表进行插值
     */
    async substituteAll(substitutions: Record<string, any>): Promise<void> {
        const sheets = this.loadSheets(this.prefix!, this.workbook, this.workbookRels);
        for (let sheet of sheets) {
            await this.substitute(sheet.id, substitutions);
        }
    }

    /**
     * 使用给定的替换数据对指定工作表进行插值
     */
    async substitute(sheetName: string | number, substitutions: Record<string, any>): Promise<void> {
        const sheet = await this.loadSheet(sheetName);
        this.sheet = sheet;
        const dimension = sheet.root.find("dimension");
        const sheetData = sheet.root.find("sheetData");
        let currentRow: number | null = null;
        let totalRowsInserted = 0;
        let totalColumnsInserted = 0;
        const namedTables = await this.loadTables(sheet.root, sheet.filename);
        const rows: any[] = [];
        let drawing: DrawingInfo | null = null;
        const rels = await this.loadSheetRels(sheet.filename);
        for (let row of sheetData.findall("row")) {
            currentRow = this.getCurrentRow(row, totalRowsInserted);
            row.attrib.r = `${currentRow}`
            rows.push(row);
            let cells: any[] = [];
            let cellsInserted = 0;
            const newTableRows: any[] = [];
            const cellsSubstituteTable: any[] = [];
            for (let cell of row.findall("c")) {
                let appendCell = true;
                cell.attrib.r = this.getCurrentCell(cell, currentRow!, cellsInserted);
                // 如果是字符串列，查找共享字符串
                if (cell.attrib.t === "s") {
                    const cellValue = cell.find("v");
                    const stringIndex = parseInt(cellValue.text.toString(), 10);
                    let strValue = this.sharedStrings[stringIndex];
                    if (strValue === undefined) {
                        break;
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
                            return;
                        }
                        if (placeholder.full && placeholder.type === "table" && substitution instanceof Array) {
                            if (placeholder.subType === 'image' && drawing == null) {
                                if (rels) {
                                    drawing = await this.loadDrawing(sheet.root, sheet.filename, rels.root);
                                } else {
                                    console.log("Need to implement initRels. Or init this with Excel");
                                }
                            }
                            cellsSubstituteTable.push(cell);
                            newCellsInserted = await this.substituteTable(
                                row, newTableRows, cells, cell, namedTables,
                                substitution, placeholder, drawing
                            );
                            if (newCellsInserted !== 0 || substitution.length) {
                                if (substitution.length === 1) {
                                    appendCell = true;
                                }
                                if (substitution[0][placeholder.key] instanceof Array) {
                                    appendCell = false;
                                }
                            }
                            if (newCellsInserted !== 0) {
                                cellsInserted += newCellsInserted;
                                this.pushRight(this.workbook, sheet.root, cell.attrib.r, newCellsInserted);
                            }
                        }
                        if (placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                            appendCell = false;
                            newCellsInserted = this.substituteArray(cells, cell, substitution);
                            if (newCellsInserted !== 0) {
                                cellsInserted += newCellsInserted;
                                this.pushRight(this.workbook, sheet.root, cell.attrib.r, newCellsInserted);
                            }
                        }
                        if (placeholder.type === "image" && placeholder.full) {
                            if (rels != null) {
                                if (drawing == null) {
                                    drawing = await this.loadDrawing(sheet.root, sheet.filename, rels.root);
                                }
                                this.substituteImage(cell, strValue, placeholder, substitution, drawing);
                            } else {
                                console.log("Need to implement initRels. Or init this with Excel");
                            }
                        }
                        if (placeholder.type === "imageincell" && placeholder.full) {
                            await this.substituteImageInCell(cell, substitution);
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
                    cells.push(cell);
                }
            }
            // 重建行的子节点
            this.replaceChildren(row, cells);
            // 更新行跨度属性
            if (cellsInserted !== 0) {
                this.updateRowSpan(row, cellsInserted);
                if (cellsInserted > totalColumnsInserted) {
                    totalColumnsInserted = cellsInserted;
                }
            }
            // 添加新插入的行
            if (newTableRows.length > 0) {
                if (this.option["moveImages"] && rels) {
                    if (drawing == null) {
                        drawing = await this.loadDrawing(sheet.root, sheet.filename, rels.root);
                    }
                    if (drawing != null) {
                        this.moveAllImages(drawing, row.attrib.r, newTableRows.length);
                    }
                }
                const cellsOverTable = row.findall("c").filter(
                    (cell: any) => !cellsSubstituteTable.includes(cell)
                );
                newTableRows.forEach((newRow: any) => {
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
                    ++totalRowsInserted;
                });
                this.pushDown(this.workbook, sheet.root, namedTables, currentRow!, newTableRows.length);
            }
        }
        // 重建 sheetData 的子节点
        this.replaceChildren(sheetData, rows);
        // 更新表格列标题中的占位符
        this.substituteTableColumnHeaders(namedTables, substitutions);
        // 更新超链接中的占位符
        this.substituteHyperlinks(rels, substitutions);
        // 更新 <dimension />
        if (dimension) {
            if (totalRowsInserted > 0 || totalColumnsInserted > 0) {
                const dimensionRange = this.splitRange(dimension.attrib.ref);
                const dimensionEndRef = this.splitRef(dimensionRange.end);
                dimensionEndRef.row += totalRowsInserted;
                dimensionEndRef.col = this.numToChar(this.charToNum(dimensionEndRef.col) + totalColumnsInserted);
                dimensionRange.end = this.joinRef(dimensionEndRef);
                dimension.attrib.ref = this.joinRange(dimensionRange);
            }
        }
        // 强制重新计算公式值
        sheetData.findall("row").forEach((row: any) => {
            row.findall("c").forEach((cell: any) => {
                const formulas = cell.findall('f');
                if (formulas && formulas.length > 0) {
                    cell.findall('v').forEach((v: any) => {
                        cell.remove(v);
                    });
                }
            });
        });
        // 写回修改后的 XML 树
        this.archive.file(sheet.filename, etree.tostring(sheet.root, {encoding: 'utf-8'}));
        this.archive.file(this.workbookPath!, etree.tostring(this.workbook, {encoding: 'utf-8'}));
        if (rels) {
            this.archive.file(rels.filename, etree.tostring(rels.root, {encoding: 'utf-8'}));
        }
        this.writeRichData();
        this.archive.file('[Content_Types].xml', etree.tostring(this.contentTypes, {encoding: 'utf-8'}));
        // 移除计算链
        if (this.calcChainPath && this.archive.file(this.calcChainPath)) {
            this.archive.remove(this.calcChainPath);
        }
        await this.writeSharedStrings();
        this.writeTables(namedTables);
        this.writeDrawing(drawing);
    }

    /**
     * 生成新的二进制 .xlsx 文件
     */
    async generate<T extends JsZip.OutputType>(options?: JsZip.JSZipGeneratorOptions<T>): Promise<OutputByType[T]> {
        return await this.archive.generateAsync(options);
    }

    /**
     *  查询占位符合数据值
     * @param substitutions 数据对象
     * @param p  占位符
     * @param full  是否fullKey
     * @return any
     */
    public valueGet(substitutions: object | Record<string, any>, p: Placeholder, full?: boolean): any {
        if (this.option.customQueryFunction === undefined) {
            if (full !== undefined && typeof full === "boolean" && full && p.key) {
                return valueDotGet(substitutions, p.name + '.' + p.key, p.default || '');
            }
            return valueDotGet(substitutions, p.name, p.default || '')
        }
        if (full !== undefined && typeof full === "boolean" && full &&
            p.key && !p.name.endsWith(`.${p.key}`)) {
            p.name = p.name + '.' + p.key
        }
        return this.option.customQueryFunction(substitutions, p)
    }

    // ==================== 辅助方法 ====================
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

    private addSharedString(s: string): number {
        const idx = this.sharedStrings.length;
        this.sharedStrings.push(s);
        this.sharedStringsLookup[s] = idx;
        return idx;
    }

    private stringIndex(s: string): number {
        let idx = this.sharedStringsLookup[s];
        if (idx === undefined) {
            idx = this.addSharedString(s);
        }
        return idx;
    }

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

    private loadSheets(prefix: string, workbook: Element, workbookRels: any): SheetInfo[] {
        const sheets: SheetInfo[] = [];
        for (const sheet of workbook.findall("sheets/sheet")) {
            const sheetId = sheet.attrib.sheetId;
            const relId = sheet.attrib['r:id'];
            const relationship = workbookRels.find(`Relationship[@Id='${relId}']`);
            const filename = prefix + "/" + relationship.attrib.Target;
            sheets.push({
                root: sheet,
                filename: filename,
                name: sheet.attrib.name,
                id: parseInt(sheetId, 10),
            });
        }
        return sheets;
    }

    async loadSheet(sheet: string | number): Promise<SheetInfoMust> {
        let info: SheetInfo | null = null;
        for (let i = 0; i < this.sheets.length; ++i) {
            if ((typeof (sheet) === "number" && this.sheets[i].id === sheet) || (this.sheets[i].name === sheet)) {
                info = this.sheets[i];
                break;
            }
        }
        if (info === null && (typeof (sheet) === "number")) {
            info = this.sheets[sheet - 1];
        }
        if (info === null) {
            throw new Error("Sheet " + sheet + " not found");
        }
        const content = await this.archive.file(info.filename).async("string");
        return {
            filename: info.filename,
            name: info.name,
            id: info.id,
            root: etree.parse(content).getroot()
        };
    }

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

    private initSheetRels(sheetFilename: string): RelsInfo {
        const sheetDirectory = path.dirname(sheetFilename);
        const sheetName = path.basename(sheetFilename);
        const relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/');
        const element = etree.Element;
        const ElementTree = etree.ElementTree;
        const root = element('Relationships');
        root.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        const relsEtree = new ElementTree(root);
        return {
            filename: relsFilename,
            root: relsEtree.getroot()
        };
    }

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

    private addContentType(partName: string, contentType: string): void {
        etree.SubElement(this.contentTypes, 'Override', {'ContentType': contentType, 'PartName': partName});
    }

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

    private writeDrawing(drawing: DrawingInfo | null): void {
        if (drawing !== null) {
            this.archive.file(drawing.filename, etree.tostring(drawing.root, {encoding: "utf-8"}));
            this.archive.file(drawing.relFilename, etree.tostring(drawing.relRoot, {encoding: "utf-8"}));
        }
    }

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

    async loadTables(sheet: Element, sheetFilename: string): Promise<TableInfo[]> {
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

    private writeTables(tables: TableInfo[]): void {
        tables.forEach((namedTable) => {
            this.archive.file(namedTable.filename, etree.tostring(namedTable.root, {encoding: 'utf-8'}));
        });
    }

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
     * 将任意类型的值转换为字符串
     * 支持扩展：使用自定义格式化器
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

    private insertCellValue(cell: any, substitution: any): string {
        const cellValue = cell.find("v");
        const stringify = this.stringify(substitution);
        if (typeof substitution === 'string' && substitution[0] === '=') {
            const formula = etree.Element("f");
            formula.text = substitution.substring(1);
            cell.insert(1, formula);
            delete cell.attrib.t;
            return formula.text.toString();
        }
        if (typeof (substitution) === "number" || substitution instanceof Date) {
            delete cell.attrib.t;
            cellValue.text = stringify;
        } else if (typeof (substitution) === "boolean") {
            cell.attrib.t = "b";
            cellValue.text = stringify;
        } else {
            cell.attrib.t = "s";
            cellValue.text = Number(this.stringIndex(stringify)).toString();
        }
        return stringify;
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
                return this.insertCellValue(cell, customResult);
            } else {
                cell.attrib.t = "s";
                return this.insertCellValue(cell, customResult);
            }
        }
        // 默认行为
        if (placeholder.full) {
            return this.insertCellValue(cell, substitution);
        } else {
            const newString = string.replace(placeholder.placeholder, this.stringify(substitution, placeholder));
            cell.attrib.t = "s";
            return this.insertCellValue(cell, newString);
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
            this.insertCellValue(newCell, element);
            newCell.attrib.r = currentCell;
            cells.push(newCell);
        });
        return newCellsInserted;
    }

    private async substituteTable(
        row: any,
        newTableRows: any[],
        cells: any[],
        cell: any,
        namedTables: TableInfo[],
        substitution: any[],
        placeholder: Placeholder,
        drawing: DrawingInfo | null
    ): Promise<number> {
        let newCellsInserted = 0;
        if (substitution.length === 0) {
            delete cell.attrib.t;
            this.replaceChildren(cell, []);
        } else {
            const parentTables = namedTables.filter((namedTable) => {
                const range = this.splitRange(namedTable.root.attrib.ref);
                return this.isWithin(cell.attrib.r, range.start, range.end);
            });
            for (const [idx, element] of substitution.entries()) {
                let newRow: any;
                let newCell: any;
                let newCellsInsertedOnNewRow = 0;
                const newCells: any[] = [];
                const value = this.valueGet(element, placeholder);
                if (idx === 0) {
                    if (value instanceof Array) {
                        newCellsInserted = this.substituteArray(cells, cell, value);
                    } else if (placeholder.subType == 'image' && value != "") {
                        this.substituteImage(cell, placeholder.placeholder, placeholder, value, drawing);
                    } else if (placeholder.subType === "imageincell" && value != "") {
                        await this.substituteImageInCell(cell, value);
                    } else {
                        // 尝试自定义替换器
                        const customResult = this.executeReplacers(cell, '', placeholder, value);
                        if (customResult !== undefined) {
                            this.insertCellValue(cell, customResult);
                        } else {
                            this.insertCellValue(cell, value);
                        }
                    }
                } else {
                    if ((idx - 1) < newTableRows.length) {
                        newRow = newTableRows[idx - 1];
                    } else {
                        newRow = this.cloneElement(row, false);
                        newRow.attrib.r = this.getCurrentRow(row, newTableRows.length + 1);
                        newTableRows.push(newRow);
                    }
                    newCell = this.cloneElement(cell);
                    newCell.attrib.r = this.joinRef({
                        row: newRow.attrib.r,
                        col: this.splitRef(newCell.attrib.r).col
                    });
                    if (value instanceof Array) {
                        newCellsInsertedOnNewRow = this.substituteArray(newCells, newCell, value);
                        newCells.forEach((nc: any) => {
                            newRow.append(nc)
                        });
                        this.updateRowSpan(newRow, newCellsInsertedOnNewRow);
                    } else if (placeholder.subType == 'image' && value != '') {
                        this.substituteImage(newCell, placeholder.placeholder, placeholder, value, drawing);
                    } else if (placeholder.subType === "imageincell" && value != "") {
                        await this.substituteImageInCell(newCell, value);
                        newRow.append(newCell);
                    } else {
                        // 尝试自定义替换器
                        const customResult = this.executeReplacers(newCell, '', placeholder, value);
                        if (customResult !== undefined) {
                            this.insertCellValue(newCell, customResult);
                        } else {
                            this.insertCellValue(newCell, value);
                        }
                        newRow.append(newCell);
                    }
                    // 检查合并单元格
                    const mergeCell = this.sheet!.root.findall("mergeCells/mergeCell")
                        .find((c: any) => this.splitRange(c.attrib.ref).start === cell.attrib.r);
                    const isMergeCell = mergeCell != null;
                    if (isMergeCell) {
                        const originalMergeRange = this.splitRange(mergeCell.attrib.ref);
                        const originalMergeStart = this.splitRef(originalMergeRange.start);
                        const originalMergeEnd = this.splitRef(originalMergeRange.end);
                        for (let column = this.charToNum(originalMergeStart.col) + 1; column <= this.charToNum(originalMergeEnd.col); column++) {
                            const data = this.sheet!.root.find('sheetData');
                            const children = data.getchildren();
                            const originalRow = children.find((f: any) => f.attrib.r == originalMergeStart.row);
                            const col = this.numToChar(column);
                            const originalCell = originalRow.getchildren().find((f: any) => f.attrib.r.startsWith(col));
                            const additionalCell = this.cloneElement(originalCell);
                            additionalCell.attrib.r = this.joinRef({
                                row: newRow.attrib.r,
                                col: this.numToChar(column)
                            });
                            newRow.append(additionalCell);
                        }
                    }
                    // 扩展命名表范围
                    parentTables.forEach((namedTable) => {
                        const tableRoot = namedTable.root;
                        const autoFilter = tableRoot.find("autoFilter");
                        const range = this.splitRange(tableRoot.attrib.ref);
                        if (!this.isWithin(newCell.attrib.r, range.start, range.end)) {
                            range.end = this.nextRow(range.end);
                            tableRoot.attrib.ref = this.joinRange(range);
                            if (autoFilter !== null) {
                                autoFilter.attrib.ref = tableRoot.attrib.ref;
                            }
                        }
                    });
                }
            }
        }
        return newCellsInserted;
    }

    private async initRichData(): Promise<void> {
        if (!this.richDataIsInit) {
            const _relsrichValueRel = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	            </Relationships>`;
            const rdrichvalue = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	            <rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="0">
	            </rvData>`;
            const rdrichvaluestructure = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	            <rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
	                <s t="_localImage">
	                    <k n="_rvRel:LocalImageIdentifier" t="i"/>
	                    <k n="CalcOrigin" t="i"/>
	                </s>
	            </rvStructures>`;
            const rdRichValueTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	            <rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
	                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"
	                xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	                <global>
	                    <keyFlags>
	                        <key name="_Self">
	                            <flag name="ExcludeFromFile" value="1"/>
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_DisplayString">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_Flags">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_Format">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_SubLabel">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_Attribution">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_Icon">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_Display">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_CanonicalPropertyNames">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                        <key name="_ClassificationId">
	                            <flag name="ExcludeFromCalcComparison" value="1"/>
	                        </key>
	                    </keyFlags>
	                </global>
	            </rvTypesInfo>`;
            const richValueRel = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	            <richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
	                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	            </richValueRels>`;
            const metadata = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	            <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	                xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
	                <metadataTypes count="1">
	                    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>
	                </metadataTypes>
	                <futureMetadata name="XLRICHVALUE" count="0">
	                </futureMetadata>
	                <valueMetadata count="0">
	                </valueMetadata>
	            </metadata>`;
            const _relsrichValueRelFileName = 'xl/richData/_rels/richValueRel.xml.rels';
            const rdrichvalueFileName = 'xl/richData/rdrichvalue.xml';
            const rdrichvaluestructureFileName = 'xl/richData/rdrichvaluestructure.xml';
            const rdRichValueTypesFileName = 'xl/richData/rdRichValueTypes.xml';
            const richValueRelFileName = 'xl/richData/richValueRel.xml';
            const metadataFileName = 'xl/metadata.xml';
            this._relsrichValueRel = etree.parse(_relsrichValueRel).getroot();
            this.rdrichvalue = etree.parse(rdrichvalue).getroot();
            this.rdrichvaluestructure = etree.parse(rdrichvaluestructure).getroot();
            this.rdRichValueTypes = etree.parse(rdRichValueTypes).getroot();
            this.richValueRel = etree.parse(richValueRel).getroot();
            this.metadata = etree.parse(metadata).getroot();
            if (this.archive.file(_relsrichValueRelFileName)) {
                const content = await this.archive.file(_relsrichValueRelFileName).async("string");
                this._relsrichValueRel = etree.parse(content).getroot();
            }
            if (this.archive.file(rdrichvalueFileName)) {
                const content = await this.archive.file(rdrichvalueFileName).async("string");
                this.rdrichvalue = etree.parse(content).getroot();
            }
            if (this.archive.file(rdrichvaluestructureFileName)) {
                const content = await this.archive.file(rdrichvaluestructureFileName).async("string");
                this.rdrichvaluestructure = etree.parse(content).getroot();
            }
            if (this.archive.file(rdRichValueTypesFileName)) {
                const content = await this.archive.file(rdRichValueTypesFileName).async("string");
                this.rdRichValueTypes = etree.parse(content).getroot();
            }
            if (this.archive.file(richValueRelFileName)) {
                const content = await this.archive.file(richValueRelFileName).async("string");
                this.richValueRel = etree.parse(content).getroot();
            }
            if (this.archive.file(metadataFileName)) {
                const content = await this.archive.file(metadataFileName).async("string");
                this.metadata = etree.parse(content).getroot();
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

    private writeRichData(): void {
        if (this.richDataIsInit) {
            const _relsrichValueRelFileName = 'xl/richData/_rels/richValueRel.xml.rels';
            const rdrichvalueFileName = 'xl/richData/rdrichvalue.xml';
            const rdrichvaluestructureFileName = 'xl/richData/rdrichvaluestructure.xml';
            const rdRichValueTypesFileName = 'xl/richData/rdRichValueTypes.xml';
            const richValueRelFileName = 'xl/richData/richValueRel.xml';
            const metadataFileName = 'xl/metadata.xml';
            const options = {encoding: "utf-8"};
            this.archive.file(_relsrichValueRelFileName, etree.tostring(this._relsrichValueRel, options));
            this.archive.file(rdrichvalueFileName, etree.tostring(this.rdrichvalue, options));
            this.archive.file(rdrichvaluestructureFileName, etree.tostring(this.rdrichvaluestructure, options));
            this.archive.file(rdRichValueTypesFileName, etree.tostring(this.rdRichValueTypes, options));
            this.archive.file(richValueRelFileName, etree.tostring(this.richValueRel, options));
            this.archive.file(metadataFileName, etree.tostring(this.metadata, options));
            const broadsideMax = this.findMaxId(this.workbookRels, 'Relationship', 'Id', /rId(\d*)/);
            let _rel: any;
            if (!this.writeRichDataAlreadyExist(this.workbookRels, 'Relationship', 'Target', "richData/rdrichvaluestructure.xml")) {
                _rel = etree.SubElement(this.workbookRels, 'Relationship');
                _rel.set('Id', 'rId' + broadsideMax);
                _rel.set('Type', "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure");
                _rel.set('Target', "richData/rdrichvaluestructure.xml");
            }
            if (!this.writeRichDataAlreadyExist(this.workbookRels, 'Relationship', 'Target', "richData/rdrichvalue.xml")) {
                _rel = etree.SubElement(this.workbookRels, 'Relationship');
                _rel.set('Id', "rId" + (broadsideMax + 1));
                _rel.set('Type', "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue");
                _rel.set('Target', "richData/rdrichvalue.xml");
            }
            if (!this.writeRichDataAlreadyExist(this.workbookRels, 'Relationship', 'Target', "richData/richValueRel.xml")) {
                _rel = etree.SubElement(this.workbookRels, 'Relationship');
                _rel.set('Id', "rId" + (broadsideMax + 2));
                _rel.set('Type', "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel");
                _rel.set('Target', "richData/richValueRel.xml");
            }
            if (!this.writeRichDataAlreadyExist(this.workbookRels, 'Relationship', 'Target', "metadata.xml")) {
                _rel = etree.SubElement(this.workbookRels, 'Relationship');
                _rel.set('Id', "rId" + (broadsideMax + 3));
                _rel.set('Type', "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata");
                _rel.set('Target', "metadata.xml");
            }
            if (!this.writeRichDataAlreadyExist(this.workbookRels, 'Relationship', 'Target', "richData/rdRichValueTypes.xml")) {
                _rel = etree.SubElement(this.workbookRels, 'Relationship');
                _rel.set('Id', "rId" + (broadsideMax + 4));
                _rel.set('Type', "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes");
                _rel.set('Target', "richData/rdRichValueTypes.xml");
            }
            if (!this.writeRichDataAlreadyExist(this.contentTypes, 'Override', 'PartName', "/xl/metadata.xml")) {
                let ctOverride = etree.SubElement(this.contentTypes, 'Override');
                ctOverride.set('PartName', "/xl/metadata.xml");
                ctOverride.set('ContentType', "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
            }
            if (!this.writeRichDataAlreadyExist(this.contentTypes, 'Override', 'PartName', "/xl/richData/richValueRel.xml")) {
                let ctOverride = etree.SubElement(this.contentTypes, 'Override');
                ctOverride.set('PartName', "/xl/richData/richValueRel.xml");
                ctOverride.set('ContentType', "application/vnd.ms-excel.richvaluerel+xml");
            }
            if (!this.writeRichDataAlreadyExist(this.contentTypes, 'Override', 'PartName', "/xl/richData/rdrichvalue.xml")) {
                let ctOverride = etree.SubElement(this.contentTypes, 'Override');
                ctOverride.set('PartName', "/xl/richData/rdrichvalue.xml");
                ctOverride.set('ContentType', "application/vnd.ms-excel.rdrichvalue+xml");
            }
            if (!this.writeRichDataAlreadyExist(this.contentTypes, 'Override', 'PartName', "/xl/richData/rdrichvaluestructure.xml")) {
                let ctOverride = etree.SubElement(this.contentTypes, 'Override');
                ctOverride.set('PartName', "/xl/richData/rdrichvaluestructure.xml");
                ctOverride.set('ContentType', "application/vnd.ms-excel.rdrichvaluestructure+xml");
            }
            if (!this.writeRichDataAlreadyExist(this.contentTypes, 'Override', 'PartName', "/xl/richData/rdRichValueTypes.xml")) {
                let ctOverride = etree.SubElement(this.contentTypes, 'Override');
                ctOverride.set('PartName', "/xl/richData/rdRichValueTypes.xml");
                ctOverride.set('ContentType', "application/vnd.ms-excel.rdrichvaluetypes+xml");
            }
            this._rebuild();
        }
    }

    private async substituteImageInCell(cell: any, substitution: any): Promise<boolean> {
        if (substitution == null || substitution == "") {
            this.insertCellValue(cell, "");
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
        this.insertCellValue(cell, "#VALUE!");
        return true;
    }

    private substituteImage(cell: any, string: string, placeholder: Placeholder, substitution: any, drawing: DrawingInfo | null): boolean {
        this.substituteScalar(cell, string, placeholder, '');
        if (substitution == null || substitution == "") {
            return true;
        }
        const maxId = this.findMaxId(drawing!.relRoot, 'Relationship', 'Id', /rId(\d*)/);
        const maxFildId = this.findMaxFileId(/xl\/media\/image\d*.jpg/, /image(\d*)\.jpg/);
        const rel = etree.SubElement(drawing!.relRoot, 'Relationship');
        rel.set('Id', 'rId' + maxId);
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
        rel.set('Target', '../media/image' + maxFildId + '.jpg');
        try {
            substitution = this.imageToBuffer(substitution);
        } catch (error) {
            if (this.option && this.option.handleImageError && typeof this.option.handleImageError === "function") {
                this.option.handleImageError(substitution, error as Error);
            } else {
                throw error;
            }
        }
        this.archive.file('xl/media/image' + maxFildId + '.jpg', toArrayBuffer(substitution), {
            binary: true,
            base64: false
        });
        const dimension = sizeOf(substitution);
        let imageWidth = this.pixelsToEMUs(dimension.width);
        let imageHeight = this.pixelsToEMUs(dimension.height);
        let imageInMergeCell = false;
        for (let mergeCell of this.sheet!.root.findall("mergeCells/mergeCell")) {
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
        }
        if (!imageInMergeCell) {
            let ratio = 100;
            if (this.option && this.option.imageRatio) {
                ratio = this.option.imageRatio;
            }
            if (ratio <= 0) {
                ratio = 100;
            }
            imageWidth = Math.floor(imageWidth * ratio / 100);
            imageHeight = Math.floor(imageHeight * ratio / 100);
        }
        const imagePart = etree.SubElement(drawing!.root, 'xdr:oneCellAnchor');
        const fromPart = etree.SubElement(imagePart, 'xdr:from');
        const fromCol = etree.SubElement(fromPart, 'xdr:col');
        fromCol.text = (this.charToNum(this.splitRef(cell.attrib.r).col) - 1).toString();
        const fromColOff = etree.SubElement(fromPart, 'xdr:colOff');
        fromColOff.text = '0';
        const fromRow = etree.SubElement(fromPart, 'xdr:row');
        fromRow.text = (this.splitRef(cell.attrib.r).row - 1).toString();
        const fromRowOff = etree.SubElement(fromPart, 'xdr:rowOff');
        fromRowOff.text = '0';
        const extImagePart = etree.SubElement(imagePart, 'xdr:ext', {cx: `${imageWidth}`, cy: `${imageHeight}`});
        const picNode = etree.SubElement(imagePart, 'xdr:pic');
        const nvPicPr = etree.SubElement(picNode, 'xdr:nvPicPr');
        const cNvPr = etree.SubElement(nvPicPr, 'xdr:cNvPr', {id: `${maxId}`, name: 'image_' + maxId, descr: ''});
        const cNvPicPr = etree.SubElement(nvPicPr, 'xdr:cNvPicPr');
        const picLocks = etree.SubElement(cNvPicPr, 'a:picLocks', {noChangeAspect: '1'});
        const blipFill = etree.SubElement(picNode, 'xdr:blipFill');
        const blip = etree.SubElement(blipFill, 'a:blip', {
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "r:embed": "rId" + maxId
        });
        const stretch = etree.SubElement(blipFill, 'a:stretch');
        const fillRect = etree.SubElement(stretch, 'a:fillRect');
        const spPr = etree.SubElement(picNode, 'xdr:spPr');
        const xfrm = etree.SubElement(spPr, 'a:xfrm');
        const off = etree.SubElement(xfrm, 'a:off', {x: "0", y: "0"});
        const ext = etree.SubElement(xfrm, 'a:ext', {cx: `${imageWidth}`, cy: `${imageHeight}`});
        const prstGeom = etree.SubElement(spPr, 'a:prstGeom', {'prst': 'rect'});
        const avLst = etree.SubElement(prstGeom, 'a:avLst');
        const clientData = etree.SubElement(imagePart, 'xdr:clientData');
        return true;
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
            }
            if (mergeStart.row == currentRow) {
                for (let i = 1; i <= numRows; i++) {
                    const newMergeCell = this.cloneElement(mergeCell);
                    mergeStart.row += 1;
                    mergeEnd.row += 1;
                    newMergeCell.attrib.ref = this.joinRange({
                        start: this.joinRef(mergeStart),
                        end: this.joinRef(mergeEnd)
                    });
                    mergeCells.attrib.count += 1;
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

    async hideCols(sheetName: string | number, hideItemIndexes: number[]): Promise<this> {
        const sheet = await this.loadSheet(sheetName);
        this.sheet = sheet;
        if (Array.isArray(hideItemIndexes) && hideItemIndexes.length) {
            const cols = sheet.root.find("cols");
            if (cols) {
                hideItemIndexes.forEach((hideIndex) => {
                    const colIndex = hideIndex + 1;
                    const col = cols.findall("col").find((c: any) => {
                        const min = parseInt(c.attrib.min, 10);
                        const max = parseInt(c.attrib.max, 10);
                        return colIndex >= min && colIndex <= max;
                    });
                    if (col) {
                        col.attrib.hidden = "1";
                    }
                });
            }
        }
        this.archive.file(sheet.filename, etree.tostring(sheet.root, {encoding: 'utf-8'}));
        this._rebuild();
        return this;
    }

    private getWidthCell(numCol: number, sheet: SheetInfo | SheetInfoMust): number {
        let defaultWidth = sheet.root.find("sheetFormatPr").attrib["defaultColWidth"];
        if (!defaultWidth) {
            defaultWidth = "11.42578125";
        }
        let finalWidth = defaultWidth;
        sheet.root.findall("cols/col").forEach((col: any) => {
            if (numCol >= col.attrib["min"] && numCol <= col.attrib["max"]) {
                if (col.attrib["width"] != undefined) {
                    finalWidth = col.attrib["width"];
                }
            }
        });
        return Number.parseFloat(finalWidth);
    }

    private getWidthMergeCell(mergeCell: any, sheet: SheetInfoMust): number {
        let mergeWidth = 0;
        const mergeRange = this.splitRange(mergeCell.attrib.ref);
        const mergeStartCol = this.charToNum(this.splitRef(mergeRange.start).col);
        const mergeEndCol = this.charToNum(this.splitRef(mergeRange.end).col);
        for (let i = mergeStartCol; i < mergeEndCol + 1; i++) {
            mergeWidth += this.getWidthCell(i, sheet);
        }
        return mergeWidth;
    }

    private getHeightCell(numRow: number, sheet: SheetInfoMust): number {
        let finalHeight = sheet.root.find("sheetFormatPr").attrib["defaultRowHeight"];
        sheet.root.findall("sheetData/row").forEach((row: any) => {
            if (numRow == row.attrib["r"]) {
                if (row.attrib["ht"] != undefined) {
                    finalHeight = row.attrib["ht"];
                }
            }
        });
        return Number.parseFloat(finalHeight);
    }

    private getHeightMergeCell(mergeCell: any, sheet: SheetInfoMust): number {
        let mergeHeight = 0;
        const mergeRange = this.splitRange(mergeCell.attrib.ref);
        const mergeStartRow = this.splitRef(mergeRange.start).row;
        const mergeEndRow = this.splitRef(mergeRange.end).row;
        for (let i = mergeStartRow; i < mergeEndRow + 1; i++) {
            mergeHeight += this.getHeightCell(i, sheet);
        }
        return mergeHeight;
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

// xlsx 模板 生成 - 函数一键调用
const generateXlsxTemplate = async function <T extends JsZip.OutputType>(data: Buffer, values: Object, options?: JsZip.JSZipGeneratorOptions<T> & FullOptions): Promise<OutputByType[T]> {
    const w = await Workbook.parse(data,options);
    await w.substituteAll(values);
    return w.generate(options);
}

export {
    Workbook,
    Placeholder,
    Ref,
    Range,
    SheetInfo,
    DrawingInfo,
    TableInfo,
    RelsInfo,
    WorkbookOptions,
    ExtensionOptions,
    FullOptions,
    CustomReplacer,
    CustomPlaceholderExtractor,
    BeforeReplaceHook,
    AfterReplaceHook,
    CustomFormatter,
    OutputByType,
    BufferType,
    defaultFormatters,
    QueryFunction,
    toArrayBuffer,
    isUrl,
    valueDotGet,
    defaultValueDotGet,
    defaultExtractPlaceholders,
    generateXlsxTemplate,
}