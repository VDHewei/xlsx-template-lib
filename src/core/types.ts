import exceljs from "exceljs";

/**
 * 占位符信息
 */
export interface Placeholder {
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

export type CellPlaceholder = {
    Row: string;
    Column: number;
    Sheet?: string | number;
    Value: string| exceljs.CellValue,
    Placeholders: Placeholder[];
};

export type CellData = {
    Row: string|number;
    Column: number;
    Value: any|exceljs.CellValue;
};

export type ImageValue = {
    imageType: 'file'|'base64';
    path?: string;
    buffer?: Buffer;
    width?:number;
    height?: number;
    zoom?: number;
}

export type PlaceholderCacheResult = {
    value:any|undefined;
    exists:boolean;
};

/**
 * 单元格引用
 */
export interface Ref {
    table?: string | null;
    colAbsolute?: boolean;
    col: string;
    rowAbsolute?: boolean;
    row: number;
}

/**
 * 范围
 */
export interface Range {
    start: string;
    end: string;
}

/**
 * 工作表信息
 */
export interface SheetInfo {
    id: number;
    name: string;
    filename?: string;
    root?: exceljs.Worksheet;
}

/**
 * 必须有根元素的工作表信息
 */
export interface SheetInfoMust {
    id: number;
    name: string;
    filename?: string;
    root: exceljs.Worksheet;
}

/**
 * 绘图信息
 */
export interface DrawingInfo {
    filename: string;
    root: any;
    relFilename?: string;
    relRoot?: any;
}

/**
 * 表格信息
 */
export interface TableInfo {
    filename: string;
    root: any;
}

/**
 * 关系文件信息
 */
export interface RelsInfo {
    filename: string;
    root: any;
}

/**
 * 基础工作簿选项
 */
export interface WorkbookOptions {
    moveImages?: boolean;
    substituteAllTableRow?: boolean;
    moveSameLineImages?: boolean;
    imageRatio?: number;
    pushDownPageBreakOnTableSubstitution?: boolean;
    useExistingRows?: boolean;
    imageRootPath?: string | null;
    handleImageError?: ((imageObj: any, error: Error) => void) | null;
}

/**
 * 生成选项 - 按输出类型映射
 */
export type OutputByType = {
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

/**
 * 自定义替换函数类型
 * @param cell - XML 单元格元素
 * @param stringValue - 当前字符串值
 * @param placeholder - 解析出的占位符信息
 * @param substitution - 替换值
 * @returns 返回替换后的字符串，或 undefined/null 表示不处理（交给下一个替换器）
 */
export type CustomReplacer = (
    cell: any,
    stringValue: string,
    placeholder: Placeholder,
    substitution: any
) => string | undefined | null;

/**
 * 自定义占位符提取函数类型
 * @param inputString - 输入字符串
 * @param options - 扩展配置选项
 * @returns 解析出的占位符数组
 */
export type CustomPlaceholderExtractor = (
    inputString: string,
    options: ExtensionOptions,
) => Placeholder[];

/**
 * 输出 Buffer 类型枚举
 */
export enum BufferType {
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
export type BeforeReplaceHook = (
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
export type AfterReplaceHook = (
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
export type CustomFormatter = (
    value: any,
    placeholder: Placeholder,
    key?: string
) => string | undefined | null;

/**
 * 自定义数据查询函数类型
 * @param obj - 数据对象
 * @param p - 占位符
 * @returns 返回对应键 key 值或者 undefined
 */
export type QueryFunction = (
    obj: Object | Record<string, any>,
    p: Placeholder,
) => any | undefined;

/**
 * 扩展配置接口
 */
export interface ExtensionOptions {
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

/**
 * 完整选项类型 - 工作簿选项与扩展选项的合并
 */
export type FullOptions = WorkbookOptions & ExtensionOptions;
