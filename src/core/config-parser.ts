import exceljs from "exceljs";
import {Stream} from "stream";

// ==================== 类型定义 ====================

/**
 * 单元格位置类型
 * @property Row - 行标识（字母格式）
 * @property Column - 列号
 * @property Sheet - 可选的工作表名称或索引
 */
export type CellPosition = {
    Row: string;
    Column: number;
    Sheet?: string | number;
};

/**
 * 合并单元格范围
 * @property top - 上边界行号
 * @property left - 左边界列号
 * @property bottom - 下边界行号
 * @property right - 右边界列号
 */
export type MergeCellRange = {
    top: number;
    left: number;
    bottom: number;
    right: number
};

/**
 * 单元格坐标点
 * @property Row - 行号
 * @property Column - 列号
 */
export type CellPoint = {
    Row: number,
    Column: number,
};

/**
 * 占位符单元格值接口
 * 定义占位符的字符串表示和合并单元格时的行为
 */
export interface PlaceholderCellValue {
    toString(): string;
    mergeCell(values: string[]): string;
}

/**
 * 扫描令牌数据结构
 * @property value - 原始值
 * @property token - 令牌字符串
 */
export type ScanTokenData = {
    value: string;
    token: string;
};

/**
 * 宏参数类型
 * @property type - 宏类型（expr/exprArr/index）
 * @property columnParam - 列参数
 * @property rowParam - 行参数（可以是单个数或数组）
 * @property formatter - 可选的格式化器
 */
export type MacroArgs = {
    type: string
    columnParam: number
    rowParam: number[] | number
    formatter?: string
};

/**
 * 宏提取参数类型
 * @property argToken - 参数分隔令牌
 * @property startToken - 起始令牌
 * @property endToken - 结束令牌
 */
export type ExtractMacroArgs = {
    argToken: string
    startToken: string
    endToken: string
};

// ==================== 常量/正则 ====================

/** 纯数字正则 */
export const isPureNumber = /^[0-9]+$/;
/** 纯大写字母正则 */
export const isPureUppercase = /^[A-Z]+$/;
/** 表达式单值类型标识 */
export const exprSingle = `expr`;
/** 表达式数组类型标识 */
export const exprArr = `exprArr`;
/** 表达式索引类型标识 */
export const exprIndex = `index`;
/** 默认键标识 */
export const defaultKey = `!!`;
/** 数字格式键标识 */
export const numberKey = `!!number`;
/** 代码键标识 */
export const codeKey = `!!codeKey`;
/** 代码别名键标识 */
export const codeAliasKey = `!!codeAliasKey`;
/** 函数命令前缀 */
export const funcCommand = "fn:";

// ==================== 枚举定义 ====================

/**
 * 规则令牌枚举
 * 定义模板配置文件中使用的各种令牌类型
 */
export enum RuleToken {
    /** 别名令牌 */
    AliasToken = 'alias',
    /** 单元格令牌 */
    CellToken = 'cell',
    /** 合并单元格令牌 */
    MergeCellToken = 'mergeCell',
    /** 行单元格令牌 */
    RowCellToken = 'rowCell',
    /** 使用别名令牌 */
    UseAliasToken = '@',
    /** 范围令牌 */
    RangeToken = '-',
    /** 位置令牌 */
    PosToken = ':',
    /** 函数模式令牌 */
    FunctionPatternToken = '<?>',
    /** 任意字符令牌 */
    AnyToken = '?',
    /** 变量模式令牌 */
    VarPatternToken = '${?}',
    /** 未定义令牌 */
    UndefinedToken = '',
    /** 等号令牌 */
    EqualToken = '=',
    /** 参数分隔令牌 */
    ArgPosToken = ',',
    /** 左括号令牌 */
    LparenToken = '(',
    /** 右括号令牌 */
    RparenToken = ')',
    /** 点获取令牌 */
    DotGetToken = '.',
    /** 编译生成单元格令牌 */
    CompileGenToken = 'compile:GenCell',
    /** 编译宏令牌 */
    CompileMacroToken = 'compile:Macro',
}

// ==================== 类型定义（续）====================

/**
 * 规则值类型
 * 解析后的单个规则表达式的结果
 */
export type RuleValue = {
    express: string;
    tokens: RuleToken[]; //  express tokens
    cells?: CellPoint[];
    key?: string; // alias key
    ref?: string[]; // alias refs
    func?: string; // express ref function name
    compileExpress?: string[];// compileExpress
    posExpr?: RuleValue; // pos Expr
    funcExpr?: RuleValue; // function funcExpr
    value: string | number[] | number | CellPoint | RangeCell | any[]; // alias value
    // extends
    [key: string]: any;
};

/**
 * 规则结果类型
 * 以令牌为键的规则值映射
 */
export type RuleResult = {
    rules: Map<RuleToken, RuleValue[]>;
};

/**
 * 过滤宏结果类型
 * @property tokens - 宏令牌列表
 * @property express - 表达式列表
 */
export type FilterMacroResult = {
    tokens: RuleToken[];
    express: string[];
};

/**
 * 编译检查器类型
 * 自定义规则检查函数
 */
export type CompileChecker = (iv: RuleResult, ctx: RuleMapOptions) => Error[] | undefined;

/**
 * 规则选项接口
 * 定义模板配置的解析和编译选项
 */
export interface RuleOptions {
    startLine?: number;
    endLine?: number;
    endColumn?: number;
    startColumn?: number;
    // rule token alias settings
    ruleKeyMap?: Map<RuleToken, string>;
    // compile settings
    compileCheckers?: CompileChecker[];
    compileSheets?: string[];

    // extends for custom
    [key: string]: any;

    /** 设置结束行 */
    setEndRow(end: number): RuleOptions;
    /** 解析令牌值 */
    parseToken(value: string): RuleToken;
    /** 设置结束列 */
    setEndColumn(end: number): RuleOptions;
    /** 设置起始行 */
    setStartRow(start: number): RuleOptions;
    /** 设置起始列 */
    setStartColumn(start: number): RuleOptions;
    /** 获取上下文映射 */
    getContextMap(): Map<RuleToken, RuleValue[]>;
    /** 添加规则映射 */
    addRuleMap(key: RuleToken, value: string): RuleOptions;
    /** 解析默认配置 */
    parseDefault(worksheet: exceljs.Worksheet): RuleOptions;
    /** 获取编译检查处理器 */
    getCompileCheckHandlers(): CompileChecker[] | undefined;
}

/**
 * 范围单元格接口
 * 定义单元格范围及其操作
 */
export interface RangeCell {
    minRow: number;
    maxRow: number;
    stepRow: number;
    minColumn: number;
    maxColumn: number;
    stepColumn: number;
    getCells(): CellPoint[];
}

/**
 * 规则映射选项类
 * 实现 RuleOptions 接口，提供默认的规则配置
 */
export class RuleMapOptions implements RuleOptions {
    // rule configure area
    startLine?: number = 1;
    endLine?: number;
    endColumn?: number;
    startColumn?: number = 1;
    // rule token alias settings
    ruleKeyMap?: Map<RuleToken, string>;
    // compile settings
    compileCheckers?: CompileChecker[];
    compileSheets?: string[];

    // extends for custom
    [key: string]: any;

    constructor(m?: Map<RuleToken, string>) {
        if (m === undefined) {
            this.ruleKeyMap = defaultRuleTokenMap;
        } else {
            this.ruleKeyMap = m;
        }
    }

    /**
     * 使用工作簿中所有工作表创建选项（排除指定工作表）
     * @param w - Excel 工作簿
     * @param excludes - 排除的工作表名称或索引列表
     */
    static withAllSheets(w: exceljs.Workbook, excludes?: string[]): RuleOptions {
        const compileSheets = [];
        const options = new RuleMapOptions();
        if (excludes === undefined) {
            excludes = [];
        }
        if (w.worksheets.length > 0 && excludes.length > 0) {
            for (const [index, sheet] of w.worksheets.entries()) {
                if (excludes.includes(index.toString()) || excludes.includes(sheet.name)) {
                    continue;
                }
                if (sheet.name.endsWith(".json")) {
                    continue;
                }
                compileSheets.push(sheet.name);
            }
        }
        options.compileSheets = compileSheets;
        return options;
    }

    /** @inheritdoc */
    parseDefault(worksheet: exceljs.Worksheet): RuleOptions {
        this.ruleKeyMap = mergeOption(this.ruleKeyMap, defaultRuleTokenMap);
        if (this.startLine === undefined || isNaN(this.startLine) || this.startLine < 0) {
            this.startLine = 1;
        }
        if (this.endLine === undefined || isNaN(this.endLine) || this.endLine < 0) {
            this.endLine = worksheet.rowCount;
        }
        if (this.startColumn === undefined || isNaN(this.startColumn) || this.startColumn < 0) {
            this.startColumn = 1;
        }
        if (this.endColumn === undefined || isNaN(this.endColumn) || this.endColumn < 0) {
            this.endColumn = worksheet.columnCount;
        }
        return this;
    }

    /** @inheritdoc */
    addRuleMap(key: RuleToken, value: string): RuleOptions {
        this.ruleKeyMap.set(key, value);
        return this;
    }

    /** @inheritdoc */
    setStartRow(start: number): RuleOptions {
        this.startLine = start;
        return this;
    }

    /** @inheritdoc */
    setStartColumn(start: number): RuleOptions {
        this.startColumn = start;
        return this;
    }

    /** @inheritdoc */
    setEndRow(end: number): RuleOptions {
        this.endLine = end;
        return this;
    }

    /** @inheritdoc */
    setEndColumn(end: number): RuleOptions {
        this.endColumn = end;
        return this;
    }

    /** @inheritdoc */
    parseToken(value: string): RuleToken {
        if (value === "") {
            return RuleToken.UndefinedToken;
        }
        for (const [token, alias] of this.ruleKeyMap.entries()) {
            if (alias === value) {
                return token;
            }
        }
        return RuleToken.UndefinedToken;
    }

    /** @inheritdoc */
    getContextMap(): Map<RuleToken, RuleValue[]> {
        const ctx = new Map<RuleToken, RuleValue[]>();
        for (const [token, expr] of this.ruleKeyMap.entries()) {
            if (!isRuleToken(token)) {
                const value: RuleValue[] = [{
                    express: expr,
                    key: expr,
                    tokens: [token],
                    value: token.toString(),
                }];
                ctx.set(token, value);
            }
        }
        return ctx;
    }

    /** @inheritdoc */
    getCompileCheckHandlers(): CompileChecker[] | undefined {
        if (this.compileCheckers !== undefined && this.compileCheckers.length > 0) {
            return this.compileCheckers;
        }
        return undefined;
    }

}

/**
 * 编译上下文类
 * 扩展 RuleMapOptions，添加别名缓存和工作表上下文
 */
export class CompileContext extends RuleMapOptions {
    private aliasMap: Map<string, string> = new Map<string, string>();
    sheet?: exceljs.Worksheet;

    constructor(m?: Map<RuleToken, string>) {
        super(m);
    }

    /**
     * 从 RuleMapOptions 创建 CompileContext
     * @param r - 规则映射选项
     */
    static create(r: RuleMapOptions): CompileContext {
        const ctx = new CompileContext(r.ruleKeyMap);
        Object.assign(ctx, {...r})
        return ctx.init();
    }

    private init(): this {
        if (this.ruleKeyMap === undefined) {
            this.ruleKeyMap = defaultRuleTokenMap;
        }
        return this;
    }

    /**
     * 加载别名映射
     * @param m - 规则结果映射
     */
    loadAlias(m: Map<RuleToken, RuleValue[]>): this {
        if (m.size <= 0 || !m.has(RuleToken.AliasToken)) {
            return this;
        }
        const values = m.get(RuleToken.AliasToken);
        for (const vs of values) {
            if (typeof vs.value === "string") {
                this.aliasMap.set(vs.key, vs.value as string);
            }
        }
        return this;
    }

    /**
     * 设置别名缓存
     * @param key - 别名的键
     * @param value - 别名的值
     */
    public setAlias(key: string, value: string): void {
        this.aliasMap.set(key, value);
    }

    /**
     * 获取别名缓存值
     * @param key - 别名的键
     * @returns 别名的值，如果不存在则返回 undefined
     */
    public getAlias(key: string): string | undefined {
        return this.aliasMap.get(key);
    }

    /**
     * 检查别名是否存在
     * @param key - 别名的键
     */
    public hasAlias(key: string): boolean {
        return this.aliasMap.has(key);
    }

    /**
     * 过滤工作表
     * @param sheetName - 工作表名称
     * @returns 如果工作表在编译列表中返回 true
     */
    public filterSheet(sheetName: string): boolean {
        if (sheetName !== "" && this.compileSheets !== undefined && this.compileSheets.length > 0) {
            return this.compileSheets.includes(sheetName)
        }
        return false;
    }

}

/**
 * 令牌解析器回复类型
 */
export type TokenParserReply = {
    ok: boolean
    values?: any,
    expr?: RuleValue,
    [key: string]: any;
};

/**
 * 令牌解析器函数类型
 */
export type TokenParser = (ctx: Map<RuleToken, RuleValue[]>, t: RuleToken, value: string) => TokenParserReply;

/**
 * 令牌解析器查找结果
 */
export type TokenParseResolver = { exists: boolean, handler?: TokenParser };

/**
 * 范围值解析选项
 */
export type RangeValueParserOptions = {
    rowRange: number;
    columnRange: number;
    row: string;
    column: string;
    express: string;
    token: RuleToken;
};

// ==================== 规则令牌默认映射 ====================

/** 默认规则令牌到字符串的映射 */
export const defaultRuleTokenMap = new Map<RuleToken, string>([
    [RuleToken.AliasToken, RuleToken.AliasToken.toString()],
    [RuleToken.AnyToken, RuleToken.AnyToken.toString()],
    [RuleToken.CellToken, RuleToken.CellToken.toString()],
    [RuleToken.RowCellToken, RuleToken.RowCellToken.toString()],
    [RuleToken.MergeCellToken, RuleToken.MergeCellToken.toString()],
    [RuleToken.UseAliasToken, RuleToken.UseAliasToken.toString()],
    [RuleToken.PosToken, RuleToken.PosToken.toString()],
    [RuleToken.RangeToken, RuleToken.RangeToken.toString()],
    [RuleToken.FunctionPatternToken, RuleToken.FunctionPatternToken.toString()],
    [RuleToken.VarPatternToken, RuleToken.VarPatternToken.toString()],
    [RuleToken.EqualToken, RuleToken.EqualToken.toString()],
    [RuleToken.ArgPosToken, RuleToken.ArgPosToken.toString()],
    [RuleToken.LparenToken, RuleToken.LparenToken.toString()],
    [RuleToken.RparenToken, RuleToken.RparenToken.toString()],
    [RuleToken.DotGetToken, RuleToken.DotGetToken.toString()],
    [RuleToken.CompileMacroToken, RuleToken.CompileMacroToken.toString()],
    [RuleToken.CompileGenToken, RuleToken.CompileGenToken.toString()],
]);

// ==================== 默认占位符单元格值类 ====================

/**
 * 默认占位符单元格值实现
 */
export class DefaultPlaceholderCellValue implements PlaceholderCellValue {
    private readonly placeholder: string;
    private readonly mergeCellPlaceholder?: string;

    constructor(p: string, merge?: string) {
        this.placeholder = p;
        this.mergeCellPlaceholder = merge === undefined ? "" : merge;
    }

    /** @inheritdoc */
    mergeCell(values: string[]): string {
        if (this.mergeCellPlaceholder !== undefined) {
            return this.mergeCellPlaceholder.replace("?", values.join(","));
        }
        return "";
    }

    /** @inheritdoc */
    toString(): string {
        return this.placeholder;
    }
}

// ==================== 内部辅助函数 ====================

/** 获取范围单元格中的所有坐标点 */
const _getCells = (thisArg: RangeCell): CellPoint[] => {
    let cells: CellPoint[] = [];
    for (let j = thisArg.minColumn; j <= thisArg.maxColumn; j += thisArg.stepColumn) {
        for (let i = thisArg.minRow; i <= thisArg.maxRow; i += thisArg.stepRow) {
            cells.push({
                Row: i,
                Column: j,
            });
        }
    }
    return cells;
};

// ==================== 令牌解析管理器 ====================

/**
 * 令牌解析管理器
 * 提供所有默认的令牌解析策略
 */
export class TokenParserManger {

    /**
     * 别名解析器
     * 处理 T=xxx 格式的别名配置
     * @example T=xxx.xxx.Tell
     */
    static aliasParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.AliasToken) {
            return {
                ok: false
            }
        }
        const parser = getTokenParser(RuleToken.EqualToken);
        const {values, expr, ok} = parser.handler(ctx, RuleToken.EqualToken, value);
        if (!ok) {
            return {
                ok: false,
            }
        }
        if (typeof expr.value !== "string" || expr.value === "") {
            return {
                ok: false,
                error: new Error(`alias express right value cannot be a empty value.`),
            }
        }
        return {
            ok: true,
            expr: {
                express: value,
                key: expr.key, // alias key
                value: expr.value, // alias value
                tokens: [token, RuleToken.EqualToken], //  express tokens
            },
            values: values,
        }
    }

    /**
     * 等号解析器
     * 处理 x=xx 格式的赋值表达式
     */
    static equalParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.EqualToken) {
            return {
                ok: false,
            }
        }
        const equalToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.EqualToken);
        const offset = equalToken.length;
        const index = value.indexOf(equalToken);
        if (index < 0) {
            return {
                ok: false,
            }
        }
        const len = value.length;
        const key = value.substring(0, index);
        const rightValue = value.substring(index + offset, len);
        if (rightValue === "") {
            return {
                ok: false,
                error: new Error(`equal express right value cannot be a empty value.`),
            }
        }
        const expr: RuleValue = {
            key: key,
            value: rightValue,
            express: value,
            tokens: [token],
        }
        return {
            ok: true,
            values: [key, rightValue],
            expr,
        }
    }

    /**
     * 单元格解析器
     * 处理 cell| X:Y=${?} 格式的单元格变量配置
     */
    static cellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.CellToken) {
            return {
                ok: false
            }
        }
        const equalToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.EqualToken);
        const eqOffset = equalToken.length;
        const posToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.PosToken);
        const eqIndex = value.indexOf(equalToken);
        const posIndex = value.indexOf(posToken);
        if (eqIndex < 0 || posIndex < 0) {
            return {
                ok: false,
            }
        }
        const parser = getTokenParser(RuleToken.PosToken);
        const varParser = getTokenParser(RuleToken.VarPatternToken);
        const posValue = value.substring(0, eqIndex);
        const varValue = value.substring(eqIndex + eqOffset);
        const posReply = parser.handler(ctx, RuleToken.PosToken, posValue);
        if (!posReply.ok) {
            return {
                ok: false,
                ...posReply,
            }
        }
        const varReply = varParser.handler(ctx, RuleToken.VarPatternToken, varValue);
        if (!varReply.ok) {
            return {
                ok: false,
                ...varReply,
            }
        }
        const expr: RuleValue = {
            ...varReply.expr,
        };
        expr.express = value
        expr.cells = posReply.expr.cells
        expr.value = varReply.expr.value // var expr value
        expr.ref = varReply.expr.ref // alias refs
        expr.tokens = [token, RuleToken.EqualToken, ...posReply.expr.tokens, ...varReply.expr.tokens]// token
        return {
            ok: true,
            expr: expr,
            values: [posReply.values, varReply.values],
        }
    }

    /**
     * 使用别名解析器
     * 处理 @X.test 或 @X.@T 格式的别名使用表达式
     */
    static useAliasParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.UseAliasToken) {
            return {
                ok: false,
            }
        }
        const useAliasToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.UseAliasToken);
        const offset = useAliasToken.length;
        const has = offset !== 0 && value.indexOf(useAliasToken) >= 0;
        if (!has) {
            return {
                ok: false,
            }
        }
        const endTokens = [RuleToken.UseAliasToken, RuleToken.LparenToken, RuleToken.ArgPosToken, RuleToken.DotGetToken];
        const values = TokenParserManger.scanToken(value, useAliasToken, TokenParserManger.toList(ctx, endTokens));
        if (values === undefined || values.length <= 0) {
            return {
                ok: false,
            }
        }
        const keys: string[] = [];
        const tokens: RuleToken[] = [];
        const expr: RuleValue = {
            express: value,
            tokens: tokens,
            value: keys,
            ref: [],
        }
        for (const v of values) {
            tokens.push(token);
            keys.push(v.token);
            expr.ref.push(v.value);
        }
        expr.value = keys
        expr.tokens = tokens
        return {
            expr,
            ok: true,
            values: values,
        }
    }

    /**
     * 范围解析器
     * 处理 A-AA 或 A-AZ 或 1-7 或 1-7,2 [step] 或 A-Z,2 [step] 格式的范围表达式
     */
    static rangeParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.RangeToken) {
            return {
                ok: false,
            }
        }
        const rangeToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.RangeToken);
        const offset = rangeToken.length;
        const index = value.indexOf(rangeToken);
        if (index < 0 || offset <= 0) {
            return {
                ok: false,
            }
        }
        let setup = 1;
        let startNumber: number = NaN;
        let endNumber: number = NaN;
        let startPos = value.substring(0, index).trim();
        let endPos = value.substring(index + offset).trim();
        const argPosToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.ArgPosToken);
        const endSetupIndex = endPos.indexOf(argPosToken);
        if (endSetupIndex > 0) {
            setup = Number.parseInt(endPos.substring(endSetupIndex).trim(), 10);
            endPos = endPos.substring(0, endSetupIndex).trim()
        }
        if (isNaN(setup)) {
            return {
                ok: false,
                error: new Error(`rangeToken parse setup NaN, ${value}`),
            }
        }
        endNumber = TokenParserManger.parsePosNumber(endPos);
        startNumber = TokenParserManger.parsePosNumber(startPos);
        if (isNaN(startNumber) || isNaN(endNumber)) {
            return {
                ok: false,
                error: new Error(`rangeToken parse start,end has NaN, ${value}`),
            }
        }
        const expr: RuleValue = {
            express: value,
            tokens: [token],
            value: [startNumber, endNumber, setup],
        };
        return {
            expr,
            ok: true,
            values: [startNumber, endNumber, setup],
        }
    }

    /**
     * 位置解析器
     * 处理 A:1 或 A-Z:1 或 A:1-10 或 A-AA:2-10 格式的位置表达式
     */
    static posParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.PosToken) {
            return {
                ok: false
            }
        }
        const posToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.PosToken);
        const offset = posToken.length;
        const index = value.indexOf(posToken);
        const len = value.length;
        const column = value.substring(0, index).trim(); // column
        const row = value.substring(index + offset, len).trim(); // row
        const rangeToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.RangeToken);
        const columnRange = column.indexOf(rangeToken);
        const rowRange = row.indexOf(rangeToken);
        if (rowRange > 0 || columnRange > 0) {
            return TokenParserManger.parseRangeValue(ctx, {rowRange, columnRange, row, column, token, express: value})
        }
        const cell: CellPoint = {
            Row: Number.parseInt(row, 10),
            Column: columnLetterToNumber(column),
        };
        const expr: RuleValue = {
            value: cell,
            cells: [cell],
            express: value,
            tokens: [token],
        }
        return {
            ok: true,
            values: [row, column],
            expr,
        }
    }

    /**
     * 合并单元格解析器
     * 处理 mergeCell 和 rowCell 类型的大范围生成表达式
     * @example A-AQ:13-15=<sum(#,[compile:Macro(exprArr,[F],[13,15],!codeKey)],compile:Marco(index),0)>
     */
    static mergeCellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.MergeCellToken && token !== RuleToken.RowCellToken) {
            return {ok: false};
        }
        const equalToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.EqualToken);
        const index = value.indexOf(equalToken);
        const offset = equalToken.length;
        if (index <= 0) return {ok: false, error: new Error(`merge cell config syntax error: ${value}`)};
        const rangeStr = value.substring(0, index).trim();
        const exprStr = value.substring(index + offset)
        const posParser = getTokenParser(RuleToken.PosToken);
        const posReply = posParser.handler(ctx, RuleToken.PosToken, rangeStr);
        const functionToken = getTokenParser(RuleToken.FunctionPatternToken);
        if (!posReply.ok) return {ok: false, ...posReply};
        const macroGenToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.CompileGenToken);
        const expr: RuleValue = {
            express: value,
            value: exprStr,
            posExpr: posReply.expr,
            tokens: [token, ...posReply.expr.tokens],
        };
        if (exprStr.startsWith(macroGenToken)) {
            const argsSplitToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.ArgPosToken);
            const args = TokenParserManger.split(exprStr, argsSplitToken, [`(`, `)`]);
            const macro = TokenParserManger.filterMacro(args);
            const aliasTokens = TokenParserManger.extractUseAliasTokens(ctx, args);
            if (aliasTokens !== undefined && aliasTokens.length > 0) {
                expr.tokens.push(...aliasTokens);
            }
            if (macro !== undefined && macro.tokens.length > 0) {
                expr.macro = macro;
                expr.tokens.push(...macro.tokens);
            }
        } else {
            const exprReply = functionToken.handler(ctx, RuleToken.FunctionPatternToken, exprStr);
            if (exprReply.error !== undefined && exprReply.error instanceof Error) {
                return {
                    ok: false,
                    ...exprReply,
                }
            }
            if (exprReply.ok && exprReply.expr !== undefined) {
                expr.funcExpr = exprReply.expr
                expr.tokens.push(...exprReply.expr.tokens);
            }
        }
        return {
            ok: true,
            expr,
            values: [rangeStr, exprStr],
        };
    }

    /**
     * 行单元格解析器
     * 处理 rowCell 类型的行生成表达式
     * @example G-AQ:12=compile:GenCell(compile:Macro(expr,F,12),'.',compile:Marco(index))
     */
    static rowCellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.RowCellToken) {
            return {ok: false};
        }
        return TokenParserManger.mergeCellParse(ctx, token, value);
    }

    /**
     * 函数模式解析器
     * 解析 $functionName($arg0:string,$arg1:string[],...)
     * @example sum(#,[F,12,13],1,0) => {func:"sum",arguments:["A",["F","12","13"],"1","0"]}
     */
    static functionPatternParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.FunctionPatternToken) {
            return {ok: false};
        }
        const wordToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.AnyToken);
        const funcToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.FunctionPatternToken);
        const splitIndex = funcToken.indexOf(wordToken);
        const splitOffset = wordToken.length;
        const funcStartToken = funcToken.substring(0, splitIndex);
        const funcEndToken = funcToken.substring(splitIndex + splitOffset);
        const rparenToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.RparenToken);
        const lparenToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.LparenToken);
        const argsSplitToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.ArgPosToken);
        if (value.startsWith(funcStartToken) && value.endsWith(funcEndToken)) {
            const content = value.substring(1, value.length - 1);
            const lparen = content.indexOf(lparenToken);
            const rparen = content.lastIndexOf(rparenToken);
            if (lparen > 0 && rparen > lparen) {
                const funcName = content.substring(0, lparen);
                const argsStr = content.substring(lparen + lparenToken.length, rparen);
                // Basic argument splitting, does not handle nested commas deeply here
                const args = TokenParserManger.split(argsStr, argsSplitToken, [`[`, `]`]);
                const macro = TokenParserManger.filterMacro(args);
                const tokens: RuleToken[] = [token];
                const expr: RuleValue = {
                    express: value,
                    func: funcName,
                    value: args, // arguments array
                    tokens: tokens,
                };
                const alias = TokenParserManger.extractUseAliasTokens(ctx, args);
                if (alias !== undefined && alias.length > 0) {
                    expr.tokens.push(...alias);
                }
                if (macro !== undefined && macro.tokens.length > 0) {
                    expr.macro = macro;
                    expr.tokens.push(...macro.tokens)
                }
                return {
                    ok: true,
                    expr: expr,
                    values: [funcName, args],
                };
            }
        }
        return {
            ok: false,
            error: new Error(`function express systax error`),
        };
    }

    /**
     * 变量模式解析器
     * 处理 ${?} 或 ${xxx} 格式的变量表达式
     */
    static varPatternParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.VarPatternToken) {
            return {ok: false};
        }
        const wordToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.AnyToken);
        const varToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.VarPatternToken);
        const index = varToken.indexOf(wordToken);
        const workTokenOffset = wordToken.length;
        const startToken = varToken.substring(0, index);
        const endToken = varToken.substring(workTokenOffset + index);
        if (!value.startsWith(startToken) || !value.endsWith(endToken)) {
            return {
                ok: false,
                error: new Error(`variable expression syntax error,\$\{ or \} flag is miss`),
            };
        }
        const innerContent = value.substring(startToken.length, value.length - endToken.length);
        if (innerContent === "") {
            return {
                ok: false,
                error: new Error("variable expression syntax error, variable name is empty"),
            }
        }
        const expr: RuleValue = {
            express: value,
            value: innerContent,
            tokens: [token],
        };
        const aliasToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.UseAliasToken);
        if (innerContent.indexOf(aliasToken) >= 0) {
            const parser = getTokenParser(RuleToken.UseAliasToken);
            const aliasReply = parser.handler(ctx, RuleToken.UseAliasToken, innerContent);
            if (aliasReply.ok) {
                expr.ref = aliasReply.expr.ref
                expr.alias = aliasReply.values
                expr.tokens.push(...aliasReply.expr.tokens);
            }
        }
        return {
            expr,
            ok: true,
            values: [innerContent],
        };
    }

    /**
     * 编译表达式提取器
     * 从表达式中提取 compile:GenCell 和 compile:Macro 令牌
     */
    static compileExprExtract(value: string): RuleToken[] {
        const results: RuleToken[] = [];
        const lp = RuleToken.LparenToken.toString();
        const args = RuleToken.ArgPosToken.toString();
        const values = value.split(lp);
        for (const v of values) {
            let items: string[] = [];
            if (v.indexOf(args) >= 0) {
                items = v.split(args);
            } else {
                items.push(v);
            }
            for (const it of items) {
                if (it === RuleToken.CompileMacroToken.toString() ||
                    it.startsWith(RuleToken.CompileMacroToken) ||
                    it.endsWith(RuleToken.CompileMacroToken)) {
                    results.push(RuleToken.CompileMacroToken);
                } else if (it === RuleToken.CompileGenToken.toString() ||
                    it.startsWith(RuleToken.CompileGenToken) ||
                    it.endsWith(RuleToken.CompileGenToken)) {
                    results.push(RuleToken.CompileGenToken);
                }
            }
        }
        return results;
    }

    /**
     * 从上下文中获取指定令牌的字符串表示
     */
    static getTokenByCtx(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken): string {
        if (ctx.size <= 0 || !ctx.has(token)) {
            return token.toString()
        }
        const values = ctx.get(token)
        if (values.length <= 0 || values[0].key === "") {
            return token.toString()
        }
        return values[0].key
    }

    /**
     * 扫描令牌
     * 从字符串中提取所有以 startToken 开头、以 endTokens 中任意一个结尾的令牌数据
     */
    static scanToken(value: string, startToken: string, endTokens: string[]): ScanTokenData[] {
        let token: string = "";
        let data: string = "";
        let end: boolean = false;
        let start: boolean = false;
        const items: ScanTokenData[] = [];
        const offset = [startToken, ...endTokens].sort((a, b) => a.length - b.length)[0].length;
        const size = value.length;
        for (let i = 0; i < size; i += offset) {
            let leftToken: string;
            if (value[i] === startToken) {
                start = true;
                if (token !== "") {
                    items.push({
                        token,
                        value: "",
                    })
                }
                token = startToken;
            } else {
                start = token === startToken;
                leftToken = `${token}${value[i]}`
            }
            let subfix = endTokens.filter(s => s === leftToken);
            end = endTokens.includes(value[i]) || subfix.length > 0;
            if (subfix.length > 0) {
                token = leftToken;
                data = "";
            }
            if (start && end) {
                continue;
            }
            if (!end) {
                token = `${token}${value[i]}`;
                data = `${data}${value[i]}`;
            }
            if (end || i + offset >= size) {
                if (data !== "") {
                    items.push({
                        token,
                        value: data,
                    });
                }
                token = "";
                data = "";
                end = false;
            }
        }
        return items;
    }

    /**
     * 分割字符串，支持忽略括号内的分隔符
     * @param argsStr - 待分割的字符串
     * @param argsSplitToken - 分隔符
     * @param ignoreTokenRange - 需要忽略的分隔范围（如括号）
     */
    static split(argsStr: string, argsSplitToken: string, ignoreTokenRange: string[]): string[] {
        let value: string = "";
        let depth = 0;
        const items: string[] = [];
        const splitLen = argsSplitToken.length;
        const startLen = ignoreTokenRange[0].length;
        const endLen = ignoreTokenRange[1].length;
        const startToken = ignoreTokenRange[0];
        const endToken = ignoreTokenRange[1];
        const isToggleMode = startToken === endToken;
        for (let i = 0; i < argsStr.length;) {
            const substr = argsStr.substring(i);
            if (depth > 0 && substr.startsWith(endToken)) {
                value += endToken;
                i += endLen;
                depth = isToggleMode ? 0 : depth - 1;
                continue;
            }
            if (substr.startsWith(startToken)) {
                if (!isToggleMode || depth === 0) {
                    value += startToken;
                    i += startLen;
                    depth++;
                    continue;
                }
            }
            if (depth === 0 && substr.startsWith(argsSplitToken)) {
                items.push(value);
                value = "";
                i += splitLen;
                continue;
            }
            value += argsStr[i];
            i++;
        }
        if (value !== "") {
            items.push(value);
        }
        return items;
    }

    /**
     * 过滤宏表达式
     * 从字符串数组中提取 compile:Macro 和 compile:GenCell 令牌
     */
    static filterMacro(values: string[]): FilterMacroResult | undefined {
        const filter: FilterMacroResult = {
            tokens: [],
            express: [],
        };
        for (const expr of values) {
            let items = TokenParserManger.compileExprExtract(expr);
            if (items === undefined || items.length <= 0) {
                continue;
            }
            filter.express.push(expr);
            filter.tokens.push(...items);
        }
        return filter.tokens.push() <= 0 ? undefined : filter;
    }

    /**
     * 提取 @ 别名令牌
     * 检查参数列表中是否包含 @ 开头的别名引用
     */
    static extractUseAliasTokens(ctx: Map<RuleToken, RuleValue[]>, args: string[]): RuleToken[] {
        const tokens: RuleToken[] = [];
        const useAliasToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.UseAliasToken);
        args.forEach(v => v.startsWith(useAliasToken) && tokens.push(RuleToken.UseAliasToken))
        return tokens;
    }

    /**
     * 将 RuleToken 列表转换为字符串列表
     */
    static toList(ctx: Map<RuleToken, RuleValue[]>, endTokens: RuleToken[]): string[] {
        const items: string[] = [];
        for (const token of endTokens) {
            items.push(TokenParserManger.getTokenByCtx(ctx, token));
        }
        return items;
    }

    /**
     * 解析范围值表达式
     * 处理行列范围组合的场景
     */
    private static parseRangeValue(ctx: Map<RuleToken, RuleValue[]>, options: RangeValueParserOptions): TokenParserReply {
        let rowReply: TokenParserReply;
        let columnReply: TokenParserReply;
        const rangeCell: RangeCell = {
            stepRow: 1,
            stepColumn: 1,
            minColumn: 0,
            minRow: 0,
            maxRow: 0,
            maxColumn: 0,
            getCells: function (): CellPoint[] {
                return _getCells(this);
            }
        };
        const {rowRange, row, column, columnRange, express, token} = options;
        const rangeParse = getTokenParser(RuleToken.RangeToken);
        if (rowRange > 0) {
            columnReply = rangeParse.handler(ctx, RuleToken.RangeToken, row);
        } else {
            const rowValue = Number.parseInt(row, 10);
            rangeCell.maxRow = rowValue;
            rangeCell.minRow = rowValue;
        }
        if (columnRange > 0) {
            rowReply = rangeParse.handler(ctx, RuleToken.RangeToken, column)
        } else {
            const columnValue = columnLetterToNumber(column);
            rangeCell.maxColumn = columnValue;
            rangeCell.minColumn = columnValue;
        }
        if (columnReply !== undefined) {
            if (!columnReply.ok || columnReply.values === undefined) {
                return columnReply
            }
            const values = columnReply.values as number[];
            let min = values[0];
            let max = values[1];
            let setup = values[2] ?? 1;
            rangeCell.minRow = min;
            rangeCell.maxRow = max;
            rangeCell.stepRow = setup;
        }
        if (rowReply !== undefined) {
            if (!rowReply.ok || rowReply.values === undefined) {
                return rowReply
            }
            const values = rowReply.values as number[];
            let min = values[0];
            let max = values[1];
            let setup = values[2] ?? 1;
            rangeCell.minColumn = min;
            rangeCell.maxColumn = max;
            rangeCell.stepColumn = setup;
        }
        const expr: RuleValue = {
            express: express,
            tokens: [token],
            value: rangeCell,
        }
        return {
            ok: true,
            expr,
            values: rangeCell,
        };
    }

    /**
     * 解析位置数值
     * 支持数字和字母两种格式
     */
    private static parsePosNumber(value: string): number {
        let num: number = NaN;
        if (isPureNumber.test(value)) {
            num = Number.parseInt(value, 10)
        } else if (isPureUppercase.test(value)) {
            num = columnLetterToNumber(value);
        }
        return !isNaN(num) && num <= 0 ? NaN : num;
    }

}


// ==================== 默认令牌解析器映射 ====================

/** 默认的规则令牌到解析函数的映射表 */
export const defaultRuleTokenParserMap = new Map<RuleToken, TokenParser>([
    [RuleToken.AliasToken, TokenParserManger.aliasParse],
    [RuleToken.CellToken, TokenParserManger.cellParse],
    [RuleToken.EqualToken, TokenParserManger.equalParse],
    [RuleToken.MergeCellToken, TokenParserManger.mergeCellParse],
    [RuleToken.RowCellToken, TokenParserManger.rowCellParse],
    [RuleToken.UseAliasToken, TokenParserManger.useAliasParse],
    [RuleToken.RangeToken, TokenParserManger.rangeParse],
    [RuleToken.PosToken, TokenParserManger.posParse],
    [RuleToken.VarPatternToken, TokenParserManger.varPatternParse],
    [RuleToken.FunctionPatternToken, TokenParserManger.functionPatternParse],
]);

/** 宏类型令牌列表 */
export const macroTokens: RuleToken[] = [RuleToken.CompileGenToken, RuleToken.CompileMacroToken,];

/**
 * 列字母转数字
 * 将 Excel 列字母标识转换为列号（A=1, B=2, ..., Z=26, AA=27, ...）
 * @param letter - 列字母标识
 */
export function columnLetterToNumber(letter: string): number {
    let num = 0;
    for (let i = 0; i < letter.length; i++) {
        num = num * 26 + (letter.charCodeAt(i) - 64);
    }
    return num;
}

/**
 * 列数字转字母
 * 将列号转换为 Excel 列字母标识（1=A, 2=B, ..., 26=Z, 27=AA, ...）
 * @param num - 列号
 */
export function columnNumberToLetter(num: number): string {
    if (num <= 0) {
        return "";
    }
    let letter = "";
    while (num > 0) {
        let remainder = (num - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        num = Math.floor((num - 1) / 26);
    }
    return letter;
}

/**
 * 判断字符串是否为 base64 编码
 * @param str - 待检查的字符串
 */
export function isBase64(str: string): boolean {
    if (str.length < 20) return false;
    if (str.includes("\\") || str.includes(" ")) return false;
    if (str.indexOf("./") >= 0 ||
        str.startsWith("file://") ||
        str.startsWith("/") ||
        str.startsWith(".")) {
        return false;
    }
    const data = str.replace(/^data:.*?;base64,/, "");
    const cleanedStr = data.replace(/[\r\n]/g, "");
    try {
        return Buffer.from(cleanedStr, 'base64').toString('utf-8') !== "";
    } catch (e) {
        return false;
    }
}

/**
 * base64 字符串转 ArrayBuffer
 * @param base64 - base64 编码的字符串
 */
export function base64ToArrayBuffer(base64: string): ArrayBuffer {
    const data = base64.replace(/^data:.*?;base64,/, "");
    const binaryString = atob(data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}

/**
 * 获取合并单元格范围
 * 检查指定行列是否在合并单元格内，并返回合并范围
 * @param ws - Excel 工作表
 * @param row - 行号
 * @param col - 列号
 * @returns 合并范围，如果不在合并单元格中则返回 null
 */
export function getMergeRange(ws: exceljs.Worksheet, row: number, col: number): MergeCellRange | null {
    const merges = ws.model.merges as string[];
    for (const mergeStr of merges) {
        const parts = mergeStr.split(":");
        const tl = ws.getCell(parts[0]);
        const br = ws.getCell(parts[1]);
        const tlRow = Number(tl.row)
        const tlCol = Number(tl.col)
        const brRow = Number(br.row)
        const brCol = Number(br.col)
        if (row >= tlRow && row <= brRow && col >= tlCol && col <= brCol) {
            return {
                top: tlRow,
                left: tlCol,
                bottom: brRow,
                right: brCol,
            };
        }
    }
    return null;
}

// ==================== 宏解析相关 ====================

/**
 * 解析编译宏生成表达式
 * 处理 compile:Gen 指令，将参数拼接成字符串
 */
export const resolveCompileMacroGen = (ctx: CompileContext, expr: string, currentCellIndex: number): string => {
    let parts: string[] = [];
    let join = ".";
    const m = ctx.getContextMap()
    const aliasToken = TokenParserManger.getTokenByCtx(m, RuleToken.UseAliasToken);
    const genToken = TokenParserManger.getTokenByCtx(m, RuleToken.CompileGenToken);
    const endTokens = [RuleToken.CompileGenToken, RuleToken.LparenToken, RuleToken.ArgPosToken, RuleToken.RparenToken];
    const values = TokenParserManger.scanToken(expr, genToken, TokenParserManger.toList(m, endTokens));
    for (const [_, item] of values.entries()) {
        if (item.value === undefined && item.value !== "") {
            continue
        }
        if (!item.token.startsWith(aliasToken) && item.value !== exprIndex) {
            parts.push(item.value);
        } else {
            parts.push(resolveAliasExpr(ctx, item.value, currentCellIndex));
        }
    }
    if (parts.length === 1) {
        return parts[0];
    }
    let end = parts[parts.length - 1];
    if (end.startsWith('"') && end.length >= 3 && end.endsWith('"')) {
        join = end;
        return parts.slice(0, parts.length - 1).join(join);
    }
    return parts.join(join);
};

/**
 * 获取宏表达式结束位置
 * 从指定位置开始查找右括号的位置
 */
export const getExprEnd = function (macroExpr: string, matchIndex: number, rparenToken: string): number {
    return macroExpr.indexOf(rparenToken, matchIndex);
};

/** 宏格式化器键名列表 */
export const macroFormatters: string[] = [numberKey, codeKey, codeAliasKey];

/**
 * 宏单元辅助函数类型
 */
export type MacroUnitHelper = (v: string, expr?: RuleValue) => string;

/**
 * 提取宏参数
 * 从 compile:Macro(type, X_arr, Y_arr, formatter) 表达式中提取参数
 */
export const extractMacro = function (expr: string, options: ExtractMacroArgs): MacroArgs {
    let end = NaN;
    const offset = options.startToken.length;
    const startIndex = expr.indexOf(options.startToken);
    const endIndex = expr.indexOf(options.endToken);
    const argValues = expr.substring(startIndex + offset, endIndex);
    const values = argValues.split(options.argToken);
    const extracResult: MacroArgs = {type: values[0], rowParam: [], columnParam: undefined as number};
    if (values.length > 1) {
        extracResult.columnParam = columnLetterToNumber(values[1])
    }
    if (macroFormatters.includes(values[values.length - 1])) {
        end = values.length - 1;
        extracResult.formatter = values[end]
    }
    if (!isNaN(end) && values.length > 2) {
        extracResult.rowParam = values.slice(2, end).map(x => Number.parseInt(x, 10));
    } else if (values.length > 2) {
        extracResult.rowParam = values.slice(2, values.length).map(x => Number.parseInt(x, 10));
    }
    if (extracResult.rowParam !== undefined && extracResult.rowParam instanceof Array && extracResult.rowParam.length === 1) {
        extracResult.rowParam = extracResult.rowParam[0];
    }
    return extracResult;
};

// ==================== 宏格式化器实现 ====================

/**
 * 代码键格式化器
 * 将字符串转换为大写且替换特殊字符为下划线，用于生成代码键
 */
export const __codeKey: MacroUnitHelper = (str: string, expr?: RuleValue): string => {
    const replaces: string[] = [" ", `-`, `/`, `,`, `'`, `&`, `.`, `(`, `)`, `{`, `}`, `@`, `\\`, `[`, `]`, `#`, `:`];
    for (const k of replaces) {
        str = str.replaceAll(k, "_").trim();
        if (str.indexOf("__") >= 0) {
            str = str.replaceAll("__", "_")
        }
    }
    if (str.startsWith("_")) {
        str = str.substring(1)
    }
    if (str.endsWith("_")) {
        str = str.substring(0, str.length - 1)
    }
    return str.toUpperCase();
};

/**
 * 数字键格式化器
 * 将字符串转换为数字值，如果非数字则返回原字符串
 */
export const __numberKey: MacroUnitHelper = (str: string, expr?: RuleValue): string => {
    if(str === "NaN" ||
        str === "Infinity" ||
        str === "null" ||
        str === "[object object]"){
        return str
    }
    if (str.startsWith("0x")) {
        return str.substring(2)
    }
    let v = Number.parseInt(str, 10);
    if (isNaN(v)) {
        return str;
    }
    return v.toString()
};

/**
 * 代码别名键格式化器
 * 将字符串转换为代码键并添加 @ 前缀，用于别名引用
 */
export const __codeAliasKey: MacroUnitHelper = (str: string, expr?: RuleValue): string => {
    const key = __codeKey(str);
    if (key !== "") {
        if (expr !== undefined && expr.tokens.length > 0) {
            expr.tokens.push(RuleToken.UseAliasToken);
        }
        return `${defaultRuleTokenMap.get(RuleToken.UseAliasToken)}${key}`;
    }
    return '';
};

/** 宏格式化器映射表 */
export const macroFormatter: Map<string, MacroUnitHelper> = new Map<string, MacroUnitHelper>([
    [codeKey, __codeKey],
    [numberKey, __numberKey],
    [codeAliasKey, __codeAliasKey],
    [defaultKey, (v: string): string => v],
]);

/**
 * 执行宏格式化
 * 根据格式化器名称对值进行格式化处理
 */
export const execMacroFormat = function (value: string, formatter: string, expr?: RuleValue): string {
    if (!macroFormatter.has(formatter)) {
        return value;
    }
    return macroFormatter.get(formatter)(value, expr)
};

// ==================== 单元格操作函数 ====================

/**
 * 将行值数组展开为实际行号列表
 * 如果参数是 [start, end] 格式的数组，则生成 start 到 end 的所有行号
 */
export const toCellRow = (rowVals: number[], setup?: number): number[] => {
    if (setup === undefined) {
        setup = 1;
    }
    if (rowVals.length == 2) {
        const values: number[] = [];
        for (let start = rowVals[0]; start <= rowVals[1]; start += setup) {
            values.push(start);
        }
        return values;
    }
    return rowVals;
};

/**
 * 将列值数组展开为实际列号列表
 * 功能同 toCellRow
 */
export const toCellColumn = (columnVals: number[], setup?: number): number[] => {
    return toCellRow(columnVals, setup);
};

/**
 * 获取单元格值的字符串表示
 * 支持普通值、富文本值和超链接值
 */
export const toCellValue = (value: exceljs.CellValue): string => {
    if (typeof value !== "string") {
        const rText = value as exceljs.CellRichTextValue;
        if (rText !== undefined && rText.richText !== undefined && rText.richText.length > 0) {
            const values: string[] = [];
            for (const [_, v] of rText.richText.entries()) {
                if (v === undefined || v === null) {
                    continue;
                }
                const text = v.text !== undefined && v.text !== null ? v.text : undefined;
                if (text !== undefined && text !== "" && text !== '[object]') {
                    values.push(text.trim());
                }
            }
            return values.join(" ");
        }
        const hText = value as exceljs.CellHyperlinkValue;
        if (hText !== undefined && hText.text !== null && hText.hyperlink !== null) {
            return `[${hText.text}](${hText.hyperlink})`;
        }
    }
    return value.toString();
};

// ==================== 核心宏解析函数 ====================

/**
 * 核心宏解析函数
 * 支持：
 * 1. compile:Macro(expr/exprArr, [X], [Y], !formatter) - 宏展开为单元格值
 * 2. compile:Gen(expr, expr1, ...) - 参数拼接生成
 */
export const resolveCompileMacroExpr = (ctx: CompileContext, macroExpr: string, macroTokens: RuleToken[], currentCellIndex: number, totalCells: number): string => {
    if (macroTokens === undefined || macroTokens.length <= 0) {
        return macroExpr;
    }

    const sheet = ctx.sheet;
    if (sheet === undefined) {
        throw new Error(`miss context worksheet`);
    }
    const m = ctx.getContextMap();
    const argToken = TokenParserManger.getTokenByCtx(m, RuleToken.ArgPosToken);
    const rparenToken = TokenParserManger.getTokenByCtx(m, RuleToken.RparenToken);
    const lparenToken = TokenParserManger.getTokenByCtx(m, RuleToken.LparenToken);
    const genToken = TokenParserManger.getTokenByCtx(m, RuleToken.CompileGenToken);
    const maroToken = TokenParserManger.getTokenByCtx(m, RuleToken.CompileMacroToken);
    const tokenMap: Map<RuleToken, number[]> = new Map<RuleToken, number[]>();
    for (const [index, token] of macroTokens.entries()) {
        let items: number[] = [];
        if (tokenMap.has(token)) {
            items = tokenMap.get(token);
        }
        items.push(index);
        tokenMap.set(token, items);
    }
    const resolveArray = (param: number | number[]): number[] => {
        if (param instanceof Array) {
            return param;
        }
        return [param];
    };
    if (tokenMap.has(RuleToken.CompileMacroToken)) {
        let offset: number = 0;
        let exprValue: string = macroExpr;
        const times = tokenMap.get(RuleToken.CompileMacroToken).length;
        for (let i = times; i > 0; i--) {
            const matchIndex: number = exprValue.indexOf(maroToken, offset);
            if (matchIndex < 0) {
                break;
            }
            offset = getExprEnd(exprValue, matchIndex, rparenToken);
            const macroCurrent = exprValue.substring(matchIndex, offset + rparenToken.length);
            const opts: ExtractMacroArgs = {startToken: lparenToken, endToken: rparenToken, argToken};
            let {type, columnParam, rowParam, formatter} = extractMacro(macroCurrent, opts);
            let rowVals: number[];
            let columnVals: number[];
            const parts: string[] = [];
            if (columnParam !== undefined && rowParam !== undefined) {
                rowVals = resolveArray(rowParam);
                columnVals = resolveArray(columnParam);
                const rowItems = toCellRow(rowVals);
                const columnItems = toCellColumn(columnVals);
                rowItems.forEach(r => {
                    columnItems.forEach(c => {
                        const cellValue = sheet.findCell(r, c);
                        if (cellValue === undefined || cellValue.value === null) {
                            return;
                        }
                        const value = toCellValue(cellValue.value);
                        let exprValue = execMacroFormat(value, formatter, ctx.currentExpr || undefined);
                        parts.push(exprValue);
                    });
                });
            }
            if (type === exprArr) {
                macroExpr = macroExpr.replace(macroCurrent, parts.join(','))
            } else if (type === exprSingle) {
                macroExpr = macroExpr.replace(macroCurrent, parts[0])
            } else if (type === exprIndex) {
                const indexValue = currentCellIndex + 1;
                macroExpr = macroExpr.replace(macroCurrent, indexValue.toString())
            }
        }
    }

    if (tokenMap.has(RuleToken.CompileGenToken)) {
        let exprValue: string = macroExpr;
        const times = tokenMap.get(RuleToken.CompileGenToken).length;
        for (let i = times; i > 0; i--) {
            const matchIndex: number = exprValue.indexOf(genToken);
            if (matchIndex < 0) {
                break;
            }
            const offset = getExprEnd(exprValue, matchIndex, rparenToken);
            const macroCurrent = exprValue.substring(matchIndex, offset + rparenToken.length);
            exprValue = resolveCompileMacroGen(ctx, macroCurrent, currentCellIndex);
        }
        macroExpr = exprValue;
    } else {
        macroExpr = resolveAliasExpr(ctx, macroExpr, currentCellIndex);
    }

    return macroExpr;
};

// ==================== 工作簿加载 ====================

/**
 * 加载 Excel 工作簿
 * 支持多种数据源：文件路径、base64 字符串、Stream、ArrayBuffer、Buffer
 * @param data - Excel 数据源
 */
export const loadWorkbook = async function <T extends ArrayBuffer | Buffer | string>(data: T): Promise<exceljs.Workbook> {
    const w = new exceljs.Workbook();
    if (typeof data === "string") {
        if (!isBase64(data)) {
            await w.xlsx.readFile(data);
        } else {
            await w.xlsx.load(base64ToArrayBuffer(data));
        }
    } else if (data instanceof Stream) {
        await w.xlsx.read(data);
    } else if (data instanceof ArrayBuffer) {
        await w.xlsx.load(data);
    } else if (data instanceof Buffer) {
        await w.xlsx.load(data as any);
    } else {
        throw new Error(`unSupport buffer type ${typeof data}`);
    }
    return w;
};

// ==================== 占位符设置 ====================

/**
 * 扫描并设置单元格占位符
 * @param excelBuffer - Excel 数据
 * @param cell - 目标单元格位置
 * @param placeholder - 占位符值
 */
export const scanCellSetPlaceholder = async function <T extends ArrayBuffer | Buffer | string>(
    excelBuffer: T,
    cell: CellPosition & { Sheet: string | number },
    placeholder: PlaceholderCellValue
): Promise<ArrayBuffer | undefined> {
    const workbook = await loadWorkbook(excelBuffer);
    const worksheet = workbook.getWorksheet(cell.Sheet);
    if (!worksheet) return undefined;
    workSheetSetPlaceholder(worksheet, cell, placeholder)
    return workbook.xlsx.writeBuffer()
};

/**
 * 在工作表中设置占位符
 * 支持合并单元格的处理
 */
export const workSheetSetPlaceholder = function (worksheet: exceljs.Worksheet, cell: CellPosition, placeholder: PlaceholderCellValue): exceljs.Worksheet {
    const colNum = columnLetterToNumber(cell.Row);
    const rowNum = cell.Column;
    const targetCell = worksheet.getCell(rowNum, colNum);
    if (targetCell.isMerged) {
        const range = getMergeRange(worksheet, rowNum, colNum);
        if (range) {
            const leftCol = colNum - 1;
            const values: string[] = [];
            if (leftCol > 0) {
                for (let r = range.top; r <= range.bottom; r++) {
                    const val = worksheet.getCell(r, leftCol).value;
                    if (val !== null && val !== undefined && val !== "") {
                        values.push(String(val));
                    }
                }
            }
            if (values.length === 0) {
                targetCell.value = placeholder.toString();
            } else {
                targetCell.value = placeholder.mergeCell(values);
            }
            return worksheet;
        }
    }
    const currentValue = targetCell.value;
    if (currentValue === null || currentValue === undefined || currentValue === "") {
        targetCell.value = placeholder.toString();
    }
    return worksheet;
};

// ==================== 规则令牌相关 ====================

/** 编译单元格令牌列表（需生成单元格的令牌类型） */
export const compileCellTokens: RuleToken[] = [RuleToken.CellToken, RuleToken.MergeCellToken, RuleToken.RowCellToken];
/** 规则令牌列表（所有可配置的令牌类型） */
export const ruleTokens: RuleToken[] = [RuleToken.AliasToken, RuleToken.CellToken, RuleToken.MergeCellToken, RuleToken.RowCellToken];

/**
 * 判断是否为规则令牌
 * 检查给定令牌是否在 ruleTokens 列表中
 */
export const isRuleToken = function (t: RuleToken): boolean {
    return ruleTokens.includes(t)
};

/**
 * 合并规则映射
 * 将默认令牌映射合并到现有映射中（不覆盖已存在的键）
 */
export const mergeOption = function (ruleKeyMap: Map<RuleToken, string>, defaultRuleTokenMap: Map<RuleToken, string>): Map<RuleToken, string> {
    for (const [key, value] of defaultRuleTokenMap.entries()) {
        if (!ruleKeyMap.has(key)) {
            ruleKeyMap.set(key, value)
        }
    }
    return ruleKeyMap;
};

/**
 * 获取令牌解析器
 * 从默认解析器映射中查找指定令牌的解析函数
 */
export const getTokenParser = function (token: RuleToken): TokenParseResolver {
    if (!defaultRuleTokenParserMap.has(token)) {
        return {
            exists: false,
        }
    }
    return {
        exists: true,
        handler: defaultRuleTokenParserMap.get(token),
    }
};

/**
 * 注册令牌解析器（不覆盖已有）
 * 如果令牌已注册则返回 false
 */
export const registerTokenParser = function (token: RuleToken, h: TokenParser): boolean {
    if (defaultRuleTokenParserMap.has(token)) {
        return false;
    }
    defaultRuleTokenParserMap.set(token, h)
    return true;
};

/**
 * 注册令牌解析器（强制覆盖）
 * 无论令牌是否已注册，都强制设置解析函数
 */
export const registerTokenParserMust = function (token: RuleToken, h: TokenParser): void {
    defaultRuleTokenParserMap.set(token, h)
};

// ==================== 工作表规则扫描与解析 ====================

/**
 * 扫描工作表中的规则配置
 * 遍历工作表的指定行范围，解析规则令牌和表达式
 */
export const scanWorkSheetRules = function (worksheet: exceljs.Worksheet, options: RuleOptions): RuleResult {
    const result = {rules: options.getContextMap()};
    for (let r = options.startLine; r <= options.endLine; r++) {
        let emptyValue = false;
        let ruleToken = RuleToken.UndefinedToken;
        for (let c = options.startColumn; c <= options.endColumn; c++) {
            const cell = worksheet.findCell(r, c);
            if (cell === undefined || cell.value === undefined || cell.value === null) {
                continue;
            }
            const cellValue = cell.value;
            const value = cellValue.toString();
            let isStartCell = c === options.startColumn;
            if (emptyValue && ruleToken === RuleToken.UndefinedToken) {
                isStartCell = true;
            }
            if (value === "" && isStartCell) {
                emptyValue = true;
                continue;
            }
            if (isStartCell) {
                ruleToken = options.parseToken(value);
                continue;
            } else {
                if (!isRuleToken(ruleToken)) {
                    break;
                }
                emptyValue = false;
            }
            const {handler, exists} = getTokenParser(ruleToken);
            if (!exists) {
                continue;
            }
            let values: RuleValue[] = [];
            if (result.rules.has(ruleToken)) {
                values = result.rules.get(ruleToken);
            }
            const {expr, ok} = handler(result.rules, ruleToken, value);
            if (ok && expr !== undefined) {
                values.push(expr);
            }
            result.rules.set(ruleToken, values);
        }
    }
    return result;
};

/**
 * 解析工作表中的规则
 * 使用给定的选项解析工作表中的规则配置
 */
export const parseWorkSheetRules = function (worksheet: exceljs.Worksheet, options?: RuleOptions): RuleResult {
    const result: RuleResult = {rules: new Map<RuleToken, RuleValue[]>()};
    if (worksheet === undefined || worksheet === null) {
        return result;
    }
    if (options === undefined) {
        options = new RuleMapOptions();
    }
    return scanWorkSheetRules(worksheet, options.parseDefault(worksheet));
};

// ==================== 编译检查 ====================

/**
 * 编译检查
 * 检查规则配置的合法性，包括别名去重和自定义检查器
 */
export const compileCheck = function (iv: RuleResult, ctx: RuleOptions): Error[] | undefined {
    if (iv.rules.size <= 0) {
        return undefined;
    }
    const errs: Error[] = [];
    const values = iv.rules.get(RuleToken.AliasToken);
    if (values !== undefined && values.length > 0) {
        const set = new Set<string>();
        for (const [idx, value] of values.entries()) {
            if (idx === 0) {
                set.add(value.key);
                continue
            }
            if (set.has(value.key)) {
                errs.push(Error(`Duplicate alias(${value.key},${value.express}) configuration`));
            }
        }
    }
    const handlers = ctx.getCompileCheckHandlers();
    if (handlers !== undefined && handlers.length > 0) {
        for (const h of handlers) {
            let err = h(iv, ctx);
            if (err !== undefined && err.length > 0) {
                errs.push(...err);
            }
        }
    }
    if (errs.length > 0) {
        return errs;
    }
    return undefined;
};

/**
 * 获取宏令牌
 * 从表达式中提取所有宏类型的令牌
 */
export const getMacroTokens = function (expr: RuleValue): RuleToken[] {
    const tokens: RuleToken[] = [];
    for (const token of expr.tokens) {
        if (macroTokens.includes(token)) {
            tokens.push(token);
        }
    }
    return tokens;
};

// ==================== 行单元格处理 ====================

/**
 * 将单元格坐标点按行分组
 * 将一维的 CellPoint 数组按行号分组为二维数组
 */
export const toRowCells = function (cells: CellPoint[]): CellPoint[][] {
    const indexes: number[] = [];
    const rows: CellPoint[][] = [];
    const rowMap: Map<number, CellPoint[]> = new Map<number, CellPoint[]>();
    for (const cell of cells) {
        let cells: CellPoint[] = [];
        if (rowMap.has(cell.Row)) {
            cells = rowMap.get(cell.Row);
        } else {
            indexes.push(cell.Row);
        }
        cells.push(cell);
        rowMap.set(cell.Row, cells);
    }
    indexes.sort((a, b) => a - b);
    for (const row of indexes) {
        const values = rowMap.get(row);
        rows.push(values);
    }
    return rows;
};

// ==================== 表达式解析 ====================

/**
 * 解析函数表达式
 * 将模板中的函数模式替换为函数命令前缀
 */
export const resolveFunctionExpr = (ctx: CompileContext, templateValue: string, expr: RuleValue): string => {
    const anyToken = defaultRuleTokenMap.get(RuleToken.AnyToken);
    const funToken = defaultRuleTokenMap.get(RuleToken.FunctionPatternToken);
    const functionTokens = expr.tokens.filter(s => s === RuleToken.FunctionPatternToken);
    if (functionTokens === undefined || functionTokens.length <= 0) {
        return templateValue;
    }
    const [start, end] = funToken.split(anyToken);
    for (let times = functionTokens.length; times > 0; times--) {
        const index = templateValue.indexOf(start);
        const offset = templateValue.indexOf(end);
        if (offset > 0 && index >= 0) {
            templateValue = templateValue.replace(start, funcCommand).replace(end, "");
        }
    }
    return templateValue;
};

/**
 * 搜索字符串中的子串
 * 在字符串中搜索多个子串中最先出现的位置
 */
export const searchIndexOf = (str: string, substr: string[], position?: number): number => {
    let index: number = NaN;
    for (let sub of substr) {
        if (position === undefined || position === null) {
            index = str.indexOf(sub);
        } else {
            index = str.indexOf(sub, position);
        }
        if (!isNaN(index) && index >= 0) {
            return index;
        }
    }
    return -1;
};

/**
 * 解析别名表达式
 * 将模板中的 @ 别名引用替换为实际的别名值
 */
export const resolveAliasExpr = (ctx: CompileContext, templateValue: string, index: number): string => {
    const expr = ctx.currentExpr as RuleValue;
    if (expr === undefined) {
        throw new Error(`miss context expr value`);
    }
    let compileValue = templateValue;
    const aliasToken = defaultRuleTokenMap.get(RuleToken.UseAliasToken);
    let aliasTokens = expr.tokens.filter(s => s === RuleToken.UseAliasToken);
    if (aliasTokens.length <= 0) {
        const num = templateValue.split(aliasToken).length - 1;
        if (num > 0) {
            aliasTokens = new Array(num).fill(aliasToken);
        }
    }
    if ((aliasTokens !== undefined && aliasTokens.length > 0)) {
        for (let i = 0; i < aliasTokens.length; ++i) {
            const token = aliasToken;
            const start = compileValue.indexOf(token);
            if (start < 0) {
                break;
            }
            const offset = aliasTokens[i].length;
            let end = searchIndexOf(compileValue, [',', '.', ')', ']'], start);
            if (end < 0) {
                end = compileValue.length
            }
            const sv = compileValue.substring(start + offset, end);
            const pl = `${token}${sv}`;
            const value = ctx.getAlias(sv);
            if (value !== undefined) {
                if (compileValue === pl) {
                    compileValue = value;
                } else {
                    compileValue = compileValue.replace(`${pl}`, value);
                }
            }
        }
    }
    return compileValue;
};

/**
 * 解析值表达式
 * 将值包裹在变量模式令牌中
 */
export const resolveValueExpr = (ctx: CompileContext, templateValue: string): string | null => {
    const expr: RuleValue = ctx.currentExpr;
    if (expr === undefined || expr.tokens.length <= 0) {
        return templateValue;
    }
    const m = ctx.getContextMap();
    const token = TokenParserManger.getTokenByCtx(m, RuleToken.VarPatternToken);
    const anyToken = TokenParserManger.getTokenByCtx(m, RuleToken.AnyToken);
    const [start, end] = token.split(anyToken);
    if (!templateValue.startsWith(start)) {
        templateValue = `${start}${templateValue}`
    }
    if (!templateValue.endsWith(end)) {
        templateValue = `${templateValue}${end}`
    }
    return templateValue;
};

// ==================== 编译行单元格 ====================

/**
 * 编译行单元格
 * 对指定行范围内的单元格执行模板编译
 */
export const compileRowCells = function (ctx: CompileContext, expr: RuleValue, cellPoints: CellPoint[], rowIndex: number, errs: Error[]) {
    if (ctx.sheet === undefined) {
        errs.push(new Error(`ctx miss worksheet`));
        return;
    }
    ctx.currentExpr = expr;
    try {
        cellPoints.forEach((cellPoint, index) => {
            const r = cellPoint.Row;
            const sheet = ctx.sheet;
            const c = cellPoint.Column;
            const cell = sheet.findCell(r, c);
            if (cell === undefined ||
                (cell.value !== undefined && cell.value !== null && cell.value !== "")) {
                return;
            }
            let templateValue = String(expr.value);
            const macroTokens = getMacroTokens(expr);
            templateValue = resolveFunctionExpr(ctx, templateValue, expr);
            templateValue = resolveCompileMacroExpr(ctx, templateValue, macroTokens, index, cellPoints.length);
            templateValue = resolveAliasExpr(ctx, templateValue, index);
            cell.value = resolveValueExpr(ctx, templateValue);
        })
    } catch (err) {
        const msg = (err as Error).message;
        throw new Error(`expr:${expr.express}, resolve error: ${msg}`);
    }
};

// ==================== 生成工作表单元格占位符 ====================

/**
 * 生成工作表中所有占位符单元格的编译结果
 * 遍历表达式中的单元格坐标，对每个单元格执行编译
 */
export const generateWorkSheetCellsPlaceholder = function (ctx: CompileContext, expr: RuleValue, sheet: exceljs.Worksheet): Error[] | undefined {
    const errs: Error[] = [];
    const posExpr: RuleValue = expr.posExpr;
    let cellsItems = expr.cells;
    if ((cellsItems === undefined || cellsItems.length <= 0) &&
        posExpr !== undefined &&
        posExpr.value !== undefined) {
        const r = posExpr.value as RangeCell;
        if (r !== undefined) {
            cellsItems = r.getCells();
        }
    }
    if (!cellsItems || cellsItems.length === 0) {
        return undefined;
    }
    ctx.sheet = sheet;
    const cells = cellsItems;
    toRowCells(cells).forEach((cellPoints, index) => compileRowCells(ctx, expr, cellPoints, index, errs));
    return errs.length > 0 ? errs : undefined;
};

/**
 * 检查令牌列表是否包含生成器令牌
 * 判断表达式是否需要生成单元格
 */
export const hasGeneratorToken = function (tokens: RuleToken[]): boolean {
    for (const t of tokens) {
        if (compileCellTokens.includes(t)) {
            return true;
        }
    }
    return false;
};

// ==================== 编译工作表占位符 ====================

/**
 * 编译工作表占位符
 * 对工作表中所有规则表达式执行编译生成
 */
export const compileWorkSheetPlaceholder = function (ctx: CompileContext, sheet: exceljs.Worksheet, result: RuleResult): Error[] | undefined {
    if (ctx.compileSheets !== undefined && ctx.compileSheets.length > 0 && !ctx.compileSheets.includes(sheet.name)) {
        return undefined;
    }
    const errs: Error[] = [];
    for (const [token, express] of result.rules.entries()) {
        if (!isRuleToken(token)) {
            continue
        }
        for (const expr of express) {
            if (expr.tokens.length <= 0 || !hasGeneratorToken(expr.tokens)) {
                continue;
            }
            let err = generateWorkSheetCellsPlaceholder(ctx, expr, sheet);
            if (err !== undefined && err.length > 0) {
                errs.push(...err);
            }
        }
    }
    if (errs.length > 0) {
        return errs;
    }
    return undefined;
};

// ==================== 编译结果类型 ====================

/**
 * 编译结果类型
 * @property workbook - 编译后的工作簿
 * @property configure - 规则配置结果
 * @property errs - 编译错误列表
 */
export type CompileResult = {
    workbook: exceljs.Workbook
    configure?: RuleResult
    errs?: Error[]
};

// ==================== 工作簿编译 ====================

/**
 * 加载编译工作表列表
 * 从工作簿中获取需要编译的工作表名称列表，排除规则配置表和 .json/.config 结尾的工作表
 */
export const loadCompileSheets = (workbook: exceljs.Workbook, ruleSheetName: string | number): string[] => {
    let first: string;
    let sheets: string[] = [];
    for (const [_, w] of workbook.worksheets.entries()) {
        if (w.name === ruleSheetName) {
            continue;
        }
        if (first === "") {
            first = w.name;
        }
        if (!w.name.endsWith(".json") &&
            !w.name.endsWith(".config")) {
            sheets.push(w.name);
        }
    }
    if (sheets.length <= 0 && first !== "") {
        sheets.push(first);
    }
    return sheets;
};

/**
 * 执行模板编译
 * 解析规则配置，对目标工作表执行占位符替换
 * @param data - Excel 数据源
 * @param ruleSheetName - 规则工作表名称或索引
 * @param options - 编译选项
 */
export const compile = async function <T extends ArrayBuffer | Buffer | string>(
    data: T,
    ruleSheetName: string | number,
    options?: RuleOptions): Promise<CompileResult> {
    const workbook = await loadWorkbook(data);
    const sheet = workbook.getWorksheet(ruleSheetName);
    if (sheet === undefined) {
        return {
            workbook,
        };
    }
    if (workbook.worksheets === undefined) {
        return {
            workbook,
            errs: [new Error(`worksheet, ${ruleSheetName} not exists!`)],
        };
    }
    if (options === undefined) {
        const excludes = [ruleSheetName as string];
        options = RuleMapOptions.withAllSheets(workbook, excludes);
    }
    options = options.parseDefault(sheet);
    if (options.compileSheets === undefined || options.compileSheets.length === 0) {
        options.compileSheets = loadCompileSheets(workbook, ruleSheetName);
    }
    const result = parseWorkSheetRules(sheet, options);
    const errs = compileCheck(result, options);
    if (errs !== undefined) {
        return {
            errs,
            workbook,
            configure: result,
        };
    }
    const compileErrs: Error[] = [];
    const ctx = CompileContext.create(options).loadAlias(result.rules);
    for (const [i, w] of workbook.worksheets.entries()) {
        if (w.name === ruleSheetName ||
            i === ruleSheetName ||
            !ctx.filterSheet(w.name)) {
            continue;
        }
        let err = compileWorkSheetPlaceholder(ctx, w, result);
        if (err !== undefined && err.length > 0) {
            compileErrs.push(...err);
        }
    }
    if (compileErrs.length > 0) {
        return {
            workbook,
            configure: result,
            errs: compileErrs,
        };
    }
    return {
        workbook,
        configure: result,
    };
};

/**
 * 编译工作表（返回 XLSX 对象）
 * @param data - Excel 数据源
 * @param sheetName - 工作表名称或索引
 * @param options - 编译选项
 * @returns 编译后的 XLSX 对象或错误列表
 */
export const compileWorkSheet = async function <T extends ArrayBuffer | Buffer | string>(
    data: T,
    sheetName: string | number,
    options?: RuleOptions): Promise<exceljs.Xlsx | Error[]> {
    const reply = await compile(data, sheetName, options);
    if (reply.errs !== undefined && reply.errs.length > 0) {
        return reply.errs;
    }
    return reply.workbook.xlsx;
};

// ==================== 别名提取 ====================

/**
 * 提取别名映射
 * 从规则结果或映射中提取所有别名的键值对
 */
export const fetchAlias = (m: Map<RuleToken, RuleValue[]> | RuleResult): Map<string, string> => {
    let sv: Map<RuleToken, RuleValue[]>;
    const alias = new Map<string, string>();
    if (m !== undefined && m !== null && !(m instanceof Map) && m?.rules !== undefined) {
        sv = m.rules;
    } else if (m instanceof Map) {
        if (m.size <= 0 || !m.has(RuleToken.AliasToken)) {
            return alias;
        }
        sv = m;
    } else {
        return alias;
    }
    const values = sv.get(RuleToken.AliasToken);
    for (const vs of values) {
        if (typeof vs.value === "string") {
            alias.set(vs.key, vs.value as string);
        }
    }
    return alias;
};

/**
 * 移除未导出的工作表
 * 根据编译选项移除不需要导出的工作表
 */
export const removeUnExportSheets = (w: exceljs.Workbook, options: RuleOptions): exceljs.Workbook => {
    let removes: string[] = [];
    if (typeof options.skipRemoveUnExportSheet === "boolean" && options.skipRemoveUnExportSheet === true) {
        return w;
    }
    if (options.compileSheets === undefined || options.compileSheets.length <= 0) {
        for (const [i, v] of w.worksheets.entries()) {
            const sheetName = v.name;
            if (sheetName.endsWith(".config") ||
                sheetName.endsWith(".json")) {
                removes.push(sheetName);
            }
        }
    } else {
        for (const [i, v] of w.worksheets.entries()) {
            if (!options.compileSheets.includes(v.name)) {
                removes.push(v.name);
            }
        }
    }
    if(removes.length === w.worksheets.length && w.worksheets[0].name === removes[0]){
        removes = removes.slice(1,removes.length);
    }
    for (const [_, name] of removes.entries()) {
        w.removeWorksheet(name)
    }
    return w;
};

/**
 * 将工作簿转换为 Buffer
 */
export const toBuffer = async (w: exceljs.Workbook): Promise<Buffer> => {
    const arrayBuffer = await w.xlsx.writeBuffer()
    return Buffer.from(arrayBuffer);
};

// ==================== 表达式解析器（静态门面类）====================

/**
 * 表达式解析器门面类
 * 提供静态方法包装所有编译相关函数
 */
export class ExprResolver {
    static toBuffer = toBuffer;
    static compile = compile;
    static toRowCells = toRowCells;
    static fetchAlias = fetchAlias;
    static getExprEnd = getExprEnd;
    static compileCheck = compileCheck;
    static extractMacro = extractMacro;
    static compileRowCells = compileRowCells;
    static searchIndexOf = searchIndexOf;
    static resolveAliasExpr = resolveAliasExpr;
    static resolveValueExpr = resolveValueExpr;
    static resolveFunctionExpr = resolveFunctionExpr;
    static removeUnExportSheets = removeUnExportSheets;
    static resolveCompileMacroGen = resolveCompileMacroGen;
    static resolveCompileMacroExpr = resolveCompileMacroExpr;
}