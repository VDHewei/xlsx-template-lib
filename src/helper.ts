import exceljs from "exceljs";
import {Stream} from "stream";

type CellPosition = {
    Row: string;
    Column: number;
    Sheet?: string | number;
};

type MergeCellRange = {
    top: number;
    left: number;
    bottom: number;
    right: number
}

type CellPoint = {
    Row: number,
    Column: number,
}

interface PlaceholderCellValue {
    toString(): string;

    mergeCell(values: string[]): string;
}

type ScanTokenData = {
    value: string;
    token: string;
}

type MacroArgs = {
    type: string
    columnParam: number
    rowParam: number[] | number
    formatter?: string
}

type ExtractMacroArgs = {
    argToken: string
    startToken: string
    endToken: string
}

const isPureNumber = /^[0-9]+$/;
const isPureUppercase = /^[A-Z]+$/;
const exprSingle = `expr`;
const exprArr = `exprArr`;
const exprIndex = `index`;
const defaultKey = `!!`;
const numberKey = `!!number`;
const codeKey = `!!codeKey`;
const codeAliasKey = `!!codeAliasKey`;
const funcCommand = "fn:";

enum RuleToken {
    AliasToken = 'alias',
    CellToken = 'cell',
    MergeCellToken = 'mergeCell',
    RowCellToken = 'rowCell',
    UseAliasToken = '@',
    RangeToken = '-',
    PosToken = ':',
    FunctionPatternToken = '<?>',
    AnyToken = '?',
    VarPatternToken = '${?}',
    UndefinedToken = '',
    EqualToken = '=',
    ArgPosToken = ',',
    LparenToken = '(',
    RparenToken = ')',
    DotGetToken = '.',
    CompileGenToken = 'compile:GenCell',
    CompileMacroToken = 'compile:Macro',

}

type RuleValue = {
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
}

type RuleResult = {
    rules: Map<RuleToken, RuleValue[]>;
}

type FilterMacroResult = {
    tokens: RuleToken[];
    express: string[];
}

type CompileChecker = (iv: RuleResult, ctx: RuleMapOptions) => Error[] | undefined

interface RuleOptions {
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

    setEndRow(end: number): RuleOptions;

    parseToken(value: string): RuleToken;

    setEndColumn(end: number): RuleOptions;

    setStartRow(start: number): RuleOptions;

    setStartColumn(start: number): RuleOptions;

    getContextMap(): Map<RuleToken, RuleValue[]>;

    addRuleMap(key: RuleToken, value: string): RuleOptions;

    parseDefault(worksheet: exceljs.Worksheet): RuleOptions;

    getCompileCheckHandlers(): CompileChecker[] | undefined;
}

interface RangeCell {
    minRow: number;
    maxRow: number;
    stepRow: number;
    minColumn: number;
    maxColumn: number;
    stepColumn: number;

    getCells(): CellPoint[];
}

class RuleMapOptions implements RuleOptions {
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

    addRuleMap(key: RuleToken, value: string): RuleOptions {
        this.ruleKeyMap.set(key, value);
        return this;
    }

    setStartRow(start: number): RuleOptions {
        this.startLine = start;
        return this;
    }

    setStartColumn(start: number): RuleOptions {
        this.startColumn = start;
        return this;
    }

    setEndRow(end: number): RuleOptions {
        this.endLine = end;
        return this;
    }

    setEndColumn(end: number): RuleOptions {
        this.endColumn = end;
        return this;
    }

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

    getCompileCheckHandlers(): CompileChecker[] | undefined {
        if (this.compileCheckers !== undefined && this.compileCheckers.length > 0) {
            return this.compileCheckers;
        }
        return undefined;
    }

}

class CompileContext extends RuleMapOptions {
    private aliasMap: Map<string, string> = new Map<string, string>();

    sheet?: exceljs.Worksheet;

    constructor(m?: Map<RuleToken, string>) {
        super(m);
    }

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
     */
    public setAlias(key: string, value: string): void {
        this.aliasMap.set(key, value);
    }

    /**
     * 获取别名缓存值
     */
    public getAlias(key: string): string | undefined {
        return this.aliasMap.get(key);
    }

    /**
     * 检查别名是否存在
     */
    public hasAlias(key: string): boolean {
        return this.aliasMap.has(key);
    }

    public filterSheet(sheetName: string): boolean {
        if (sheetName !== "" && this.compileSheets !== undefined && this.compileSheets.length > 0) {
            return this.compileSheets.includes(sheetName)
        }
        return false;
    }

}

type TokenParserReply = {
    ok: boolean
    values?: any,
    expr?: RuleValue,
    [key: string]: any;
}

type TokenParser = (ctx: Map<RuleToken, RuleValue[]>, t: RuleToken, value: string) => TokenParserReply;

type TokenParseResolver = { exists: boolean, handler?: TokenParser }

type RangeValueParserOptions = {
    rowRange: number;
    columnRange: number;
    row: string;
    column: string;
    express: string;
    token: RuleToken;
};

const defaultRuleTokenMap = new Map<RuleToken, string>([
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

class DefaultPlaceholderCellValue implements PlaceholderCellValue {
    private readonly placeholder: string;
    private readonly mergeCellPlaceholder?: string;

    constructor(p: string, merge?: string) {
        this.placeholder = p;
        this.mergeCellPlaceholder = merge === undefined ? "" : merge;
    }

    mergeCell(values: string[]): string {
        if (this.mergeCellPlaceholder !== undefined) {
            return this.mergeCellPlaceholder.replace("?", values.join(","));
        }
        return "";
    }

    toString(): string {
        return this.placeholder;
    }
}

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
}

class TokenParserManger {

    // T=xxx
    static aliasParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.AliasToken) {
            return {
                ok: false
            }
        }
        // value eg: T=xxx.xxx.Tell
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

    // x=xx
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

    //cell| X:Y=${?}
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

    static useAliasParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.UseAliasToken) {
            return {
                ok: false,
            }
        }
        // @X.test or @X.@T or @U@T
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

    static rangeParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // A-AA or A-AZ or 1-7 or 1-7,2 [step] or A-Z,2 [step]
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

    static posParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.PosToken) {
            return {
                ok: false
            }
        }
        // A:1 or A-Z:1 or A:1-10 or A-AA:2-10
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

    // value input: A-AQ:13-15=<sum(#,[compile:Macro(exprArr,[F],[13,15],!codeKey)],compile:Marco(index),0)>
    // output: {
    //      A:13 => sum(#,[ codeKey(F:13.value),codeKey(F:14.value),codeKey(F:15.value)],1,0),
    //      B:13 => sum(#,[ codeKey(F:13.value),codeKey(F:14.value),codeKey(F:15.value)],2,0),
    //      C:13 => sum(#,[ codeKey(F:13.value),codeKey(F:14.value),codeKey(F:15.value)],3,0),
    //     ....,
    //     A:14 => sum(#,[ codeKey(F:13.value),codeKey(F:14.value),codeKey(F:15.value)],1,0),
    //     B:14 => sum(#,[ codeKey(F:13.value),codeKey(F:14.value),codeKey(F:15.value)],2,0),
    //     C:14 => sum(#,[ codeKey(F:13.value),codeKey(F:14.value),codeKey(F:15.value)],3,0),
    //  }
    static mergeCellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // Logic: Range=Expression. Eg: A-AQ:13-15=<sum(...)>
        if (token !== RuleToken.MergeCellToken && token !== RuleToken.RowCellToken) {
            return {ok: false};
        }
        const equalToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.EqualToken);
        const index = value.indexOf(equalToken);
        const offset = equalToken.length;
        if (index <= 0) return {ok: false, error: new Error(`merge cell config syntax error: ${value}`)};
        const rangeStr = value.substring(0, index).trim();
        const exprStr = value.substring(index + offset)
        // Parse Range (PosToken logic usually handles X:Y, here we wrap it)
        // We reuse PosToken parsing logic which handles ranges via RangeToken
        const posParser = getTokenParser(RuleToken.PosToken);
        const posReply = posParser.handler(ctx, RuleToken.PosToken, rangeStr);
        const functionToken = getTokenParser(RuleToken.FunctionPatternToken);
        if (!posReply.ok) return {ok: false, ...posReply};
        const macroGenToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.CompileGenToken);
        // Parse Expression (could be function, var, or raw string)
        // We treat the right side as a raw string or simple token for the generator to handle
        const expr: RuleValue = {
            express: value,
            value: exprStr, // The template expression
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

    // G-AQ:12=compile:GenCell(compile:Macro(expr,F,12),'.',compile:Marco(index))
    static rowCellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // Logic: G-AQ:12=compile:GenCell(...)
        if (token !== RuleToken.RowCellToken) {
            return {ok: false};
        }
        // Reuse logic similar to mergeCellParse for range/expression split
        return TokenParserManger.mergeCellParse(ctx, token, value);
    }

    // $functionName($arg0:string,$arg1:string[],$arg2:string|number,$arg3:string|number)
    // eg: func(A,[xx1,xx2],xxx3), sum(#,[F,12,13],1,0)
    // $functionName not in compile:Macro,compile:GenCell
    // extract => {func:$functionName, arguments:[$arg0,$arg1,$arg2,$argN...]}
    // output: {func:"sum",arguments:["A",["F","12","13"],"1","0"]}
    static functionPatternParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // <?> or <functionName(args)>
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
        // Check structure <?>
        if (value.startsWith(funcStartToken) && value.endsWith(funcEndToken)) {
            const content = value.substring(1, value.length - 1);
            // Simple parse: functionName(arg1,arg2)
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

    static varPatternParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // ${?} or ${xxx}
        if (token !== RuleToken.VarPatternToken) {
            return {ok: false};
        }
        // Check if value matches the pattern like ${...}
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

    // compile:GenCell(expr1,expr2,...,) => {G:13=>`${expr1}${expr2}..`,G:14=>`${expr1}${expr2}..`,G:15=>`${expr1}${expr2}..`,...}
    // compile:Macro($exprType,$X,$Y,$formater)
    // $exprType: enums [ expr:(single value), exprArr:(array value) ]
    // $X: column index value, number|string or number[]|string[]
    // $Y: row index value, number or number[]
    // $formater: enums [ codeKey(a function for format string),number(a function for string to number) ]
    // compile:Macro(exprArr,[X1,X2],[Y1,Y2],!codeKey) =>
    // codeKey(X1:Y1.value),codeKey(X1:Y2.value),codeKey(X2:Y1.value),codeKey(X2:Y2.value)
    // compile:Macro(exprArr,[X1],[Y1,Y2],!codeKey) => codeKey(X1:Y1.value),codeKey(X1:Y2.value)
    // compile:Marco(index,!number) => number(i),number(i+1),number(i+2) ,i=1
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

    static split(argsStr: string, argsSplitToken: string, ignoreTokenRange: string[]): string[] {
        let value: string = "";
        // 使用 depth 处理嵌套情况（如括号嵌套）或简单的开关状态（如引号）
        // depth > 0 表示当前处于“不分割”的区域内
        let depth = 0;
        const items: string[] = [];
        // 预计算 Token 长度，避免循环中重复计算
        const splitLen = argsSplitToken.length;
        const startLen = ignoreTokenRange[0].length;
        const endLen = ignoreTokenRange[1].length;
        const startToken = ignoreTokenRange[0];
        const endToken = ignoreTokenRange[1];
        // 判断是否为同一种符号（如单引号、双引号）
        const isToggleMode = startToken === endToken;
        for (let i = 0; i < argsStr.length;) {
            // 截取当前位置开始的子串用于匹配
            const substr = argsStr.substring(i);
            // 1. 检测是否遇到“结束忽略”标记 (必须在检测“开始标记”之前，否则同符号模式会瞬间开关)
            if (depth > 0 && substr.startsWith(endToken)) {
                value += endToken;
                i += endLen;
                // 如果是同符号模式（如引号），直接关闭；否则减少嵌套深度
                depth = isToggleMode ? 0 : depth - 1;
                continue;
            }
            // 2. 检测是否遇到“开始忽略”标记
            if (substr.startsWith(startToken)) {
                // 如果是同符号模式，只有当前处于关闭状态(depth=0)才打开
                // 如果是不同符号模式（如括号），则增加深度
                if (!isToggleMode || depth === 0) {
                    value += startToken;
                    i += startLen;
                    depth++;
                    continue;
                }
            }
            // 3. 检测分隔符 (仅当不在忽略区域内时生效)
            if (depth === 0 && substr.startsWith(argsSplitToken)) {
                items.push(value);
                value = "";
                i += splitLen; // 跳过分隔符
                continue;
            }
            // 4. 默认情况：追加当前字符，索引+1
            value += argsStr[i];
            i++;
        }
        // 将最后一段内容推入结果
        // 如果字符串以分隔符结尾，此时 value 为空，应 push 空字符串以保持 split 行为一致性
        if (value !== "") {
            items.push(value);
        }
        return items;
    }

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

    static extractUseAliasTokens(ctx: Map<RuleToken, RuleValue[]>, args: string[]): RuleToken[] {
        const tokens: RuleToken[] = [];
        const useAliasToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.UseAliasToken);
        args.forEach(v => v.startsWith(useAliasToken) && tokens.push(RuleToken.UseAliasToken))
        return tokens;
    }

    static toList(ctx: Map<RuleToken, RuleValue[]>, endTokens: RuleToken[]): string[] {
        const items: string[] = [];
        for (const token of endTokens) {
            items.push(TokenParserManger.getTokenByCtx(ctx, token));
        }
        return items;
    }

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
            tokens: [token], //  express tokens
            value: rangeCell, // alias value
        }
        return {
            ok: true,
            expr,
            values: rangeCell,
        };
    }

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

const defaultRuleTokenParserMap = new Map<RuleToken, TokenParser>([
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

const macroTokens: RuleToken[] = [RuleToken.CompileGenToken, RuleToken.CompileMacroToken,];

function columnLetterToNumber(letter: string): number {
    let num = 0;
    for (let i = 0; i < letter.length; i++) {
        num = num * 26 + (letter.charCodeAt(i) - 64);
    }
    return num;
}

function columnNumberToLetter(num: number): string {
    // 处理非正整数情况
    if (num <= 0) {
        return "";
    }
    let letter = "";
    while (num > 0) {
        // 核心逻辑：Excel 列标是从1开始的26进制数
        // 1 -> A, 26 -> Z, 27 -> AA
        // 计算当前位的字符：先减1以映射到 0-25 的区间
        let remainder = (num - 1) % 26;
        // A 的 ASCII 码是 65
        letter = String.fromCharCode(65 + remainder) + letter;
        // 移动到下一位：由于是整数除法，需减去当前位再除以26
        // 或者更简单的写法：num = Math.floor((num - 1) / 26);
        num = Math.floor((num - 1) / 26);
    }
    return letter;
}

function isBase64(str: string): boolean {
    if (str.length < 20) return false;
    if (str.includes("\\") || str.includes(" ")) return false;
    if (str.indexOf("./") >= 0 ||
        str.startsWith("file://") ||
        str.startsWith("/") ||
        str.startsWith(".")) {
        return false;
    }
    // 剔除可能存在的换行符后再进行纯字符集校验
    const data = str.replace(/^data:.*?;base64,/, "");
    const cleanedStr = data.replace(/[\r\n]/g, "");
    try {
        return Buffer.from(cleanedStr, 'base64').toString('utf-8') !== "";
    } catch (e) {
        return false;
    }
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
    const data = base64.replace(/^data:.*?;base64,/, "");
    const binaryString = atob(data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}

function getMergeRange(ws: exceljs.Worksheet, row: number, col: number): MergeCellRange | null {
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

const resolveCompileMacroGen = (ctx: CompileContext, expr: string, currentCellIndex: number): string => {
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
}

const getExprEnd = function (macroExpr: string, matchIndex: number, rparenToken: string): number {
    return macroExpr.indexOf(rparenToken, matchIndex);
}

const macroFormatters: string[] = [numberKey, codeKey, codeAliasKey];
type MacroUnitHelper = (v: string, expr?: RuleValue) => string;

// compile:Macro(type, X_arr, Y_arr, formatter)
const extractMacro = function (expr: string, options: ExtractMacroArgs): MacroArgs {
    let end = NaN;
    const offset = options.startToken.length;
    const startIndex = expr.indexOf(options.startToken);
    const endIndex = expr.indexOf(options.endToken);
    const argValues = expr.substring(startIndex + offset, endIndex);
    const values = argValues.split(options.argToken);
    const extracResult: MacroArgs = {type: values[0], rowParam: [], columnParam: undefined};
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
}

const __codeKey: MacroUnitHelper = (str: string, expr?: RuleValue): string => {
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
}

const __numberKey: MacroUnitHelper = (str: string, expr?: RuleValue): string => {
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
}

const __codeAliasKey: MacroUnitHelper = (str: string, expr?: RuleValue): string => {
    const key = __codeKey(str);
    if (key !== "") {
        if (expr !== undefined && expr.tokens.length > 0) {
            expr.tokens.push(RuleToken.UseAliasToken);
        }
        return `${defaultRuleTokenMap.get(RuleToken.UseAliasToken)}${key}`;
    }
    return '';
}

const macroFormatter: Map<string, MacroUnitHelper> = new Map<string, MacroUnitHelper>([
    [codeKey, __codeKey],
    [numberKey, __numberKey],
    [codeAliasKey, __codeAliasKey],
    [defaultKey, (v: string): string => v],
]);

const execMacroFormat = function (value: string, formatter: string, expr?: RuleValue): string {
    if (!macroFormatter.has(formatter)) {
        return value;
    }
    return macroFormatter.get(formatter)(value, expr)
}

const toCellRow = (rowVals: number[], setup?: number): number[] => {
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
}

const toCellColumn = (columnVals: number[], setup?: number): number[] => {
    return toCellRow(columnVals, setup);
}

const toCellValue = (value: exceljs.CellValue): string => {
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
}

/**
 * 核心宏解析函数
 * 支持：
 * 1. compile:Macro(expr/exprArr, [X], [Y], !formatter)
 * 2. compile:Gen(expr, expr1, ...)
 */
const resolveCompileMacroExpr = (ctx: CompileContext, macroExpr: string, macroTokens: RuleToken[], currentCellIndex: number, totalCells: number): string => {
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
    // 1. 处理 compile:Macro -> gen.expr | gen.exprArr
    // 语法: compile:Macro(type, X_arr, Y_arr, formatter)
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
            // compile:Macro(type, X_arr, Y_arr, formatter)
            let {type, columnParam, rowParam, formatter} = extractMacro(macroCurrent, opts);
            let rowVals: number[];
            let columnVals: number[];
            const parts: string[] = [];
            // 辅助函数：解析参数数组
            if (columnParam !== undefined && rowParam !== undefined) {
                rowVals = resolveArray(rowParam);
                columnVals = resolveArray(columnParam);
                const rowItems = toCellRow(rowVals);
                const columnItems = toCellColumn(columnVals);
                rowItems.forEach(r => {
                    columnItems.forEach(c => {
                        // 构建坐标引用，如 F:13
                        //const cellRef = `${columnNumberToLetter(r)}:${c}`;
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
            // 处理 exprArr (笛卡尔积展开)
            if (type === exprArr) {
                macroExpr = macroExpr.replace(macroCurrent, parts.join(','))
            } else if (type === exprSingle) {
                // 处理 expr (单一值或索引)
                // 如果是 expr 类型，通常用于取特定值
                macroExpr = macroExpr.replace(macroCurrent, parts[0])
            } else if (type === exprIndex) {
                // 处理 exprIndex == index
                const indexValue = currentCellIndex + 1;
                macroExpr = macroExpr.replace(macroCurrent, indexValue.toString())
            }
        }
    }

    // 2. 处理 compile:Gen -> expr,expr1,expr2...
    // 逻辑: 递归解析参数并拼接
    // 语法: compile:Gen(expr1, expr2, ...)
    if (tokenMap.has(RuleToken.CompileGenToken)) {
        let exprValue: string = macroExpr;
        const times = tokenMap.get(RuleToken.CompileGenToken).length;
        for (let i = times; i > 0; i--) {
            const matchIndex: number = exprValue.indexOf(genToken);
            if (matchIndex < 0) {
                break;
            }
            // compile:Gen(expr1, expr2, ...)
            const offset = getExprEnd(exprValue, matchIndex, rparenToken);
            const macroCurrent = exprValue.substring(matchIndex, offset + rparenToken.length);
            // args.join('.')
            exprValue = resolveCompileMacroGen(ctx, macroCurrent, currentCellIndex);
        }
        macroExpr = exprValue;
    } else {
        macroExpr = resolveAliasExpr(ctx, macroExpr, currentCellIndex);
    }

    return macroExpr;
};

const loadWorkbook = async function <T extends ArrayBuffer | Buffer | string>(data: T): Promise<exceljs.Workbook> {
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
}

const scanCellSetPlaceholder = async function <T extends ArrayBuffer | Buffer | string>(
    excelBuffer: T,
    cell: CellPosition & { Sheet: string | number },
    placeholder: PlaceholderCellValue
): Promise<ArrayBuffer | undefined> {
    const workbook = await loadWorkbook(excelBuffer);
    const worksheet = workbook.getWorksheet(cell.Sheet);
    if (!worksheet) return undefined;
    workSheetSetPlaceholder(worksheet, cell, placeholder)
    return workbook.xlsx.writeBuffer()
}

const workSheetSetPlaceholder = function (worksheet: exceljs.Worksheet, cell: CellPosition, placeholder: PlaceholderCellValue): exceljs.Worksheet {
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
}

const compileCellTokens: RuleToken[] = [RuleToken.CellToken, RuleToken.MergeCellToken, RuleToken.RowCellToken];
const ruleTokens: RuleToken[] = [RuleToken.AliasToken, RuleToken.CellToken, RuleToken.MergeCellToken, RuleToken.RowCellToken];

const isRuleToken = function (t: RuleToken): boolean {
    return ruleTokens.includes(t)
}

const mergeOption = function (ruleKeyMap: Map<RuleToken, string>, defaultRuleTokenMap: Map<RuleToken, string>): Map<RuleToken, string> {
    for (const [key, value] of defaultRuleTokenMap.entries()) {
        if (!ruleKeyMap.has(key)) {
            ruleKeyMap.set(key, value)
        }
    }
    return ruleKeyMap;
}

const getTokenParser = function (token: RuleToken): TokenParseResolver {
    if (!defaultRuleTokenParserMap.has(token)) {
        return {
            exists: false,
        }
    }
    return {
        exists: true,
        handler: defaultRuleTokenParserMap.get(token),
    }
}

const registerTokenParser = function (token: RuleToken, h: TokenParser): boolean {
    if (defaultRuleTokenParserMap.has(token)) {
        return false;
    }
    defaultRuleTokenParserMap.set(token, h)
    return true;
}

const registerTokenParserMust = function (token: RuleToken, h: TokenParser): void {
    defaultRuleTokenParserMap.set(token, h)
}

const scanWorkSheetRules = function (worksheet: exceljs.Worksheet, options: RuleOptions): RuleResult {
    const result = {rules: options.getContextMap()};
    // row
    for (let r = options.startLine; r <= options.endLine; r++) {
        let emptyValue = false;
        let ruleToken = RuleToken.UndefinedToken;
        // column
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
}

const parseWorkSheetRules = function (worksheet: exceljs.Worksheet, options?: RuleOptions): RuleResult {
    const result: RuleResult = {rules: new Map<RuleToken, RuleValue[]>()};
    if (worksheet === undefined || worksheet === null) {
        return result;
    }
    if (options === undefined) {
        options = new RuleMapOptions();
    }
    return scanWorkSheetRules(worksheet, options.parseDefault(worksheet));
}

const compileCheck = function (iv: RuleResult, ctx: RuleOptions): Error[] | undefined {
    if (iv.rules.size <= 0) {
        return undefined;
    }
    const errs: Error[] = [];
    // alias configuration check
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
    // custom compile checker
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
}

const getMacroTokens = function (expr: RuleValue): RuleToken[] {
    const tokens: RuleToken[] = [];
    for (const token of expr.tokens) {
        if (macroTokens.includes(token)) {
            tokens.push(token);
        }
    }
    return tokens;
}

const toRowCells = function (cells: CellPoint[]): CellPoint[][] {
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
}

const resolveFunctionExpr = (ctx: CompileContext, templateValue: string, expr: RuleValue): string => {
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
}

const searchIndexOf = (str: string, substr: string[], position?: number): number => {
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
}

const resolveAliasExpr = (ctx: CompileContext, templateValue: string, index: number): string => {
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
}

const resolveValueExpr = (ctx: CompileContext, templateValue: string): string | null => {
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
}

const compileRowCells = function (ctx: CompileContext, expr: RuleValue, cellPoints: CellPoint[], rowIndex: number, errs: Error[]) {
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
            // 我们的解析逻辑中 X 已经转换为数字索引
            const cell = sheet.findCell(r, c);
            if (cell === undefined ||
                (cell.value !== undefined && cell.value !== null && cell.value !== "")) {
                return;
            }
            // 获取表达式模板
            let templateValue = String(expr.value);
            // 步骤 1: 解析宏指令
            // 提取所有 compile: 开头的指令
            const macroTokens = getMacroTokens(expr);
            // 递归替换宏
            templateValue = resolveFunctionExpr(ctx, templateValue, expr); // <?>
            templateValue = resolveCompileMacroExpr(ctx, templateValue, macroTokens, index, cellPoints.length); // compile:Macro,compile:Gen
            templateValue = resolveAliasExpr(ctx, templateValue, index); // @
            // 写入单元格
            cell.value = resolveValueExpr(ctx, templateValue);
        })
    } catch (err) {
        const msg = (err as Error).message;
        throw new Error(`expr:${expr.express}, resolve error: ${msg}`);
    }
}

const generateWorkSheetCellsPlaceholder = function (ctx: CompileContext, expr: RuleValue, sheet: exceljs.Worksheet): Error[] | undefined {
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
    // 遍历目标单元格
    toRowCells(cells).forEach((cellPoints, index) => compileRowCells(ctx, expr, cellPoints, index, errs));
    return errs.length > 0 ? errs : undefined;
};

const hasGeneratorToken = function (tokens: RuleToken[]): boolean {
    for (const t of tokens) {
        if (compileCellTokens.includes(t)) {
            return true;
        }
    }
    return false;
}

const compileWorkSheetPlaceholder = function (ctx: CompileContext, sheet: exceljs.Worksheet, result: RuleResult): Error[] | undefined {
    // check need compile sheet setting
    if (ctx.compileSheets !== undefined && ctx.compileSheets.length > 0 && !ctx.compileSheets.includes(sheet.name)) {
        return undefined;
    }
    const errs: Error[] = [];
    for (const [token, express] of result.rules.entries()) {
        if (!isRuleToken(token)) {
            continue
        }
        // expr generat worksheet cells
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
}

type CompileResult = {
    workbook: exceljs.Workbook
    configure?: RuleResult
    errs?: Error[]
}

const loadCompileSheets = (workbook: exceljs.Workbook, ruleSheetName: string | number): string[] => {
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
}

const compile = async function <T extends ArrayBuffer | Buffer | string>(
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
    // parse rules
    const result = parseWorkSheetRules(sheet, options);
    const errs = compileCheck(result, options);
    if (errs !== undefined) {
        return {
            errs,
            workbook,
            configure: result,
        };
    }
    // compile
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
}

const compileWorkSheet = async function <T extends ArrayBuffer | Buffer | string>(
    data: T,
    sheetName: string | number,
    options?: RuleOptions): Promise<exceljs.Xlsx | Error[]> {
    const reply = await compile(data, sheetName, options);
    if (reply.errs !== undefined && reply.errs.length > 0) {
        return reply.errs;
    }
    return reply.workbook.xlsx;
}

const fetchAlias = (m: Map<RuleToken, RuleValue[]> | RuleResult): Map<string, string> => {
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
}

const removeUnExportSheets = (w: exceljs.Workbook, options: RuleOptions): exceljs.Workbook => {
    const removes: string[] = [];
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
    for (const [_, name] of removes.entries()) {
        w.removeWorksheet(name)
    }
    return w;
}

const toBuffer = async (w: exceljs.Workbook): Promise<Buffer> => {
    const arrayBuffer = await w.xlsx.writeBuffer()
    return Buffer.from(arrayBuffer);
}

class ExprResolver {
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

export {
    CellPosition,
    CompileResult,
    DefaultPlaceholderCellValue,
    PlaceholderCellValue,
    exceljs,
    RuleToken,
    RuleMapOptions,
    TokenParserManger,
    RuleResult,
    RuleOptions,
    CompileContext,
    FilterMacroResult,
    MacroUnitHelper,
    MacroArgs,
    ExtractMacroArgs,
    ExprResolver,
    toCellValue,
    scanCellSetPlaceholder,
    workSheetSetPlaceholder,
    parseWorkSheetRules,
    columnLetterToNumber,
    columnNumberToLetter,
    isRuleToken,
    hasGeneratorToken,
    getTokenParser,
    registerTokenParser,
    registerTokenParserMust,
    compileWorkSheet,
    compileWorkSheetPlaceholder,
    loadWorkbook,
    loadCompileSheets,
};