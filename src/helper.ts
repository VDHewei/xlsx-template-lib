import exceljs from "exceljs";
import {Stream} from "stream";

type CellPosition = {
    X: string;
    Y: number;
    Sheet?: string | number;
};

type MergeCellRange = {
    top: number;
    left: number;
    bottom: number;
    right: number
}

type CellPoint = {
    X: number,
    Y: number,
}

interface PlaceholderCellValue {
    toString(): string;

    mergeCell(values: string[]): string;
}

type ScanTokenData = {
    value: string;
    token: string;
}

const isPureNumber = /^[0-9]+$/;
const isPureUppercase = /^[A-Z]+$/;

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
    minX: number;
    maxX: number;
    stepX: number;
    minY: number;
    maxY: number;
    stepY: number;

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

    parseDefault(worksheet: exceljs.Worksheet): RuleOptions {
        this.ruleKeyMap = mergeOption(this.ruleKeyMap, defaultRuleTokenMap);
        if (this.startLine === undefined) {
            this.startLine = 1;
        }
        if (this.endLine === undefined) {
            this.endLine = worksheet.rowCount;
        }
        if (this.startColumn === undefined) {
            this.startColumn = 1;
        }
        if (this.endColumn === undefined) {
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
    xRange: number;
    yRange: number;
    x: string;
    y: string;
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
    for (let j = thisArg.minY; j < thisArg.maxY; j += thisArg.stepY) {
        for (let i = thisArg.minX; i < thisArg.maxX; i += thisArg.stepX) {
            cells.push({
                X: i,
                Y: j,
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
        const x = value.substring(0, index).trim();
        const y = value.substring(index + offset, len).trim();
        const rangeToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.RangeToken);
        const xRange = x.indexOf(rangeToken);
        const yRange = y.indexOf(rangeToken);
        if (xRange > 0 || yRange > 0) {
            return TokenParserManger.parseRangeValue(ctx, {xRange, yRange, x, y, express: value, token})
        }
        const cell: CellPoint = {
            X: columnLetterToNumber(x),
            Y: Number.parseInt(y, 10),
        };
        const expr: RuleValue = {
            value: cell,
            cells: [cell],
            express: value,
            tokens: [token],
        }
        return {
            ok: true,
            values: [x, y],
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
        const macroGenToken = TokenParserManger.getTokenByCtx(ctx,RuleToken.CompileGenToken);
        // Parse Expression (could be function, var, or raw string)
        // We treat the right side as a raw string or simple token for the generator to handle
        const expr: RuleValue = {
            express: value,
            value: exprStr, // The template expression
            posExpr: posReply.expr,
            tokens: [token, ...posReply.expr.tokens],
        };
        if(exprStr.startsWith(macroGenToken)){
            const argsSplitToken = TokenParserManger.getTokenByCtx(ctx,RuleToken.ArgPosToken);
            const args = TokenParserManger.split(exprStr, argsSplitToken, [`(`, `)`]);
            const macro = TokenParserManger.filterMacro(args);
            if(macro!==undefined && macro.tokens.length >0){
                expr.macro = macro;
                expr.tokens.push(...macro.tokens);
            }
        }else {
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
        const reply = TokenParserManger.mergeCellParse(ctx, token, value);
        return reply;
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

    private static parseRangeValue(ctx: Map<RuleToken, RuleValue[]>, options: RangeValueParserOptions): TokenParserReply {
        let xReply: TokenParserReply;
        let yReply: TokenParserReply;
        const rangeCell: RangeCell = {
            stepX: 1,
            stepY: 1,
            minY: 0,
            minX: 0,
            maxX: 0,
            maxY: 0,
            getCells: function (): CellPoint[] {
                return _getCells(this);
            }
        };
        const {xRange, x, y, yRange, express, token} = options;
        const rangeParse = getTokenParser(RuleToken.RangeToken);
        if (xRange > 0) {
            xReply = rangeParse.handler(ctx, RuleToken.RangeToken, x);
        } else {
            rangeCell.maxX = rangeCell.minX = columnLetterToNumber(x);
        }
        if (yRange > 0) {
            yReply = rangeParse.handler(ctx, RuleToken.RangeToken, y)
        } else {
            rangeCell.maxY = rangeCell.minY = Number.parseInt(y, 10);
        }
        if (xReply !== undefined) {
            if (!xReply.ok || xReply.values === undefined) {
                return xReply
            }
            const values = xReply.values as number[];
            let min = values[0];
            let max = values[1];
            let setup = values[2] ?? 1;
            rangeCell.minX = min;
            rangeCell.maxX = max;
            rangeCell.stepX = setup;
        }
        if (yReply !== undefined) {
            if (!yReply.ok || yReply.values === undefined) {
                return yReply
            }
            const values = yReply.values as number[];
            let min = values[0];
            let max = values[1];
            let setup = values[2] ?? 1;
            rangeCell.minY = min;
            rangeCell.maxY = max;
            rangeCell.stepY = setup;
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

    private static scanToken(value: string, startToken: string, endTokens: string[]): ScanTokenData[] {
        let token: string = "";
        let data: string = "";
        let end: boolean = false;
        let start: boolean = false;
        const items: ScanTokenData[] = [];
        const offset = startToken.length;
        const size = value.length;
        for (let i = 0; i < size; i += offset) {
            if (value[i] === startToken) {
                start = true;
                token = startToken;
            } else {
                start = false;
            }
            if (endTokens.includes(value[i])) {
                end = true;
            } else {
                end = false;
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

    private static toList(ctx: Map<RuleToken, RuleValue[]>, endTokens: RuleToken[]): string[] {
        const items: string[] = [];
        for (const token of endTokens) {
            items.push(TokenParserManger.getTokenByCtx(ctx, token));
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
            tokens:[],
            express:[],
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

/**
 * 核心宏解析函数
 * 支持：
 * 1. compile:Macro(expr/exprArr, [X], [Y], !formatter)
 * 2. compile:Gen(expr, expr1, ...)
 */
const resolveCompileMacro = (ctx: CompileContext, macroStr: string, currentCellIndex: number, totalCells: number): string => {
    // 1. 处理 compile:Gen
    // 语法: compile:Gen(expr1, expr2, ...)
    // 逻辑: 递归解析参数并拼接
    const m = ctx.getContextMap();
    const aliasToken = TokenParserManger.getTokenByCtx(m, RuleToken.AliasToken);
    const genToken = TokenParserManger.getTokenByCtx(m, RuleToken.CompileGenToken);
    const maroToken = TokenParserManger.getTokenByCtx(m, RuleToken.CompileMacroToken);
    if (macroStr.startsWith(genToken)) {
        const match = macroStr.match(/compile:Gen\((.+)\)/);
        if (match) {
            const argsStr = match[1];
            // 简单的参数分割 (注意：如果参数内部包含逗号，需更复杂的解析器，此处假设简单结构)
            // 这里我们需要递归解析每个参数，因为它可能是另一个 macro
            // 由于正则难以处理嵌套，此处采用简单的平衡组或深度遍历（简化版实现）
            // 我们假设参数由 compile:Macro 或普通字符串组成

            // 提取参数列表 (简易版)
            const args = argsStr.split(/,(?![^\[]*\])/); // 忽略 [] 内的逗号进行分割 (粗略)

            const resolvedParts = args.map(arg => {
                arg = arg.trim();
                if (arg.startsWith('compile:')) {
                    return resolveCompileMacro(ctx, arg, currentCellIndex, totalCells);
                }
                // 解析参数中的别名引用 @T
                if (arg.startsWith(aliasToken)) {
                    return ctx.getAlias(arg.substring(1)) || arg;
                }
                return arg;
            });

            return resolvedParts.join(''); // Gen 通常用于拼接
        }
    }

    // 2. 处理 compile:Macro
    // 语法: compile:Macro(type, X_arr, Y_arr, formatter)
    if (macroStr.startsWith(maroToken)) {
        const match = macroStr.match(/compile:Macro\(([^,]+),([^,]+),([^,]+),([^)]+)\)/);
        if (!match) return macroStr;

        let [, type, xParam, yParam, formatter] = match;

        // 辅助函数：解析参数数组，支持别名 @T
        const resolveArray = (param: string): (string | number)[] => {
            // 去除 []
            const clean = param.replace(/[\[\]]/g, '').trim();
            // 按 , 分割
            return clean.split(',').map(item => {
                item = item.trim();
                if (item.startsWith(aliasToken)) {
                    // 使用 CompileContext 提供的 getAlias 方法
                    const val = ctx.getAlias(item.substring(1));
                    // 如果别名是数字字符串，转换为数字，否则返回字符串
                    return val && !isNaN(Number(val)) ? Number(val) : (val || item);
                }
                return !isNaN(Number(item)) ? Number(item) : item;
            });
        };

        const xVals = resolveArray(xParam);
        const yVals = resolveArray(yParam);

        // 处理 exprArr (笛卡尔积展开)
        if (type === 'exprArr') {
            const parts: string[] = [];
            xVals.forEach(x => {
                yVals.forEach(y => {
                    // 构建坐标引用，如 F:13
                    const cellRef = `${x}:${y}`;

                    // 根据 formatter 生成内容
                    if (formatter === '!codeKey') {
                        // 规则：codeKey(X:Y.value) -> 模拟生成代码引用
                        // 假设 codeKey 是一种特定的代码生成格式
                        parts.push(`codeKey(${cellRef}.value)`);
                    } else if (formatter === '!number') {
                        // 规则：转化为数字索引或值
                        // 这里假设直接输出坐标对应的数字索引或者格式化后的值
                        // 示例中未明确 !number 的具体输出格式，暂时保留数字转换逻辑
                        parts.push(String(currentCellIndex));
                    } else {
                        parts.push(cellRef);
                    }
                });
            });
            return parts.join(',');
        }

        // 处理 expr (单一值或索引)
        else if (type === 'expr' || type === 'index') {
            if (formatter === '!number') {
                return String(currentCellIndex);
            }
            // 如果是 expr 类型，通常用于取特定值
            return `${xVals[0]}:${yVals[0]}`;
        }
    }

    return macroStr;
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
    const colNum = columnLetterToNumber(cell.X);
    const rowNum = cell.Y;
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

const generateWorkSheetCellsPlaceholder = function (ctx: CompileContext, expr: RuleValue, sheet: exceljs.Worksheet): Error[] | undefined {
    const errs: Error[] = [];
    // 获取解析出的 cells 列表
    const cells = expr.cells;

    if (!cells || cells.length === 0) {
        return undefined;
    }
    ctx.sheet = sheet;
    // 遍历目标单元格
    cells.forEach((cellPoint, index) => {
        const {X, Y} = cellPoint;
        // exceljs 中 row 是 1-based，col 也是 1-based
        // 我们的解析逻辑中 X 已经转换为数字索引
        const cell = sheet.getCell(Y, X);

        // 获取表达式模板
        let templateValue = String(expr.value);

        // 步骤 1: 解析宏指令
        // 提取所有 compile: 开头的指令
        const macroRegex = /compile:(Gen|Macro)\([^)]+\)/g;
        // 递归替换宏
        templateValue = templateValue.replace(macroRegex, (match) => {
            return resolveCompileMacro(ctx, match, index + 1, cells.length);
        });

        // 步骤 2: 解析变量表达式 ${...}
        const varRegex = /\$\{([^}]+)\}/g;
        templateValue = templateValue.replace(varRegex, (match, p1) => {
            // p1 可能是 key 或 @T.key
            // 尝试从 aliasMap 获取
            if (p1.startsWith('@')) {
                return ctx.getAlias(p1.substring(1)) || match;
            }
            return ctx.getAlias(p1) || match;
        });

        // 步骤 3: 处理纯别名引用 @T (如果外部直接写 @T 而没有 ${})
        const aliasRegex = /@([a-zA-Z0-9_]+)/g;
        templateValue = templateValue.replace(aliasRegex, (match, p1) => {
            return ctx.getAlias(p1) || match;
        });

        // 步骤 4: 处理合并单元格特殊逻辑
        // 如果当前单元格是合并区域的一部分，且规则定义了 mergeCell 逻辑
        if (cell.isMerged) {
            const range = getMergeRange(sheet, Y, X);
            if (range) {
                // 如果是合并单元格，且规则要求合并替换 (如 mergeCellToken)
                // 这里简单处理：只在合并区域的左上角单元格写入值
                if (range.left === X && range.top === Y) {
                    cell.value = templateValue;
                } else {
                    // 其他部分清空或保持，避免覆盖
                    // cell.value = null;
                }
                return;
            }
        }
        // 写入单元格
        cell.value = templateValue;
    });

    return errs.length > 0 ? errs : undefined;
};

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

const compileWorkSheet = async function <T extends ArrayBuffer | Buffer | string>(
    data: T,
    ruleSheetName: string | number,
    options?: RuleMapOptions): Promise<exceljs.Xlsx | Error[]> {
    const workbook = await loadWorkbook(data);
    const sheet = workbook.getWorksheet(ruleSheetName);
    if (sheet === undefined) {
        throw new Error(`compile error, ${ruleSheetName} not exists!`);
    }
    if (options === undefined) {
        options = new RuleMapOptions();
    }
    // parse rules
    const result = parseWorkSheetRules(sheet, options);
    const errs = compileCheck(result, options);
    if (errs !== undefined) {
        return errs;
    }
    if (workbook.worksheets === undefined) {
        return workbook.xlsx;
    }
    // compile
    const compileErrs: Error[] = [];
    const ctx = CompileContext.create(options).loadAlias(result.rules);
    for (const [i, w] of workbook.worksheets.entries()) {
        if (w.name === ruleSheetName || i === ruleSheetName) {
            continue;
        }
        let err = compileWorkSheetPlaceholder(ctx, w, result);
        if (err !== undefined && err.length > 0) {
            compileErrs.push(...err);
        }
    }
    if (compileErrs.length > 0) {
        return compileErrs;
    }
    return workbook.xlsx;
}

export {
    scanCellSetPlaceholder,
    workSheetSetPlaceholder,
    CellPosition,
    DefaultPlaceholderCellValue,
    PlaceholderCellValue,
    exceljs,
    RuleToken,
    parseWorkSheetRules,
    columnLetterToNumber,
    columnNumberToLetter,
    TokenParserManger,
    RuleResult,
    RuleOptions,
    CompileContext,
    FilterMacroResult,
    isRuleToken,
    getTokenParser,
    registerTokenParser,
    registerTokenParserMust,
    compileWorkSheet,
    compileWorkSheetPlaceholder,
    loadWorkbook,

};