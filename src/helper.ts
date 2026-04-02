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
    CompileGenToken = 'compile:Gen',
    CompileMacroToken = 'compile:Macro',

}

type CompileChecker = (iv: RuleResult, ctx: RuleMapOptions) => Error[] | undefined

class RuleMapOptions {
    // rule configure area
    startLine?: number = 1;
    endLine?: number = 20;
    endColumn?: number = 30;
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

    parseDefault(worksheet: exceljs.Worksheet): RuleMapOptions {
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

    addRuleMap(key: RuleToken, value: string): RuleMapOptions {
        this.ruleKeyMap.set(key, value);
        return this;
    }

    setStartRow(start: number): RuleMapOptions {
        this.startLine = start;
        return this;
    }

    setStartColumn(start: number): RuleMapOptions {
        this.startColumn = start;
        return this;
    }

    setEndRow(end: number): RuleMapOptions {
        this.endLine = end;
        return this;
    }

    setEndColumn(end: number): RuleMapOptions {
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
    private aliasMap: Map<string, string>;

    constructor(m?: Map<RuleToken, string>) {
        super(m);
    }

    static create(r: RuleMapOptions): CompileContext {
        const ctx = new CompileContext(r.ruleKeyMap);
        Object.assign(ctx, {...r})
        return ctx.init();
    }

    private init(): this {
        this.aliasMap = new Map<string, string>();
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

}

type CellPoint = {
    X: number,
    Y: number,
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

type RuleValue = {
    express: string;
    tokens: RuleToken[]; //  express tokens
    cells?: CellPoint[];
    key?: string; // alias key
    ref?: string[]; // alias refs
    func?: string; // express ref function name
    compileExpress?: string[];// compileExpress
    value: string | number[] | number | CellPoint | RangeCell | any[]; // alias value
    // extends
    [key: string]: any;
}

type RuleResult = {
    rules: Map<RuleToken, RuleValue[]>;
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
        return {
            ok: false,
            expr: {
                express: value,
                key: expr.key, // alias key
                value: expr.value, // alias value
                tokens: [token], //  express tokens
            },
            values: values,
        }
    }

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
        const key = value.substring(0, index - offset);
        const alias = value.substring(index + offset, len);
        const expr: RuleValue = {
            key: key,
            value: alias,
            express: value,
            tokens: [token],
        }
        return {
            ok: true,
            values: [key, alias],
            expr,
        }
    }

    static cellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        if (token !== RuleToken.CellToken) {
            return {
                ok: false
            }
        }
        // X:Y=${?}
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
        const posValue = value.substring(0, eqIndex + eqOffset);
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
            express: value,
            cells: posReply.expr.cells,
            value: varReply.expr.value, // var expr value
            ref: varReply.expr.ref, // alias refs
            tokens: [token, ...posReply.expr.tokens, ...varReply.expr.tokens],// token
            ...varReply.expr,
        };
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
        const endTokens =  [RuleToken.UseAliasToken,RuleToken.LparenToken,RuleToken.ArgPosToken,RuleToken.DotGetToken];
        const values = TokenParserManger.scanToken(value, useAliasToken,TokenParserManger.toList(ctx,endTokens));
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
        let startPos = value.substring(0, index - offset).trim();
        let endPos = value.substring(index + offset).trim();
        const argPosToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.ArgPosToken);
        const argPosOffset = argPosToken.length;
        const endSetupIndex = endPos.indexOf(argPosToken);
        if (endSetupIndex > 0) {
            setup = Number.parseInt(endPos.substring(endSetupIndex + argPosOffset).trim(), 10);
            endPos = endPos.substring(0, endSetupIndex - argPosOffset).trim()
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
        const x = value.substring(0, index - offset).trim();
        const y = value.substring(index + offset, len).trim();
        const rangeToken = TokenParserManger.getTokenByCtx(ctx, RuleToken.RangeToken);
        const xRange = x.indexOf(rangeToken);
        const yRange = y.indexOf(rangeToken);
        if (xRange > 0 || yRange > 0) {
            return TokenParserManger.parseRangeValue(ctx, {xRange, yRange, x, y, express: value, token})
        }
        const expr: RuleValue = {
            value: {
                X: columnLetterToNumber(x),
                Y: Number.parseInt(y, 10),
            },
            express: value,
            tokens: [token],
        }
        return {
            ok: true,
            values: [x, y],
            expr,
        }
    }

    static mergeCellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        //
        return {
            ok: false,
        }
    }

    static rowCellParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // rowCell
        return {
            ok: false,
        }
    }

    static functionPatternParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // <?>
        return {
            ok: false,
        }
    }

    static varPatternParse(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken, value: string): TokenParserReply {
        // ${xxx}
        return {
            ok: false,
        }
    }

    static compileExprExtract(value: string): string[] {

        return;
    }

    private static getTokenByCtx(ctx: Map<RuleToken, RuleValue[]>, token: RuleToken): string {
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
        let items: ScanTokenData[] = [];
        const offset = startToken.length;
        for (let i = 0; i < value.length; i += offset) {
            if (value[i] === startToken) {
                token = startToken;
            }else {
                token = `${token}${value[i]}`;
            }
            if(endTokens.includes(value[i])){
                end = true;
            }
            if(!end) {
                data  = `${data}${value[i]}`;
                if( i+offset >= value.length){
                    end = !end;
                }
            }
            if(end){
                items.push({
                    token,
                    value:data,
                });
                token = "";
                data = "";
            }
        }
        return items;
    }

    private static toList(ctx: Map<RuleToken, RuleValue[]>,endTokens: RuleToken[]): string[] {
        const items : string[] = [];
        for(const token of endTokens){
            items.push(TokenParserManger.getTokenByCtx(ctx, token));
        }
        return items;
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

function isBase64(str: string): boolean {
    if (str.length < 20) return false;
    if (str.includes("\\") || str.includes(" ")) return false;
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

const scanWorkSheetRules = function (worksheet: exceljs.Worksheet, options: RuleMapOptions): RuleResult {
    const result = {rules: options.getContextMap()};
    for (let i = options.startLine; i < options.endLine; i++) {
        let emptyValue = false;
        let ruleToken = RuleToken.UndefinedToken;
        for (let j = options.startColumn; j < options.endColumn; j++) {
            const cellValue = worksheet.getCell(i, j).value;
            const value = cellValue.toString();
            let isStartCell = i === options.startLine && j === options.startColumn;
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
            if (ok) {
                values.push(expr);
            }
            result.rules.set(ruleToken, values);
        }
    }
    return result;
}

const parseWorkSheetRules = function (worksheet: exceljs.Worksheet, options?: RuleMapOptions): RuleResult {
    const result: RuleResult = {rules: new Map<RuleToken, RuleValue[]>()};
    if (worksheet === undefined || worksheet === null) {
        return result;
    }
    if (options === undefined) {
        options = new RuleMapOptions();
    }
    return scanWorkSheetRules(worksheet, options.parseDefault(worksheet));
}

const compileCheck = function (iv: RuleResult, ctx: RuleMapOptions): Error[] | undefined {
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
    if (handlers !== undefined) {
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
    // TODO
    return undefined;
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
    TokenParserManger,
    RuleResult,
    CompileContext,
    isRuleToken,
    getTokenParser,
    registerTokenParser,
    registerTokenParserMust,
    compileWorkSheet,
    compileWorkSheetPlaceholder,
    loadWorkbook,
};