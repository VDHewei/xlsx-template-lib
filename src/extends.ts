import JsZip from "jszip";
import {Placeholder,Workbook,FullOptions,OutputByType, defaultValueDotGet, valueDotGet, QueryFunction} from "./core";

type Argument = {
    root: string;
    alias?: string;
    groups?: string[];
    suffix?: string;
    default: number | string | any;
    func: string;
    p: Placeholder;
}

type CmdFunction = (
    obj: Object | Record<string, any>,
    argument: Argument,
) => any | undefined;

const tokenNextIter = [`root`, `groups`, `suffix`, `default`];

class ArgumentData {
    root: string;
    p: Placeholder;
    default: number | string | any = '';
    alias?: string;
    suffix?: string;
    groups?: string[] = [];
    tokenIterIndex: number = 0;
    private readonly func: string;

    constructor(private readonly fn: string, p: Placeholder) {
        this.p = p;
        this.func = fn;
        this.groups = [];
    }

    To(): Argument {
        return {
            p: this.p,
            root: this.root,
            alias: this.alias,
            groups: this.groups,
            suffix: this.suffix,
            default: this.default,
            func: this.func,
        };
    }

    Add(startToken: string, value: string|undefined) {
        if (value === undefined){
            this.tokenIterIndex++;
            return;
        }
        switch (startToken) {
            case `(`:
                if (tokenNextIter[this.tokenIterIndex] === "root") {
                    this.root = value;
                    this.tokenIterIndex++;
                }
                break;
            case `[`:
                if (tokenNextIter[this.tokenIterIndex] === "groups") {
                    this.groups.push(value)
                }
                break;
            case `]`:
                if (tokenNextIter[this.tokenIterIndex] === "groups") {
                    this.groups.push(value);
                }
                break;
            case `,`:
                const token = tokenNextIter[this.tokenIterIndex];
                if (token === "root") {
                    this.root = value;
                    this.tokenIterIndex++;
                }else if (token === "groups") {
                    this.groups.push(value);
                }else if (token=== "suffix") {
                    this.suffix = value;
                    this.tokenIterIndex++;
                }else if (token === "default") {
                    this.default = value;
                    this.tokenIterIndex++;
                }
                break;
            case `)`:
                this.tokenIterIndex++;
                break;
        }
    }

    ParseAlias(alias: any | object | Record<string, string>) {
        if (alias === undefined || this.root === undefined || this.root === "") {
            return;
        }
        const value = valueDotGet(alias, this.root, "");
        if (value === undefined || typeof value !== "string") {
            return;
        }
        this.alias = this.root;
        this.root = value as string;
    }
}

const ArgumentValueLoader = (values: Object | Record<string, any>, args: Argument): any[] => {
    let all: string[] = [];
    for (let v of args.groups) {
        let key = `${args.root}.${v}`;
        if (args.suffix !== undefined && args.suffix !== "") {
            key = `${key}.${args.suffix}`
        }
        all.push(key);
    }
    if (all.length <= 0) {
        return args.default || '';
    }
    const items: any[] = [];
    for (let k of all) {
        let vs = valueDotGet(values, k, args.default);
        if (vs === undefined) {
            continue
        }
        items.push(vs)
    }
    return items;
}

// ${fn:sum_all(#,[C308,C342,C321,C3016,C309_C409],1,0)}
const sum_all: CmdFunction = (values: Object | Record<string, any>, argument: Argument,): any | undefined => {
    let sum: number = NaN;
    let items = ArgumentValueLoader(values, argument);
    for (let num of items) {
        if (num === undefined) {
            num = argument.default;
        }
        if (isNaN(sum)) {
            sum = Number(num)
        } else {
            sum = sum + Number(num);
        }
    }
    if (isNaN(sum)) {
        throw new Error(`parse ${argument.p.name} NaN error`);
    }
    return sum;
}

// ${fn:sub_all(#,[C308,C342,C321,C3016,C309_C409],1,0)}
const sub_value: CmdFunction = (values: Object | Record<string, any>, argument: Argument,): any | undefined => {
    let sub: number = NaN;
    let items = ArgumentValueLoader(values, argument);
    for (let num of items) {
        if (num === undefined) {
            num = argument.default;
        }
        if (isNaN(Number(num))) {
            continue;
        }
        if (isNaN(sub)) {
            sub = num;
        } else {
            sub = sub - Number(num);
        }
    }
    if (isNaN(sub)) {
        throw new Error(`parse ${argument.p.name} NaN error`);
    }
    return sub;
}

const defaultCommands = new Map<string, CmdFunction>([
    ["sum_all", sum_all],
    ["sub", sub_value],
]);

const resolveFunc = function (value: string): string {
    if (value.indexOf("(") > 0 && value.endsWith(")")) {
        const names = value.split("(")
        return names[0]
    }
    return ""
}

const resolveArgument = function (p: Placeholder, data: object | Record<string, any>): Argument {
    const value = p.name;
    const fn = resolveFunc(value);
    const args = new ArgumentData(fn, p);
    if (fn !== "") {
        let key: string = "";
        let startT: string = "";
        const endToken = [`)`,`,`, `]`];
        const startToken = [`(`, `,`, `[`];
        const tokenRow = value.split(`${fn}`)[1]
        const len = tokenRow.length;
        for (let i = 0; i < len ; i++) {
            let start = startToken.includes(tokenRow[i]);
            let end = endToken.includes(tokenRow[i]);
            if (start) {
                startT = tokenRow[i];
            }
            if (startT !== "" && tokenRow[i] !== startT && !end) {
                key = `${key}${tokenRow[i]}`;
            }
            if (end) {
                if(key===""){
                    args.Add(startT,undefined);
                }else {
                    args.Add(startT, key);
                    key = "";
                }
            }
        }
    }
    const alias = valueDotGet(data, `__alias`);
    if (alias !== undefined) {
        args.ParseAlias(alias)
    }
    return args.To();
}

const commandExtendQuery: QueryFunction = function (values: object | Record<string, any>, p: Placeholder): any | undefined {
    if (p.type !== "fn") {
        return defaultValueDotGet(values, p)
    }
    const argument = resolveArgument(p, values);
    if (argument.func !== "" && defaultCommands.has(argument.func)) {
        return defaultCommands.get(argument.func)(values, argument)
    }
    return defaultValueDotGet(values, p)
}

/**
 * 安全添加扩展函数支持
 * @param key
 * @param h
 * @constructor
 */
const AddCommand = (key: string, h: CmdFunction): boolean => {
    if (defaultCommands.has(key)) {
        return false;
    }
    defaultCommands.set(key, h);
    return true;
}

/**
 * 强制添加扩展函数支持
 * @param key
 * @param h
 * @constructor
 */
const AddCommandMust = (key: string, h: CmdFunction): void => {
    defaultCommands.set(key, h);
}

// xlsx 模板 生成 - 函数一键调用
const generateCommandsXlsxTemplate = async function <T extends JsZip.OutputType>(data: Buffer, values: Object, options?: JsZip.JSZipGeneratorOptions<T> & FullOptions): Promise<OutputByType[T]> {
    const w = await Workbook.parse(data,options);
    w.setQueryFunctionHandler(commandExtendQuery);
    await w.substituteAll(values);
    return w.generate(options);
}

export {
    commandExtendQuery,
    defaultCommands,
    resolveArgument,
    CmdFunction,
    AddCommand,
    AddCommandMust,
    ArgumentValueLoader,
    generateCommandsXlsxTemplate,
}