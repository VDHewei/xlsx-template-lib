import {
    Workbook,
    FullOptions,
    SheetInfo, BufferType, valueDotGet,
} from "./core";
import {extname} from "node:path";
import {clone} from "lodash"
import JsZip from "jszip";
import AdmZip from "adm-zip";
import {AutoOptions, compileAll, commandExtendQuery, compileRuleSheetName, CmdFunction, Argument, AddCommand} from "./extends"
import {RuleMapOptions} from "./helper";

type CustomChecker =(data: Buffer, options: FullOptions & { [key: string]: any}, values: Object, fileName?: string)=> Promise<Buffer>

type CustomCheckerOptions = {
    checker?: CustomChecker;
    options?: FullOptions;
    [key: string]: any;
}

class XlsxRender extends Workbook {
    constructor(option?: FullOptions) {
        super(option);
    }

    static async create(data: Buffer, option?: FullOptions): Promise<XlsxRender> {
        const w = await super.parse(data, option);
        w.setQueryFunctionHandler(commandExtendQuery);
        const app = new XlsxRender(option);
        Object.assign(app, {...w})
        return app;
    }

    public async render(values: Object, sheetName: string): Promise<void> {
        await this.substitute(sheetName, values);
    }

    public getSheets(): SheetInfo[] {
        if (this.sheets.length === 0 && this.workbook) {
            this.sheets = this.loadSheets(this.workbook);
        }
        return this.sheets;
    }

}

class ZipXlsxTemplateApp {
    zipBuffer?: Buffer;
    private zip: AdmZip;
    private xlsxEntries: Map<string, Buffer>;
    private records: Map<string, XlsxRender> = new Map<string, XlsxRender>();

    constructor(data?: Buffer) {
        this.zipBuffer = data;
        if (data !== undefined) {
            this.xlsxEntries = this.parse(data);
        }
    }

    public loadZipBuffer(data: Buffer): ZipXlsxTemplateApp {
        this.zipBuffer = data;
        this.zip = new AdmZip(data);
        this.xlsxEntries = this.parse(data);
        return this;
    }

    public parse(data: Buffer): Map<string, Buffer> {
        const zip = new AdmZip(data);
        const result = new Map<string, Buffer>();
        const entries = zip.getEntries();
        for (let fd of entries) {
            if (fd.isDirectory) {
                continue
            }
            let ext = extname(fd.entryName).substring(1).toLowerCase();
            if (ext !== "xlsx") {
                continue
            }
            result.set(fd.entryName, fd.getData());
        }
        this.zip = zip;
        return result;
    }

    public getEntries(): Map<string, Buffer> {
        if (this.xlsxEntries !== undefined && this.xlsxEntries.size > 0) {
            return this.xlsxEntries;
        } else {
            if (this.zipBuffer !== undefined) {
                return this.parse(this.zipBuffer);
            }
        }
        return new Map<string, Buffer>();
    }

    static async compileAll(files: Map<string, Buffer>, renderData?: Object, compileOpts?: AutoOptions): Promise<Map<string, Buffer>> {
        const records = new Map<string, Buffer>();
        if (compileOpts !== undefined && (compileOpts.sheetName === undefined ||
            compileOpts.sheetName === "")) {
            compileOpts.sheetName = compileRuleSheetName;
        }
        let values = clone(renderData);
        for (let [key, buf] of files.entries()) {
            buf = await compileAll(buf,clone(compileOpts),clone(values));
            records.set(key, buf);
        }
        return records;
    }

    public async substituteAll(renderData: Object, compileOpts?: AutoOptions, renderOpts?: FullOptions): Promise<ZipXlsxTemplateApp> {
        const files = await ZipXlsxTemplateApp.compileAll(this.xlsxEntries, renderData, compileOpts);
        for (const [k, buf] of files.entries()) {
            const xlsx = await XlsxRender.create(buf, renderOpts);
            await xlsx.substituteAll(renderData);
            this.records.set(k, xlsx);
        }
        return this;
    }

    /**
     * 渲染 zip 中所有 xlsx 文件的指定 sheet
     * @param sheetName - 要渲染的 sheet 名称
     * @param renderData - 渲染数据
     * @param compileOpts - 可选编译选项（如 { sheetName: compileRuleSheetName, remove: true }）
     * @param renderOpts - 可选渲染选项
     */
    public async renderSheet(
        sheetName: string,
        renderData: Object,
        compileOpts?: AutoOptions,
        renderOpts?: FullOptions
    ): Promise<ZipXlsxTemplateApp> {
        // 编译（如需要）
        let entries = this.xlsxEntries;
        if (compileOpts) {
            entries = await ZipXlsxTemplateApp.compileAll(entries, renderData, compileOpts);
        }
        // 渲染指定 sheet
        for (const [key, buf] of entries.entries()) {
            const xlsx = await XlsxRender.create(buf, renderOpts);
            await xlsx.render(renderData, sheetName);
            this.records.set(key, xlsx);
        }
        return this;
    }

    public async generate(options?: JsZip.JSZipGeneratorOptions<BufferType.NodeBuffer> & FullOptions): Promise<Buffer> {
        if (this.records === undefined || this.records.size <= 0) {
            return this.zipBuffer;
        }
        if (this.zip === undefined) {
            this.zip = new AdmZip();
        }
        if (options === undefined || options === null) {
            options = {
                type: BufferType.NodeBuffer,
                compression: "DEFLATE",
                compressionOptions: {
                    level: 9
                }
            }
        }
        for (const [key, xlsx] of this.records) {
            const buf = await xlsx.generate(options);
            let entry = this.zip.getEntry(key);
            if (entry !== null) {
                entry.setData(Buffer.from(buf));
            } else {
                this.zip.addFile(key, Buffer.from(buf));
            }
        }
        return this.zip.toBuffer();
    }

    static async compileTo(data: Buffer,opts:CustomCheckerOptions, values?: Record<string, any> | Object): Promise<Buffer> {
        const zip = new AdmZip(data);
        const entries = zip.getEntries();
        let files = new Map<string, Buffer>();
        if(values === undefined){
            values = new Map<string, any>();
        }
        for (let fd of entries) {
            if (fd.isDirectory) {
                continue
            }
            let ext = extname(fd.entryName).substring(1).toLowerCase();
            if (ext !== "xlsx") {
                continue
            }
            let buf = fd.getData();
            if(opts.checker!==undefined){
               await opts.checker(buf,opts.options,values,opts.fileName||undefined)
            }
            files.set(fd.entryName, buf);
        }
        const compileOpts = new RuleMapOptions();
        compileOpts.remove = true;
        if (files.size > 0) {
            files = await ZipXlsxTemplateApp.compileAll(files, values, compileOpts);
        } else {
            throw new Error(`empty xlsx file zip file`);
        }
        if (files.size > 0) {
            for (const [k, data] of files.entries()) {
                zip.getEntry(k).setData(data);
            }
        }
        return zip.toBuffer();
    }
}

// formStatusImage(histories,statusesData,"statusCode","xxx")
const formStatusImage: CmdFunction = (values: Object | Record<string, any>, argument: Argument): any | undefined => {
    const histories = valueDotGet(values, argument.root);
    const statusesData = valueDotGet(values, argument.groups[0] || '');
    const matchField = argument.groups[1]?.replace(/^"|"$/g, '');
    const matchValue = argument.groups[2]?.replace(/^"|"$/g, '');

    if (!Array.isArray(histories) || !matchValue) return undefined;

    // 从 formStatuses 中查找 identifier
    let identifier: string | undefined;
    if (statusesData && typeof statusesData === 'object') {
        if (statusesData[matchValue] != null) {
            identifier = typeof statusesData[matchValue] === 'object'
                ? statusesData[matchValue].identifier
                : statusesData[matchValue];
        } else if (Array.isArray(statusesData)) {
            const found = statusesData.find((s: any) => s[matchField] === matchValue);
            if (found) identifier = found.identifier;
        }
    }
    if (!identifier) return undefined;

    // 过滤 formStatusHistories 中匹配 identifier 的条目，返回最后一条的签名图片
    const matched = histories.filter((h: any) => h.formStatusIdentifier === identifier);
    if (matched.length === 0) return undefined;
    return matched[matched.length - 1].actionSignatureBase64;
};

AddCommand("formStatusImage", formStatusImage);

export {
    ZipXlsxTemplateApp,
    XlsxRender,
    CustomChecker,
    CustomCheckerOptions,
    formStatusImage,
}