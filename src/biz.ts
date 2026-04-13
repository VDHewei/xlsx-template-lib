import {
    Workbook,
    FullOptions,
    SheetInfo, BufferType,
} from "./core";
import {extname} from "node:path";
import {clone} from "lodash"
import JsZip from "jszip";
import AdmZip from "adm-zip";
import {AutoOptions, compileAll, commandExtendQuery, compileRuleSheetName} from "./extends"
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

export {
    ZipXlsxTemplateApp,
    XlsxRender,
    CustomChecker,
}