import {isNumber, isObject, isString} from "lodash";
import * as fs from "node:fs/promises";
import exceljs from "exceljs";
import {imageSize as sizeOf} from 'image-size';
import {ImageValue} from "./types";

/**
 * URL 匹配正则表达式
 * 用于判断字符串是否为有效的 URL 地址
 */
const pattern = new RegExp('^(https?:\\/\\/)?' +
    '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' +
    '((\\d{1,3}\\.){3}\\d{1,3}))' +
    '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' +
    '(\\?[;&a-z\\d%_.~+=-]*)?' +
    '(\\#[-a-z\\d_]*)?$', 'i');

/**
 * 判断字符串是否为 URL
 * @param str - 待检测的字符串
 * @returns 是否为有效的 URL
 */
const isUrl = function (str: string): boolean {
    return !!pattern.test(str);
}

/**
 * 将 Buffer 转换为 ArrayBuffer
 * @param buffer - 源 Buffer 数据
 * @returns 转换后的 ArrayBuffer
 */
const toArrayBuffer = function (buffer: Buffer): ArrayBuffer {
    const ab = new ArrayBuffer(buffer.length);
    const view = new Uint8Array(ab);
    for (let i = 0; i < buffer.length; ++i) {
        view[i] = buffer[i];
    }
    return ab;
}

const toDate = function (v: any): Date | undefined {
    if (v === undefined || v === null) {
        return undefined;
    }
    if (v instanceof Date) {
        return v as Date;
    }
    if (isNumber(v) || isNaN(v)) {
        let timestamp = Number(v);
        if (timestamp < 10000000000) {
            timestamp *= 1000;
        }
        return new Date(timestamp);
    }
    if (isString(v)) {
        const timestamp = Date.parse(v as string);
        if (isNaN(timestamp)) {
            return undefined;
        } else {
            return new Date(timestamp);
        }
    }
    return undefined;
}

const updateFormulaCell = function (v: exceljs.CellValue, newValue: any): exceljs.CellValue {
    if (v && typeof v === 'object' && 'formula' in (v as any)) {
        const formulaObj = v as any;
        return { ...formulaObj, formula: String(newValue ?? '') };
    }
    if (typeof newValue === 'string' && newValue.startsWith('=')) {
        return { formula: newValue.substring(1) } as any;
    }
    return v;
}

const updateRichTextCell = function (v: exceljs.CellValue, newValue: any): exceljs.CellValue {
    if (newValue === undefined || newValue === null) return v;
    const text = String(newValue);
    if (v && typeof v === 'object' && 'richText' in (v as any)) {
        const rt = (v as any).richText;
        if (Array.isArray(rt) && rt.length > 0) {
            rt.forEach((part: any) => { if (part.text !== undefined) part.text = text; });
            if (rt.length > 0 && rt[0].text === undefined) rt[0].text = text;
            return v;
        }
    }
    return [{ font: (v as any)?.richText?.[0]?.font || {}, text }] as any;
}

const updateBooleanCell = function (v: exceljs.CellValue, newValue: any): exceljs.CellValue {
    if (typeof newValue === 'boolean') return newValue;
    if (typeof newValue === 'string') {
        const lower = newValue.trim().toLowerCase();
        if (lower === 'true' || lower === '1' || lower === 'yes') return true;
        if (lower === 'false' || lower === '0' || lower === 'no') return false;
    }
    if (typeof newValue === 'number') return newValue !== 0;
    return v;
}

const updateHyperlinkCell = function (v: exceljs.CellValue, newValue: any): exceljs.CellValue {
    if (newValue === undefined || newValue === null) return v;
    if (v && typeof v === 'object' && 'hyperlink' in (v as any)) {
        const h = v as any;
        if (typeof newValue === 'string') {
            h.hyperlink = newValue;
            h.text = h.text || newValue;
        }
        return v;
    }
    if (typeof newValue === 'string') {
        return { hyperlink: newValue, text: newValue } as any;
    }
    return v;
}

const updateImageCell = async function (v: exceljs.Cell, newValue: any, sheet: exceljs.Worksheet, w: exceljs.Workbook): Promise<boolean> {
    const val = newValue as ImageValue;
    if (!val || !val.imageType) return false;

    let imageBuffer: Buffer;
    if (val.imageType === 'file' && val.path) {
        try {
            imageBuffer = (await fs.readFile(val.path)) as any as Buffer;
        } catch {
            return false;
        }
    } else if (val.imageType === 'base64' && val.buffer) {
        imageBuffer = val.buffer;
    } else {
        return false;
    }

    let ext: 'png' | 'jpeg' | 'gif' = 'png';
    try {
        const dim = sizeOf(imageBuffer);
        if (dim.type === 'png' || dim.type === 'jpeg' || dim.type === 'gif') {
            ext = dim.type;
        }
    } catch { /* use default */ }

    const imgId = w.addImage({
        buffer: imageBuffer as any,
        extension: ext,
    });

    // 从 Cell 对象获取 0-based 的行列索引
    const cellRow = (v as any).row as number ?? 1;
    const cellCol = (v as any).col as number ?? 1;

    const wd = val.width || 200;
    const ht = val.height || 200;
    const zoom = val.zoom || 1;

    sheet.addImage(imgId, {
        tl: { col: cellCol - 1, row: cellRow - 1 },
        ext: { width: Math.round(wd * zoom), height: Math.round(ht * zoom) },
    });
    v.value = undefined; // 清除单元格原有值
    return true;
}

const isImageValue = function (value: any|ImageValue): boolean {
    // 判断是否 ImageVale 类型 是否 返回 true
    if (isObject(value)) {
        return (value as ImageValue).imageType !== undefined;
    }
    return false;
}

export {
    isUrl,
    toDate,
    isImageValue,
    toArrayBuffer,
    updateImageCell,
    updateHyperlinkCell,
    updateBooleanCell,
    updateRichTextCell,
    updateFormulaCell,
};