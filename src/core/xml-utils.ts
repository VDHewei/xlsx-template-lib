import {isNumber, isString} from "lodash";
import exceljs from "exceljs";
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

    return v;
}

const updateRichTextCell = function (v: exceljs.CellValue, newValue: any): exceljs.CellValue {

    return v;
}

const updateBooleanCell = function (v: exceljs.CellValue, newValue: any): exceljs.CellValue {

    return v;
}

const updateHyperlinkCell = function (v: exceljs.CellValue, newValue: any): exceljs.CellValue {

    return v;
}

const updateImageCell =async function (v: exceljs.Cell, newValue: any,w:exceljs.Worksheet):Promise<boolean> {

    return false;
}

const isImageValue = function (value: any|ImageValue): boolean {
    //
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
