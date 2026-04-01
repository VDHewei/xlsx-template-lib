import exceljs from "exceljs";
import {Stream} from "stream";

type CellPosition = {
    X: string;
    Y: number;
    Sheet?: string|number;
};

type MergeCellRange =  {
    top: number;
    left: number;
    bottom: number;
    right: number
}

interface PlaceholderCellValue {
    toString(): string;
    mergeCell(values: string[]): string;
}

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
    try{
        atob(cleanedStr)
        return /^[A-Za-z0-9+/]*={0,2}$/.test(cleanedStr);
    }catch (e) {
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

function getMergeRange(ws: exceljs.Worksheet, row: number, col: number): MergeCellRange| null {
    const merges = ws.model.merges as string[];
    for (const mergeStr of merges) {
        const parts = mergeStr.split(":");
        const tl = ws.getCell(parts[0]);
        const br = ws.getCell(parts[1]);
        const tlRow = Number(tl.row)
        const tlCol  = Number(tl.col)
        const brRow =   Number(br.row)
        const brCol  = Number(br.col)
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

async function scanCellSetPlaceholder<T extends ArrayBuffer|Buffer | string>(
    excelBuffer: T,
    cell: CellPosition & {Sheet: string|number},
    placeholder: PlaceholderCellValue
): Promise<ArrayBuffer|undefined> {
    const w = new exceljs.Workbook();
    if (typeof excelBuffer === "string") {
        if (!isBase64(excelBuffer)) {
            await w.xlsx.readFile(excelBuffer);
        } else {
            await w.xlsx.load(base64ToArrayBuffer(excelBuffer));
        }
    } else if(excelBuffer instanceof  Stream) {
        await w.xlsx.read(excelBuffer);
    }else if(excelBuffer instanceof  ArrayBuffer) {
        await w.xlsx.load(excelBuffer);
    }
    const worksheet = w.getWorksheet(cell.Sheet);
    if (!worksheet) return undefined;
    workSheetSetPlaceholder(worksheet,cell,placeholder)
    return w.xlsx.writeBuffer()
}

function workSheetSetPlaceholder(worksheet: exceljs.Worksheet, cell: CellPosition, placeholder: PlaceholderCellValue) :exceljs.Worksheet {
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

export { scanCellSetPlaceholder,workSheetSetPlaceholder, CellPosition,DefaultPlaceholderCellValue,PlaceholderCellValue,exceljs};