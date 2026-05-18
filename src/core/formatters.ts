import {isMap} from "node:util/types";
import {isArray} from "lodash";
import {Placeholder, CustomFormatter} from "./types";
import { fromUnixTime,format, parseISO } from 'date-fns'

/**
 * 从对象中按单键名获取值（支持数组索引语法 `key[index]`）
 * @param obj - 源对象
 * @param key - 键名，支持 `key` 或 `key[0]` 语法
 * @param arrayList - 如果为 true 且 obj 为数组，则遍历数组收集每个元素的 key 值
 * @returns 获取到的值
 */
function _getSimple(obj: any | object | Record<string, any> | Record<string, string>, key: string, arrayList?: boolean): any {
    if (key.includes("[")) {
        // 修正正则：匹配 [ 和 ] 并进行拆分
        // 例如：'list[0]' -> ['list', '0', '']
        const parts = key.split(/[\[\]]/);
        const property = parts[0];
        const index = parts[1];
        if (property && index !== undefined) {
            return obj?.[property]?.[index];
        }
    }
    if (isMap(obj)) {
        return obj.get(key)
    }
    if (arrayList && isArray(obj)) {
        let list = [];
        for (const item of obj) {
            list.push(_getSimple(item, key) || "")
        }
        return list;
    }
    return obj?.[key];
}

type PathImpl<T, Key extends string> =
    T extends object
        ? Key extends `${infer K}.${infer Rest}`
            ? K extends keyof T
                ? PathImpl<T[K], Rest>
                : never
            : Key extends keyof T
                ? T[Key]
                : never
        : any;

type PathType<T, Key extends string> = string extends Key ? any : PathImpl<T, Key>;

/**
 * 基于路径从对象中获取值
 * 模拟 lodash 的 get 方法
 * @param obj - 源对象
 * @param path - 点分隔的路径字符串
 * @param defaultValue - 默认值
 * @param valueType - 值类型（可传入 "table" 以启用数组列表模式）
 * @returns 路径对应的值
 */
function valueDotGet<T extends Record<string, any> & object, P extends string>(
    obj: T,
    path: P,
    defaultValue?: PathType<T, P>,
    valueType: string = null,
): PathType<T, P> {
    if (!path || !obj) return defaultValue as PathType<T, P>;
    const keys = path.split('.');
    const size = keys.length;
    let current: any = obj;
    for (const [index, key] of keys.entries()) {
        if (current === null || current === undefined) return defaultValue as PathType<T, P>;
        if (index < size - 1) {
            current = _getSimple(current, key);
        } else {
            current = _getSimple(current, key, valueType != null && valueType === "table");
        }
    }
    return current === undefined ? defaultValue as PathType<T, P> : current;
}

/**
 * 基于路径从对象中获取值，默认方法
 * 模拟 lodash 的 get 方法，使用占位符中的默认值
 * @param obj - 源对象
 * @param p - 占位符信息
 * @returns 获取到的值
 */
function defaultValueDotGet<T extends Record<string, any> & object>(obj: T, p: Placeholder): PathType<T, string> {
    return valueDotGet(obj, p.name, p.default || '');
}

/**
 * 从 placeholder.placeholder 原始字符串中解析出完整的数据路径。
 * 默认正则只捕获第一个点分隔的段，对于 ${type:a.b.c.d} 只能得到 name="a", key="b"。
 * 此函数提取 type 前缀后，返回完整路径 "a.b.c.d"。
 * @param placeholder - 占位符信息
 * @returns 完整的数据路径字符串
 */
function resolveFullDataPath(placeholder: Placeholder): string {
    if (!placeholder.placeholder) return placeholder.name;
    let innerText = placeholder.placeholder.substring(2, placeholder.placeholder.length - 1);
    const typePrefix = placeholder.type + ':';
    if (innerText.startsWith(typePrefix)) {
        innerText = innerText.substring(typePrefix.length);
    }
    return innerText;
}

/**
 * 日期格式化器 - 将 Date 对象转换为 Excel 日期序号
 * @param value - 待格式化的值
 * @param _placeholder - 占位符信息（未使用）
 * @param _key - 可选键名（未使用）
 * @returns 格式化后的字符串，或 undefined
 */
const dateFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if(!_key || _key !== "date"){
        return undefined;
    }
    if (value instanceof Date) {
        return format(value as Date,'yyyy-MM-dd');
    }
    if (typeof value === 'number') {
		return format(fromUnixTime(value),'yyyy-MM-dd');
	} else if (typeof value === 'string') {
        let timestamp= parseInt(value as string, 10);
        if(!isNaN(timestamp)) {
            return format(fromUnixTime(timestamp),'yyyy-MM-dd');
        }
        // 验证是否为 ISO 格式日期字符串
        const result = parseISO(value as string);
        if (!isNaN(result.getTime())) {
            return format(result,'yyyy-MM-dd');
        }
		const date = new Date()
        const [hour, minute, second = '00'] = value.split(':')
		date.setHours(parseInt(hour, 10))
		date.setMinutes(parseInt(minute, 10))
		date.setSeconds(parseInt(second, 10))
		return format(date,'yyyy-MM-dd');
	}
    return undefined;
};


/**
 * 日月-格式化器 - 将 Date 对象转换为 Excel 月日 - 序号
 * @param value - 待格式化的值
 * @param _placeholder - 占位符信息（未使用）
 * @param _key - 可选键名（未使用）
 * @returns 格式化后的字符串，或 undefined
 */
const dayFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if(!_key || _key !== "day"){
        return undefined;
    }
    if (value instanceof Date) {
        return format(value as Date,'MM-dd');
    }
    if (typeof value === 'number') {
		return format(fromUnixTime(value),'MM-dd');
	} else if (typeof value === 'string') {
        let timestamp= parseInt(value as string, 10);
        if(!isNaN(timestamp)) {
            return format(fromUnixTime(timestamp),'MM-dd');
        }
        // 验证是否为 ISO 格式日期字符串
        const result = parseISO(value as string);
        if (!isNaN(result.getTime())) {
            return format(result,'MM-dd');
        }
		const date = new Date()
        const [hour, minute, second = '00'] = value.split(':')
		date.setHours(parseInt(hour, 10))
		date.setMinutes(parseInt(minute, 10))
		date.setSeconds(parseInt(second, 10))
		return format(date,'MM-dd');
	}
    return undefined;
};


/**
 * 数字格式化器 - 将数字转换为字符串
 * @param value - 待格式化的值
 * @param _placeholder - 占位符信息（未使用）
 * @param _key - 可选键名（未使用）
 * @returns 格式化后的字符串，或 undefined
 */
const numberFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if(!_key || _key !== "number"){
        return undefined;
    }
    if (typeof value === "number") {
        return value.toString();
    }
    if (typeof value === "string") {
        const num = parseFloat(value);
        if (!isNaN(num)) {
            return value as string;
        }
    }
    return undefined;
};

/**
 * 布尔格式化器 - 将布尔值转换为数字字符串（0 或 1）
 * @param value - 待格式化的值
 * @param _placeholder - 占位符信息（未使用）
 * @param _key - 可选键名（未使用）
 * @returns 格式化后的字符串，或 undefined
 */
const booleanFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if(!_key || (_key !== "boolean" && _key !== "bool")){
        return undefined;
    }
    if (typeof value === "boolean") {
        return Number(value).toString();
    }
    if (typeof value === "string") {
        const lower = value.toLowerCase();
        if (lower === "true") {
            return "1";
        } else if (lower === "false") {
            return "0";
        }
    }
    return undefined;
};

/**
 * 字符串格式化器（默认） - 将字符串值原样返回
 * @param value - 待格式化的值
 * @param _placeholder - 占位符信息（未使用）
 * @param _key - 可选键名（未使用）
 * @returns 格式化后的字符串，或 undefined
 */
const stringFormatter: CustomFormatter = (value: any, _placeholder: Placeholder, _key?: string): string | undefined => {
    if(!_key || _key !== "string"){
        return undefined;
    }
    if (typeof value === "string") {
        return value as string;
    }
    if (value !== null && value !== undefined) {
        return String(value);
    }
    return undefined;
};

/** 默认格式化器列表 */
const defaultFormatters: CustomFormatter[] = [
    dateFormatter,
    numberFormatter,
    booleanFormatter,
    stringFormatter,
    dayFormatter,
];

export {
    _getSimple,
    valueDotGet,
    defaultValueDotGet,
    resolveFullDataPath,
    dateFormatter,
    dayFormatter,
    numberFormatter,
    booleanFormatter,
    stringFormatter,
    defaultFormatters,
};
