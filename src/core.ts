// ==================== Barrel 文件 ====================
// 此文件从子模块重新导出所有公开类型/函数/类
// 详见 src/core/ 目录下的各模块文件

import { Workbook } from './core/workbook';

// 类型/接口/枚举
export type {
    Placeholder,
    Ref,
    Range,
    SheetInfo,
    SheetInfoMust,
    DrawingInfo,
    TableInfo,
    RelsInfo,
    WorkbookOptions,
    OutputByType,
    FullOptions,
    ExtensionOptions,
    CustomReplacer,
    CustomPlaceholderExtractor,
    BeforeReplaceHook,
    AfterReplaceHook,
    CustomFormatter,
    QueryFunction,
} from './core/types';

export {
    BufferType,
} from './core/types';

// 工具函数
export {
    isUrl,
    toArrayBuffer,
} from './core/xml-utils';

// 值获取/格式化
export {
    valueDotGet,
    defaultValueDotGet,
    defaultFormatters,
} from './core/formatters';

// 占位符提取
export {
    defaultExtractPlaceholders,
} from './core/placeholders';

// Workbook 类
export { Workbook };

// ==================== generateXlsxTemplate 快捷函数 ====================
import JsZip from "jszip";
import { type OutputByType, type FullOptions } from './core/types';

/**
 * xlsx 模板一键生成函数
 * 解析模板、替换数据、生成最终文件
 * @param data - 模板文件的 Buffer
 * @param values - 替换数据对象
 * @param options - JSZip 生成选项和 FullOptions
 * @returns 根据 type 参数返回不同格式的输出
 */
const generateXlsxTemplate = async function <T extends JsZip.OutputType>(
    data: Buffer,
    values: Object,
    options?: JsZip.JSZipGeneratorOptions<T> & FullOptions
): Promise<OutputByType[T]> {
    const w = await Workbook.parse(data, options);
    await w.substituteAll(values);
    return w.generate(options);
};

export {
    generateXlsxTemplate,
};
