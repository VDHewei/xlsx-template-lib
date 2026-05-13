import {Placeholder, CustomPlaceholderExtractor, ExtensionOptions} from "./types";

/**
 * 默认占位符正则表达式
 * 匹配格式：${type:name.key:subType}，其中 type、key、subType 均为可选项
 * 示例：${name}、${table:items.key}、${image:photo:jpg}
 */
const defaultRe = /\${(?:([^{}:]+?):)?([^{}:]+?)(?:\.([^{}:]+?))?(?::([^{}:]+?))??}/g;

/**
 * 默认占位符提取器
 * 从字符串中解析出所有符合占位符格式的标记
 * @param inputString - 输入字符串
 * @param options - 扩展配置选项（支持自定义正则表达式和默认解析）
 * @returns 解析出的占位符数组
 */
const defaultExtractPlaceholders: CustomPlaceholderExtractor = (inputString: string, options: ExtensionOptions): Placeholder[] => {
    const matches: Placeholder[] = [];
    // 默认正则表达式
    // 使用自定义正则表达式（如果提供）
    const re = options.customPlaceholderRegex || defaultRe;
    // 如果启用了默认解析且使用了自定义正则，先执行默认解析
    if (options.enableDefaultParsing && options.customPlaceholderRegex) {
        let match: RegExpExecArray | null;
        while ((match = defaultRe.exec(inputString)) !== null) {
            matches.push({
                placeholder: match[0],
                type: match[1] || 'normal',
                name: match[2],
                key: match[3],
                subType: match[4],
                full: match[0].length === inputString.length
            });
        }
    }
    // 执行当前正则匹配
    let match: RegExpExecArray | null;
    // 重置 lastIndex（如果正则不是全局的）
    re.lastIndex = 0;
    while ((match = re.exec(inputString)) !== null) {
        // 如果已经启用了默认解析，检查是否重复
        if (options.enableDefaultParsing && options.customPlaceholderRegex) {
            const isDuplicate = matches.some(m => m.placeholder === match![0]);
            if (isDuplicate) continue;
        }
        matches.push({
            placeholder: match[0],
            type: match[1] || 'normal',
            name: match[2],
            key: match[3],
            subType: match[4],
            full: match[0].length === inputString.length
        });
    }
    return matches;
}

export {
    defaultRe,
    defaultExtractPlaceholders,
};
