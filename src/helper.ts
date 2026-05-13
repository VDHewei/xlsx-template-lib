import exceljs from "exceljs";

// ==================== 重新导出核心配置解析模块 ====================
// 配置解析相关代码已迁移至 src/core/config-parser.ts
// 此处通过重新导出保持向后兼容性

export {
    CellPosition,
    CompileResult,
    DefaultPlaceholderCellValue,
    PlaceholderCellValue,
    RuleToken,
    RuleMapOptions,
    TokenParserManger,
    RuleResult,
    RuleOptions,
    CompileContext,
    FilterMacroResult,
    MacroUnitHelper,
    MacroArgs,
    ExtractMacroArgs,
    ExprResolver,
    toCellValue,
    scanCellSetPlaceholder,
    workSheetSetPlaceholder,
    parseWorkSheetRules,
    columnLetterToNumber,
    columnNumberToLetter,
    isRuleToken,
    hasGeneratorToken,
    getTokenParser,
    registerTokenParser,
    registerTokenParserMust,
    compileWorkSheet,
    compileWorkSheetPlaceholder,
    loadWorkbook,
    loadCompileSheets,
} from './core/config-parser';

export { exceljs };
