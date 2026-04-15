# xlsx-template-lib

基于 ExcelJS 的强大 XLSX 模板渲染库，支持基于模板的 Excel 文件生成和数据替换。

## 功能特性

- **模板渲染**：使用带有占位符的模板渲染 Excel 文件
- **数据替换**：支持表达式和函数的动态数据替换
- **规则配置**：通过专用规则工作表配置渲染规则
- **自定义命令**：通过自定义命令函数扩展功能（如 `fn:sum`, `fn:sub`）
- **ZIP 支持**：批量处理 ZIP 压缩包中的多个 Excel 文件
- **TypeScript 支持**：完整的 TypeScript 支持和类型定义
- **CLI 工具**：提供命令行工具快速处理

## 安装

```bash
npm install @vdhewei/xlsx-template-lib
```

## 模板语法

### 占位符格式

在 Excel 单元格中使用 `${变量名}` 格式进行数据替换：

| 模板（编译前） | 渲染后结果 |
|:-------------|:-----------|
| `${contract.contractCode}` | `CTR-2024-001` |
| `${exportData.LRR.mothOrYear}` | `2024-01` |
| `${contract.contractTitle}` | `Construction Project A` |

### 编译规则配置

在规则工作表（如 `export.metadata.config`）中配置渲染规则，支持以下语法：

| 规则类型 | 语法 | 说明 |
|:---------|:-----|:-----|
| **alias** | `alias: @#key => use aliasKey: @# => @#` | 字段别名映射 |
| **mergeCell** | `G-AQ(1-17)=sum(#LT[compile Macro]#err@F113.17[#codeKey],[compile Macro]#index@0)` | 合并单元格计算 |
| **mergeCell（续）** | `G-AQ(18-35)=sum(#LT[compile Macro]#err@F118.20[#codeKey],[compile Macro]#index@0)` | 连续合并单元格 |
| **cell** | `D-7=@#[@D.MY]` | 单元格值赋值 |
| **rowCell** | `G-AQ117=compile GenCell(#LT[compile Macro]#err@F118[#codeKey],[compile Macro]#index@0)` | 行单元格生成 |

### 渲染函数

内置和自定义的数据处理函数：

| 函数 | 语法 | 示例 |
|:-----|:-----|:-----|
| **sum** | `fn:sum(...values)` | `fn:sum(10, 20, 30)` => `60` |
| **sub** | `fn:sub(a, b)` | `fn:sub(100, 30)` => `70` |
| **自定义函数** | `fn:customName(...args)` | 用户自定义逻辑 |

## 快速开始

### 基础用法

```typescript
import { ZipXlsxTemplateApp } from '@vdhewei/xlsx-template-lib';
import * as fs from 'node:fs/promises';

// 从 buffer 加载模板
const templateBuffer = await fs.readFile('template.xlsx');
const app = new ZipXlsxTemplateApp(templateBuffer);

// 使用数据渲染
const data = {
  contract: {
    contractCode: 'CTR-2024-001',
    contractTitle: 'Construction Project A'
  },
  exportData: {
    LRR: {
      mothOrYear: '2024-01'
    }
  }
};

await app.substituteAll(data);

// 生成输出
const output = await app.generate();
await fs.writeFile('output.xlsx', output);
```

### 使用 XlsxRender

```typescript
import { XlsxRender } from '@vdhewei/xlsx-template-lib';

const templateBuffer = await fs.readFile('template.xlsx');
const xlsx = await XlsxRender.create(templateBuffer);

// 渲染特定工作表
await xlsx.render({ 
  contract: { contractCode: 'CTR-001' } 
}, 'Sheet1');

// 生成输出
const buffer = await xlsx.generate();
```

### 带规则配置的编译和渲染

```typescript
import { 
  ZipXlsxTemplateApp, 
  compileAll,
  AddCommand 
} from '@vdhewei/xlsx-template-lib';

// 添加自定义渲染函数
AddCommand('sum', (obj, args) => {
  return args.groups.reduce((acc, val) => acc + Number(val), 0);
});

// 配置编译选项
const compileOpts = {
  sheetName: 'export.metadata.config',  // 规则配置工作表
  remove: true                           // 编译后移除规则工作表
};

const zipBuffer = await fs.readFile('template.xlsx');
const result = await ZipXlsxTemplateApp.compileTo(zipBuffer, {
  checker: async (buf, opts, values, fileName) => {
    // 自定义验证逻辑
    return buf;
  },
  options: compileOpts
}, renderData);
```

### CLI 工具

```bash
# 编译模板
xlsx-cli compile template.xlsx -s compiled.xlsx -r

# 渲染模板
xlsx-cli render template.xlsx -s output.xlsx -c
```

## 高级功能

### 自定义命令（渲染函数）

```typescript
import { AddCommand, generateCommandsXlsxTemplate } from '@vdhewei/xlsx-template-lib';

// 添加自定义命令
AddCommand('multiply', (obj, args) => {
  const values = args.groups.map(g => valueDotGet(obj, g));
  return values.reduce((a, b) => a * b, 1);
});

// 使用自定义命令生成模板
const buffer = await generateCommandsXlsxTemplate(data, options);
```

### 批量处理

```typescript
// 处理包含多个 XLSX 文件的 ZIP 文件
const zipBuffer = await fs.readFile('templates.zip');
const app = new ZipXlsxTemplateApp(zipBuffer);

const compileOpts = {
  sheetName: 'export.metadata.config',
  remove: true
};

const renderOpts = {
  // 渲染选项
};

await app.substituteAll(renderData, compileOpts, renderOpts);
const output = await app.generate();
```

## API 参考

### ZipXlsxTemplateApp

处理 ZIP 压缩包中 Excel 文件的主类。

| 方法 | 说明 |
|:-----|:-----|
| `constructor(data?: Buffer)` | 使用 ZIP buffer 初始化 |
| `loadZipBuffer(data: Buffer)` | 加载 ZIP buffer |
| `parse(data: Buffer)` | 解析 ZIP 并提取 XLSX 条目 |
| `getEntries()` | 获取所有 XLSX 文件条目 |
| `substituteAll(renderData, compileOpts?, renderOpts?)` | 替换所有占位符 |
| `generate(options?)` | 生成输出 buffer |
| `static compileAll(files, renderData?, compileOpts?)` | 编译多个文件 |
| `static compileTo(data, opts, values?)` | 使用自定义检查器编译 ZIP 中的 XLSX |

### XlsxRender

渲染单个 Excel 文件的主类。

| 方法 | 说明 |
|:-----|:-----|
| `static create(data: Buffer, option?)` | 从 buffer 创建 |
| `render(values: Object, sheetName: string)` | 渲染特定工作表 |
| `getSheets()` | 获取所有工作表信息 |
| `generate(options?)` | 生成输出 buffer |

### 辅助函数

| 函数 | 说明 |
|:-----|:-----|
| `ExprResolver` | 用于复杂表达式的表达式解析器 |
| `compileRuleSheetName` | 默认规则工作表名称 |
| `generateXlsxTemplate` | 生成 XLSX 模板 |
| `generateCommandsXlsxTemplate` | 使用自定义命令生成 |
| `AddCommand(name, fn)` | 添加自定义渲染函数 |

## 完整示例

### 模板结构

```
template.xlsx
├── Sheet1 (数据工作表，包含占位符)
│   ├── A1: ${contract.contractCode}
│   ├── B1: ${contract.contractTitle}
│   └── C1: ${exportData.LRR.mothOrYear}
└── export.metadata.config (规则配置工作表)
    ├── 第1行: alias @#key => use aliasKey: @# => @#
    ├── 第2行: mergeCell G-AQ(1-17)=sum(...)
    └── 第3行: cell D-7=@#[@D.MY]
```

### 编译与渲染流程

| 步骤 | 输入 | 输出 | 说明 |
|:-----|:-----|:-----|:-----|
| 1. 加载 | `template.xlsx` Buffer | `ZipXlsxTemplateApp` | 加载模板文件 |
| 2. 编译 | 规则配置工作表 | 编译后的规则 | 解析 mergeCell/cell/rowCell 规则 |
| 3. 替换 | 数据对象 | 渲染后的工作表 | 替换 `${...}` 占位符 |
| 4. 生成 | - | `output.xlsx` Buffer | 最终输出文件 |

```typescript
import { ZipXlsxTemplateApp, AddCommand } from '@vdhewei/xlsx-template-lib';
import * as fs from 'node:fs/promises';

// 定义自定义函数
AddCommand('calculateTotal', (obj, args) => {
  const base = Number(args.root);
  const multiplier = args.groups.length > 0 ? Number(args.groups[0]) : 1;
  return base * multiplier;
});

// 主处理流程
async function processTemplate() {
  const templateBuffer = await fs.readFile('template.xlsx');
  
  const app = new ZipXlsxTemplateApp(templateBuffer);
  
  const renderData = {
    contract: {
      contractCode: 'CTR-2024-001',
      contractTitle: 'Monthly Return of Site Labour'
    },
    exportData: {
      LRR: {
        mothOrYear: '2024-01',
        workCode: 'WC-001'
      }
    }
  };
  
  const compileOpts = {
    sheetName: 'export.metadata.config',
    remove: true  // 在输出中移除规则工作表
  };
  
  await app.substituteAll(renderData, compileOpts);
  
  const output = await app.generate();
  await fs.writeFile('output.xlsx', output);
}

processTemplate();
```

## 注意事项

- `test_data` 目录包含内部测试模板，仅供私有使用，不可外传
- 规则配置工作表支持：`alias`、`mergeCell`、`cell`、`rowCell` 规则类型
- 可通过 `AddCommand(name, handler)` 注册自定义函数

## 许可证

MIT

## 作者

VDHewei

## 仓库

https://github.com/VDHewei/xlsx-template-lib

## 贡献

欢迎贡献！请随时提交 Pull Request。
