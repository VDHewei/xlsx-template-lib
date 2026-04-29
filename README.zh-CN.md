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

### 占位符类型

#### 1. 简单值（标量）

使用单个值替换占位符。

**Excel 模板：**
```
A              B
1 提取日期：${extractDate}
```

**代码：**
```typescript
const values = {
    extractDate: new Date('2024-01-15')
};
template.substitute(1, values);
```

**结果：**
```
A              B
1 提取日期：Jan-15-2024
```

**注意事项：**
- 占位符可以单独在单元格中，也可以作为文本的一部分："总计：${amount}"
- Excel 单元格格式（日期、数字、货币）会被保留

#### 2. 数组索引

直接在模板中访问特定数组元素。

**Excel 模板：**
```
A            B
1 第一个日期：${dates[0]}
2 第二个日期：${dates[1]}
```

**代码：**
```typescript
const values = {
    dates: [new Date('2024-01-01'), new Date('2024-02-01')]
};
template.substitute(1, values);
```

**结果：**
```
A            B
1 第一个日期：Jan-01-2024
2 第二个日期：Feb-01-2024
```

#### 3. 列数组

在列中水平展开数组。

**Excel 模板：**
```
A
1 ${dates}
```

**代码：**
```typescript
const values = {
    dates: [
        new Date('2024-01-01'),
        new Date('2024-02-01'),
        new Date('2024-03-01')
    ]
};
template.substitute(1, values);
```

**结果：**
```
A            B            C
1 Jan-01-2024 Feb-01-2024 Mar-01-2024
```

**注意：** 占位符必须是单元格中的唯一内容

#### 4. 表格行

从对象数组生成多行。

**Excel 模板：**
```
A               B    C
1 姓名           年龄  部门
2 ${team.name}   ${team.age}  ${team.dept}
```

**代码：**
```typescript
const values = {
    team: [
        { name: 'Alice Johnson', age: 28, dept: 'Engineering' },
        { name: 'Bob Smith', age: 34, dept: 'Marketing' },
        { name: 'Carol White', age: 25, dept: 'Sales' }
    ]
};
template.substitute(1, values);
```

**结果：**
```
A               B    C
1 姓名           年龄  部门
2 Alice Johnson  28   Engineering
3 Bob Smith      34   Marketing
4 Carol White    25   Sales
```

**语法：** `${table:数组名.属性名}`
- 数组中的每个对象创建一个新行
- 如果属性是数组，则水平展开

#### 5. 图片

在单元格中插入图片。

**Excel 模板：**
```
A          B
1 Logo：${image:companyLogo}
```

**代码：**
```typescript
const values = {
    companyLogo: '/path/to/logo.png'  // 或 Base64, Buffer
};
template.substitute(1, values);
```

**结果：**
```
A          B
1 Logo：🖼️
```

**支持的图片格式：**
- 文件路径（绝对或相对）：'/path/to/image.png'
- Base64 字符串：'data:image/png;base64,iVBORw0KG...'
- Buffer：fs.readFileSync('image.png')

**图片选项：**
```typescript
const template = new XlsxTemplate(data, {
    imageRootPath: '/absolute/path/to/images',  // 相对图片路径的基础路径
    imageRatio: 75                               // 图片缩放比例为 75%
});
```

**表格图片：**
```
A               B
1 产品          照片
2 ${products.name} ${products.photo:image}
```

**代码：**
```typescript
const values = {
    products: [
        { name: 'Product 1', photo: 'product1.jpg' },
        { name: 'Product 2', photo: 'product2.jpg' }
    ]
};
```

**结果：**
```
A               B
1 产品          照片
2 Product 1      🖼️
3 Product 2      🖼️
```

| 模板（编译前） | 渲染后结果 |
|:-------------|:-----------|
| `${contract.contractCode}` | `CTR-2024-001` |
| `${exportData.LRR.mothOrYear}` | `2024-01` |
| `${contract.contractTitle}` | `Construction Project A` |

### 编译规则配置

在规则工作表（如 `export.metadata.config`）中配置渲染规则，支持以下语法：

**⚠️ 重要规则：**
- **同一个工作表中的相同规则不能重复配置，重复配置无效或会导致编译解析异常**
- 每种规则类型在同一个工作表中必须唯一

| 规则类型 | 语法 | 说明 |
|:---------|:-----|:-----|
| **alias** | `alias: @#key => use aliasKey: @# => @#` | 字段别名映射 |
| **rowCell** | `G-AP:12=compile GenCell(...)` | 行规则配置 |
| **mergeCell** | `G-AQ:13-17=sum(...)` | 合并单元格计算 |
| **cell** | `D-7=@#[@D.MY]` | 单个单元格值赋值 |

#### Alias 规则

为变量名或变量取值表达式定义别名缩写。支持多行配置。

**语法：** `alias 缩写=原变量/原表达式`

**规则：**
- 别名缩写在同一工作表中必须唯一
- 别名可在表达式中使用 `@别名` 引用

**示例：**

| 别名配置 | 说明 |
|:---------|:-----|
| `T=template` | 将 `T` 映射到 `template` |
| `LLR=exportData.LRR` | 将 `LLR` 映射到 `exportData.LRR` |

**在表达式中使用：**

```
使用前: ${exportData.LRR.value}
使用后:  ${@LLR.value}
```

#### RowCell 规则

配置行规则，为单元格范围赋值。支持多行配置。

**语法：** `列起始号-列结束号:行号=值表达式`

**范围格式：**
- `列起始号-列结束号:行号`
- 示例：`G-AP:12`（G 到 AP 列，第 12 行）

**值表达式：**
- 通常使用 `compile:GenCell` 宏替换或 `compile:Macro` 展开

**示例：**

| 规则 | 说明 |
|:-----|:-----|
| `G-AP:12=compile GenCell(@#item,[compile Macro]#index@0)` | 为第 12 行的 G-AP 列赋值生成的值 |
| `A-Z:5=compile Macro(@#data,2,5,!!codeKey)` | 为第 5 行的 A-Z 列赋值格式化后的单元格值 |

#### MergeCell 规则

合并单元格并应用计算函数。支持多行配置。

**语法：** `列起始号-列结束号:行起始号-行结束号=函数表达式`

**范围格式：**
- `列起始号-列结束号:行起始号-行结束号`
- 示例：`G-AQ:13-17`（G 到 AQ 列，第 13 到 17 行）

**函数表达式：**
- 通常使用 `sum` 或 `sub` 函数配合宏替换

**示例：**

| 规则 | 说明 |
|:-----|:-----|
| `G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)` | 合并并计算第 13-17 行的和 |
| `G-AQ:18-35=sub(@LT,[compile:Macro(exprArr,F,18,35,!!codeKey)],compile:Macro(index),0)` | 合并并计算第 18-35 行的差 |

#### Cell 规则

为单个单元格赋值。

**语法：** `列号:行号=值表达式`

**坐标格式：**
- `列号:行号`
- 示例：`D:7`（D 列，第 7 行）

**值表达式：**
- `compile:Macro` 展开
- 变量占位符：`${变量}`
- 带别名的变量占位符：`${@别名}`

**示例：**

| 规则 | 说明 |
|:-----|:-----|
| `D:7=@#[@D.MY]` | 从表达式赋值 |
| `A:1=${contractCode}` | 从变量占位符赋值 |
| `B:1=${@LLR.value}` | 从带别名的变量赋值 |

#### 计算函数

**sum 函数**

计算多个值的和。

**语法：** `sum(值根,[值子项...],值后缀,默认值)`

**参数：**
- `值根`: 所有值表达式的共同父级
- `值子项`: 各级值子项数组
- `值后缀`: 每个值的结尾值
- `默认值`: 当统计值之和为 0 时返回的值（undefined 不会返回默认值）

**示例：**
```
sum(orders,[cat,food,game],1,0)
// 相当于: orders.cat.1 + orders.food.1 + orders.game.1
```

**sub 函数**

计算多个值的差。

**语法：** `sub(值根,[值子项...],值后缀,默认值)`

**参数：**
- `值根`: 所有值表达式的共同父级
- `值子项`: 各级值子项数组
- `值后缀`: 每个值的结尾值
- `默认值`: 当统计值之差为 0 时返回的值（undefined 不会返回默认值）

**示例：**
```
sub(orders,[money,food,game],1,0)
// 相当于: orders.money.1 - orders.food.1 - orders.game.1
```

### 宏替换规则

本库支持强大的宏替换函数，用于动态生成单元格内容：

#### GenCell 宏

通过连接多个部分生成单元格表达式：

| 语法 | 说明 | 示例 | 结果 |
|:-----|:-----|:-----|:-----|
| `compile:GenCell(expr1,expr2,...,exprN)` | 使用默认分隔符 `·` 连接 | `GenCell(test,1,2)` | `test·1·2` |
| `compile:GenCell(expr1,expr2,...,exprN,"sep")` | 使用自定义分隔符连接 | `GenCell(test,1,2,"_")` | `test_1_2` |
| `compile:GenCell(expr1,expr2,...,exprN,"")` | 不使用分隔符连接 | `GenCell(test,1,2,"")` | `test12` |

#### Macro 宏展开

##### 单个单元格宏

展开为单个单元格的值：

| 语法 | 说明 | 示例 |
|:-----|:-----|:-----|
| `compile:Macro(expr,columnNum,rowNum)` | 获取 (columnNum, rowNum) 位置单元格的值 | `Macro(data,2,5)` |
| `compile:Macro(expr,columnNum,rowNum,MacroFormatter)` | 获取格式化后的值 | `Macro(data,2,5,!!codeKey)` |

**参数说明：**
- `expr`: 基础表达式
- `columnNum`: 列号（从 1 开始）
- `rowNum`: 行号（从 1 开始）
- `MacroFormatter`: 可选的格式化器（见下方）

##### 多个单元格宏

展开为多个单元格的值：

| 语法 | 说明 | 示例 |
|:-----|:-----|:-----|
| `compile:Macro(exprArr,columnNum,rowStartNum,rowEndNum)` | 获取单元格范围内的值 | `Macro(data,1,1,5)` |
| `compile:Macro(exprArr,columnNum,rowStartNum,rowEndNum,MacroFormatter)` | 获取格式化后的值 | `Macro(data,1,1,5,!!number)` |

**参数说明：**
- `exprArr`: 基础表达式数组
- `columnNum`: 列号（从 1 开始）
- `rowStartNum`: 起始行号（从 1 开始）
- `rowEndNum`: 结束行号（从 1 开始）
- `MacroFormatter`: 可选的格式化器

##### Index 宏

生成从 1 开始的迭代序列：

| 语法 | 说明 | 使用示例 | 结果 |
|:-----|:-----|:---------|:-----|
| `compile:Macro(index)` | 自动递增索引（从 1 开始） | 第 1 行: `Macro(index)` | `1` |
| | | 第 2 行: `Macro(index)` | `2` |
| | | 第 3 行: `Macro(index)` | `3` |

#### Macro 格式化器

使用以 `!!` 开头的特殊格式化器格式化宏输出：

| 格式化器 | 说明 | 输入 | 输出 |
|:---------|:-----|:-----|:-----|
| `!!codeKey` | 将特殊字符（`@-[]{}\/'.`）转换为 `_`，删除多余 `__`，去除首尾 `_`，转为大写 | `test..x` | `TEST_X` |
| | | `@data-value` | `DATA_VALUE` |
| | | `[item].name` | `ITEM_NAME` |
| `!!codeKeyAlias` | 与 `!!codeKey` 相同，但添加前缀（默认 `@`） | `test..x` | `@TEST_X` |
| | （默认前缀 `@`）| `data.value` | `@DATA_VALUE` |
| `!!number` | 转换为十进制整数，支持 `0x` 十六进制前缀 | `123` | `123` |
| | | `0xFF` | `255` |
| | | `abc` | `abc`（保持不变，NaN） |

**CodeKey 转换规则：**
- 替换的特殊字符：`@`, `-`, 空格, `[`, `]`, `{`, `}`, `\`, `/`, `'`, `.`
- 连续的多个 `__` 合并为单个 `_`
- 删除开头和结尾的 `_`
- 最终结果转为大写

#### Macro 使用示例

**示例 1: 使用行单元格生成 CodeKey**

```
规则: G-AQ:117=compile:GenCell(#LT,compile:Macro(index),0)
结果: errValue·1, errValue·2, errValue·3, ...
```

**示例 2: 使用 CodeKey 格式化单元格值**

```
规则: D-7=compile:Macro(expr,5,7,!!codeKey)
如果 cell(5,7) = "project-alpha-2024"
结果: PROJECT_ALPHA_2024
```

**示例 3: 生成 CodeKeyAlias**

```
规则: cell F-10=compile:Marco(expr,3,10,!!codeKeyAlias)
如果 cell(3,10) = "test..data"
结果: @TEST_DATA
```

**示例 4: 数字转换**

```
规则: row-5=compile:Macro(expr,2,5,!!number)
如果 cell(2,5) = "42"
结果: 42

规则: row-6=compile Macro(expr,4,6,!!number)
如果 cell(4,6) = "0x1A"
结果: 26
```

**示例 5: 使用 Index 迭代**

```
第 1 行: Code-compile:Macro(index)  →  Code-1
第 2 行: Code-compile:Macro(index)  →  Code-2
第 3 行: Code-compile:Macro(index)  →  Code-3
```

#### 完整规则配置示例

规则工作表（`export.metadata.config`）的完整示例，包含所有规则类型：

```
# Alias 规则（为长表达式定义快捷方式）
T=template
LLR=exportData.LRR
CTR=contract.contractCode

# RowCell 规则（为单元格范围赋值）
G-AQ:12=compile:GenCell(@#item,G,compile:Macro(index),0)
A-Z:5=compile:Macro(@#data,2,5,!!codeKey)

# MergeCell 规则（合并单元格并应用计算）
G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)
G-AQ:18-35=sub(@LT,[compile:Macro(exprArr,F,18,35,!!codeKey)],compile:Macro(index),0)

# Cell 规则（为单个单元格赋值）
D:7=@#[@D.MY]
A:1=${@CTR}
B:1=${@LLR.value}
```

**⚠️ 重要注意事项：**
- 每种规则类型（alias、rowCell、mergeCell、cell）可以出现多次
- 但**相同规则的重复配置无效或会导致编译解析异常**
- 别名缩写在同一工作表中必须唯一
- 行/列范围不能以冲突的方式重叠

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

CLI 工具 `xlsx-cli` 提供了命令行接口用于快速处理 Excel 模板。

#### 安装

```bash
npm install -g @vdhewei/xlsx-template-lib
```

或直接使用 `npx`:

```bash
npx @vdhewei/xlsx-template-lib <命令> [选项]
```

#### 命令

##### 1. compile 命令

编译带有规则配置的 Excel 文件。

```bash
xlsx-cli compile <xlsx-文件> [选项]
```

**参数:**
- `<xlsx-文件>` - Excel 文件路径

**选项:**
- `-s, --save <string>` - 将编译后的文件保存到指定目录（默认：当前目录）
- `-n, --sheet-name <string>` - 要编译的工作表名称（默认：第一个工作表）
- `-r, --remove` - 编译后移除配置规则工作表（默认：false）

**示例:**

```bash
# 使用默认设置编译
xlsx-cli compile template.xlsx

# 编译并保存到指定位置
xlsx-cli compile template.xlsx -s ./output/

# 编译指定工作表
xlsx-cli compile template.xlsx -n Sheet1

# 编译并移除配置工作表
xlsx-cli compile template.xlsx -r

# 完整示例
xlsx-cli compile template.xlsx -s ./output/ -n Sheet1 -r
```

**输出:**
- 编译后的 Excel 文件保存为 `<文件名>_<时间戳>.xlsx`
- 成功消息以绿色显示
- 错误以红色显示并返回退出码 1

##### 2. render 命令

使用数据替换渲染 Excel 模板。

```bash
xlsx-cli render <xlsx-文件> [选项]
```

**参数:**
- `<xlsx-文件>` - Excel 模板文件路径

**选项:**
- `-c, --compile` - 渲染前自动编译规则（默认：false）
- `-n, --sheet-name <string>` - 要渲染的工作表名称（默认：第一个工作表）
- `-s, --save <string>` - 将渲染后的文件保存到指定目录（默认：当前目录）
- `-d, --data <string>` - 渲染数据源（JSON 字符串、文件路径或 URL）
- `--header <string>` - 远程数据获取的 HTTP 请求头（可指定多次，格式：`Key:Value`）
- `--body <string>` - POST 请求的 HTTP 请求体

**示例:**

```bash
# 使用空数据基本渲染
xlsx-cli render template.xlsx

# 使用 JSON 字符串渲染
xlsx-cli render template.xlsx -d '{"name":"张三","age":30}'

# 使用 JSON 文件渲染
xlsx-cli render template.xlsx -d ./data.json

# 使用远程 JSON URL 渲染
xlsx-cli render template.xlsx -d 'https://api.example.com/data.json'

# 渲染并自动编译
xlsx-cli render template.xlsx -c -d './data.json'

# 渲染指定工作表
xlsx-cli render template.xlsx -n Sheet1 -d './data.json'

# 使用自定义 HTTP 请求头渲染
xlsx-cli render template.xlsx -d 'https://api.example.com/data.json' --header 'Authorization:Bearer token123' --header 'Content-Type:application/json'

# 使用 POST 请求体渲染
xlsx-cli render template.xlsx -d 'https://api.example.com/api/query' --body '{"query":"SELECT * FROM users"}' --header 'Content-Type:application/json'

# 通过 header 指定 POST 方法
xlsx-cli render template.xlsx -d 'https://api.example.com/api/create' --body '{"name":"测试"}' --header 'Content-Type:application/json' --header 'method:POST'

# 完整示例
xlsx-cli render template.xlsx -c -n Sheet1 -s ./output/ -d './data.json'
```

**数据源:**
- **JSON 字符串**: 直接使用单引号括起来的 JSON 字符串
- **本地文件**: `.json` 文件的路径（相对或绝对）
- **远程 URL**: 返回 JSON 的 HTTP/HTTPS URL

**HTTP 请求选项（用于远程 URL）:**
- **请求头**: 使用 `--header` 添加自定义 HTTP 请求头（可指定多次）
  - 格式：`--header 'Key:Value'`
  - 示例：`--header 'Authorization:Bearer token123' --header 'Content-Type:application/json'`
  - 特殊请求头：`method:POST` 可设置 HTTP 方法为 POST
- **请求体**: 使用 `--body` 发送请求体（通常用于 POST 请求）
  - 格式：`--body '{"key":"value"}'`
  - 提供请求体时自动使用 POST 方法
- **默认行为**: 无请求头的 GET 请求

**HTTP 请求示例:**

```bash
# 带自定义请求头的 GET 请求
xlsx-cli render template.xlsx \
  -d 'https://api.example.com/data.json' \
  --header 'Authorization:Bearer your-token' \
  --header 'X-API-Key:api-key-123'

# 带 JSON 请求体的 POST 请求
xlsx-cli render template.xlsx \
  -d 'https://api.example.com/api/query' \
  --body '{"query":"SELECT * FROM users LIMIT 10"}' \
  --header 'Content-Type:application/json'

# 通过 header 指定 POST 方法
xlsx-cli render template.xlsx \
  -d 'https://api.example.com/api/create' \
  --body '{"name":"新记录","value":100}' \
  --header 'Content-Type:application/json' \
  --header 'method:POST'

# 复杂示例：带认证和查询请求体
xlsx-cli render template.xlsx \
  -d 'https://api.example.com/v1/export' \
  --header 'Authorization:Bearer eyJhbGc...' \
  --header 'Content-Type:application/json' \
  --body '{"format":"xlsx","filter":{"status":"active"}}' \
  -c -n Sheet1 -s ./output/
```

**HTTP 请求详细说明:**

1. **HTTP 方法确定**:
   - 默认：`GET`
   - 使用 `--body`：自动变为 `POST`
   - 使用 `method:POST` 请求头：显式设置为 `POST`
   - 使用 `method:GET` 请求头：显式设置为 `GET`

2. **请求头格式**:
   - 请求头按 `Key:Value` 格式解析
   - 可以使用多个 `--header` 选项
   - 示例：`--header 'Accept:application/json' --header 'User-Agent:MyApp/1.0'`

3. **错误处理**:
   - 非 200 状态码返回 `undefined` 并显示错误消息
   - 网络错误会被捕获并以红色显示
   - 缺少 `node-fetch`（Node.js < 18）会显示错误消息

4. **支持的数据格式**:
   - JSON 对象：`{"key":"value"}`
   - JSON 数组：`[{"id":1},{"id":2}]`
   - 嵌套结构：`{"user":{"name":"张三","age":30}}`

**输出:**
- 渲染后的 Excel 文件保存为 `<文件名>_<时间戳>.xlsx`
- 检查工作表是否存在
- 使用适当的颜色显示成功/错误消息

##### 3. rules 命令

向 Excel 文件添加规则配置。

```bash
xlsx-cli rules <xlsx-文件> [选项]
```

**参数:**
- `<xlsx-文件>` - Excel 文件路径

**选项:**

**模式 1：命令行规则**
- `-t, --type <string>` - 规则类型：`cell`、`alias`、`rowCell`、`mergeCell`（使用 -r 时必需）
- `-r, --rule <string>` - 规则表达式字符串（可指定多次）

**模式 2：文件规则**
- `-f, --file <string>` - 从文件读取规则（格式：每行 `<类型> 规则表达式`）
  - 以 `#` 开头的行被视为注释
  - 空行将被跳过
  - 规则类型：`cell`、`alias`、`rowCell`、`mergeCell`

**通用选项:**
- `-s, --save <string>` - 保存编译后的文件到指定目录（默认：当前目录）

**示例:**

**单个规则（命令行）：**
```bash
# 添加 alias 规则
xlsx-cli rules template.xlsx -t alias -r 'T=template'

# 添加 cell 规则
xlsx-cli rules template.xlsx -t cell -r 'D:7=${@LLR.value}'

# 添加 rowCell 规则
xlsx-cli rules template.xlsx -t rowCell -r 'G-AQ:12=compile GenCell(@#item,[compile Macro]#index@0)'

# 添加 mergeCell 规则
xlsx-cli rules template.xlsx -t mergeCell -r 'G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)'
```

**多个规则（命令行）：**
```bash
# 添加同类型的多个规则
xlsx-cli rules template.xlsx -t cell -r 'D:7=${@LLR.value}' -r 'A:1=${@T}' -r 'B:1=${@LLR.value}'
```

**从文件读取规则：**
```bash
# 从文件读取规则
xlsx-cli rules template.xlsx -f rules.txt

# 创建 rules.txt 文件：
# 这是注释行
alias T=template
alias LLR=exportData.LRR
cell D:7=${@T}
cell A:1=${@LLR.value}
rowCell G-AQ:12=compile GenCell(@#item,[compile Macro]#index@0)
mergeCell G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)
```

**保存到指定目录：**
```bash
xlsx-cli rules template.xlsx -f rules.txt -s ./output/
xlsx-cli rules template.xlsx -t cell -r 'D:7=${@LLR.value}' -s ./output/
```

**文件格式（-f 模式）：**
```bash
# 格式：<类型> 规则表达式
# 注释行以 # 开头
# 有效类型：cell、alias、rowCell、mergeCell

cell D:7=${@LLR.value}
alias T=template
rowCell G-AQ:12=compile GenCell(@#item,[compile Macro]#index@0)
mergeCell G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)
```

**行为:**
- 如果不存在则创建 `export_metadata.config` 工作表
- 添加规则并应用样式：类型字段加粗+居中，表达式居中
- 根据内容自动调整列宽
- 每种规则类型（cell、alias、rowCell、mergeCell）每行支持最多 4 个规则
- 如果为同一类型添加超过 4 个规则，自动创建新行
- 支持从命令行或文件批量添加规则
- 输出带时间戳的新文件

#### 通用特性

**环境变量:**
- CLI 自动从当前目录加载 `.env` 文件（如果存在）

**文件路径解析:**
- 支持绝对和相对路径
- 解析相对于当前工作目录的路径
- 处理前验证文件是否存在

**错误处理:**
- 所有错误使用 chalk 以红色显示
- 出错时返回非零退出码
- 提供详细的错误消息用于调试

**跨平台支持:**
- 在 Windows、Linux 和 macOS 上运行
- 使用平台无关的路径处理

**输出文件名格式:**
- 默认：`<输入文件名>_<时间戳>.xlsx`
- 时间戳为自纪元以来的毫秒数
- 保留原始文件名

**详细日志:**
- 灰色信息消息显示处理步骤
- 绿色成功消息
- 红色错误消息
- 黄色警告

## 高级功能

### 自定义命令（渲染函数）

```typescript
import { AddCommand, generateCommandsXlsxTemplate } from '@vdhewei/xlsx-template-lib';

// 添加自定义命令
AddCommand('multiply', (obj, args) => {
  const values = args.groups.map(g => valueDotGet(obj, g));
  return values.reduce((a, b) => a * b, 1);
});
// 读取xlsx模板
const data = await fs.readFile('simple.xlsx');
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

| 步骤 | 输入                    | 输出 | 说明 |
|:-----|:----------------------|:-----|:-----|
| 1. 加载 | `template.xlsx` Buffer | `ZipXlsxTemplateApp` | 加载模板文件 |
| 2. 编译 | 规则配置工作表               | 编译后的规则 | 解析 mergeCell/cell/rowCell 规则 |
| 3. 替换 | 数据对象                  | 渲染后的工作表 | 替换 `${...}` 占位符 |
| 4. 生成 | -                     | `output.xlsx` Buffer | 最终输出文件 |

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

## 致谢

本项目受到了优秀的开源项目 [xlsx-template](https://github.com/optilude/xlsx-template)（由 optilude 开发）的启发。

**xlsx-template** 为基于模板的 Excel 报表生成和动态数据替换提供了坚实的基础。本库的许多概念和设计模式都受到了 xlsx-template 的影响，包括：

- 基于模板的 Excel 文件生成
- 占位符替换语法
- 数组和表格展开
- 图片插入和定位
- 单元格格式保留

我们向 xlsx-template 团队和贡献者致以诚挚的感谢，感谢他们在开源社区中的宝贵工作。

**原 xlsx-template 仓库：** https://github.com/optilude/xlsx-template
