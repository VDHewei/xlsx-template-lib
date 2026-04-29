# xlsx-template-lib

[中文文档](./README.zh-CN.md)

A powerful XLSX template rendering library based on ExcelJS, supporting template-based Excel file generation and data substitution.

## Features

- **Template Rendering**: Render Excel files using templates with placeholders
- **Data Substitution**: Dynamic data substitution with expressions and functions
- **Rule Configuration**: Configure rendering rules via dedicated rule sheets
- **Custom Commands**: Extend functionality with custom command functions (e.g., `fn:sum`, `fn:sub`)
- **ZIP Support**: Batch process multiple Excel files within a ZIP archive
- **TypeScript Support**: Full TypeScript support with type definitions
- **CLI Tool**: Command-line interface for quick processing

## Installation

```bash
npm install @vdhewei/xlsx-template-lib
```

## Template Syntax

### Placeholder Format

Use `${variableName}` format in Excel cells for data substitution:

| Template (Before) | Rendered (After) |
|:------------------|:-----------------|
| `${contract.contractCode}` | `CTR-2024-001` |
| `${exportData.LRR.mothOrYear}` | `2024-01` |
| `${contract.contractTitle}` | `Construction Project A` |

### Placeholder Types

#### 1. Simple Values (Scalars)

Replace a placeholder with a single value.

**Excel template:**
```
A              B
1 Extracted on: ${extractDate}
```

**Code:**
```typescript
const values = {
    extractDate: new Date('2024-01-15')
};
template.substitute(1, values);
```

**Result:**
```
A              B
1 Extracted on: Jan-15-2024
```

**Notes:**
- Placeholders can be standalone in a cell or part of text: "Total: ${amount}"
- Excel cell formatting (date, number, currency) is preserved

#### 2. Array Indexing

Access specific array elements directly in templates.

**Excel template:**
```
A            B
1 First date: ${dates[0]}
2 Second date: ${dates[1]}
```

**Code:**
```typescript
const values = {
    dates: [new Date('2024-01-01'), new Date('2024-02-01')]
};
template.substitute(1, values);
```

**Result:**
```
A            B
1 First date: Jan-01-2024
2 Second date: Feb-01-2024
```

#### 3. Column Arrays

Expand an array horizontally across columns.

**Excel template:**
```
A
1 ${dates}
```

**Code:**
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

**Result:**
```
A            B            C
1 Jan-01-2024 Feb-01-2024 Mar-01-2024
```

**Note:** The placeholder must be the only content in its cell

#### 4. Table Rows

Generate multiple rows from an array of objects.

**Excel template:**
```
A               B    C
1 Name           Age  Department
2 ${team.name}   ${team.age}  ${team.dept}
```

**Code:**
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

**Result:**
```
A               B    C
1 Name           Age  Department
2 Alice Johnson  28   Engineering
3 Bob Smith      34   Marketing
4 Carol White    25   Sales
```

**Syntax:** `${table:arrayName.propertyName}`
- Each object in the array creates a new row
- If a property is an array, it expands horizontally

#### 5. Images

Insert images into cells.

**Excel template:**
```
A          B
1 Logo: ${image:companyLogo}
```

**Code:**
```typescript
const values = {
    companyLogo: '/path/to/logo.png'  // or Base64, Buffer
};
template.substitute(1, values);
```

**Result:**
```
A          B
1 Logo: 🖼️
```

**Supported image formats:**
- File path (absolute or relative): '/path/to/image.png'
- Base64 string: 'data:image/png;base64,iVBORw0KG...'
- Buffer: fs.readFileSync('image.png')

**Image options:**
```typescript
const template = new XlsxTemplate(data, {
    imageRootPath: '/absolute/path/to/images',  // Base path for relative image paths
    imageRatio: 75                               // Scale images to 75%
});
```

**Table images:**
```
A               B
1 Product        Photo
2 ${products.name} ${products.photo:image}
```

**Code:**
```typescript
const values = {
    products: [
        { name: 'Product 1', photo: 'product1.jpg' },
        { name: 'Product 2', photo: 'product2.jpg' }
    ]
};
```

**Result:**
```
A               B
1 Product        Photo
2 Product 1      🖼️
3 Product 2      🖼️
```

| Template (Before) | Rendered (After) |
|:------------------|:-----------------|
| `${contract.contractCode}` | `CTR-2024-001` |
| `${exportData.LRR.mothOrYear}` | `2024-01` |
| `${contract.contractTitle}` | `Construction Project A` |

### Compile Rules Configuration

Configure rendering rules in a rule sheet (e.g., `export.metadata.config`) with the following syntax:

**⚠️ Important Rule:**
- **Duplicate rules within the same sheet are invalid and may cause compilation errors**
- Each rule type must have unique configurations within the same sheet

| Rule Type | Syntax | Description |
|:----------|:-------|:------------|
| **alias** | `alias: @#key => use aliasKey: @# => @#` | Alias for field mapping |
| **rowCell** | `G-AP:12=compile:GenCell(...)` | Row rule configuration |
| **mergeCell** | `G-AQ:13-17=sum(...)` | Merge cell calculation |
| **cell** | `D-7=@#[@D.MY]` | Single cell value assignment |

#### Alias Rules

Define abbreviations for variable names or variable expression values. Multiple alias configurations are supported.

**Syntax:** `alias abbreviation=originalVariable/originalExpression`

**Rules:**
- Alias abbreviations must be unique within the same sheet
- Aliases can be referenced using `@aliasName` in expressions

**Examples:**

| Alias Configuration | Description |
|:--------------------|:------------|
| `T=template` | Map `T` to `template` |
| `LLR=exportData.LRR` | Map `LLR` to `exportData.LRR` |

**Usage in Expressions:**

```
Before: ${exportData.LRR.value}
After:  ${@LLR.value}
```

#### RowCell Rules

Configure row rules to assign values to cell ranges. Multiple rowCell configurations are supported.

**Syntax:** `columnStartNum-columnEndNum:rowNum=valueExpression`

**Range Format:**
- `columnStartNum-columnEndNum:rowNum`
- Example: `G-AP:12` (columns G to AP, row 12)

**Value Expression:**
- Typically uses `compile:GenCell` macro replacement or `compile:Macro` expansion

**Examples:**

| Rule | Description |
|:-----|:------------|
| `G-AP:12=compile:GenCell(@#item,[compile:Macro]#index@0)` | Assign generated values to row 12, columns G-AP |
| `A-Z:5=compile:Macro(@#data,2,5,!!codeKey)` | Assign formatted cell value to row 5, columns A-Z |

#### MergeCell Rules

Merge cells and apply calculation functions. Multiple mergeCell configurations are supported.

**Syntax:** `columnStartNum-columnEndNum:rowStartNum-rowEndNum=functionExpression`

**Range Format:**
- `columnStartNum-columnEndNum:rowStartNum-rowEndNum`
- Example: `G-AQ:13-17` (columns G to AQ, rows 13 to 17)

**Function Expression:**
- Typically uses `sum` or `sub` functions with macro replacements

**Examples:**

| Rule | Description |
|:-----|:------------|
| `G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)` | Merge and calculate sum for rows 13-17 |
| `G-AQ:18-35=sub(@LT,[compile:Macro(exprArr,F,18,35,!!codeKey)],compile:Macro(index),0)` | Merge and calculate difference for rows 18-35 |

#### Cell Rules

Assign values to single cells.

**Syntax:** `columnNum:rowNum=valueExpression`

**Coordinate Format:**
- `columnNum:rowNum`
- Example: `D:7` (column D, row 7)

**Value Expression:**
- `compile:Macro` expansion
- Variable placeholder: `${variable}`
- Variable placeholder with alias: `${@alias}`

**Examples:**

| Rule | Description |
|:-----|:------------|
| `D:7=@#[@D.MY]` | Assign value from expression |
| `A:1=${contractCode}` | Assign from variable placeholder |
| `B:1=${@LLR.value}` | Assign from aliased variable |

#### Calculation Functions

**sum Function**

Calculate the sum of multiple values.

**Syntax:** `sum(valueRoot,[valueItems...],valueSuffix,defaultValue)`

**Parameters:**
- `valueRoot`: Common parent of all value expressions
- `valueItems`: Array of child value items
- `valueSuffix`: Suffix for each value
- `defaultValue`: Value to return when sum equals 0 (undefined will not return default)

**Example:**
```
sum(orders,[cat,food,game],1,0)
// Equivalent to: orders.cat.1 + orders.food.1 + orders.game.1
```

**sub Function**

Calculate the difference of multiple values.

**Syntax:** `sub(valueRoot,[valueItems...],valueSuffix,defaultValue)`

**Parameters:**
- `valueRoot`: Common parent of all value expressions
- `valueItems`: Array of child value items
- `valueSuffix`: Suffix for each value
- `defaultValue`: Value to return when difference equals 0 (undefined will not return default)

**Example:**
```
sub(orders,[money,food,game],1,0)
// Equivalent to: orders.money.1 - orders.food.1 - orders.game.1
```

### Macro Replacement Rules

The library supports powerful macro replacement functions for dynamic cell content generation:

#### GenCell Macro

Generate cell expressions by concatenating multiple parts:

| Syntax | Description | Example | Result |
|:-------|:------------|:--------|:--------|
| `compile:GenCell(expr1,expr2,...,exprN)` | Concatenate with default separator `·` | `GenCell(test,1,2)` | `test·1·2` |
| `compile:GenCell(expr1,expr2,...,exprN,"sep")` | Concatenate with custom separator | `GenCell(test,1,2,"_")` | `test_1_2` |
| `compile:GenCell(expr1,expr2,...,exprN,"")` | Concatenate without separator | `GenCell(test,1,2,"")` | `test12` |

#### Macro Expansion

##### Single Cell Macro

Expand to a single cell's value:

| Syntax | Description | Example |
|:-------|:------------|:--------|
| `compile:Macro(expr,columnNum,rowNum)` | Get value from cell at (columnNum, rowNum) | `Macro(data,2,5)` |
| `compile:Macro(expr,columnNum,rowNum,MacroFormatter)` | Get formatted value | `Macro(data,2,5,!!codeKey)` |

**Parameters:**
- `expr`: Base expression
- `columnNum`: Column number (1-indexed)
- `rowNum`: Row number (1-indexed)
- `MacroFormatter`: Optional formatter (see below)

##### Multiple Cells Macro

Expand to multiple cell values:

| Syntax | Description | Example |
|:-------|:------------|:--------|
| `compile:Macro(exprArr,columnNum,rowStartNum,rowEndNum)` | Get values from cell range | `Macro(data,1,1,5)` |
| `compile:Macro(exprArr,columnNum,rowStartNum,rowEndNum,MacroFormatter)` | Get formatted values | `Macro(data,1,1,5,!!number)` |

**Parameters:**
- `exprArr`: Base expression array
- `columnNum`: Column number (1-indexed)
- `rowStartNum`: Start row number (1-indexed)
- `rowEndNum`: End row number (1-indexed)
- `MacroFormatter`: Optional formatter

##### Index Macro

Generate iteration sequence starting from 1:

| Syntax | Description | Example Usage | Result |
|:-------|:------------|:--------------|:--------|
| `compile:Macro(index)` | Auto-increment index (1-based) | Row 1: `Macro(index)` | `1` |
| | | Row 2: `Macro(index)` | `2` |
| | | Row 3: `Macro(index)` | `3` |

#### Macro Formatters

Format macro output using special formatters starting with `!!`:

| Formatter | Description | Input | Output |
|:----------|:------------|:-------|:--------|
| `!!codeKey` | Convert special chars (`@-[]{}\/'.`) to `_`, remove extra `__`, trim leading/trailing `_`, convert to uppercase | `test..x` | `TEST_X` |
| | | `@data-value` | `DATA_VALUE` |
| | | `[item].name` | `ITEM_NAME` |
| `!!codeKeyAlias` | Same as `!!codeKey` but adds prefix (default `@`) | `test..x` | `@TEST_X` |
| | (with default prefix `@`) | `data.value` | `@DATA_VALUE` |
| `!!number` | Convert to decimal integer, supports `0x` hex prefix | `123` | `123` |
| | | `0xFF` | `255` |
| | | `abc` | `abc` (unchanged, NaN) |

**CodeKey Conversion Rules:**
- Special characters replaced: `@`, `-`, space, `[`, `]`, `{`, `}`, `\`, `/`, `'`, `.`
- Multiple consecutive `__` collapsed to single `_`
- Leading and trailing `_` removed
- Final result converted to uppercase

#### Macro Examples

**Example 1: Generate CodeKey with Row Cell**

```
Rule: G-AQ:117=compile:GenCell(#LT[compile:Macro]#err@F118[#codeKey],[compile:Macro]#index@0)
Result: errValue·1, errValue·2, errValue·3, ...
```

**Example 2: Format Cell Value with CodeKey**

```
Rule: D-7=compile:Macro(@#[@D.MY],5,7,!!codeKey)
If cell(5,7) = "project-alpha-2024"
Result: PROJECT_ALPHA_2024
```

**Example 3: Generate CodeKeyAlias**

```
Rule: cell F-10=compile:Macro(@#key,3,10,!!codeKeyAlias)
If cell(3,10) = "test..data"
Result: @TEST_DATA
```

**Example 4: Number Conversion**

```
Rule: row-5=compile:Macro(@#value,2,5,!!number)
If cell(2,5) = "42"
Result: 42

Rule: row-6=compile:Macro(@#hex,4,6,!!number)
If cell(4,6) = "0x1A"
Result: 26
```

**Example 5: Iteration with Index**

```
Row 1: Code-${compile:Macro(index)}  →  Code-1
Row 2: Code-${compile:Macro(index)}  →  Code-2
Row 3: Code-${compile:Macro(index)}  →  Code-3
```

#### Complete Rule Configuration Example

A complete example of a rule sheet (`export.metadata.config`) with all rule types:

```
# Alias Rules (define shortcuts for long expressions)
T=template
LLR=exportData.LRR
CTR=contract.contractCode

# RowCell Rules (assign values to cell ranges)
G-AQ:12=compile:GenCell(@#item,[compile:Macro]#index@0)
A-Z:5=compile:Macro(@#data,2,5,!!codeKey)

# MergeCell Rules (merge cells and apply calculations)
G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)
G-AQ:18-35=sub(@LT,[compile:Macro(exprArr,F,18,35,!!codeKey)],compile:Macro(index),0)

# Cell Rules (assign values to single cells)
D:7=@#[@D.MY]
A:1=${@CTR}
B:1=${@LLR.value}
```

**⚠️ Important Notes:**
- Each rule type (alias, rowCell, mergeCell, cell) can appear multiple times
- But **duplicate rules of the same type with the same configuration are invalid**
- Alias abbreviations must be unique within the sheet
- Row/column ranges must not overlap in conflicting ways

### Render Functions

Built-in and custom functions for data processing:

| Function | Syntax | Example |
|:---------|:-------|:--------|
| **sum** | `fn:sum(...values)` | `fn:sum(10, 20, 30)` => `60` |
| **sub** | `fn:sub(a, b)` | `fn:sub(100, 30)` => `70` |
| **Custom Function** | `fn:customName(...args)` | User-defined logic |

## Quick Start

### Basic Usage

```typescript
import { ZipXlsxTemplateApp } from '@vdhewei/xlsx-template-lib';
import * as fs from 'node:fs/promises';

// Load zip template from buffer,zip file has [a.xlsx,b.xlsx...]
const templateBuffer = await fs.readFile('template.zip');
const app = new ZipXlsxTemplateApp(templateBuffer);

// Render with data
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

// Generate output
const output = await app.generate();
await fs.writeFile('output.xlsx', output);
```

### Using XlsxRender

```typescript
import { XlsxRender } from '@vdhewei/xlsx-template-lib';

const templateBuffer = await fs.readFile('template.zip');
const xlsx = await XlsxRender.create(templateBuffer);

// Render a specific sheet
await xlsx.render({ 
  contract: { contractCode: 'CTR-001' } 
}, 'Sheet1');

// Generate output
const buffer = await xlsx.generate();
```

### Compile and Render with Rules

```typescript
import { 
  ZipXlsxTemplateApp, 
  compileAll,
  AddCommand 
} from '@vdhewei/xlsx-template-lib';

// Add custom render function
AddCommand('sum', (obj, args) => {
  return args.groups.reduce((acc, val) => acc + Number(val), 0);
});

// Process with compile options
const compileOpts = {
  sheetName: 'export.metadata.config',  // Rule configuration sheet
  remove: true                           // Remove rule sheet after compile
};

const zipBuffer = await fs.readFile('template.zip');
const result = await ZipXlsxTemplateApp.compileTo(zipBuffer, {
  checker: async (buf, opts, values, fileName) => {
    // Custom validation logic
    return buf;
  },
  options: compileOpts
}, renderData);
```

### CLI Tool

The CLI tool `xlsx-cli` provides command-line interface for quick Excel template processing.

#### Installation

```bash
npm install -g @vdhewei/xlsx-template-lib
```

Or use directly from `npx`:

```bash
npx @vdhewei/xlsx-template-lib <command> [options]
```

Or git clone source code use bun compile to local Native CLI
```bash
git clone https://github.com/VDHewei/xlsx-template-lib.git
cd xlsx-template-lib
pnpm i
npm install -g bun
pnpm run complie-cli # default output bin/xlsx-cli 
# or use -o compile to user local dir
pnpm run compile-cli -o your-path/
```

#### Commands

##### 1. Compile Command

Compile Excel files with rule configurations.

```bash
xlsx-cli compile <xlsx-file> [options]
```

**Arguments:**
- `<xlsx-file>` - Path to the Excel file

**Options:**
- `-s, --save <string>` - Save compiled file to specified directory (default: current directory)
- `-n, --sheet-name <string>` - Sheet name to compile (default: first sheet)
- `-r, --remove` - Remove configure rules sheet after compilation (default: false)

**Examples:**

```bash
# Basic compile with default settings
xlsx-cli compile template.zip

# Compile and save to specific location
xlsx-cli compile template.zip -s ./output/

# Compile specific sheet
xlsx-cli compile template.zip -n Sheet1

# Compile and remove config sheet
xlsx-cli compile template.zip -r

# Full example
xlsx-cli compile template.zip -s ./output/ -n Sheet1 -r
```

**Output:**
- Compiled Excel file saved as `<filename>_<timestamp>.xlsx`
- Success messages displayed in green
- Errors displayed in red with process exit code 1

##### 2. Render Command

Render Excel templates with data substitution.

```bash
xlsx-cli render <xlsx-file> [options]
```

**Arguments:**
- `<xlsx-file>` - Path to the Excel template file

**Options:**
- `-c, --compile` - Auto-compile rules before rendering (default: false)
- `-n, --sheet-name <string>` - Sheet name to render (default: first sheet)
- `-s, --save <string>` - Save rendered file to specified directory (default: current directory)
- `-d, --data <string>` - Render data source (JSON string, file path, or URL)
- `--header <string>` - HTTP headers for remote data fetch (can be specified multiple times, format: `Key:Value`)
- `--body <string>` - HTTP request body for POST requests

**Examples:**

```bash
# Basic render with empty data
xlsx-cli render template.zip

# Render with JSON data string
xlsx-cli render template.zip -d '{"name":"John","age":30}'

# Render with JSON file
xlsx-cli render template.zip -d ./data.json

# Render with remote JSON URL
xlsx-cli render template.zip -d 'https://api.example.com/data.json'

# Render with auto-compile
xlsx-cli render template.zip -c -d './data.json'

# Render specific sheet
xlsx-cli render template.zip -n Sheet1 -d './data.json'

# Render with custom HTTP headers
xlsx-cli render template.zip -d 'https://api.example.com/data.json' --header 'Authorization:Bearer token123' --header 'Content-Type:application/json'

# Render with POST request body
xlsx-cli render template.zip -d 'https://api.example.com/api/query' --body '{"query":"SELECT * FROM users"}' --header 'Content-Type:application/json'

# Render with POST method via header
xlsx-cli render template.zip -d 'https://api.example.com/api/create' --body '{"name":"Test"}' --header 'Content-Type:application/json' --header 'method:POST'

# Full example
xlsx-cli render template.zip -c -n Sheet1 -s ./output/ -d './data.json'
```

**Data Sources:**
- **JSON String**: Direct JSON string enclosed in single quotes
- **Local File**: Path to `.json` file (relative or absolute)
- **Remote URL**: HTTP/HTTPS URL returning JSON

**HTTP Request Options (for Remote URL):**
- **Headers**: Use `--header` to add custom HTTP headers (can be specified multiple times)
  - Format: `--header 'Key:Value'`
  - Example: `--header 'Authorization:Bearer token123' --header 'Content-Type:application/json'`
  - Special header: `method:POST` can set the HTTP method to POST
- **Body**: Use `--body` to send request body (typically for POST requests)
  - Format: `--body '{"key":"value"}'`
  - Automatically uses POST method when body is provided
- **Default Behavior**: GET request with no headers

**HTTP Request Examples:**

```bash
# GET request with custom headers
xlsx-cli render template.zip \
  -d 'https://api.example.com/data.json' \
  --header 'Authorization:Bearer your-token' \
  --header 'X-API-Key:api-key-123'

# POST request with JSON body
xlsx-cli render template.zip \
  -d 'https://api.example.com/api/query' \
  --body '{"query":"SELECT * FROM users LIMIT 10"}' \
  --header 'Content-Type:application/json'

# POST request with method specified in header
xlsx-cli render template.zip \
  -d 'https://api.example.com/api/create' \
  --body '{"name":"New Record","value":100}' \
  --header 'Content-Type:application/json' \
  --header 'method:POST'

# Complex example with authentication and query body
xlsx-cli render template.zip \
  -d 'https://api.example.com/v1/export' \
  --header 'Authorization:Bearer eyJhbGc...' \
  --header 'Content-Type:application/json' \
  --body '{"format":"xlsx","filter":{"status":"active"}}' \
  -c -n Sheet1 -s ./output/
```

**HTTP Request Details:**

1. **Method Determination**:
   - Default: `GET`
   - With `--body`: Automatically becomes `POST`
   - With `method:POST` header: Explicitly set to `POST`
   - With `method:GET` header: Explicitly set to `GET`

2. **Header Format**:
   - Headers are parsed as `Key:Value` pairs
   - Multiple `--header` options can be used
   - Example: `--header 'Accept:application/json' --header 'User-Agent:MyApp/1.0'`

3. **Error Handling**:
   - Non-200 status codes return `undefined` and display error message
   - Network errors are caught and displayed in red
   - Missing `node-fetch` (Node.js < 18) displays error message

4. **Supported Data Formats**:
   - JSON objects: `{"key":"value"}`
   - JSON arrays: `[{"id":1},{"id":2}]`
   - Nested structures: `{"user":{"name":"John","age":30}}`

**Output:**
- Rendered Excel file saved as `<filename>_<timestamp>.xlsx`
- Validation checks for sheet existence
- Success/error messages with appropriate colors

##### 3. Rules Command

Add rule configurations to Excel files.

```bash
xlsx-cli rules <xlsx-file> [options]
```

**Arguments:**
- `<xlsx-file>` - Path to the Excel file

**Options:**

**Mode 1: Command Line Rules**
- `-t, --type <string>` - Rule type: `cell`, `alias`, `rowCell`, `mergeCell` (required when using -r)
- `-r, --rule <string>` - Rule expression string (can be specified multiple times)

**Mode 2: File Rules**
- `-f, --file <string>` - Read rules from file (format: `<type> ruleExpr` per line)
  - Lines starting with `#` are treated as comments
  - Empty lines are skipped
  - Rule types: `cell`, `alias`, `rowCell`, `mergeCell`

**Common Options:**
- `-s, --save <string>` - Save compiled file to specified directory (default: current directory)

**Examples:**

**Single Rule (Command Line):**
```bash
# Add alias rule
xlsx-cli rules template.zip -t alias -r 'T=template'

# Add cell rule
xlsx-cli rules template.zip -t cell -r 'D:7=${@LLR.value}'

# Add rowCell rule
xlsx-cli rules template.zip -t rowCell -r 'G-AQ:12=compile:GenCell(@#item,[compile:Macro]#index@0)'

# Add mergeCell rule
xlsx-cli rules template.zip -t mergeCell -r 'G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)'
```

**Multiple Rules (Command Line):**
```bash
# Add multiple rules with same type
xlsx-cli rules template.zip -t cell -r 'D:7=${@LLR.value}' -r 'A:1=${@T}' -r 'B:1=${@LLR.value}'
```

**Rules from File:**
```bash
# Read rules from file
xlsx-cli rules template.zip -f rules.txt

# Create rules.txt file:
# This is a comment
alias T=template
alias LLR=exportData.LRR
cell D:7=${@T}
cell A:1=${@LLR.value}
rowCell G-AQ:12=compile:GenCell(@#item,[compile:Macro]#index@0)
mergeCell G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)
```

**Save to Specific Directory:**
```bash
xlsx-cli rules template.zip -f rules.txt -s ./output/
xlsx-cli rules template.zip -t cell -r 'D:7=${@LLR.value}' -s ./output/
```

**File Format (-f mode):**
```bash
# Format: <type> ruleExpr
# Comments start with #
# Valid types: cell, alias, rowCell, mergeCell

cell D:7=${@LLR.value}
alias T=template
rowCell G-AQ:12=compile:GenCell(@#item,[compile:Macro]#index@0)
mergeCell G-AQ:13-17=sum(@LT,[compile:Macro(exprArr,F,13,17,!!codeKey)],compile:Macro(index),0)
```

**Behavior:**
- Creates `export_metadata.config` sheet if not exists
- Adds rule with proper styling: bold + center alignment for type, center alignment for expression
- Auto-adjusts column widths based on content
- Each rule type (cell, alias, rowCell, mergeCell) supports up to 4 rules per row
- If more than 4 rules are added for same type, automatically creates a new row
- Outputs new file with timestamp

#### Common Features

**Environment Variables:**
- CLI automatically loads `.env` file from current directory if present

**File Path Resolution:**
- Supports absolute and relative paths
- Resolves paths relative to current working directory
- Validates file existence before processing

**Error Handling:**
- All errors displayed in red using chalk
- Non-zero exit code on errors
- Detailed error messages for debugging

**Cross-Platform Support:**
- Works on Windows, Linux, and macOS
- Uses platform-independent path handling

**Output Filename Format:**
- Default: `<input-filename>_<timestamp>.xlsx`
- Timestamp in milliseconds since epoch
- Preserves original file name

**Verbose Logging:**
- Gray informational messages for process steps
- Green success messages
- Red error messages
- Yellow warnings

## Advanced Features

### Custom Commands (Render Functions)

```typescript
import { AddCommand, generateCommandsXlsxTemplate } from '@vdhewei/xlsx-template-lib';

// Add custom command
AddCommand('multiply', (obj, args) => {
  const values = args.groups.map(g => valueDotGet(obj, g));
  return values.reduce((a, b) => a * b, 1);
});
const data = await fs.readFile('simple.xlsx');
// Generate template with custom commands
const buffer = await generateCommandsXlsxTemplate(data, options);
```

### Batch Processing

```typescript
// Process ZIP file containing multiple XLSX files
const zipBuffer = await fs.readFile('templates.zip');
const app = new ZipXlsxTemplateApp(zipBuffer);

const compileOpts = {
  sheetName: 'export.metadata.config',
  remove: true
};

const renderOpts = {
  // Rendering options
};

await app.substituteAll(renderData, compileOpts, renderOpts);
const output = await app.generate();
```

## API Reference

### ZipXlsxTemplateApp

Main class for processing Excel files in ZIP archives.

| Method | Description |
|:-------|:------------|
| `constructor(data?: Buffer)` | Initialize with ZIP buffer |
| `loadZipBuffer(data: Buffer)` | Load ZIP buffer |
| `parse(data: Buffer)` | Parse ZIP and extract XLSX entries |
| `getEntries()` | Get all XLSX file entries |
| `substituteAll(renderData, compileOpts?, renderOpts?)` | Substitute all placeholders |
| `generate(options?)` | Generate output buffer |
| `static compileAll(files, renderData?, compileOpts?)` | Compile multiple files |
| `static compileTo(data, opts, values?)` | Compile XLSX in ZIP with custom checker |

### XlsxRender

Main class for rendering single Excel files.

| Method | Description |
|:-------|:------------|
| `static create(data: Buffer, option?)` | Create from buffer |
| `render(values: Object, sheetName: string)` | Render specific sheet |
| `getSheets()` | Get all sheet information |
| `generate(options?)` | Generate output buffer |

### Helper Functions

| Function | Description |
|:---------|:------------|
| `ExprResolver` | Expression resolver for complex expressions |
| `compileRuleSheetName` | Default rule sheet name |
| `generateXlsxTemplate` | Generate XLSX template |
| `generateCommandsXlsxTemplate` | Generate with custom commands |
| `AddCommand(name, fn)` | Add custom render function |

## Complete Example

### Template Structure

```
template.zip
├── Sheet1 (Data Sheet with placeholders)
│   ├── A1: ${contract.contractCode}
│   ├── B1: ${contract.contractTitle}
│   └── C1: ${exportData.LRR.mothOrYear}
└── export.metadata.config (Rule Configuration Sheet)
    ├── Row 1: alias @#key => use aliasKey: @# => @#
    ├── Row 2: mergeCell G-AQ(1-17)=sum(...)
    └── Row 3: cell D-7=@#[@D.MY]
```

### Compile & Render Flow

| Step | Input                 | Output | Description |
|:-----|:----------------------|:-------|:------------|
| 1. Load | `template.zip` Buffer | `ZipXlsxTemplateApp` | Load template file |
| 2. Compile | Rule Config Sheet     | Compiled Rules | Parse mergeCell/cell/rowCell rules |
| 3. Substitute | Data Object           | Rendered Sheets | Replace `${...}` placeholders |
| 4. Generate | -                     | `output.xlsx` Buffer | Final output file |

```typescript
import { ZipXlsxTemplateApp, AddCommand } from '@vdhewei/xlsx-template-lib';
import * as fs from 'node:fs/promises';

// Define custom function
AddCommand('calculateTotal', (obj, args) => {
  const base = Number(args.root);
  const multiplier = args.groups.length > 0 ? Number(args.groups[0]) : 1;
  return base * multiplier;
});

// Main processing
async function processTemplate() {
  const templateBuffer = await fs.readFile('template.zip');
  
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
    remove: true  // Remove rule sheet in output
  };
  
  await app.substituteAll(renderData, compileOpts);
  
  const output = await app.generate();
  await fs.writeFile('output.xlsx', output);
}

processTemplate();
```

## Note

- The `test_data` directory contains internal test templates and is for private use only
- Rule configuration sheet supports: `alias`, `mergeCell`, `cell`, `rowCell` rule types
- Custom functions can be registered via `AddCommand(name, handler)`

## License

MIT

## Author

VDHewei

## Repository

https://github.com/VDHewei/xlsx-template-lib

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

This project was inspired by the excellent open-source project [xlsx-template](https://github.com/optilude/xlsx-template) by optilude.

**xlsx-template** provides a robust foundation for Excel template-based report generation with dynamic data substitution. Many concepts and design patterns from xlsx-template have influenced this library, including:

- Template-based Excel file generation
- Placeholder substitution syntax
- Array and table expansion
- Image insertion and positioning
- Cell formatting preservation

We extend our gratitude to the xlsx-template team and contributors for their valuable work in the open-source community.

**Original xlsx-template repository:** https://github.com/optilude/xlsx-template
