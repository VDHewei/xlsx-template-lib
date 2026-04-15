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

### Compile Rules Configuration

Configure rendering rules in a rule sheet (e.g., `export.metadata.config`) with the following syntax:

| Rule Type | Syntax | Description |
|:----------|:-------|:------------|
| **alias** | `alias: @#key => use aliasKey: @# => @#` | Alias for field mapping |
| **mergeCell** | `G-AQ(1-17)=sum(#LT[compile Macro]#err@F113.17[#codeKey],[compile Macro]#index@0)` | Merge cell calculation |
| **mergeCell (continued)** | `G-AQ(18-35)=sum(#LT[compile Macro]#err@F118.20[#codeKey],[compile Macro]#index@0)` | Merge with continuation |
| **cell** | `D-7=@#[@D.MY]` | Cell value assignment |
| **rowCell** | `G-AQ117=compile GenCell(#LT[compile Macro]#err@F118[#codeKey],[compile Macro]#index@0)` | Row cell generation |

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

// Load template from buffer
const templateBuffer = await fs.readFile('template.xlsx');
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

const templateBuffer = await fs.readFile('template.xlsx');
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

const zipBuffer = await fs.readFile('template.xlsx');
const result = await ZipXlsxTemplateApp.compileTo(zipBuffer, {
  checker: async (buf, opts, values, fileName) => {
    // Custom validation logic
    return buf;
  },
  options: compileOpts
}, renderData);
```

### CLI Tool

```bash
# Compile template
xlsx-cli compile template.xlsx -s compiled.xlsx -r

# Render template
xlsx-cli render template.xlsx -s output.xlsx -c
```

## Advanced Features

### Custom Commands (Render Functions)

```typescript
import { AddCommand, generateCommandsXlsxTemplate } from '@vdhewei/xlsx-template-lib';

// Add custom command
AddCommand('multiply', (obj, args) => {
  const values = args.groups.map(g => valueDotGet(obj, g));
  return values.reduce((a, b) => a * b, 1);
});

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
template.xlsx
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

| Step | Input | Output | Description |
|:-----|:------|:-------|:------------|
| 1. Load | `template.xlsx` Buffer | `ZipXlsxTemplateApp` | Load template file |
| 2. Compile | Rule Config Sheet | Compiled Rules | Parse mergeCell/cell/rowCell rules |
| 3. Substitute | Data Object | Rendered Sheets | Replace `${...}` placeholders |
| 4. Generate | - | `output.xlsx` Buffer | Final output file |

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
