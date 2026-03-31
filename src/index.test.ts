import { describe, it, expect } from 'vitest'
import * as fs from  "node:fs/promises";
import {BufferType, generateXlsxTemplate} from './core'
import {generateCommandsXlsxTemplate} from './extends'

describe('generateXlsxTemplate', () => {
    it('should generate a template', async () => {
        const columns = [
            { header: 'Age', key: 'age' },
            { header: 'Name', key: 'name' },
        ]
       const xlsx = await  fs.readFile("./test_data/test.xlsx")
       const data = {columns,"name":"test"};
       const buffer = await generateXlsxTemplate(xlsx,data, {type:BufferType.NodeBuffer})
       await fs.writeFile(`./test_data/test_${new Date().valueOf()}.xlsx`,buffer)
       expect(buffer).toBeInstanceOf(Buffer)
    })

   //it('should generate a template with data', async () => {
   //    // ... 与上述类似，但包含数据 ...
   //   //expect(sheet!.rowCount).toBe(2)
   //   //expect(sheet!.getRow(2).getCell(1).value).toBe('John')
   //})
})


describe('generateCommandsXlsxTemplate', () => {
    it('should generate a template', async () => {
        const columns = [
            { header: 'Age', key: 'age' },
            { header: 'Name', key: 'name' },
        ]
        const xlsx = await  fs.readFile("./test_data/test.xlsx")
        const data = {columns,"name":"test"};
        const buffer = await generateCommandsXlsxTemplate(xlsx,data, {type:BufferType.NodeBuffer})
        await fs.writeFile(`./test_data/test_${new Date().valueOf()}.xlsx`,buffer)
        expect(buffer).toBeInstanceOf(Buffer)
    })
   // it('should generate a template with data', async () => {
   //     // ... 与上述类似，但包含数据 ...
   //     //expect(sheet!.rowCount).toBe(2)
   //     //expect(sheet!.getRow(2).getCell(1).value).toBe('John')
   // })
})