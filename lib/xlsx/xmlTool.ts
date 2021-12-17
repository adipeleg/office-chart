import xml2js from 'xml2js';
import JSZip from "jszip";
import fs from 'fs';

export class XmlTool {
    private zip: JSZip;
    private parser: xml2js.Parser;
    private builder: xml2js.Builder;
    private parts: any;

    constructor() {
        this.zip = new JSZip();
        this.parser = new xml2js.Parser({ explicitArray: false });
        this.builder = new xml2js.Builder();
    }

    public getZip = (): JSZip => {
        return this.zip;
    }

    public readXlsx = async () => {
        let path = __dirname + "/templates/template.xlsx";

        await new Promise((resolve, reject) => fs.readFile(path, async (err, data) => {
            if (err) {
                console.error(`Template ${path} not read: ${err}`);
                reject(err);
                return;
            };
            return await this.zip.loadAsync(data).then(d => {
                resolve(d);
            })
        }));
    }

    public readXml = async (file: string) => {
        return this.zip.file(file).async('string').then(data => {
            return this.parser.parseStringPromise(data);
        })
    }

    public write = async (filename: string, data: any) => {
        var xml = this.builder.buildObject(data);
        this.zip.file(filename, Buffer.from(xml), { base64: true });
    }

    public writeStr = async (filename: string, data: string) => {
        // var xml = this.builder.buildObject(data);
        this.zip.file(filename, Buffer.from(data), { base64: true });
    }

    public addSheetToWb = async (name: string) => {
        const wb = await this.readXml('xl/workbook.xml');
        let count: string;

        if (!Array.isArray(wb.workbook.sheets.sheet)) {
            count = '5';
            wb.workbook.sheets = {
                sheet: [
                    { '$': wb.workbook.sheets.sheet.$ },
                    { '$': { state: 'visible', name: 'Sheet2', sheetId: '2', 'r:id': 'rId' + count } }
                ]
            }

        } else {
            count = wb.workbook.sheets.sheet.length + 4;
            wb.workbook.sheets.sheet.push({
                '$': { state: 'visible', name: name, sheetId: count, 'r:id': 'rId' + count }
            });
        }
        await this.addSheetToParts(count);
        this.write('xl/workbook.xml', wb);
        return count;
    }

    public createSheet = async (id: string) => {
        const resSheet = await this.readXml('xl/worksheets/sheet1.xml');

        delete resSheet.worksheet.drawing

        const WbRel = await this.readXml('xl/_rels/workbook.xml.rels');
        WbRel.Relationships.Relationship.push({
            '$':
            {
                Id: 'rId' + id,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: `worksheets/sheet${id}.xml`
            }
        })


        this.write(`xl/worksheets/sheet${id}.xml`, resSheet);
        this.write('xl/_rels/workbook.xml.rels', WbRel);
        return resSheet;
    }

    public generateBuffer = async (): Promise<Buffer> => {
        return this.zip.generateAsync({ type: 'nodebuffer' });
    }

    public generateFile = async (name: string) => {
        const buf = await this.generateBuffer();
        fs.writeFileSync(name + '.xlsx', buf);
        return buf;
    }

    public removeTemplateSheets = async () => {
        const wb = await this.readXml('xl/workbook.xml');

        wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
            return 'SheetTemplate' !== it.$.name.toString() &&
                'barTemplate' !== it.$.name.toString() &&
                'lineTemplate' !== it.$.name.toString() &&
                'pieTemplate' !== it.$.name.toString() &&
                'scatterTemplate' !== it.$.name.toString();
        })

        return this.write('xl/workbook.xml', wb);
    }

    public writeTable = async (sheet: any, data: any[][], id: string) => {
        const sheetWithTable = await this.readXml('xl/worksheets/sheet2.xml');
        const rowTemplate = sheetWithTable.worksheet.sheetData.row[0];
        const header = data.shift();

        const rows: any[] = [];
        rows.push(this.addRow(header, JSON.parse(JSON.stringify(rowTemplate)), 1));
        data.forEach((data, idx) => {
            rows.push(this.addRow(data, JSON.parse(JSON.stringify(rowTemplate)), idx + 2));
        })
        sheet.worksheet.sheetData = { row: rows };

        return this.write(`xl/worksheets/sheet${id}.xml`, sheet);
    }

    private addRow(rowData: any[], rowTemplate: any, index: number) {
        rowTemplate.$.r = index;
        const cols: any[] = [];
        rowData.forEach((data, col) => {
            const type = typeof data === 'string' ? 's' : '';
            const c = { '$': { r: this.getColName(col) + (index), s: '1' }, v: data }
            if (type === 's') {
                c.$['t'] = 'str';
            }
            cols.push(c)
        })

        rowTemplate.c = cols;
        return rowTemplate;
    }

    private addSheetToParts = async (id: string) => {
        const parts = await this.readXml('[Content_Types].xml');

        parts['Types']['Override'].push({
            '$':
            {
                ContentType:
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                PartName: `/xl/worksheets/sheet${id}.xml`
            }
        })

        return this.write(`[Content_Types].xml`, parts);
    }

    private getColName = (n: number) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc[n] || abc[n % 26];
    }


    private ColToNum = (char: string) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc.indexOf(char);
    }
}