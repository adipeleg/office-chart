import xml2js from 'xml2js';
import JSZip from "jszip";
import fs from 'fs';

export class XmlTool {
    private zip: JSZip;
    // private zipData: JSZip;
    private parser: xml2js.Parser;
    private builder: xml2js.Builder;

    constructor() {
        this.zip = new JSZip();
        this.parser = new xml2js.Parser({ explicitArray: false });
        this.builder = new xml2js.Builder();
        // this.readXlsx();
    }
    public readXlsx = async () => {
        let path = __dirname + "/templates/empty.xlsx";
        // let path = __dirname + "/templates/spreadsheet.xlsx";
        // let path = __dirname + "/templates/test.xlsx";
        // let path = 'xl/worksheets/_rels/sheet1.xml';

        await new Promise((resolve, reject) => fs.readFile(path, async (err, data) => {
            if (err) {
                console.error(`Template ${path} not read: ${err}`);
                reject(err);
                return;
            };
            return await this.zip.loadAsync(data).then(d => {
                console.log(d)
                resolve(d);
            })
        }));
    }

    public readXml = async (file: string) => {
        return this.zip.file(file).async('string').then(data => {
            // console.log('data', data)
            return this.parser.parseStringPromise(data);
        })
    }

    public readXmlStr = async (file: string): Promise<string> => {
        return this.zip.file(file).async('string').then(data => {
            console.log('data', data)
            return data;
            // return this.parser.parseStringPromise(data);
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
        console.log(wb.workbook.sheets.sheet, Array.isArray(wb.workbook.sheets.sheet))
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

        // console.log(wb.workbook.sheets.sheet)
        this.write('xl/workbook.xml', wb);
        return count;
    }

    public createSheet = async (name: string, id: string) => {
        // await this.readXlsx();

        const resSheet = await this.readXml('xl/worksheets/sheet1.xml');

        // console.log(resSheet.worksheet.drawing);
        delete resSheet.worksheet.drawing

        const WbRel = await this.readXml('xl/_rels/workbook.xml.rels');
        WbRel.Relationships.Relationship.push({
            '$':
            {
                Id: 'rId' + id,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: `worksheets/${name}.xml`
            }
        })


        this.write(`xl/worksheets/${name}.xml`, resSheet);
        this.write('xl/_rels/workbook.xml.rels', WbRel);
        // const buf = await this.generateBuffer();
        // fs.writeFileSync('test3.xlsx', buf);
    }

    public generateBuffer = async (): Promise<Buffer> => {
        return this.zip.generateAsync({ type: 'nodebuffer' });
    }

    public generate = async (): Promise<string> => {
        return this.zip.generateAsync({ type: 'string' });
    }

    public generateFile = async (name: string) => {
        const buf = await this.generateBuffer();
        fs.writeFileSync(name + '.xlsx', buf);
    }

    public writeTable = async () => {

    }

}