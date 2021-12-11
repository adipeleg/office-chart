import xml2js from 'xml2js';
import JSZip from "jszip";
import fs from 'fs';

export class XmlTool {
    private zip: JSZip;
    // private zipData: JSZip;
    private parser: xml2js.Parser;
    private chartZip: JSZip;
    private builder: xml2js.Builder;
    
    constructor() {
        this.zip = new JSZip();
        this.chartZip = new JSZip();
        this.parser = new xml2js.Parser({ explicitArray: false });
        this.builder = new xml2js.Builder();
    }

    public readXlsx = async (fileName?: string) => {
        let path =  __dirname + "/templates/template.xlsx";

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
        const resSheet = await this.readXml('xl/worksheets/SheetTemplate.xml');

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
        return resSheet;
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

    public writeTable = async (sheet: any, name: string, data: any[][]) => {
        var rows: any = [{
            $: {
                r: 1,
                spans: "1:" + (data[0].length)
            },
            c: data[0].map((t, x) => {
                return {
                    $: {
                        r: this.getColName(x + 1) + 1,
                        // t: "s"
                    },
                    v: t.toString()
                }
            })
        }];

        const header = data.shift();

        data.forEach((f, y) => {
            var r: any = {
                $: {
                    r: y + 2,
                    spans: "1:" + (header.length)
                }
            };
            const c = [];
            f.forEach((t, x) => {
                c.push({
                    $: {
                        r: this.getColName(x + 1) + (y + 2),
                    },
                    v: t.toString()
                });
            });
            r.c = c;
            rows.push(r);
        });
        sheet.worksheet.sheetData = { row: rows };
        console.log(sheet)
        return this.write(`xl/worksheets/${name}.xml`, sheet);
    }

    public addChart = async (sheet: any, name: string, data: any[][], range: string) => {
        // let path = __dirname + "/templates/charts/chart1.xml";
        // const read = this.readXml(path);
        // const chartTemplate = await new Promise((resolve, reject) => fs.readFile(path, 'utf8', async (err, data) => {
        //     if (err) {
        //         console.error(`Template ${path} not read: ${err}`);
        //         reject(err);
        //         return;
        //     };
        //     console.log(data)
        //     return this.parser.parseStringPromise(data);
        //     // return await this.chartZip.loadAsync(data).then(d => {
        //     //     console.log(d)
        //     //     resolve(d);
        //     // })
        // }));

        // console.log(read)
        // console.log(JSON.parse(await this.parser.parseStringPromise(chartTemplate)));
        // const chartTemplate = this.readXml(__dirname + "/templates/charts/chart1.xml");
    }

    public getColName = (n) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        n--;
        if (n < 26) {
            return abc[n];
        } else {
            return abc[(n / 26 - 1) | 0] + abc[n % 26];
        }
    }

}