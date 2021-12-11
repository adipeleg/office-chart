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
        let path = __dirname + "/templates/template.xlsx";

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
        const resSheet = await this.readXml('xl/worksheets/sheet1.xml');

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

    public removeTemplateSheets = async () => {
        const wb = await this.readXml('xl/workbook.xml');

        wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
            return 'SheetTemplate' !== it.$.name.toString() && 'ChartTemplate' !== it.$.name.toString();
        })

        return this.write('xl/workbook.xml', wb);
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
                        r: this.getColName(x) + 1,
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
                        r: this.getColName(x) + (y + 2),
                    },
                    v: t.toString()
                });
            });
            r.c = c;
            rows.push(r);
        });
        sheet.worksheet.sheetData = { row: rows };

        return this.write(`xl/worksheets/${name}.xml`, sheet);
    }

    public addChart = async (sheet: any, sheetName: string, title: string, data: any[][], range: string) => {
        // let path = __dirname + "/templates/charts/chart1.xml";
        const readChart = await this.readXml('xl/charts/chart1.xml');
        // console.log(readChart['c:chartSpace']['c:chart']['c:plotArea']['c:valAx']);
        readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:t'] = title;
        readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']['c:ser']['c:cat']['c:strRef']['c:f'] = sheetName + '!$A$1:$A$3';
        readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']['c:ser']['c:val']['c:numRef']['c:f'] = sheetName + '!$B$1:$B$3';
        const id = await this.addDrawingRel(sheet, sheetName);
        await this.addChartToDraw(sheetName, id);
        await this.addChartToSheet(sheetName, sheet, id);
        await this.addChartToSheetRel(sheetName, id);
        //might not work because file name
        return this.write(`xl/charts/chart${sheetName}.xml`, readChart);
    }

    private addDrawingRel = async (sheet, sheetName: string) => {
        const drawRel = await this.readXml('xl/drawings/_rels/drawing2.xml.rels'); //add new chart rel
        // console.log(drawRel.Relationships.Relationship);

        let id = 2
        // if (!Array.isArray(drawRel.Relationships.Relationship)) {
        drawRel.Relationships.Relationship =
        // { '$': drawRel.Relationships.Relationship.$ },
        {
            '$': {
                Id: 'rId' + id,
                Type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                Target: `../charts/chart${sheetName}.xml`
            }
        }


        // } else {
        //     id = drawRel.Relationships.Relationship.length;
        //     drawRel.Relationships.Relationship.push({
        //         '$':
        //         {
        //             Id: 'rId' + id,
        //             Type:
        //                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        //             Target: `../charts/chart${sheetName}.xml`
        //         }
        //     })

        // }

        await this.write(`xl/drawings/_rels/drawing${sheetName}.xml.rels`, drawRel);
        return id;

        //add to sheet relationships as drawing
        //add to sheet as drawing;
    }

    private addChartToDraw = async (sheetName, id) => {
        const draw = await this.readXml('xl/drawings/drawing2.xml'); // add new chart draw
        // console.log(draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id']);
        draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = 'rId' + id;
        return this.write(`xl/drawings/chart${sheetName}.xml`, draw);
    }

    private addChartToSheetRel = async (sheetName, id) => {
        const draw = await this.readXml('xl/worksheets/_rels/sheet2.xml.rels'); // add new chart to sheet rel
        // console.log(draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id']);
        console.log(draw['Relationships']['Relationship'])
        draw['Relationships']['Relationship'] =
        {
            '$':
            {
                Id: 'rId' + id,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                Target: `../drawings/chart${sheetName}.xml`
            }
        }
        // draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = 'rId' + id;
        return this.write(`xl/worksheets/_rels/${sheetName}.xml.rels`, draw);
    }

    private addChartToSheet = async (sheetName, sheet, id) => {
        sheet['worksheet']['drawing'] = {
            $: {
                'r:id': "rId" + id
            }
        };
        return this.write(`xl/worksheets/${sheetName}.xml`, sheet);
    }

    public getColName = (n: number) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc[n] || abc[(n / 26 - 1) | 0] + abc[n % 26];
    }

}