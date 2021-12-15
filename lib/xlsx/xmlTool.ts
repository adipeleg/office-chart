import xml2js from 'xml2js';
import JSZip, { forEach } from "jszip";
import fs from 'fs';

export class XmlTool {
    private zip: JSZip;
    private parser: xml2js.Parser;
    private builder: xml2js.Builder;

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
        this.addSheetToParts(count);
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

    public generate = async (): Promise<string> => {
        return this.zip.generateAsync({ type: 'string' });
    }

    public generateFile = async (name: string) => {
        const buf = await this.generateBuffer();
        fs.writeFileSync(name + '.xlsx', buf);
        return buf;
    }

    public removeTemplateSheets = async () => {
        const wb = await this.readXml('xl/workbook.xml');

        wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
            return 'SheetTemplate' !== it.$.name.toString() && 'ChartTemplate' !== it.$.name.toString();
        })

        return this.write('xl/workbook.xml', wb);
    }

    public writeTable = async (sheet: any, data: any[][], id: string) => {
        const header = data.shift();

        var rows: any = [{
            $: {
                r: 1,
                spans: "1:" + (header.length)
            },
            c: header.map((t, x) => {
                return {
                    $: {
                        r: this.getColName(x) + 1,
                        // t: "s"
                    },
                    v: t.toString()
                }
            })
        }];


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

        return this.write(`xl/worksheets/sheet${id}.xml`, sheet);
    }

    public addChart = async (sheet: any, sheetName: string, title: string, range: string, id: string, type: 'line' | 'bar') => {
        let readChart = await this.readXml('xl/charts/chart1.xml');
        readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:t'] = title;

        const chartType = type === 'line' ? 'c:lineChart' : 'c:barChart';
        if (type === 'line') {
            readChart['c:chartSpace']['c:chart']['c:plotArea']['c:lineChart'] = JSON.parse(JSON.stringify(readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']));
            delete readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart'];
        }

        let rowNum = 1;
        try {
            rowNum = this.ColToNum(range.split(':')[1][0])
        } catch {
            console.log('range is not right');
        }
        const ser = { ...readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'] };
        readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'] = [];

        for (let i = 1; i < rowNum + 1; i++) {
            let d = JSON.parse(JSON.stringify(ser));;
            d['c:idx'] = i;
            d['c:order'] = i;
            d['c:cat']['c:strRef']['c:f'] = sheetName + '!$A$1:$C$1';
            d['c:val']['c:numRef']['c:f'] = sheetName + '!$A$' + (i + 1) + ':$C$' + (i + 1) + '';
            readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'].push(d)
        }

        await this.addDrawingRel(sheet, sheetName, id);
        await this.addChartToDraw(id);
        await this.addChartToSheet(sheet, id);
        await this.addChartToSheetRel(id);
        await this.addChartToParts(id);

        return this.write(`xl/charts/chart${id}.xml`, readChart);
    }

    private addDrawingRel = async (sheet, sheetName: string, id: string) => {
        const drawRel = await this.readXml('xl/drawings/_rels/drawing2.xml.rels'); //add new chart rel
        drawRel.Relationships.Relationship =

        {
            '$': {
                Id: 'rId' + id,
                Type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                Target: `../charts/chart${id}.xml`
            }
        }

        await this.write(`xl/drawings/_rels/drawing${id}.xml.rels`, drawRel);
        return id;

    }

    private addChartToDraw = async (id) => {
        const draw = await this.readXml('xl/drawings/drawing2.xml'); // add new chart draw
        draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = 'rId' + id;
        return this.write(`xl/drawings/drawing${id}.xml`, draw);
    }

    private addChartToSheetRel = async (id: string) => {
        const draw = await this.readXml('xl/worksheets/_rels/sheet2.xml.rels'); // add new chart to sheet rel

        draw['Relationships']['Relationship'] =
        {
            '$':
            {
                Id: 'rId' + id,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                Target: `../drawings/drawing${id}.xml`
            }
        }

        return this.write(`xl/worksheets/_rels/sheet${id}.xml.rels`, draw);
    }

    private addChartToSheet = async (sheet, id: string) => {
        sheet['worksheet']['drawing'] = {
            $: {
                'r:id': "rId" + id
            }
        };
        return this.write(`xl/worksheets/sheet${id}.xml`, sheet);
    }

    private addChartToParts = async (id: string) => {
        const parts = await this.readXml('[Content_Types].xml');

        parts['Types']['Override'].push({
            '$':
            {
                ContentType:
                    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
                PartName: `/xl/charts/chart${id}.xml`
            }
        })
        parts['Types']['Override'].push({
            '$':
            {
                ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
                PartName: `/xl/drawings/drawing${id}.xml`
            }
        })

        return this.write(`[Content_Types].xml`, parts);

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

    public getColName = (n: number) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc[n] || abc[n % 26];
    }


    public ColToNum = (char: string) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc.indexOf(char);
    }


}