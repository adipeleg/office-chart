import xml2js from 'xml2js';
import JSZip from "jszip";
import fs from 'fs';
import { Chart } from './chart';

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
        this.addSheetToParts(count);
        // console.log(wb.workbook.sheets.sheet)
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
    }

    public removeTemplateSheets = async () => {
        const wb = await this.readXml('xl/workbook.xml');

        wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
            return 'SheetTemplate' !== it.$.name.toString() && 'ChartTemplate' !== it.$.name.toString();
        })

        return this.write('xl/workbook.xml', wb);
    }

    public writeTable = async (sheet: any, data: any[][], id: string) => {
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

        return this.write(`xl/worksheets/sheet${id}.xml`, sheet);
    }

    // public addChart = async (sheet: any, sheetName: string, title: string, data: any[][], range: string, id: string) => {

    // }

    public addChart = async (sheet: any, sheetName: string, title: string, data: any[][], range: string, id: string) => {
        // let path = __dirname + "/templates/charts/chart1.xml";
        let readChart = await this.readXml('xl/charts/chart1.xml');
        // console.log(readChart['c:chartSpace']['c:chart']['c:plotArea']['c:valAx']);
        readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:t'] = title;
        readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']['c:ser']['c:cat']['c:strRef']['c:f'] = sheetName + '!$A$1:$A$2';
        readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']['c:ser']['c:val']['c:numRef']['c:f'] = sheetName + '!$B$1:$B$2';
        // console.log(readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']['c:ser'])
        // const c = new Chart();
        // var opts = {
        //     chart: "bar",
        //     titles: [
        //         "Price"
        //     ],
        //     fields: [
        //         "Apple",
        //         "Blackberry",
        //         "Strawberry",
        //         "Cowberry"
        //     ],
        //     data: {
        //         "Price": {
        //             "Apple": 10,
        //             "Blackberry": 5,
        //             "Strawberry": 15,
        //             "Cowberry": 20
        //         }
        //     },
        //     chartTitle: "Bar chart"
        // };
        // const d = c.getChart(sheetName, opts.titles, 1, 1, opts.fields, opts.data, opts.chart, readChart)
        await this.addDrawingRel(sheet, sheetName, id);
        await this.addChartToDraw(id);
        await this.addChartToSheet(sheet, id);
        await this.addChartToSheetRel(id);
        await this.addChartToParts(id);

        return this.write(`xl/charts/chart${id}.xml`, readChart);
    }

    public getChart = () => {
        return {
            'c:chartSpace': {
                '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                '@xmlns:r':
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'c:lang': { '@val': 'en-US' },
                'c:date1904': { '@val': '1' },
                'c:chart': {
                    'c:plotArea': {
                        'c:layout': {},
                        'c:barChart': {
                            'c:barDir': { '@val': 'col' },
                            'c:grouping': { '@val': 'clustered' },
                            // 'c:overlap': { '@val': options.overlap || '0' },
                            // 'c:gapWidth': { '@val': options.gapWidth || '150' },

                            '#text': [
                                { 'c:axId': { '@val': '64451712' } },
                                { 'c:axId': { '@val': '64453248' } }
                            ]
                        },
                        'c:catAx': {
                            'c:axId': { '@val': '64451712' },
                            'c:scaling': {
                                'c:orientation': {
                                    // '@val': options.catAxisReverseOrder ? 'maxMin' : 'minMax'
                                }
                            },
                            'c:axPos': { '@val': 'l' },
                            'c:tickLblPos': { '@val': 'nextTo' },
                            'c:crossAx': { '@val': '64453248' },
                            'c:crosses': { '@val': 'autoZero' },
                            'c:auto': { '@val': '1' },
                            'c:lblAlgn': { '@val': 'ctr' },
                            'c:lblOffset': { '@val': '100' }
                        },
                        'c:valAx': {
                            'c:axId': { '@val': '64453248' },
                            'c:scaling': {
                                'c:orientation': { '@val': 'minMax' }
                            },
                            'c:axPos': { '@val': 'b' },
                            //              "c:majorGridlines": {},
                            'c:numFmt': {
                                '@formatCode': 'General',
                                '@sourceLinked': '1'
                            },
                            'c:tickLblPos': { '@val': 'nextTo' },
                            'c:crossAx': { '@val': '64451712' },
                            'c:crosses': {
                                // '@val': options.valAxisCrossAtMaxCategory ? 'max' : 'autoZero'
                            },
                            'c:crossBetween': { '@val': 'between' }
                        }
                    },
                    'c:legend': {
                        'c:legendPos': { '@val': 'r' },
                        'c:layout': {}
                    },
                    'c:plotVisOnly': { '@val': '1' }
                },
                'c:txPr': {
                    'a:bodyPr': {},
                    'a:lstStyle': {},
                    'a:p': {
                        'a:pPr': {
                            'a:defRPr': { '@sz': '1800' }
                        },
                        'a:endParaRPr': { '@lang': 'en-US' }
                    }
                },
                // 'c:externalData': { '@r:id': 'rId1' }
            }
        }
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
        // console.log(draw['Relationships']['Relationship'])
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
        console.log(parts['Types']['Override'])
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
        console.log(parts['Types']['Override'])

        return this.write(`[Content_Types].xml`, parts);

    }

    private addSheetToParts = async (id: string) => {
        const parts = await this.readXml('[Content_Types].xml');
        console.log(parts['Types']['Override'])
        parts['Types']['Override'].push({
            '$':
            {
                ContentType:
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                PartName: `/xl/worksheets/sheet${id}.xml`
            }
        })

        console.log(parts['Types']['Override'])

        return this.write(`[Content_Types].xml`, parts);

    }

    public getColName = (n: number) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc[n] || abc[(n / 26 - 1) | 0] + abc[n % 26];
    }

}