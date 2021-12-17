import { XmlTool } from "./xmlTool";

export class ChartTool {
    private parts: any;

    constructor(private xmlTool: XmlTool) { }

    public addChart = async (sheet: any, sheetName: string, title: string, range: string, id: string, type: 'line' | 'bar') => {
        let readChart = await this.xmlTool.readXml('xl/charts/chart1.xml');
        readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:t'] = title;

        const chartType = type === 'line' ? 'c:lineChart' : 'c:barChart';
        if (type === 'line') {
            readChart['c:chartSpace']['c:chart']['c:plotArea']['c:lineChart'] = JSON.parse(JSON.stringify(readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']));
            delete readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart'];
            delete readChart['c:chartSpace']['c:chart']['c:plotArea']['c:lineChart']['c:barDir'];
        }

        let rowNum = 1;
        let lastCol = 'A';
        let firstCol = 'A';
        try {
            const splitRange: string[] = range.split(':');
            const first = splitRange[0]
            firstCol = first[0]
            const sec = splitRange[1];
            lastCol = sec[0];
            rowNum = parseInt(sec.substring(1));
        } catch {
            console.log('range is not right');
            throw Error('range is not right');
        }

        const ser = { ...readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'] };
        readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'] = [];

        for (let i = 1; i < rowNum; i++) {
            let d = JSON.parse(JSON.stringify(ser));;
            d['c:idx'] = i;
            d['c:order'] = i;
            d['c:cat']['c:strRef']['c:f'] = sheetName + `!$${firstCol}$1:$${lastCol}$1`;
            d['c:val']['c:numRef']['c:f'] = sheetName + `!$${firstCol}$${(i + 1)}:$${lastCol}$${(i + 1)}`;
            readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'].push(d)
        }

        await this.addDrawingRel(sheet, sheetName, id);
        await this.addChartToDraw(id);
        await this.addChartToSheet(sheet, id);
        await this.addChartToSheetRel(id);
        await this.addChartToParts(id);

        return this.xmlTool.write(`xl/charts/chart${id}.xml`, readChart);
    }

    private addDrawingRel = async (sheet, sheetName: string, id: string) => {
        const drawRel = await this.xmlTool.readXml('xl/drawings/_rels/drawing2.xml.rels'); //add new chart rel
        drawRel.Relationships.Relationship =

        {
            '$': {
                Id: 'rId' + id,
                Type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                Target: `../charts/chart${id}.xml`
            }
        }

        await this.xmlTool.write(`xl/drawings/_rels/drawing${id}.xml.rels`, drawRel);
        return id;

    }

    private addChartToDraw = async (id) => {
        const draw = await this.xmlTool.readXml('xl/drawings/drawing2.xml'); // add new chart draw
        draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = 'rId' + id;
        return this.xmlTool.write(`xl/drawings/drawing${id}.xml`, draw);
    }

    private addChartToSheetRel = async (id: string) => {
        const draw = await this.xmlTool.readXml('xl/worksheets/_rels/sheet2.xml.rels'); // add new chart to sheet rel

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

        return this.xmlTool.write(`xl/worksheets/_rels/sheet${id}.xml.rels`, draw);
    }

    private addChartToSheet = async (sheet, id: string) => {
        sheet['worksheet']['drawing'] = {
            $: {
                'r:id': "rId" + id
            }
        };
        return this.xmlTool.write(`xl/worksheets/sheet${id}.xml`, sheet);
    }

    private addChartToParts = async (id: string) => {
        const parts = this.parts || await this.xmlTool.readXml('[Content_Types].xml');

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


        this.parts = parts;
        return this.xmlTool.write(`[Content_Types].xml`, parts);
    }

}