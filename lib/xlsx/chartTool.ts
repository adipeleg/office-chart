import { IData } from "./models/data.model";
import { XmlTool } from "./xmlTool";

export class ChartTool {
    private parts: any;

    constructor(private xmlTool: XmlTool) { }

    public addChart = async (sheet: any, sheetName: string, opt: IData, id: string) => {
        let readChart = await this.xmlTool.readXml(`xl/charts/chart${this.getChartNum(opt)}.xml`);
        readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:t'] = opt.title.name;
        if (opt.title.color) {
            readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:rPr']['a:solidFill']['a:srgbClr'].$.val = opt.title.color
        }
        if (opt.title.size) {
            readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:rPr'].$.sz = opt.title.size;
        }

        const chartType = `c:${opt.type}Chart`;

        let rowNum = 1;
        let lastCol = 'A';
        let firstCol = 'A';
        try {
            const splitRange: string[] = opt.range.split(':');
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
            const data = JSON.parse(JSON.stringify(ser));
            let d = data[0] || data;

            d['c:idx'] = i;
            d['c:order'] = i;

            if (opt.type !== 'scatter') {
                d['c:cat']['c:strRef']['c:f'] = sheetName + `!$${firstCol}$1:$${lastCol}$1`;
                d['c:val']['c:numRef']['c:f'] = sheetName + `!$${firstCol}$${(i + 1)}:$${lastCol}$${(i + 1)}`;
            } else {
                d['c:xVal']['c:numRef']['c:f'] = sheetName + `!$${firstCol}$1:$${lastCol}$1`;
                d['c:yVal']['c:numRef']['c:f'] = sheetName + `!$${firstCol}$${(i + 1)}:$${lastCol}$${(i + 1)}`;
            }

            if (opt.rgbColors && opt.rgbColors[i - 1] && (opt.type === 'line' || opt.type === 'bar')) {
                d['c:spPr']['a:ln']['a:solidFill']['a:srgbClr'].$.val = opt.rgbColors[i - 1];
                if (d['c:spPr']['a:solidFill']) {
                    d['c:spPr']['a:solidFill'] = {
                        'a:srgbClr': {
                            $: {
                                val: opt.rgbColors[i - 1]
                            }
                        }
                    }
                }
                if (d['c:marker']) {
                    d['c:marker']['c:size'].$.val = opt?.marker?.size || '4';
                    d['c:marker']['c:symbol'].$.val = opt?.marker?.shape || 'circle';
                    d['c:marker']['c:spPr']['a:solidFill']['a:srgbClr'].$.val = opt.rgbColors[i - 1];
                    d['c:marker']['c:spPr']['a:ln']['a:solidFill']['a:srgbClr'].$.val = opt.rgbColors[i - 1];
                }
            }

            if (opt.labels) {
                d['c:tx'] = {
                    'c:strRef': {
                        'c:f': sheetName + '!$A$' + i
                    }
                }
            } else {
                delete d['c:tx'];
            }

            readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'].push(d)
        }

        await this.addDrawingRel(sheet, sheetName, id);
        await this.addChartToDraw(id);
        await this.addChartToSheet(sheet, id);
        await this.addChartToSheetRel(id);
        await this.addChartToParts(id);

        return this.xmlTool.write(`xl/charts/chart${id}.xml`, readChart);
    }

    private getChartNum = (opt: IData) => {
        switch (opt.type) {
            case 'bar':
                return 1;
            case 'line':
                return 2;
            case 'pie':
                return 3;
            case 'scatter':
                return 4;
        }
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