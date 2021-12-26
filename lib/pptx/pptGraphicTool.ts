import { ChartTool } from './../xlsx/chartTool';
import { XlsxGenerator } from './../xlsx/xlsxGenerator';
import { IData, IPPTChartData, IPPTChartDataVal } from './../xlsx/models/data.model';
import { XmlTool } from "../xmlTool";

export class PptGraphicTool {
    constructor(private xmlTool: XmlTool,
        private xlsxGenerator: XlsxGenerator,
        private chartTool: ChartTool) { }

    public writeTable = async (id: number, slide: any, data: any[][]) => {
        const slideWithTable = await this.xmlTool.readXml('ppt/slides/slide2.xml');

        const rowTemplate = slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tr'][1];
        const colTemplate = rowTemplate['a:tc'][1];


        const header = data.shift();

        const rows: any[] = [];
        rows.push(this.addRow(header, JSON.parse(JSON.stringify(rowTemplate)), colTemplate));

        data.forEach((row, idx) => {
            rows.push(this.addRow(row, JSON.parse(JSON.stringify(rowTemplate)), colTemplate));
        })

        slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tr'] = rows;

        slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tblGrid']['a:gridCol'] = [];

        for (let i = 0; i < header.length; i++) {
            slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tblGrid']['a:gridCol'].push({ '$': { w: '2381250' } })
        }

        slide['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'] = slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'];
        return this.xmlTool.write(`ppt/slides/slide${id}.xml`, slide);
    }

    private addRow(rowData: any[], rowTemplate: any, colTemplate: any) {
        const cols: any[] = [];
        rowData.forEach((data, col) => {
            colTemplate['a:txBody']['a:p']['a:r']['a:t'] = data
            cols.push(JSON.parse(JSON.stringify(colTemplate)));
        })

        rowTemplate['a:tc'] = JSON.parse(JSON.stringify(cols));
        return rowTemplate;
    }

    public addChart = async (slide, chartOpt: IPPTChartData, slideId: number) => {
        const data = JSON.parse(JSON.stringify(this.buildData(chartOpt.data)));
        chartOpt.data = JSON.parse(JSON.stringify(data));
        
        chartOpt.range = `A1:${this.getColName(data[0].length - 1)}${data.length}`;

        const chartId = await this.addContentTypeChart();
        await this.addChartRef(chartId);
        await this.createXlsxWithTableAndChart(data, chartId);

        await this.buildChart(chartOpt, chartId);

        const slideWithChart = await this.xmlTool.readXml('ppt/slides/slide3.xml');

        const graphicFrame = slideWithChart['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'];
        graphicFrame['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = "rId" + chartId;
        slide['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'] = graphicFrame;

        this.xmlTool.write(`ppt/slides/slide${slideId}.xml`, slide);
        await this.addSlideChartRel(slideId, chartId);

    }

    private buildData = (data: any[][] | IPPTChartDataVal[]): any[][] => {
        if (data && data[0] && data[0].hasOwnProperty('values')) {
            const dataAsTable = [];
            (data as IPPTChartDataVal[]).forEach((value: IPPTChartDataVal) => {
                dataAsTable[0] = ['labels', ...value.labels];
                dataAsTable.push([value.name, ...value.values])
            })

            return dataAsTable;
        }

        return data as any[][];
    }

    private buildChart = async (chartOpt: IPPTChartData, chartId: number) => {
        let readChart = chartOpt.type === 'line' ? await this.xmlTool.readXml(`ppt/charts/chart1.xml`) : await this.xmlTool.readXml(`ppt/charts/chart2.xml`);
        const chartData = this.chartTool.buildChart(readChart, chartOpt, 'chart' + chartId);
        chartData['c:chartSpace']['c:externalData'].$['r:id'] = "rId" + chartId;

        this.xmlTool.write(`ppt/charts/chart${chartId}.xml`, chartData)
    }

    private addContentTypeChart = async (): Promise<number> => {
        const pptParts = await this.xmlTool.readXml('[Content_Types].xml');
        const charts = pptParts['Types']['Override'].filter(part => {
            return part.$.ContentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
        })

        const chartsIds = charts.map(chart => {
            return parseInt(chart.$.PartName.split('/ppt/charts/chart')[1].split('.xml')[0], 10);
        })

        const id = Math.max(...chartsIds) + 1;

        pptParts['Types']['Override'].push(
            {
                '$':
                {
                    ContentType:
                        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
                    PartName: `/ppt/charts/chart${id}.xml`
                }
            }
        )

        await this.xmlTool.write('[Content_Types].xml', pptParts);

        return id;
    }

    private addSlideChartRel = async (slideId: number, chartId: number) => {
        const slideRels = await this.xmlTool.readXml('ppt/slides/_rels/slide3.xml.rels');

        slideRels['Relationships']['Relationship'][1] = {
            '$':
            {
                Id: 'rId' + chartId,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                Target: `../charts/chart${chartId}.xml`
            }
        }

        await this.xmlTool.write(`ppt/slides/_rels/slide${slideId}.xml.rels`, slideRels);

    }

    private addChartRef = async (id: number) => {
        const chartRel = await this.xmlTool.readXml('ppt/charts/_rels/chart1.xml.rels');

        chartRel['Relationships']['Relationship'] = {
            '$':
            {
                Id: 'rId' + id,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
                Target: `../embeddings/Microsoft_Excel_Sheet${id}.xlsx`
            }
        }
        return this.xmlTool.write(`ppt/charts/_rels/chart${id}.xml.rels`, chartRel);
    }

    private createXlsxWithTableAndChart = async (data: any[][], chartId: number) => {
        await this.xlsxGenerator.createWorkbook();
        const sheet = await this.xlsxGenerator.createWorksheet('chart' + chartId);
        await sheet.addTable(data);
        const bf = await this.xlsxGenerator.generate('', 'buffer');

        this.xmlTool.writeBuffer(`ppt/embeddings/Microsoft_Excel_Sheet${chartId}.xlsx`, bf);
    }

    private getColName = (n: number) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc[n] || abc[n % 26];
    }

}