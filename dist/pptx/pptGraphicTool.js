"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.PptGraphicTool = void 0;
class PptGraphicTool {
    constructor(xmlTool, xlsxGenerator, chartTool) {
        this.xmlTool = xmlTool;
        this.xlsxGenerator = xlsxGenerator;
        this.chartTool = chartTool;
        this.writeTable = (id, slide, data, opt) => __awaiter(this, void 0, void 0, function* () {
            const slideWithTable = yield this.xmlTool.readXml('ppt/slides/slide2.xml');
            const rowTemplate = slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tr'][1];
            const colTemplate = rowTemplate['a:tc'][1];
            if (opt === null || opt === void 0 ? void 0 : opt.rowHeight) {
                rowTemplate.$.h = opt === null || opt === void 0 ? void 0 : opt.rowHeight;
            }
            if (opt === null || opt === void 0 ? void 0 : opt.colWidth) {
                const gridColVals = slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tblGrid']['a:gridCol'];
                gridColVals.forEach(col => {
                    col.$.w = opt.colWidth;
                });
            }
            this.addLocationGraphicElements(slideWithTable, opt);
            const header = data.shift();
            const rows = [];
            rows.push(this.addRow(header, JSON.parse(JSON.stringify(rowTemplate)), colTemplate));
            data.forEach((row, idx) => {
                rows.push(this.addRow(row, JSON.parse(JSON.stringify(rowTemplate)), colTemplate));
            });
            slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tr'] = rows;
            slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tblGrid']['a:gridCol'] = [];
            for (let i = 0; i < header.length; i++) {
                slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tblGrid']['a:gridCol'].push({ '$': { w: '2381250' } });
            }
            slide['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'] = slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'];
            return this.xmlTool.write(`ppt/slides/slide${id}.xml`, slide);
        });
        this.addLocationGraphicElements = (slide, opt) => {
            const locationElement = slide['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['p:xfrm'];
            slide['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['p:xfrm'] = {
                'a:off': {
                    $: {
                        x: (opt === null || opt === void 0 ? void 0 : opt.x) || locationElement['a:off'].$.x,
                        y: (opt === null || opt === void 0 ? void 0 : opt.y) || locationElement['a:off'].$.y,
                    }
                },
                'a:ext': {
                    $: {
                        cx: (opt === null || opt === void 0 ? void 0 : opt.cx) || locationElement['a:ext'].$.cx,
                        cy: (opt === null || opt === void 0 ? void 0 : opt.cy) || locationElement['a:ext'].$.cy
                    }
                }
            };
        };
        this.addChart = (slide, chartOpt, slideId) => __awaiter(this, void 0, void 0, function* () {
            const data = JSON.parse(JSON.stringify(this.buildData(chartOpt.data)));
            chartOpt.data = JSON.parse(JSON.stringify(data));
            chartOpt.labels = data[0].hasOwnProperty('values') ? true : chartOpt.labels;
            chartOpt.range = `B1:${this.getColName(data[0].length - 1)}${data.length}`;
            const chartId = yield this.addContentTypeChart();
            yield this.addChartRef(chartId);
            yield this.createXlsxWithTableAndChart(data, chartId);
            yield this.buildChart(chartOpt, chartId);
            const slideWithChart = yield this.xmlTool.readXml('ppt/slides/slide3.xml');
            const graphicFrame = slideWithChart['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'];
            graphicFrame['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = "rId" + chartId;
            if (chartOpt === null || chartOpt === void 0 ? void 0 : chartOpt.location) {
                this.addLocationGraphicElements(slideWithChart, chartOpt.location);
            }
            slide['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'] = graphicFrame;
            this.xmlTool.write(`ppt/slides/slide${slideId}.xml`, slide);
            yield this.addSlideChartRel(slideId, chartId);
        });
        this.buildData = (data) => {
            if (data && data[0] && data[0].hasOwnProperty('values')) {
                const dataAsTable = [];
                data.forEach((value) => {
                    dataAsTable[0] = ['labels', ...value.labels];
                    dataAsTable.push([value.name, ...value.values]);
                });
                return dataAsTable;
            }
            return data;
        };
        this.buildChart = (chartOpt, chartId) => __awaiter(this, void 0, void 0, function* () {
            let readChart = yield this.xmlTool.readXml(`ppt/charts/chart${chartOpt.type === 'line' ? 1 : 2}.xml`);
            const chartData = this.chartTool.buildChart(readChart, chartOpt, 'chart' + chartId);
            chartData['c:chartSpace']['c:externalData'].$['r:id'] = "rId" + chartId;
            this.xmlTool.write(`ppt/charts/chart${chartId}.xml`, chartData);
        });
        this.addContentTypeChart = () => __awaiter(this, void 0, void 0, function* () {
            const pptParts = yield this.xmlTool.readXml('[Content_Types].xml');
            const charts = pptParts['Types']['Override'].filter(part => {
                return part.$.ContentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
            });
            const chartsIds = charts.map(chart => {
                return parseInt(chart.$.PartName.split('/ppt/charts/chart')[1].split('.xml')[0], 10);
            });
            const id = Math.max(...chartsIds) + 1;
            pptParts['Types']['Override'].push({
                '$': {
                    ContentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
                    PartName: `/ppt/charts/chart${id}.xml`
                }
            });
            yield this.xmlTool.write('[Content_Types].xml', pptParts);
            return id;
        });
        this.addSlideChartRel = (slideId, chartId) => __awaiter(this, void 0, void 0, function* () {
            const slideRels = yield this.xmlTool.readXml('ppt/slides/_rels/slide3.xml.rels');
            slideRels['Relationships']['Relationship'][1] = {
                '$': {
                    Id: 'rId' + chartId,
                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                    Target: `../charts/chart${chartId}.xml`
                }
            };
            yield this.xmlTool.write(`ppt/slides/_rels/slide${slideId}.xml.rels`, slideRels);
        });
        this.addChartRef = (id) => __awaiter(this, void 0, void 0, function* () {
            const chartRel = yield this.xmlTool.readXml('ppt/charts/_rels/chart1.xml.rels');
            chartRel['Relationships']['Relationship'] = {
                '$': {
                    Id: 'rId' + id,
                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
                    Target: `../embeddings/Microsoft_Excel_Sheet${id}.xlsx`
                }
            };
            return this.xmlTool.write(`ppt/charts/_rels/chart${id}.xml.rels`, chartRel);
        });
        this.createXlsxWithTableAndChart = (data, chartId) => __awaiter(this, void 0, void 0, function* () {
            yield this.xlsxGenerator.createWorkbook();
            const sheet = yield this.xlsxGenerator.createWorksheet('chart' + chartId);
            yield sheet.addTable(data);
            const bf = yield this.xlsxGenerator.generate('', 'buffer');
            this.xmlTool.writeBuffer(`ppt/embeddings/Microsoft_Excel_Sheet${chartId}.xlsx`, bf);
        });
        this.getColName = (n) => {
            var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return abc[n] || abc[n % 26];
        };
    }
    addRow(rowData, rowTemplate, colTemplate) {
        const cols = [];
        rowData.forEach((data, col) => {
            colTemplate['a:txBody']['a:p']['a:r']['a:t'] = data;
            cols.push(JSON.parse(JSON.stringify(colTemplate)));
        });
        rowTemplate['a:tc'] = JSON.parse(JSON.stringify(cols));
        return rowTemplate;
    }
}
exports.PptGraphicTool = PptGraphicTool;
