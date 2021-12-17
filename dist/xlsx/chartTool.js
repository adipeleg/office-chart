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
exports.ChartTool = void 0;
class ChartTool {
    constructor(xmlTool) {
        this.xmlTool = xmlTool;
        this.addChart = (sheet, sheetName, title, range, id, type) => __awaiter(this, void 0, void 0, function* () {
            let readChart = yield this.xmlTool.readXml('xl/charts/chart1.xml');
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
                const splitRange = range.split(':');
                const first = splitRange[0];
                firstCol = first[0];
                const sec = splitRange[1];
                lastCol = sec[0];
                rowNum = parseInt(sec.substring(1));
            }
            catch (_a) {
                console.log('range is not right');
                throw Error('range is not right');
            }
            const ser = Object.assign({}, readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser']);
            readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'] = [];
            for (let i = 1; i < rowNum; i++) {
                let d = JSON.parse(JSON.stringify(ser));
                ;
                d['c:idx'] = i;
                d['c:order'] = i;
                d['c:cat']['c:strRef']['c:f'] = sheetName + `!$${firstCol}$1:$${lastCol}$1`;
                d['c:val']['c:numRef']['c:f'] = sheetName + `!$${firstCol}$${(i + 1)}:$${lastCol}$${(i + 1)}`;
                readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'].push(d);
            }
            yield this.addDrawingRel(sheet, sheetName, id);
            yield this.addChartToDraw(id);
            yield this.addChartToSheet(sheet, id);
            yield this.addChartToSheetRel(id);
            yield this.addChartToParts(id);
            return this.xmlTool.write(`xl/charts/chart${id}.xml`, readChart);
        });
        this.addDrawingRel = (sheet, sheetName, id) => __awaiter(this, void 0, void 0, function* () {
            const drawRel = yield this.xmlTool.readXml('xl/drawings/_rels/drawing2.xml.rels'); //add new chart rel
            drawRel.Relationships.Relationship =
                {
                    '$': {
                        Id: 'rId' + id,
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                        Target: `../charts/chart${id}.xml`
                    }
                };
            yield this.xmlTool.write(`xl/drawings/_rels/drawing${id}.xml.rels`, drawRel);
            return id;
        });
        this.addChartToDraw = (id) => __awaiter(this, void 0, void 0, function* () {
            const draw = yield this.xmlTool.readXml('xl/drawings/drawing2.xml'); // add new chart draw
            draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = 'rId' + id;
            return this.xmlTool.write(`xl/drawings/drawing${id}.xml`, draw);
        });
        this.addChartToSheetRel = (id) => __awaiter(this, void 0, void 0, function* () {
            const draw = yield this.xmlTool.readXml('xl/worksheets/_rels/sheet2.xml.rels'); // add new chart to sheet rel
            draw['Relationships']['Relationship'] =
                {
                    '$': {
                        Id: 'rId' + id,
                        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                        Target: `../drawings/drawing${id}.xml`
                    }
                };
            return this.xmlTool.write(`xl/worksheets/_rels/sheet${id}.xml.rels`, draw);
        });
        this.addChartToSheet = (sheet, id) => __awaiter(this, void 0, void 0, function* () {
            sheet['worksheet']['drawing'] = {
                $: {
                    'r:id': "rId" + id
                }
            };
            return this.xmlTool.write(`xl/worksheets/sheet${id}.xml`, sheet);
        });
        this.addChartToParts = (id) => __awaiter(this, void 0, void 0, function* () {
            const parts = this.parts || (yield this.xmlTool.readXml('[Content_Types].xml'));
            parts['Types']['Override'].push({
                '$': {
                    ContentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
                    PartName: `/xl/charts/chart${id}.xml`
                }
            });
            parts['Types']['Override'].push({
                '$': {
                    ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
                    PartName: `/xl/drawings/drawing${id}.xml`
                }
            });
            this.parts = parts;
            return this.xmlTool.write(`[Content_Types].xml`, parts);
        });
    }
}
exports.ChartTool = ChartTool;
