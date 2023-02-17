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
        this.addChart = (sheet, sheetName, opt, id) => __awaiter(this, void 0, void 0, function* () {
            let readChart = yield this.xmlTool.readXml(`xl/charts/chart${this.getChartNum(opt)}.xml`);
            this.buildChart(readChart, opt, sheetName);
            if (sheet) {
                yield this.addDrawingRel(id);
                yield this.addChartToDraw(id);
                yield this.addChartToSheetRel(id);
                yield this.addChartToParts(id);
                yield this.addChartToSheet(sheet, id);
                return this.xmlTool.write(`xl/charts/chart${id}.xml`, readChart);
            }
            return readChart;
        });
        this.getChartNum = (opt) => {
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
        };
        this.buildChart = (readChart, opt, sheetName) => {
            var _a, _b;
            sheetName = `'${sheetName}'`;
            readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:t'] = opt.title.name;
            if (opt.title.color) {
                readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:rPr']['a:solidFill']['a:srgbClr'].$.val = opt.title.color;
            }
            if (opt.title.size) {
                readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:rPr'].$.sz = opt.title.size;
            }
            const chartType = `c:${opt.type}Chart`;
            let rowNum = '';
            let lastCol = '';
            let firstCol = '';
            try {
                const splitRange = opt.range.split(':');
                Array.from(splitRange[0]).forEach(letter => {
                    const notNumCheck = isNaN(parseInt(letter));
                    if (notNumCheck) {
                        firstCol += letter;
                    }
                });
                Array.from(splitRange[1]).forEach(letter => {
                    const letterNum = parseInt(letter);
                    if (isNaN(letterNum)) {
                        lastCol += letter;
                    }
                    else {
                        rowNum += letter;
                    }
                });
            }
            catch (_c) {
                console.log('range is not right');
                throw Error('range is not right');
            }
            const ser = Object.assign({}, readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser']);
            readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'] = [];
            // delete readChart['c:chartSpace']['c:chart']['c:plotArea']['c:layout']
            for (let i = 1; i < parseInt(rowNum); i++) {
                const data = JSON.parse(JSON.stringify(ser));
                let d = data[0] || data;
                d['c:idx'] = { $: { val: i - 1 } };
                d['c:order'] = { $: { val: i - 1 } };
                if (opt.type !== 'scatter') {
                    d['c:cat']['c:strRef']['c:f'] = sheetName + `!$${firstCol}$1:$${lastCol}$1`;
                    d['c:val']['c:numRef']['c:f'] = sheetName + `!$${firstCol}$${(i + 1)}:$${lastCol}$${(i + 1)}`;
                    if (opt.hasOwnProperty('data')) {
                        d['c:cat']['c:strRef']['c:strCache'] = this.buildCache(opt['data'][0], opt.labels);
                        d['c:val']['c:numRef']['c:numCache'] = this.buildCache(opt['data'][i], opt.labels);
                    }
                }
                else {
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
                        };
                    }
                }
                if (opt.type === 'line') {
                    d['c:spPr']['a:ln'].$.w = opt.lineWidth || 30000;
                }
                if (d['c:marker'] && (opt === null || opt === void 0 ? void 0 : opt.marker) && (opt.type === 'line')) {
                    d['c:marker']['c:size'].$.val = ((_a = opt === null || opt === void 0 ? void 0 : opt.marker) === null || _a === void 0 ? void 0 : _a.size) || '4';
                    d['c:marker']['c:symbol'].$.val = ((_b = opt === null || opt === void 0 ? void 0 : opt.marker) === null || _b === void 0 ? void 0 : _b.shape) || 'circle';
                    delete d['c:marker']['c:spPr']['a:noFill'];
                    if (opt.rgbColors && opt.rgbColors[i - 1]) {
                        d['c:marker']['c:spPr']['a:solidFill'] = { 'a:srgbClr': { $: { val: opt.rgbColors[i - 1] } } };
                        d['c:marker']['c:spPr']['a:ln']['a:solidFill']['a:srgbClr'].$.val = opt.rgbColors[i - 1];
                    }
                }
                if (opt.labels) {
                    d['c:tx'] = {
                        'c:strRef': {
                            'c:f': sheetName + `!$A$${i + 1}`
                        }
                    };
                    if (opt.hasOwnProperty('data')) {
                        d['c:tx']['c:strRef']['c:strCache'] = {
                            'c:ptCount': { $: { val: 1 } },
                            'c:pt': {
                                $: { idx: 0 },
                                'c:v': opt['data'][i][0]
                            }
                        };
                    }
                }
                else {
                    delete d['c:tx'];
                }
                readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'].push(d);
                if (!!readChart['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'] && opt.type !== 'scatter') {
                    readChart['c:chartSpace']['c:chart']['c:plotArea']['c:valAx']['c:majorUnit'] = {};
                    readChart['c:chartSpace']['c:chart']['c:plotArea']['c:valAx']['c:minorUnit'] = {};
                    if (readChart['c:chartSpace']['c:chart']['c:legend']['c:layout']) {
                        delete readChart['c:chartSpace']['c:chart']['c:legend']['c:layout'];
                        // readChart['c:chartSpace']['c:chart']['c:legend']['c:layout']['c:manualLayout']['c:h'] = {
                        //     $: { val: 0.08 * rowNum + 1 }
                        // }
                    }
                }
            }
            return readChart;
        };
        this.buildCache = (rowData, labels) => {
            const rowDataCopy = JSON.parse(JSON.stringify(rowData));
            if (labels) {
                rowDataCopy.shift();
            }
            const cache = { 'c:ptCount': { $: { val: `${rowDataCopy.length}` } } };
            cache['c:pt'] = [];
            for (let i = 0; i < rowDataCopy.length; i++) {
                cache['c:pt'].push({
                    $: { idx: `${i}` },
                    'c:v': rowDataCopy[i]
                });
            }
            return cache;
        };
        this.addDrawingRel = (id) => __awaiter(this, void 0, void 0, function* () {
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
