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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.XmlTool = void 0;
const xml2js_1 = __importDefault(require("xml2js"));
const jszip_1 = __importDefault(require("jszip"));
const fs_1 = __importDefault(require("fs"));
class XmlTool {
    constructor() {
        this.readXlsx = () => __awaiter(this, void 0, void 0, function* () {
            let path = __dirname + "/templates/template.xlsx";
            yield new Promise((resolve, reject) => fs_1.default.readFile(path, (err, data) => __awaiter(this, void 0, void 0, function* () {
                if (err) {
                    console.error(`Template ${path} not read: ${err}`);
                    reject(err);
                    return;
                }
                ;
                return yield this.zip.loadAsync(data).then(d => {
                    resolve(d);
                });
            })));
        });
        this.readXml = (file) => __awaiter(this, void 0, void 0, function* () {
            return this.zip.file(file).async('string').then(data => {
                return this.parser.parseStringPromise(data);
            });
        });
        this.write = (filename, data) => __awaiter(this, void 0, void 0, function* () {
            var xml = this.builder.buildObject(data);
            this.zip.file(filename, Buffer.from(xml), { base64: true });
        });
        this.writeStr = (filename, data) => __awaiter(this, void 0, void 0, function* () {
            // var xml = this.builder.buildObject(data);
            this.zip.file(filename, Buffer.from(data), { base64: true });
        });
        this.addSheetToWb = (name) => __awaiter(this, void 0, void 0, function* () {
            const wb = yield this.readXml('xl/workbook.xml');
            let count;
            if (!Array.isArray(wb.workbook.sheets.sheet)) {
                count = '5';
                wb.workbook.sheets = {
                    sheet: [
                        { '$': wb.workbook.sheets.sheet.$ },
                        { '$': { state: 'visible', name: 'Sheet2', sheetId: '2', 'r:id': 'rId' + count } }
                    ]
                };
            }
            else {
                count = wb.workbook.sheets.sheet.length + 4;
                wb.workbook.sheets.sheet.push({
                    '$': { state: 'visible', name: name, sheetId: count, 'r:id': 'rId' + count }
                });
            }
            this.addSheetToParts(count);
            this.write('xl/workbook.xml', wb);
            return count;
        });
        this.createSheet = (id) => __awaiter(this, void 0, void 0, function* () {
            const resSheet = yield this.readXml('xl/worksheets/sheet1.xml');
            delete resSheet.worksheet.drawing;
            const WbRel = yield this.readXml('xl/_rels/workbook.xml.rels');
            WbRel.Relationships.Relationship.push({
                '$': {
                    Id: 'rId' + id,
                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                    Target: `worksheets/sheet${id}.xml`
                }
            });
            this.write(`xl/worksheets/sheet${id}.xml`, resSheet);
            this.write('xl/_rels/workbook.xml.rels', WbRel);
            return resSheet;
        });
        this.generateBuffer = () => __awaiter(this, void 0, void 0, function* () {
            return this.zip.generateAsync({ type: 'nodebuffer' });
        });
        this.generate = () => __awaiter(this, void 0, void 0, function* () {
            return this.zip.generateAsync({ type: 'string' });
        });
        this.generateFile = (name) => __awaiter(this, void 0, void 0, function* () {
            const buf = yield this.generateBuffer();
            fs_1.default.writeFileSync(name + '.xlsx', buf);
        });
        this.removeTemplateSheets = () => __awaiter(this, void 0, void 0, function* () {
            const wb = yield this.readXml('xl/workbook.xml');
            wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
                return 'SheetTemplate' !== it.$.name.toString() && 'ChartTemplate' !== it.$.name.toString();
            });
            return this.write('xl/workbook.xml', wb);
        });
        this.writeTable = (sheet, data, id) => __awaiter(this, void 0, void 0, function* () {
            const header = data.shift();
            var rows = [{
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
                        };
                    })
                }];
            data.forEach((f, y) => {
                var r = {
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
        });
        this.addChart = (sheet, sheetName, title, range, id, type) => __awaiter(this, void 0, void 0, function* () {
            let readChart = yield this.readXml('xl/charts/chart1.xml');
            readChart['c:chartSpace']['c:chart']['c:title']['c:tx']['c:rich']['a:p']['a:r']['a:t'] = title;
            const chartType = type === 'line' ? 'c:lineChart' : 'c:barChart';
            if (type === 'line') {
                readChart['c:chartSpace']['c:chart']['c:plotArea']['c:lineChart'] = JSON.parse(JSON.stringify(readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart']));
                delete readChart['c:chartSpace']['c:chart']['c:plotArea']['c:barChart'];
            }
            let rowNum = 1;
            try {
                rowNum = this.ColToNum(range.split(':')[1][0]);
            }
            catch (_a) {
                console.log('range is not right');
            }
            const ser = Object.assign({}, readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser']);
            readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'] = [];
            for (let i = 1; i < rowNum + 1; i++) {
                let d = JSON.parse(JSON.stringify(ser));
                ;
                d['c:idx'] = i;
                d['c:order'] = i;
                d['c:cat']['c:strRef']['c:f'] = sheetName + '!$A$1:$C$1';
                d['c:val']['c:numRef']['c:f'] = sheetName + '!$A$' + (i + 1) + ':$C$' + (i + 1) + '';
                readChart['c:chartSpace']['c:chart']['c:plotArea'][chartType]['c:ser'].push(d);
            }
            yield this.addDrawingRel(sheet, sheetName, id);
            yield this.addChartToDraw(id);
            yield this.addChartToSheet(sheet, id);
            yield this.addChartToSheetRel(id);
            yield this.addChartToParts(id);
            return this.write(`xl/charts/chart${id}.xml`, readChart);
        });
        this.addDrawingRel = (sheet, sheetName, id) => __awaiter(this, void 0, void 0, function* () {
            const drawRel = yield this.readXml('xl/drawings/_rels/drawing2.xml.rels'); //add new chart rel
            drawRel.Relationships.Relationship =
                {
                    '$': {
                        Id: 'rId' + id,
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                        Target: `../charts/chart${id}.xml`
                    }
                };
            yield this.write(`xl/drawings/_rels/drawing${id}.xml.rels`, drawRel);
            return id;
        });
        this.addChartToDraw = (id) => __awaiter(this, void 0, void 0, function* () {
            const draw = yield this.readXml('xl/drawings/drawing2.xml'); // add new chart draw
            draw['xdr:wsDr']['xdr:oneCellAnchor']['xdr:graphicFrame']['a:graphic']['a:graphicData']['c:chart'].$['r:id'] = 'rId' + id;
            return this.write(`xl/drawings/drawing${id}.xml`, draw);
        });
        this.addChartToSheetRel = (id) => __awaiter(this, void 0, void 0, function* () {
            const draw = yield this.readXml('xl/worksheets/_rels/sheet2.xml.rels'); // add new chart to sheet rel
            draw['Relationships']['Relationship'] =
                {
                    '$': {
                        Id: 'rId' + id,
                        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                        Target: `../drawings/drawing${id}.xml`
                    }
                };
            return this.write(`xl/worksheets/_rels/sheet${id}.xml.rels`, draw);
        });
        this.addChartToSheet = (sheet, id) => __awaiter(this, void 0, void 0, function* () {
            sheet['worksheet']['drawing'] = {
                $: {
                    'r:id': "rId" + id
                }
            };
            return this.write(`xl/worksheets/sheet${id}.xml`, sheet);
        });
        this.addChartToParts = (id) => __awaiter(this, void 0, void 0, function* () {
            const parts = yield this.readXml('[Content_Types].xml');
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
            return this.write(`[Content_Types].xml`, parts);
        });
        this.addSheetToParts = (id) => __awaiter(this, void 0, void 0, function* () {
            const parts = yield this.readXml('[Content_Types].xml');
            parts['Types']['Override'].push({
                '$': {
                    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                    PartName: `/xl/worksheets/sheet${id}.xml`
                }
            });
            return this.write(`[Content_Types].xml`, parts);
        });
        this.getColName = (n) => {
            var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return abc[n] || abc[n % 26];
        };
        this.ColToNum = (char) => {
            var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return abc.indexOf(char);
        };
        this.zip = new jszip_1.default();
        this.parser = new xml2js_1.default.Parser({ explicitArray: false });
        this.builder = new xml2js_1.default.Builder();
    }
}
exports.XmlTool = XmlTool;
