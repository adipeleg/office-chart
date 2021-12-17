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
        this.getZip = () => {
            return this.zip;
        };
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
            yield this.addSheetToParts(count);
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
        this.generateFile = (name) => __awaiter(this, void 0, void 0, function* () {
            const buf = yield this.generateBuffer();
            fs_1.default.writeFileSync(name + '.xlsx', buf);
            return buf;
        });
        this.removeTemplateSheets = () => __awaiter(this, void 0, void 0, function* () {
            const wb = yield this.readXml('xl/workbook.xml');
            wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
                return 'SheetTemplate' !== it.$.name.toString() &&
                    'barTemplate' !== it.$.name.toString() &&
                    'lineTemplate' !== it.$.name.toString() &&
                    'pieTemplate' !== it.$.name.toString() &&
                    'scatterTemplate' !== it.$.name.toString();
            });
            return this.write('xl/workbook.xml', wb);
        });
        this.writeTable = (sheet, data, id) => __awaiter(this, void 0, void 0, function* () {
            const sheetWithTable = yield this.readXml('xl/worksheets/sheet2.xml');
            const rowTemplate = sheetWithTable.worksheet.sheetData.row[0];
            const header = data.shift();
            const rows = [];
            rows.push(this.addRow(header, JSON.parse(JSON.stringify(rowTemplate)), 1));
            data.forEach((data, idx) => {
                rows.push(this.addRow(data, JSON.parse(JSON.stringify(rowTemplate)), idx + 2));
            });
            sheet.worksheet.sheetData = { row: rows };
            return this.write(`xl/worksheets/sheet${id}.xml`, sheet);
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
    addRow(rowData, rowTemplate, index) {
        rowTemplate.$.r = index;
        const cols = [];
        rowData.forEach((data, col) => {
            const type = typeof data === 'string' ? 's' : '';
            const c = { '$': { r: this.getColName(col) + (index), s: '1' }, v: data };
            if (type === 's') {
                c.$['t'] = 'str';
            }
            cols.push(c);
        });
        rowTemplate.c = cols;
        return rowTemplate;
    }
}
exports.XmlTool = XmlTool;
