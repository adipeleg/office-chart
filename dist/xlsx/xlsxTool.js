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
exports.XlsxTool = void 0;
class XlsxTool {
    constructor(xmlTool) {
        this.xmlTool = xmlTool;
        this.addSheetToWb = (name) => __awaiter(this, void 0, void 0, function* () {
            const wb = yield this.xmlTool.readXml('xl/workbook.xml');
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
            this.xmlTool.write('xl/workbook.xml', wb);
            return count;
        });
        this.createSheet = (id) => __awaiter(this, void 0, void 0, function* () {
            const resSheet = yield this.xmlTool.readXml('xl/worksheets/sheet1.xml');
            delete resSheet.worksheet.drawing;
            const WbRel = yield this.xmlTool.readXml('xl/_rels/workbook.xml.rels');
            WbRel.Relationships.Relationship.push({
                '$': {
                    Id: 'rId' + id,
                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                    Target: `worksheets/sheet${id}.xml`
                }
            });
            this.xmlTool.write(`xl/worksheets/sheet${id}.xml`, resSheet);
            this.xmlTool.write('xl/_rels/workbook.xml.rels', WbRel);
            return resSheet;
        });
        this.removeTemplateSheets = () => __awaiter(this, void 0, void 0, function* () {
            const wb = yield this.xmlTool.readXml('xl/workbook.xml');
            wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
                return 'SheetTemplate' !== it.$.name.toString() &&
                    'barTemplate' !== it.$.name.toString() &&
                    'lineTemplate' !== it.$.name.toString() &&
                    'pieTemplate' !== it.$.name.toString() &&
                    'scatterTemplate' !== it.$.name.toString();
            });
            return this.xmlTool.write('xl/workbook.xml', wb);
        });
        this.writeTable = (sheet, data, id) => __awaiter(this, void 0, void 0, function* () {
            const sheetWithTable = yield this.xmlTool.readXml('xl/worksheets/sheet2.xml');
            const rowTemplate = sheetWithTable.worksheet.sheetData.row[0];
            const header = data.shift();
            const rows = [];
            rows.push(this.addRow(header, JSON.parse(JSON.stringify(rowTemplate)), 1));
            data.forEach((data, idx) => {
                rows.push(this.addRow(data, JSON.parse(JSON.stringify(rowTemplate)), idx + 2));
            });
            sheet.worksheet.sheetData = { row: rows };
            yield this.addSharedStrings(data);
            return this.xmlTool.write(`xl/worksheets/sheet${id}.xml`, sheet);
        });
        this.addSharedStrings = (data) => __awaiter(this, void 0, void 0, function* () {
            const str = yield this.xmlTool.readXml('xl/sharedStrings.xml');
            data.forEach(row => {
                row.forEach(element => {
                    if (typeof element === 'string') {
                        const inside = str['sst']['si'].find(it => {
                            return it.t === element;
                        });
                        if (!inside) {
                            str['sst']['si'].push({
                                t: element
                            });
                            str['sst'].$.uniqueCount++;
                        }
                    }
                });
            });
            yield this.xmlTool.write('xl/sharedStrings.xml', str);
        });
        this.addSheetToParts = (id) => __awaiter(this, void 0, void 0, function* () {
            const parts = yield this.xmlTool.readXml('[Content_Types].xml');
            parts['Types']['Override'].push({
                '$': {
                    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                    PartName: `/xl/worksheets/sheet${id}.xml`
                }
            });
            return this.xmlTool.write(`[Content_Types].xml`, parts);
        });
        this.getColName = (n) => {
            var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return abc[n] || abc[n % 26];
        };
        this.ColToNum = (char) => {
            var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return abc.indexOf(char);
        };
    }
    addRow(rowData, rowTemplate, index) {
        rowTemplate.$.r = index;
        const cols = [];
        rowData.forEach((data, col) => {
            const type = typeof data === 'string' ? 's' : '';
            const c = { '$': { r: this.getColName(col) + (col <= 22 ? index : this.getColName(col)), s: '1' }, v: data };
            if (type === 's') {
                c.$['t'] = 'str';
            }
            cols.push(c);
        });
        rowTemplate.c = cols;
        return rowTemplate;
    }
}
exports.XlsxTool = XlsxTool;
