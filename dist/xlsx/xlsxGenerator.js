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
exports.XlsxGenerator = void 0;
const xlsxTool_1 = require("./xlsxTool");
const xmlTool_1 = require("../xmlTool");
const chartTool_1 = require("./chartTool");
class XlsxGenerator {
    constructor() {
        this.xmlTool = new xmlTool_1.XmlTool();
        this.chartTool = new chartTool_1.ChartTool(this.xmlTool);
        this.xlsxTool = new xlsxTool_1.XlsxTool(this.xmlTool);
        this.createWorkbook = () => __awaiter(this, void 0, void 0, function* () {
            return this.xmlTool.readOriginal('xlsx');
        });
        this.createWorksheet = (name) => __awaiter(this, void 0, void 0, function* () {
            const id = yield this.xlsxTool.addSheetToWb(name);
            const sheet = yield this.xlsxTool.createSheet(id);
            return {
                data: sheet,
                name: name,
                id: id,
                addTable: (data) => {
                    return this.xlsxTool.writeTable(sheet, data, id);
                },
                addChart: (opt) => __awaiter(this, void 0, void 0, function* () { return this.chartTool.addChart(sheet, name, opt, id); })
            };
        });
        this.generate = (file, type) => __awaiter(this, void 0, void 0, function* () {
            yield this.xlsxTool.removeTemplateSheets();
            if (type === 'file') {
                return this.xmlTool.generateFile(file, 'xlsx');
            }
            else {
                return this.xmlTool.generateBuffer();
            }
        });
    }
}
exports.XlsxGenerator = XlsxGenerator;
