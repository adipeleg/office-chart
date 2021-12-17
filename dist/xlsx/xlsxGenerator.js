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
const chartTool_1 = require("./chartTool");
const xmlTool_1 = require("./xmlTool");
class XlsxGenerator {
    constructor() {
        this.xmlTool = new xmlTool_1.XmlTool();
        this.chartTool = new chartTool_1.ChartTool(this.xmlTool);
        this.createWorkbook = () => __awaiter(this, void 0, void 0, function* () {
            return this.xmlTool.readXlsx();
        });
        this.createWorksheet = (name) => __awaiter(this, void 0, void 0, function* () {
            const id = yield this.xmlTool.addSheetToWb(name);
            const sheet = yield this.xmlTool.createSheet(id);
            return {
                data: sheet,
                name: name,
                id: id,
                addTable: (data) => {
                    return this.xmlTool.writeTable(sheet, data, id);
                },
                addChart: (range, title, type) => __awaiter(this, void 0, void 0, function* () { return yield this.chartTool.addChart(sheet, name, title, range, id, type); })
            };
        });
        this.generate = (file, type) => __awaiter(this, void 0, void 0, function* () {
            yield this.xmlTool.removeTemplateSheets();
            if (type === 'file') {
                return this.xmlTool.generateFile(file);
            }
            else {
                return this.xmlTool.generateBuffer();
            }
        });
    }
}
exports.XlsxGenerator = XlsxGenerator;
