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
exports.PptxGenetator = void 0;
const xlsxGenerator_1 = require("./../xlsx/xlsxGenerator");
const pptTool_1 = require("./pptTool");
const xmlTool_1 = require("../xmlTool");
const pptGraphicTool_1 = require("./pptGraphicTool");
const chartTool_1 = require("../xlsx/chartTool");
class PptxGenetator {
    constructor() {
        this.xmlTool = new xmlTool_1.XmlTool();
        this.pptTool = new pptTool_1.PptTool(this.xmlTool);
        this.pptGraphicTool = new pptGraphicTool_1.PptGraphicTool(this.xmlTool, new xlsxGenerator_1.XlsxGenerator(), new chartTool_1.ChartTool(this.xmlTool));
        this.createPresentation = () => __awaiter(this, void 0, void 0, function* () {
            return this.xmlTool.readOriginal('pptx');
        });
        this.createSlide = () => __awaiter(this, void 0, void 0, function* () {
            const id = yield this.pptTool.addSlidePart();
            const slide = yield this.pptTool.createSlide(id);
            return {
                data: slide,
                id: id,
                tData: [],
                addTitle: (text, opt) => __awaiter(this, void 0, void 0, function* () { return yield this.pptTool.addTitle(slide, id, text, opt); }),
                addSubTitle: (text, opt) => __awaiter(this, void 0, void 0, function* () { return yield this.pptTool.addSubTitle(slide, id, text, opt); }),
                addText: (text, opt) => __awaiter(this, void 0, void 0, function* () { return yield this.pptTool.addText(slide, id, text, opt); }),
                addTable: (data, opt) => __awaiter(this, void 0, void 0, function* () { return this.pptGraphicTool.writeTable(id, slide, data, opt); }),
                addChart: (opt) => __awaiter(this, void 0, void 0, function* () { return yield this.pptGraphicTool.addChart(slide, opt, id); })
            };
        });
        this.generate = (file, type) => __awaiter(this, void 0, void 0, function* () {
            yield this.pptTool.removeTemplateSlide();
            if (type === 'file') {
                return this.xmlTool.generateFile(file, 'pptx');
            }
            else {
                return this.xmlTool.generateBuffer();
            }
        });
    }
}
exports.PptxGenetator = PptxGenetator;
