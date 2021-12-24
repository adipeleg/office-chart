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
const pptxGenerator_1 = require("./../pptxGenerator");
describe('create pptx', () => {
    it('create pptx with text', () => __awaiter(void 0, void 0, void 0, function* () {
        const gen = new pptxGenerator_1.PptxGenetator();
        yield gen.createPresentation();
        const slide = yield gen.createSlide();
        slide.addTitle('this is title', {
            x: '0',
            y: '0',
            color: 'FF0000',
            size: 4000
        });
        slide.addSubTitle('this is subtitle');
        slide.addText('this is text', {
            color: 'FF0000'
        });
        slide.addText('this is text 2', {
            x: '0',
            y: '0',
            size: 5000
        });
        slide.addText('this is text 3', {
            x: '0',
            y: '0',
            color: 'FF0000'
        });
        slide.addText('this is text 4', {
            x: '0',
            y: '0'
        });
        const slide2 = yield gen.createSlide();
        yield slide2.addTable(getShotData());
        const slide3 = yield gen.createSlide();
        yield slide3.addTable(getShotData2());
        const slide4 = yield gen.createSlide();
        yield slide4.addTable(getShotData3());
        const slide5 = yield gen.createSlide();
        const opt = {
            title: {
                name: 'testChart line',
                color: '8ab4f8',
                size: 3000
            },
            type: 'line',
            data: getShotData(),
            rgbColors: ['8ab4f8', 'ff7769'],
            labels: false,
        };
        yield slide5.addChart(opt);
        yield slide5.addTitle('line chart', {
            x: '0',
            y: '0',
        });
        const slide6 = yield gen.createSlide();
        opt.data = getShotDataLabels();
        opt.labels = true;
        opt.title.name = 'with labels';
        yield slide6.addChart(opt);
        yield slide6.addTitle(null);
        yield slide6.addSubTitle(null);
        const slide7 = yield gen.createSlide();
        opt.data = getShotDataLabels();
        opt.labels = true;
        opt.type = 'bar';
        opt.title.name = 'bar with labels';
        yield slide7.addChart(opt);
        yield slide7.addTitle(null);
        yield slide7.addSubTitle(null);
        yield gen.generate(__dirname + '/test11', 'file');
    }));
});
const getShotData = () => {
    return [['h', 'b', 'c', 'd', 'e'], [1, 2, 3, 4, 5], [4, 5, 6, 7, 8]];
};
const getShotDataLabels = () => {
    return [['h', 'b', 'c', 'd', 'e'], ['label1', 2, 3, 4, 5], ['label2', 5, 6, 7, 8], ['label3', 4, 6, 8, 10]];
};
const getShotData2 = () => {
    return [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]];
};
const getShotData3 = () => {
    return [[1, 2], [5, 6], [9, 10]];
};
