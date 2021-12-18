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
        const slide3 = yield gen.createSlide();
        const slide4 = yield gen.createSlide();
        yield gen.generate(__dirname + '/test10', 'file');
    }));
});
