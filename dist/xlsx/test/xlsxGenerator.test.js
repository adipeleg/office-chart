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
const xlsxGenerator_1 = require("./../xlsxGenerator");
describe('check xlsxGenerator', () => {
    it('test create xlsx with multi sheets and charts', () => __awaiter(void 0, void 0, void 0, function* () {
        const gen = new xlsxGenerator_1.XlsxGenerator();
        yield gen.createWorkbook();
        const sheet1 = yield gen.createWorksheet("sheet1");
        const sheet2 = yield gen.createWorksheet("sheetWithChart2");
        yield sheet2.addTable(getShotData());
        const opt = {
            title: {
                name: 'testChart line',
                color: '8ab4f8',
                size: 3000
            },
            range: 'B1:D4',
            type: 'line',
            rgbColors: ['8ab4f8', 'ff7769'],
            labels: true,
            marker: {
                size: 4,
                shape: 'square'
            },
            lineWidth: 20000
        };
        yield sheet2.addChart(opt);
        const sheet3 = yield gen.createWorksheet("sheet3");
        yield sheet3.addTable(getLongData());
        const opt2 = {
            title: {
                name: 'testChart bar',
                color: '2d2e30'
            },
            range: 'A1:B4',
            type: 'bar',
            rgbColors: ['8ab4f8', 'ff7769', '1d9f08']
        };
        yield sheet3.addChart(opt2);
        const sheet4 = yield gen.createWorksheet("sheet4");
        yield sheet4.addTable(getShotData());
        const optPie = {
            title: {
                name: 'testChart pie',
                color: '8ab4f8',
                size: 3000
            },
            range: 'B1:D3',
            type: 'pie',
            rgbColors: ['8ab4f8', 'ff7769'],
            labels: true,
            marker: {
                size: 4,
                shape: 'square'
            }
        };
        yield sheet4.addChart(optPie);
        const sheet5 = yield gen.createWorksheet("sheet5");
        yield sheet5.addTable(getShotData());
        const optScatter = {
            title: {
                name: 'testChart scatter',
                color: '8ab4f8',
                size: 3000
            },
            range: 'B1:D3',
            type: 'scatter',
            rgbColors: ['8ab4f8', 'ff7769'],
            labels: true,
            marker: {
                size: 4,
                shape: 'square'
            }
        };
        yield sheet5.addChart(optScatter);
        yield gen.generate(__dirname + '/test10', 'file');
        // const buffer = await gen.generate(__dirname + '/test9', 'file');
        // console.log(buffer);
    }));
});
const getLongData = () => {
    const data = []; //[['h1', 'h2', 'h3']];
    for (let i = 0; i < 1000; i++) {
        data.push([i, i + 1, i + 2000]);
    }
    return data;
};
const getShotData = () => {
    return [['h', 'b', 'c', 'd'], ['tot', 2, -300, 4], ['sos', 5, 6000, 7], ['bos', 2, 1000, 10]];
    // return [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]];
};
