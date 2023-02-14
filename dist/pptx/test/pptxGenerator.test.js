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
            cx: '4000000',
            cy: '600000',
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
        yield slide2.addTable(getShotData(), {
            x: '1000',
            y: '1000',
            colWidth: 1081250,
            rowHeight: 1059279
        });
        const slide3 = yield gen.createSlide();
        yield slide3.addTable(getShotData2());
        const slide4 = yield gen.createSlide();
        yield slide4.addTable(complexData());
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
            marker: {
                shape: 'circle',
                size: 4
            },
            lineWidth: 20000,
            location: {
                x: '1000',
                y: '1000'
            }
        };
        yield slide5.addChart(opt);
        yield slide5.addTitle('line chart', {
            x: '0',
            y: '0',
        });
        yield slide5.addSubTitle(null);
        const slide6 = yield gen.createSlide();
        opt.data = getShotDataLabels();
        opt.labels = true;
        opt.title.name = 'with labels';
        slide6.addTitle(null);
        slide6.addSubTitle(null);
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
        const slide8 = yield gen.createSlide();
        opt.data = getDataIPPTChartDataVal();
        opt.type = 'line';
        opt.title.name = 'line with labels - new format';
        yield slide8.addChart(opt);
        yield slide8.addTitle(null);
        yield slide8.addSubTitle(null);
        const slide9 = yield gen.createSlide();
        opt.data = getDataIPPTChartDataVal();
        opt.type = 'line';
        opt.title.name = 'line with labels - new format';
        yield slide9.addChart(addNewData());
        yield slide9.addTitle(null);
        yield slide9.addSubTitle(null);
        yield gen.generate(__dirname + '/test11', 'file');
    }));
});
const getShotData = () => {
    return [['', 'b', 'c', 'd', 'e'], [1, 2, 30, 40, 5], [4, 5, -600, 70, 800], [4, 5, -5500, 70, 20]];
};
const getShotDataLabels = () => {
    return [['', 'b', 'c', 'd', 'e'], ['label1', 2, 3, 4, 5], ['label2', 5, 6, 7, 8], ['label3', 4, 6, 8, 10]];
};
const getShotData2 = () => {
    return [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]];
};
const getShotData3 = () => {
    return [[1, 2], [5, 6], [9, 10]];
};
const getDataIPPTChartDataVal = () => {
    return [
        {
            name: 'lab1 test',
            values: [1, 2, 3, 4, 5],
            labels: ['h', 'b', 'c', 'd', 'e']
        }, {
            name: 'lab2 test',
            values: [4, 5, 6, 7, 8],
            labels: ['h', 'b', 'c', 'd', 'e']
        }, {
            name: 'lab3 test',
            values: [9, 1, 2, 4, -10],
            labels: ['h', 'b', 'c', 'd', 'e']
        }, {
            name: 'lab4 test',
            values: [9, 1, 2, 4, 10],
            labels: ['h', 'b', 'c', 'd', 'e']
        }, {
            name: 'lab5 test',
            values: [9, 1, 2, -4, 10],
            labels: ['h', 'b', 'c', 'd', 'e']
        }
    ];
};
const addNewData = () => {
    return {
        title: { name: 'c over time' },
        type: 'line',
        labels: true,
        marker: {
            size: 4,
            shape: 'circle'
        },
        data: [{
                name: 'c1_long_long_long',
                labels: ['May', 'Aug', 'Nov'],
                values: [17362, 28283, 12842]
            }, {
                name: 'c2_long_long_long long_longlong_longlong_long',
                labels: ['May', 'Aug', 'Nov'],
                values: [-29.548774549586106, -72.19879488464903, -33.88251042578386]
            }, {
                name: 'c3_long_long_long',
                labels: ['May', 'Aug', 'Nov'],
                values: [17362, -282830, 12842]
            }, {
                name: 'c4_long_long_long',
                labels: ['May', 'Aug', 'Nov'],
                values: [173620, 28283, 12842]
            }]
    };
};
const complexData = () => {
    return [
        [
            't1',
            't2',
            't3',
            't4',
            't5',
            't6',
            't7'
        ],
        [
            'd1',
            1402,
            5042,
            -33.2332,
            '2021-10-01',
            '2021-07-01',
            '2021-10-01 Vs 2021-07-01'
        ],
        [
            'd1',
            1544,
            1900,
            -1.2395293,
            '2021-04-01',
            '2021-01-01',
            '2021-04-01 Vs 2021-01-01'
        ],
        [
            'd1',
            3456,
            12345,
            -34.34632423,
            '2021-07-01',
            '2021-04-01',
            '2021-07-01 Vs 2021-04-01'
        ],
        [
            'd2',
            10521,
            15963,
            -34.09133621499718,
            '2021-10-01',
            '2021-07-01',
            '2021-10-01 Vs 2021-07-01'
        ],
        [
            'd2',
            14879,
            21763,
            -31.631668428065982,
            '2021-04-01',
            '2021-01-01',
            '2021-04-01 Vs 2021-01-01'
        ],
        [
            'd2',
            23383,
            84289,
            -72.25853907390051,
            '2021-07-01',
            '2021-04-01',
            '2021-07-01 Vs 2021-04-01'
        ],
        [
            'd3',
            807,
            1172,
            -31.14334470989761,
            '2021-04-01',
            '2021-01-01',
            '2021-04-01 Vs 2021-01-01'
        ],
        [
            'd3',
            1040,
            1095,
            -5.0228310502283104,
            '2021-10-01',
            '2021-07-01',
            '2021-10-01 Vs 2021-07-01'
        ],
        [
            'd3',
            1579,
            4476,
            -64.72296693476318,
            '2021-07-01',
            '2021-04-01',
            '2021-07-01 Vs 2021-04-01'
        ]
    ];
};
