import { IPPTChartData, IPPTChartDataVal } from './../../xlsx/models/data.model';
import { PptxGenetator } from './../pptxGenerator';
describe('create pptx', () => {

    it('create pptx with text', async () => {
        const gen = new PptxGenetator();
        await gen.createPresentation();
        const slide = await gen.createSlide();
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
        const slide2 = await gen.createSlide();
        await slide2.addTable(getShotData());
        const slide3 = await gen.createSlide();
        await slide3.addTable(getShotData2());
        const slide4 = await gen.createSlide();
        await slide4.addTable(getShotData3());
        const slide5 = await gen.createSlide();
        const opt: IPPTChartData = {
            title: {
                name: 'testChart line',
                color: '8ab4f8',
                size: 3000
            },
            type: 'line',
            data: getShotData(),
            rgbColors: ['8ab4f8', 'ff7769'],
            labels: false,
        }
        await slide5.addChart(opt);
        await slide5.addTitle('line chart', {
            x: '0',
            y: '0',
        });
        const slide6 = await gen.createSlide();
        opt.data = getShotDataLabels();
        opt.labels = true;
        opt.title.name = 'with labels';
        await slide6.addChart(opt);
        await slide6.addTitle(null);
        await slide6.addSubTitle(null);

        const slide7 = await gen.createSlide();
        opt.data = getShotDataLabels();
        opt.labels = true;
        opt.type = 'bar';
        opt.title.name = 'bar with labels';
        await slide7.addChart(opt);
        await slide7.addTitle(null);
        await slide7.addSubTitle(null);

        const slide8 = await gen.createSlide();
        opt.data = getDataIPPTChartDataVal();
        opt.labels = true;
        opt.type = 'line';
        opt.title.name = 'line with labels - new format';
        await slide8.addChart(opt);
        await slide8.addTitle(null);
        await slide8.addSubTitle(null);
        await gen.generate(__dirname + '/test11', 'file');
    })
})

const getShotData = () => {
    return [['h', 'b', 'c', 'd', 'e'], [1, 2, 3, 4, 5], [4, 5, 6, 7, 8]];
}


const getShotDataLabels = () => {
    return [['h', 'b', 'c', 'd', 'e'], ['label1', 2, 3, 4, 5], ['label2', 5, 6, 7, 8], ['label3', 4, 6, 8, 10]];
}

const getShotData2 = () => {
    return [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]];
}

const getShotData3 = () => {
    return [[1, 2], [5, 6], [9, 10]];
}

const getDataIPPTChartDataVal = (): IPPTChartDataVal[] => {
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
            values: [9, 1, 2, 4, 10],
            labels: ['h', 'b', 'c', 'd', 'e']
        }
    ]
}