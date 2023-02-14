import { PptxGenetator } from 'office-chart';
import { IData } from 'office-chart';
import { XlsxGenerator } from "office-chart";
import { IPPTChartData } from 'office-chart/dist/xlsx/models/data.model';

const create = async () => {
    const gen = new XlsxGenerator();

    await gen.createWorkbook();

    const sheet1 = await gen.createWorksheet("sheet1");

    const sheet2 = await gen.createWorksheet("sheetWithChart2");

    const header = ['h', 'b', 'c', 'd'];
    const row1 = ['label1', 2, 3, 4];
    const row2 = ['label2', 5, 6, 7];

    await sheet2.addTable([header, row1, row2]);

    const opt: IData = {
        title: {
            name: 'testChart line',
            color: '8ab4f8',
            size: 5000
        },
        range: 'B1:D3',
        type: 'line',
        rgbColors: ['8ab4f8', 'ff7769'],
        labels: true, //table contains labels
        marker: {
            size: 4,
            shape: 'square' //marker shapes, can be circle, diamond, star
        }
    }

    await sheet2.addChart(opt)

    const sheet3 = await gen.createWorksheet("sheet3");

    await gen.generate(__dirname + '/test', 'file');

    
}
create();

const slide = async () =>  {
    const gen = new PptxGenetator();

    await gen.createPresentation();

    const slide = await gen.createSlide();

    slide.addTitle("this is title", {
        x: "0",
        y: "0",
        color: "FF0000",
        size: 4000,
    });

    slide.addSubTitle("this is subtitle");

    slide.addText("this is text", {
        color: "FF0000",
    });

    const header = ["h", "b", "c", "d"];
    const row1 = ["label1", 2, 3, 4];
    const row2 = ["label2", 5, 6, 7];

    await slide.addTable([header, row1, row2]);

    const opt: IPPTChartData = {
        title: {
            name: "testChart line",
            color: "8ab4f8",
            size: 3000,
        },
        type: "line",
        data: [header, row1, row2],
        rgbColors: ["8ab4f8", "ff7769"],
        labels: true,
        marker: {
            shape: 'circle',
            size: 4
        }
    };

    const slide2 = await gen.createSlide();

    await slide2.addChart(opt);
    await slide2.addTitle(null); //remove title
    await slide2.addSubTitle(null); //remove subtitle

    await gen.generate(__dirname + "/test3", "file");

}

slide()

const getShotData = () => {
    return [['h', 'b', 'c', 'd'], ['1', 2, 3, 4], ['2', 5, 6, 7]];
    // return [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]];
}

const getLongData = () => {
    const data: any[][] = [];//[['h1', 'h2', 'h3']];
    for (let i = 0; i < 1000; i++) {
        data.push([i, i + 1, i + 2])
    }
    return data;
}

const createppt = async () => {
    const gen = new PptxGenetator();
    await gen.createPresentation();
    const s1 = await gen.createSlide();
    s1.addTitle('hello');
    gen.generate(__dirname + '/testppt', 'file')
}

createppt();