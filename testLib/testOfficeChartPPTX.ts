import { PptxGenetator } from 'office-chart';
import { IPPTChartData } from 'office-chart/dist/xlsx/models/data.model';


const slide = async () => {
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
    await slide2.addTitle(); //remove title
    await slide2.addSubTitle(); //remove subtitle

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