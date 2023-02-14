import { IData } from 'office-chart';
import { XlsxGenerator } from "office-chart";

const addData = async (sheet2, chartName?: string) => {
    const header = ['h', 'b', 'c', 'd'];
    const row1 = ['label1', 2, 3, 4];
    const row2 = ['label2', 5, 6, 7];

    await sheet2.addTable([header, row1, row2]);

    const opt: IData = {
        title: {
            name: chartName || 'testChart line',
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

}

const create = async () => {
    const gen = new XlsxGenerator();

    await gen.createWorkbook();

    const sheet1 = await gen.createWorksheet("sheet1");

    const sheet2 = await gen.createWorksheet("sheetWithChart 2");

    await addData(sheet2);
    const sheet3 = await gen.createWorksheet("sheet3");
    await gen.createWorksheet("sheet4");
    await gen.createWorksheet("sheet5");
    const sheet6 = await gen.createWorksheet("sheet6");
    await addData(sheet6, 'sheet 6 chart');
    await gen.generate(__dirname + '/test', 'file');


}
create();
