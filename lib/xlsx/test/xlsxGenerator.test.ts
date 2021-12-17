import { IData } from '../models/data.model';
import { XlsxGenerator } from './../xlsxGenerator';
describe('check xlsxGenerator', () => {
    it('', async () => {
        const gen = new XlsxGenerator();
        await gen.createWorkbook();
        const sheet1 = await gen.createWorksheet("sheet1");
        const sheet2 = await gen.createWorksheet("sheetWithChart2");
        await sheet2.addTable(getShotData());
        const opt: IData = {
            title: {
                name: 'testChart line',
                color: '8ab4f8',
                size: 3000
            },
            range: 'B1:D3',
            type: 'line',
            rgbColors: ['8ab4f8', 'ff7769'],
            labels: true,
            marker: {
                size: 4,
                shape: 'square'
            }
        }
        await sheet2.addChart(opt)
        const sheet3 = await gen.createWorksheet("sheet3");

        await sheet3.addTable(getLongData());
        const opt2: IData = {
            title: {
                name: 'testChart bar',
                color: '2d2e30'
            },
            range: 'A1:B4',
            type: 'bar',
            rgbColors: ['8ab4f8', 'ff7769', '1d9f08']

        }
        await sheet3.addChart(opt2)
        const sheet4 = await gen.createWorksheet("sheet4");
        await sheet4.addTable(getShotData());
        const optPie: IData = {
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
        }
        await sheet4.addChart(optPie);

        const sheet5 = await gen.createWorksheet("sheet5");
        await sheet5.addTable(getShotData());
        const optScatter: IData = {
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
        }
        await sheet5.addChart(optScatter);
        await gen.generate(__dirname + '/test10', 'file');
        // const buffer = await gen.generate(__dirname + '/test9', 'file');
        // console.log(buffer);
    })
})

const getLongData = () => {
    const data: any[][] = [];//[['h1', 'h2', 'h3']];
    for (let i = 0; i < 1000; i++) {
        data.push([i, i + 1, i + 2])
    }
    return data;
}
const getShotData = () => {
    return [['h', 'b', 'c', 'd'], ['tot', 2, 3, 4], ['sos', 5, 6, 7]];
    // return [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]];
}
