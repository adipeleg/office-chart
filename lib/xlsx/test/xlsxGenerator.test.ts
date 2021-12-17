import { XlsxGenerator } from './../xlsxGenerator';
describe('check xlsxGenerator', () => {
    it('', async () => {
        const gen = new XlsxGenerator();
        await gen.createWorkbook();
        const sheet1 = await gen.createWorksheet("sheet1");
        const sheet2 = await gen.createWorksheet("sheetWithChart2");
        await sheet2.addTable(getShotData());
        await sheet2.addChart("B1:D3", 'testChart line', 'line')
        const sheet3 = await gen.createWorksheet("sheet3");
        // await sheet3.addTable(getShotData());
        await sheet3.addTable(getLongData());
        await sheet3.addChart("A1:C100", 'testChart bar', 'bar')
        const sheet4 = await gen.createWorksheet("sheet4");
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
    return [['h', 'b', 'c', 'd'], [1, 2, 3, 4], [4, 5, 6, 7]];
    // return [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]];
}
