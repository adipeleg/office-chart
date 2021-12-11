import { XlsxGenerator } from './../xlsxGenerator';
describe('check xlsxGenerator', () => {
    it('', async () => {
        const gen = new XlsxGenerator();
        await gen.createWorkbook();
        const sheet2 = await gen.createWorksheet("sheet2");
        await sheet2.addTable([[0, 1, 2], [1, 2, 3], [4, 5, 6]])
        // sheet2.addChart("", [['h1', 'h2', 'h3'], [1, 2, 3], [4, 5, 6]])
        const sheet3 = await gen.createWorksheet("sheet3");
        // sheet3.addTable(getLongData());
        const sheet4 = await gen.createWorksheet("sheet4");
        await gen.generate('test5');
    })
})

const getLongData = () => {
    const data: any[][] = [['h1', 'h2', 'h3']];
    for (let i = 0; i < 100000; i++) {
        data.push([i, i + 1, i + 2])
    }
    return data;
}
