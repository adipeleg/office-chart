import { XlsxGenerator } from './../xlsxGenerator';
describe('check xlsxGenerator', () => {
    it('', async() => {
        const gen = new XlsxGenerator();
        await gen.createWorkbook();
        await gen.createWorksheet("sheet2");
        await gen.createWorksheet("sheet3");
        await gen.createWorksheet("sheet4");
        await gen.generate('test5');
    })
})
