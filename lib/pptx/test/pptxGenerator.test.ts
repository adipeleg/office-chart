import { PptxGenetator } from './../pptxGenerator';
describe('create pptx', () => {

    it('create pptx with text', async () => {
        const gen = new PptxGenetator();
        await gen.createPresentation();
        const slide = await gen.createSlide();
        const slide2 = await gen.createSlide();
        const slide3 = await gen.createSlide();
        const slide4 = await gen.createSlide();
        await gen.generate(__dirname + '/test10', 'file');
    })
})