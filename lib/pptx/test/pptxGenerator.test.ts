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
        const slide3 = await gen.createSlide();
        const slide4 = await gen.createSlide();
        await gen.generate(__dirname + '/test10', 'file');
    })
})