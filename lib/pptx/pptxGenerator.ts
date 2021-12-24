import { PptTool } from './pptTool';
import { XmlTool } from "../xmlTool";
import { ITextModel } from './models/text.model';
import { PptGraphicTool } from './pptGraphicTool';

export class PptxGenetator {
    private xmlTool: XmlTool = new XmlTool();
    private pptTool: PptTool = new PptTool(this.xmlTool);
    private pptGraphicTool: PptGraphicTool = new PptGraphicTool(this.xmlTool);

    public createPresentation = async () => {
        return this.xmlTool.readOriginal('pptx');
    }

    public createSlide = async () => {
        const id = await this.pptTool.addSlidePart();
        const slide = await this.pptTool.createSlide(id);
        console.log(id)
        return {
            data: slide,
            id: id,
            addTitle: async (text: string, opt?: ITextModel) => await this.pptTool.addTitle(slide, id, text, opt),
            addSubTitle: async (text: string, opt?: ITextModel) => await this.pptTool.addSubTitle(slide, id, text, opt),
            addText: async (text: string, opt?: ITextModel) => await this.pptTool.addText(slide, id, text, opt),
            addTable: async (data: any[][]) => await this.pptGraphicTool.writeTable(id, slide, data)
            // addChart: async (opt: IData) => await this.chartTool.addChart(sheet, name, opt, id)
        }
    }

    public generate = async (file: string, type: 'file' | 'buffer') => {
        await this.pptTool.removeTemplateSlide();
        if (type === 'file') {
            return this.xmlTool.generateFile(file, 'pptx');
        } else {
            return this.xmlTool.generateBuffer();
        }
    }

}