import { PptTool } from './pptTool';
import { XmlTool } from "../xmlTool";

export class PptxGenetator {
    private xmlTool: XmlTool = new XmlTool();
    private pptTool: PptTool = new PptTool(this.xmlTool);

    public createPresentation = async () => {
        return this.xmlTool.readOriginal('pptx');
    }

    public createSlide = async () => {
        const id = await this.pptTool.addSlidePart();
        const slide = await this.pptTool.createSlide(id);
        return {
            data: slide,
            id: id,
            addText: async (text: string) => await this.pptTool.addText(text)
            // addTable: (data: any[][]) => {
            //     return this.xmlTool.writeTable(sheet, data, id)
            // },
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