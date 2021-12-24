import { XmlTool } from "../xmlTool";
import { ITextModel } from "./models/text.model";
export declare class PptTool {
    private xmlTool;
    constructor(xmlTool: XmlTool);
    addSlidePart: () => Promise<number>;
    private addSlidePPTRels;
    private addSlideRels;
    private addSlideToPPT;
    removeTemplateSlide: () => Promise<void>;
    createSlide: (id: number) => Promise<any>;
    addTitle: (slide: any, id: number, text: string, opt?: ITextModel) => Promise<void>;
    addSubTitle: (slide: any, id: number, text: string, opt?: ITextModel) => Promise<void>;
    addText: (slide: any, id: number, text: string, opt?: ITextModel) => Promise<void>;
    private addColorAndSize;
}
