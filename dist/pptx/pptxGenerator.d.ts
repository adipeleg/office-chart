/// <reference types="node" />
import { ITextModel } from './models/text.model';
export declare class PptxGenetator {
    private xmlTool;
    private pptTool;
    createPresentation: () => Promise<void>;
    createSlide: () => Promise<{
        data: any;
        id: number;
        addTitle: (text: string, opt?: ITextModel) => Promise<void>;
        addSubTitle: (text: string, opt?: ITextModel) => Promise<void>;
        addText: (text: string, opt?: ITextModel) => Promise<void>;
    }>;
    generate: (file: string, type: 'file' | 'buffer') => Promise<Buffer>;
}
