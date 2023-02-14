/// <reference types="node" />
import { IPPTChartData, IPptTableOpt } from './../xlsx/models/data.model';
import { ITextModel } from './models/text.model';
export declare class PptxGenetator {
    private xmlTool;
    private pptTool;
    private pptGraphicTool;
    createPresentation: () => Promise<void>;
    createSlide: () => Promise<{
        data: any;
        id: number;
        tData: any[];
        addTitle: (text?: string, opt?: ITextModel) => Promise<void>;
        addSubTitle: (text?: string, opt?: ITextModel) => Promise<void>;
        addText: (text?: string, opt?: ITextModel) => Promise<void>;
        addTable: (data: any[][], opt?: IPptTableOpt) => Promise<void>;
        addChart: (opt: IPPTChartData) => Promise<void>;
    }>;
    generate: (file: string, type: 'file' | 'buffer') => Promise<Buffer>;
}
