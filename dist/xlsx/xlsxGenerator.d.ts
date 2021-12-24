/// <reference types="node" />
import { IData } from './models/data.model';
export declare class XlsxGenerator {
    private xmlTool;
    private chartTool;
    private xlsxTool;
    createWorkbook: () => Promise<void>;
    createWorksheet: (name: string) => Promise<{
        data: any;
        name: string;
        id: string;
        addTable: (data: any[][]) => Promise<void>;
        addChart: (opt: IData) => Promise<any>;
    }>;
    generate: (file: string, type: 'file' | 'buffer') => Promise<Buffer>;
}
