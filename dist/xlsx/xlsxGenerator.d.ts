/// <reference types="node" />
import { IData } from './models/data.model';
export declare class XlsxGenerator {
    private xmlTool;
    private chartTool;
    createWorkbook: () => Promise<void>;
    createWorksheet: (name: string) => Promise<{
        data: any;
        name: string;
        id: string;
        addTable: (data: any[][]) => Promise<void>;
        addChart: (opt: IData) => Promise<void>;
    }>;
    generate: (file: string, type: 'file' | 'buffer') => Promise<Buffer>;
}
