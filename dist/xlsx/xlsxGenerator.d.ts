/// <reference types="node" />
export declare class XlsxGenerator {
    private xmlTool;
    createWorkbook: () => Promise<void>;
    createWorksheet: (name: string) => Promise<{
        data: any;
        name: string;
        id: string;
        addTable: (data: any[][]) => Promise<void>;
        addChart: (range: string, title: string, type: 'line' | 'bar') => Promise<void>;
    }>;
    generate: (file: string, type: 'file' | 'buffer') => Promise<Buffer>;
}
