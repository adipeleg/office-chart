/// <reference types="node" />
export declare class XmlTool {
    private zip;
    private parser;
    private builder;
    constructor();
    readXlsx: () => Promise<void>;
    readXml: (file: string) => Promise<any>;
    write: (filename: string, data: any) => Promise<void>;
    writeStr: (filename: string, data: string) => Promise<void>;
    addSheetToWb: (name: string) => Promise<string>;
    createSheet: (id: string) => Promise<any>;
    generateBuffer: () => Promise<Buffer>;
    generate: () => Promise<string>;
    generateFile: (name: string) => Promise<void>;
    removeTemplateSheets: () => Promise<void>;
    writeTable: (sheet: any, data: any[][], id: string) => Promise<void>;
    addChart: (sheet: any, sheetName: string, title: string, range: string, id: string, type: 'line' | 'bar') => Promise<void>;
    private addDrawingRel;
    private addChartToDraw;
    private addChartToSheetRel;
    private addChartToSheet;
    private addChartToParts;
    private addSheetToParts;
    getColName: (n: number) => string;
    ColToNum: (char: string) => number;
}
