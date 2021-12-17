/// <reference types="node" />
import JSZip from "jszip";
export declare class XmlTool {
    private zip;
    private parser;
    private builder;
    private parts;
    constructor();
    getZip: () => JSZip;
    readXlsx: () => Promise<void>;
    readXml: (file: string) => Promise<any>;
    write: (filename: string, data: any) => Promise<void>;
    writeStr: (filename: string, data: string) => Promise<void>;
    addSheetToWb: (name: string) => Promise<string>;
    createSheet: (id: string) => Promise<any>;
    generateBuffer: () => Promise<Buffer>;
    generateFile: (name: string) => Promise<Buffer>;
    removeTemplateSheets: () => Promise<void>;
    writeTable: (sheet: any, data: any[][], id: string) => Promise<void>;
    private addRow;
    private addSheetToParts;
    private getColName;
    private ColToNum;
}
