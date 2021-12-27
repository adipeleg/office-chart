import { XmlTool } from "../xmlTool";
export declare class XlsxTool {
    private xmlTool;
    constructor(xmlTool: XmlTool);
    addSheetToWb: (name: string) => Promise<string>;
    createSheet: (id: string) => Promise<any>;
    removeTemplateSheets: () => Promise<void>;
    writeTable: (sheet: any, data: any[][], id: string) => Promise<void>;
    private addRow;
    private addSharedStrings;
    private addSheetToParts;
    private getColName;
    private ColToNum;
}
