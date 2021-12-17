import { XmlTool } from "./xmlTool";
export declare class ChartTool {
    private xmlTool;
    private parts;
    constructor(xmlTool: XmlTool);
    addChart: (sheet: any, sheetName: string, title: string, range: string, id: string, type: 'line' | 'bar') => Promise<void>;
    private addDrawingRel;
    private addChartToDraw;
    private addChartToSheetRel;
    private addChartToSheet;
    private addChartToParts;
}
