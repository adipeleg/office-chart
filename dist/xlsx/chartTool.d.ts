import { IData } from "./models/data.model";
import { XmlTool } from "./xmlTool";
export declare class ChartTool {
    private xmlTool;
    private parts;
    constructor(xmlTool: XmlTool);
    addChart: (sheet: any, sheetName: string, opt: IData, id: string) => Promise<void>;
    private addDrawingRel;
    private addChartToDraw;
    private addChartToSheetRel;
    private addChartToSheet;
    private addChartToParts;
}
