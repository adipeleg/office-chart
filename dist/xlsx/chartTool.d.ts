import { IData, IPPTChartData } from "./models/data.model";
import { XmlTool } from "../xmlTool";
export declare class ChartTool {
    private xmlTool;
    private parts;
    constructor(xmlTool: XmlTool);
    addChart: (sheet: any, sheetName: string, opt: IData, id: string) => Promise<any>;
    private getChartNum;
    buildChart: (readChart: any, opt: IData | IPPTChartData, sheetName: string) => any;
    private buildCache;
    private addDrawingRel;
    private addChartToDraw;
    private addChartToSheetRel;
    private addChartToSheet;
    private addChartToParts;
}
