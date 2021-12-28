import { ChartTool } from './../xlsx/chartTool';
import { XlsxGenerator } from './../xlsx/xlsxGenerator';
import { IPPTChartData, IPptTableOpt } from './../xlsx/models/data.model';
import { XmlTool } from "../xmlTool";
export declare class PptGraphicTool {
    private xmlTool;
    private xlsxGenerator;
    private chartTool;
    constructor(xmlTool: XmlTool, xlsxGenerator: XlsxGenerator, chartTool: ChartTool);
    writeTable: (id: number, slide: any, data: any[][], opt: IPptTableOpt) => Promise<void>;
    private addRow;
    private addLocationGraphicElements;
    addChart: (slide: any, chartOpt: IPPTChartData, slideId: number) => Promise<void>;
    private buildData;
    private buildChart;
    private addContentTypeChart;
    private addSlideChartRel;
    private addChartRef;
    private createXlsxWithTableAndChart;
    private getColName;
}
