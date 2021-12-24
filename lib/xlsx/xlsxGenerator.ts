import { XlsxTool } from './xlsxTool';
import { XmlTool } from '../xmlTool';
import { ChartTool } from './chartTool';
import { IData } from './models/data.model';
export class XlsxGenerator {
    private xmlTool: XmlTool = new XmlTool();
    private chartTool: ChartTool = new ChartTool(this.xmlTool);
    private xlsxTool: XlsxTool = new XlsxTool(this.xmlTool);

    public createWorkbook = async () => {
        return this.xmlTool.readOriginal('xlsx')
    }

    public createWorksheet = async (name: string) => {
        const id = await this.xlsxTool.addSheetToWb(name);
        const sheet = await this.xlsxTool.createSheet(id);
        return {
            data: sheet,
            name: name,
            id: id,
            addTable: (data: any[][]) => {
                return this.xlsxTool.writeTable(sheet, data, id)
            },
            addChart: async (opt: IData) => { return this.chartTool.addChart(sheet, name, opt, id) }
        }
    }

    public generate = async (file: string, type: 'file' | 'buffer') => {
        await this.xlsxTool.removeTemplateSheets();
        if (type === 'file') {
            return this.xmlTool.generateFile(file, 'xlsx');
        } else {
            return this.xmlTool.generateBuffer();
        }
    }

}