import { ChartTool } from './chartTool';
import { XmlTool } from './xmlTool';
export class XlsxGenerator {
    private xmlTool: XmlTool = new XmlTool();
    private chartTool: ChartTool = new ChartTool(this.xmlTool);

    public createWorkbook = async () => {
        return this.xmlTool.readXlsx()
    }

    public createWorksheet = async (name: string) => {
        const id = await this.xmlTool.addSheetToWb(name);
        const sheet = await this.xmlTool.createSheet(id);
        return {
            data: sheet,
            name: name,
            id: id,
            addTable: (data: any[][]) => {
                return this.xmlTool.writeTable(sheet, data, id)
            },
            addChart: async (range: string, title: string, type: 'line' | 'bar') => await this.chartTool.addChart(sheet, name, title, range, id, type)
        }
    }

    public generate = async (file: string, type: 'file' | 'buffer') => {
        await this.xmlTool.removeTemplateSheets();
        if (type === 'file') {
            return this.xmlTool.generateFile(file);
        } else {
            return this.xmlTool.generateBuffer();
        }
    }

}