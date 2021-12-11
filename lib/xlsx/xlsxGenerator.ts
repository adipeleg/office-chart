import { XmlTool } from './xmlTool';
export class XlsxGenerator {
    private workbook;
    private xmlTool: XmlTool = new XmlTool();

    public createWorkbook = async () => {
        this.workbook = await this.xmlTool.readXlsx()
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
            addChart: (range: string, data: any[][], title: string) => this.xmlTool.addChart(sheet, name, title, data, range, id)
        }
    }

    public generate = async (file: string) => {
        await this.xmlTool.removeTemplateSheets();
        await this.xmlTool.generateFile(file);
    }

}