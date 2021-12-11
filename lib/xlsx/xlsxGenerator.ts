import { XmlTool } from './xmlTool';
export class XlsxGenerator {
    private workbook;
    private xmlTool: XmlTool = new XmlTool();

    public createWorkbook = async () => {
        this.workbook = await this.xmlTool.readXlsx()
    }

    public createWorksheet = async (name: string) => {
        const id = await this.xmlTool.addSheetToWb(name);
        const sheet = await this.xmlTool.createSheet(name, id);
        return {
            data: sheet,
            name: name,
            addTable: (data: any[][]) =>{
                return this.xmlTool.writeTable(sheet, name, data)
            },
            addChart: (range: string, data: any[][]) => this.xmlTool.addChart(sheet, name, data, range)
        }
    }

    public generate = async (file: string) => {
        await this.xmlTool.generateFile(file);
    }

}