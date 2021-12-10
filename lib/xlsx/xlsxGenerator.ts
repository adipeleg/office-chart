import { XmlTool } from './xmlTool';
export class XlsxGenerator {
    private workbook;
    private xmlTool: XmlTool = new XmlTool();

    public createWorkbook = async () => {
        this.workbook = await this.xmlTool.readXlsx()
    }

    public createWorksheet = async (name: string) => {
        // const sheet1 = this.xmlTool.readXml('xl/worksheets/sheet1.xml');
        // this.xmlTool.write('sheet2', sheet1);
        const id = await this.xmlTool.addSheetToWb(name);
        await this.xmlTool.createSheet(name, id);
    }

    public generate = async (file: string) => {
        await this.xmlTool.generateFile(file);
        // this.createWorksheet
    }


}