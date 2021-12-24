import { XmlTool } from "../xmlTool";

export class PptGraphicTool {
    constructor(private xmlTool: XmlTool) { }

    public writeTable = async (id: number, slide: any, data: any[][]) => {
        const slideWithTable = await this.xmlTool.readXml('ppt/slides/slide2.xml');

        const rowTemplate = slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tr'][1];
        const colTemplate = rowTemplate['a:tc'][1];


        const header = data.shift();

        const rows: any[] = [];
        rows.push(this.addRow(header, JSON.parse(JSON.stringify(rowTemplate)), colTemplate, 1));
        data.forEach((row, idx) => {
            rows.push(this.addRow(row, JSON.parse(JSON.stringify(rowTemplate)), colTemplate, idx + 2));
            // console.log(slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tblGrid']['a:gridCol']);
        })
        slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tr'] = rows;

        for (let i = 0; i < header.length - 4; i++) {
            slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame']['a:graphic']['a:graphicData']['a:tbl']['a:tblGrid']['a:gridCol'].push({ '$': { w: '2381250' } })
        }

        slide['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'] = slideWithTable['p:sld']['p:cSld']['p:spTree']['p:graphicFrame'];
        return this.xmlTool.write(`ppt/slides/slide${id}.xml`, slide);
    }

    private addRow(rowData: any[], rowTemplate: any, colTemplate: any, index: number) {
        const cols: any[] = [];
        rowData.forEach((data, col) => {
            colTemplate['a:txBody']['a:p']['a:r']['a:t'] = data
            cols.push(JSON.parse(JSON.stringify(colTemplate)));
        })

        rowTemplate['a:tc'] = JSON.parse(JSON.stringify(cols));
        return rowTemplate;
    }

    public addChart = () => {

    }

}