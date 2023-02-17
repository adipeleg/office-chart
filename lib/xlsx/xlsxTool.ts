import { XmlTool } from "../xmlTool";

export class XlsxTool {
    constructor(private xmlTool: XmlTool) { }

    public addSheetToWb = async (name: string) => {
        const wb = await this.xmlTool.readXml('xl/workbook.xml');
        let count: string;

        if (!Array.isArray(wb.workbook.sheets.sheet)) {
            count = '5';
            wb.workbook.sheets = {
                sheet: [
                    { '$': wb.workbook.sheets.sheet.$ },
                    { '$': { state: 'visible', name: 'Sheet2', sheetId: '2', 'r:id': 'rId' + count } }
                ]
            }

        } else {
            count = wb.workbook.sheets.sheet.length + 4;
            wb.workbook.sheets.sheet.push({
                '$': { state: 'visible', name: name, sheetId: count, 'r:id': 'rId' + count }
            });
        }
        await this.addSheetToParts(count);
        this.xmlTool.write('xl/workbook.xml', wb);
        return count;
    }

    public createSheet = async (id: string) => {
        const resSheet = await this.xmlTool.readXml('xl/worksheets/sheet1.xml');

        delete resSheet.worksheet.drawing

        const WbRel = await this.xmlTool.readXml('xl/_rels/workbook.xml.rels');
        WbRel.Relationships.Relationship.push({
            '$':
            {
                Id: 'rId' + id,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: `worksheets/sheet${id}.xml`
            }
        })


        this.xmlTool.write(`xl/worksheets/sheet${id}.xml`, resSheet);
        this.xmlTool.write('xl/_rels/workbook.xml.rels', WbRel);
        return resSheet;
    }

    public removeTemplateSheets = async () => {
        const wb = await this.xmlTool.readXml('xl/workbook.xml');

        wb.workbook.sheets.sheet = wb.workbook.sheets.sheet.filter(it => {
            return 'SheetTemplate' !== it.$.name.toString() &&
                'barTemplate' !== it.$.name.toString() &&
                'lineTemplate' !== it.$.name.toString() &&
                'pieTemplate' !== it.$.name.toString() &&
                'scatterTemplate' !== it.$.name.toString();
        })

        return this.xmlTool.write('xl/workbook.xml', wb);
    }

    public writeTable = async (sheet: any, data: any[][], id: string) => {
        const sheetWithTable = await this.xmlTool.readXml('xl/worksheets/sheet2.xml');
        const rowTemplate = sheetWithTable.worksheet.sheetData.row[0];
        const header = data.shift();

        const rows: any[] = [];
        rows.push(this.addRow(header, JSON.parse(JSON.stringify(rowTemplate)), 1));
        data.forEach((data, idx) => {
            rows.push(this.addRow(data, JSON.parse(JSON.stringify(rowTemplate)), idx + 2));
        })
        sheet.worksheet.sheetData = { row: rows };

        await this.addSharedStrings(data);

        return this.xmlTool.write(`xl/worksheets/sheet${id}.xml`, sheet);
    }

    private addRow(rowData: any[], rowTemplate: any, index: number) {
        rowTemplate.$.r = index;
        const cols: any[] = [];
        rowData.forEach((data, col) => {
            const type = typeof data === 'string' ? 's' : '';
            const c = { '$': { r: this.getColName(col) + (col <= 22 ? index : this.getColName(col)), s: '1' }, v: data }
            if (type === 's') {
                c.$['t'] = 'str';
            }
            cols.push(c)
        })

        rowTemplate.c = cols;
        return rowTemplate;
    }

    private addSharedStrings = async (data: any[][]) => {
        const str = await this.xmlTool.readXml('xl/sharedStrings.xml');
        data.forEach(row => {
            row.forEach(element => {
                if (typeof element === 'string') {
                    const inside = str['sst']['si'].find(it => {
                        return it.t === element;
                    })
                    if (!inside) {
                        str['sst']['si'].push({
                            t: element
                        })
                        str['sst'].$.uniqueCount++;
                    }
                }
            })
        })

        await this.xmlTool.write('xl/sharedStrings.xml', str);
    }

    private addSheetToParts = async (id: string) => {
        const parts = await this.xmlTool.readXml('[Content_Types].xml');

        parts['Types']['Override'].push({
            '$':
            {
                ContentType:
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                PartName: `/xl/worksheets/sheet${id}.xml`
            }
        })

        return this.xmlTool.write(`[Content_Types].xml`, parts);
    }

    private getColName = (n: number) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc[n] || abc[n % 26];
    }


    private ColToNum = (char: string) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc.indexOf(char);
    }

}