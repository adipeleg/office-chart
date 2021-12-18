import fs from 'fs';
import { XmlTool } from '../../xmlTool';
xdescribe('check xmlTool', () => {

    it('', async () => {
        const tool = new XmlTool();
        await tool.readOriginal('xlsx');
        const res = await tool.readXml('xl/workbook.xml');

        res.workbook.sheets = {
            sheet: [
                {'$' :res.workbook.sheets.sheet.$},
                { '$': { state: 'visible', name: 'Sheet2', sheetId: '2', 'r:id': 'rId5' } }
            ]
        }

        const resSheet = await tool.readXml('xl/worksheets/sheet1.xml');

        console.log(resSheet.worksheet.drawing);
        delete resSheet.worksheet.drawing

        const WbRel = await tool.readXml('xl/_rels/workbook.xml.rels');
        WbRel.Relationships.Relationship.push({
            '$':
            {
                Id: 'rId5',
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: 'worksheets/sheet2.xml'
            }
        })

        tool.write('xl/workbook.xml', res);
        tool.write('xl/worksheets/sheet2.xml', resSheet);
        tool.write('xl/_rels/workbook.xml.rels', WbRel);
        const buf = await tool.generateBuffer();
        fs.writeFileSync('test2.xlsx', buf);
        console.log('done')
    })
    // it('', async () => {
    //     const tool = new XmlTool();
    //     await tool.readXlsx();
    //     const reswb = await tool.readXmlStr('xl/workbook.xml');
    //     const buf = await tool.generateBuffer();
    //     fs.writeFileSync('test.xlsx', buf);
    //     console.log("xlsx created.");
    // })

    // xit('2', async () => {
    //     const tool = new XmlTool();
    //     await tool.readXlsx();
    //     const reswb = await tool.readXmlStr('xl/workbook.xml');
    //     let sheets = reswb.split('<sheets>')[1].split('</sheets>')[0]
    //     sheets += '<sheet state="visible" name="Sheet2" sheetId="2" r:id="rId5"/>'
    //     const beforeSheets = reswb.split('<sheets>')[0];
    //     const afterSheets = reswb.split('</sheets>')[1];

    //     const wb = beforeSheets + sheets + afterSheets;

    //     console.log('wb after', wb)
    //     tool.writeStr('xl/workbook.xml', wb);

    //     const reswbrels = await tool.readXmlStr('xl/_rels/workbook.xml.rels');
    //     const relstr: string[] = reswbrels.split('</Relationships>');
    //     const newRels = relstr[0] + '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    //     + relstr[1];
    //     console.log('rel', reswbrels)
    //     tool.writeStr('xl/_rels/workbook.xml.rels', newRels);


    //     const sheet = await tool.readXmlStr('xl/worksheets/sheet1.xml');
    //     const sheetrels = await tool.readXmlStr('xl/worksheets/_rels/sheet1.xml.rels');

    //     tool.writeStr('xl/worksheets/sheet2.xml', sheet);

    //     console.log('sheet', sheet);
    //     console.log(sheetrels)


    //     const buf = await tool.generateBuffer();
    //     fs.writeFileSync('test3.xlsx', buf);
    //     console.log("xlsx created.");

    //     // tool.readXlsx('test.xlsx');

    // })
})