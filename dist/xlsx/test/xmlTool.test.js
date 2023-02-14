"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const fs_1 = __importDefault(require("fs"));
const xmlTool_1 = require("../../xmlTool");
describe('check xmlTool', () => {
    it('', () => __awaiter(void 0, void 0, void 0, function* () {
        const tool = new xmlTool_1.XmlTool();
        yield tool.readOriginal('xlsx');
        const res = yield tool.readXml('xl/workbook.xml');
        res.workbook.sheets = {
            sheet: [
                { '$': res.workbook.sheets.sheet.$ },
                { '$': { state: 'visible', name: 'Sheet2', sheetId: '2', 'r:id': 'rId5' } }
            ]
        };
        const resSheet = yield tool.readXml('xl/worksheets/sheet1.xml');
        delete resSheet.worksheet.drawing;
        const WbRel = yield tool.readXml('xl/_rels/workbook.xml.rels');
        WbRel.Relationships.Relationship.push({
            '$': {
                Id: 'rId5',
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: 'worksheets/sheet2.xml'
            }
        });
        tool.write('xl/workbook.xml', res);
        tool.write('xl/worksheets/sheet2.xml', resSheet);
        tool.write('xl/_rels/workbook.xml.rels', WbRel);
        const buf = yield tool.generateBuffer();
        fs_1.default.writeFileSync('test2.xlsx', buf);
        console.log('done');
    }));
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
});
