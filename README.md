create xlsx with multi worksheets and charts

npm install office-chart

#example
const gen = new XlsxGenerator();
await gen.createWorkbook();
const sheet1 = await gen.createWorksheet("sheet1");
const sheet2 = await gen.createWorkshee("sheetWithChart2";
await sheet2.addTable([['h', 'b', 'c', 'd'], [1, 2, 3, 4], [4, 5, 6, 7]]);
await sheet2.addChart("B1:D3", 'testChart line', 'line')
const sheet3 = await gen.createWorksheet("sheet3");
await gen.generate(\_\_dirname + '/test', 'file');
