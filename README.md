<h1>Create xlsx with multi worksheets and charts</h1>
<!-- <h2></h2> -->

#how to download? <br/>
npm install office-chart

example </br>
const gen = new XlsxGenerator(); </br>

await gen.createWorkbook(); </br>
const sheet1 = await gen.createWorksheet("sheet1"); </br>
const sheet2 = await gen.createWorkshee("sheetWithChart2"; </br>
await sheet2.addTable([['h', 'b', 'c', 'd'], [1, 2, 3, 4], [4, 5, 6, 7]]);</br>
await sheet2.addChart("B1:D3", 'testChart line', 'line')</br>
const sheet3 = await gen.createWorksheet("sheet3");</br>
await gen.generate(\_\_dirname + '/test', 'file');
