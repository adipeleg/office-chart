# Create xlsx with multi worksheets and charts

Node.js excel chart builder

## Quick start

Install

```bash
npm install office-chart
```

Generate and write chart to file

```js

const gen = new XlsxGenerator();

await gen.createWorkbook();

const sheet1 = await gen.createWorksheet("sheet1");

const sheet2 = await gen.createWorkshee("sheetWithChart2";

await sheet2.addTable([['h', 'b', 'c', 'd'], [1, 2, 3, 4], [4, 5, 6, 7]]);

await sheet2.addChart("B1:D3", 'testChart line', 'line');

const sheet3 = await gen.createWorksheet("sheet3");

await gen.generate(\_\_dirname + '/test', 'file');
// you can also generate buffer
```

#### This is an open source project, you can contribute by going to: https://github.com/adipeleg/office-chart

#### Enjoy:)