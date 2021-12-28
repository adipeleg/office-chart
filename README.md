# Create XLSX and PPTX with multi worksheets and charts

#

Node.js excel chart builder

## Quick start

Install

```bash
npm install office-chart
```

Generate xlsx and write chart to file

```js
import { IData } from "office-chart";
import { XlsxGenerator } from "office-chart";

const gen = new XlsxGenerator();

await gen.createWorkbook();

const sheet1 = await gen.createWorksheet("sheet1");

const sheet2 = await gen.createWorksheet("sheetWithChart2");

const header = ["h", "b", "c", "d"];
const row1 = ["label1", 2, 3, 4];
const row2 = ["label2", 5, 6, 7];

await sheet2.addTable([header, row1, row2]);

const opt: IData = {
  title: {
    name: "testChart line",
    color: "8ab4f8",
    size: 5000,
  },
  range: "B1:D3",
  type: "line",
  rgbColors: ["8ab4f8", "ff7769"],
  labels: true, //table contains labels
  marker: {
    size: 4,
    shape: "square", //marker shapes, can be circle, diamond, star
  },
  lineWidth: 20000,
};

await sheet2.addChart(opt);

const sheet3 = await gen.createWorksheet("sheet3");

await gen.generate(__dirname + "/test", "file");
// you can also generate buffer
```

#

Generate ppt with slides and text

#

```js
import { PptxGenetator } from "office-chart";
import { IPPTChartData } from "office-chart/dist/xlsx/models/data.model";

const gen = new PptxGenetator();

await gen.createPresentation();

const slide = await gen.createSlide();

slide.addTitle("this is title", {
  x: "0",
  y: "0",
  color: "FF0000",
  size: 4000,
});

slide.addSubTitle("this is subtitle");

slide.addText("this is text", {
  color: "FF0000",
});

const header = ["h", "b", "c", "d"];
const row1 = ["label1", 2, 3, 4];
const row2 = ["label2", 5, 6, 7];

await slide.addTable([header, row1, row2], {
  x: "1000", // left
  y: "1000", // top
  colWidth: 1081250,
  rowHeight: 1059279,
});

const opt: IPPTChartData = {
  title: {
    name: "testChart line",
    color: "8ab4f8",
    size: 3000,
  },
  type: "line",
  data: [header, row1, row2], // can also be:
  // [
  //       {
  //           name: 'lab1 test', //label
  //           values: [1, 2, 3, 4, 5], //yvalues
  //           labels: ['h', 'b', 'c', 'd', 'e'] //xvalue
  //       }, {
  //           name: 'lab2 test',
  //           values: [4, 5, 6, 7, 8],
  //           labels: ['h', 'b', 'c', 'd', 'e']
  //       }, {
  //           name: 'lab3 test',
  //           values: [9, 1, 2, 4, 10],
  //           labels: ['h', 'b', 'c', 'd', 'e']
  //       }
  //   ]
  rgbColors: ["8ab4f8", "ff7769"],
  lineWidth: 20000,
  marker: {
    shape: "circle",
    size: 4,
  },
  labels: true,
};

const slide2 = await gen.createSlide();

await slide2.addChart(opt);
await slide2.addTitle(null); //remove title
await slide2.addSubTitle(null); //remove subtitle

await gen.generate(__dirname + "/test2", "file");
```

#### This is an open source project, you can contribute by going to: https://github.com/adipeleg/office-chart.

#

#### currently only column, line, pie and scatter charts are supported in Xlsx.

#### currently only column, line chart are supported in PPTX.

#

#### Enjoy and don't forget to add a star :)
