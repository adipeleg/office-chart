export class Chart {
    public getChart(sheetName, titles: string[], chartN, row, fields: string[], data: any, chart: string, o) {
        var ser = {};
        titles.forEach((t, i) => {
            // var chart = me.data[t].chart || me.chart;
            var r = {
                "c:idx": {
                    $: {
                        val: i
                    }
                },
                "c:order": {
                    $: {
                        val: i
                    }
                },
                "c:tx": {
                    "c:strRef": {
                        "c:f": sheetName + "!$" + this.getColName(i + 2) + "$" + row,
                        "c:strCache": {
                            "c:ptCount": {
                                $: {
                                    val: 1
                                }
                            },
                            "c:pt": {
                                $: {
                                    idx: 0
                                },
                                "c:v": t
                            }
                        }
                    }
                },
                "c:cat": {
                    "c:strRef": {
                        "c:f": sheetName + "!$A$" + (row + 1) + ":$A$" + (fields.length + row),
                        "c:strCache": {
                            "c:ptCount": {
                                $: {
                                    val: fields.length
                                }
                            },
                            "c:pt": fields.map((f, j) => {
                                return {
                                    $: {
                                        idx: j
                                    },
                                    "c:v": f
                                };
                            })
                        }
                    }
                },
                "c:val": {
                    "c:numRef": {
                        "c:f": sheetName + "!$" + this.getColName(i + 2) + "$" + (row + 1) + ":$" + this.getColName(i + 2) + "$" + (fields.length + row),
                        "c:numCache": {
                            "c:formatCode": "General",
                            "c:ptCount": {
                                $: {
                                    val: fields.length
                                }
                            },
                            "c:pt": fields.map((f, j) => {
                                return {
                                    $: {
                                        idx: j
                                    },
                                    "c:v": data[t][f]
                                };
                            })
                        }
                    }
                }
            };
            if (chart == "scatter") {
                r["c:xVal"] = r["c:cat"];
                delete r["c:cat"];
                r["c:yVal"] = r["c:val"];
                delete r["c:val"];
                r["c:spPr"] = {
                    "a:ln": {
                        $: {
                            w: 28575
                        },
                        "a:noFill": ""
                    }
                };
            };
            // ser[chart] = ser[chart] || [];
            ser = r;
        });
        // _.each(ser, function (ser, chart) {
        if (chart == "column") {
            // if (me.tplName == "charts") {
            //     o["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][0]["c:ser"] = ser;
            // } else {
            o["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"] = ser;
            // };
        } else
            if (chart == "bar") {
                // if (me.tplName == "charts") {
                //     o["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"][1]["c:ser"] = ser;
                // } else {
                o["c:chartSpace"]["c:chart"]["c:plotArea"]["c:barChart"]["c:ser"] = ser;
                // };
            } else {
                o["c:chartSpace"]["c:chart"]["c:plotArea"]["c:" + chart + "Chart"]["c:ser"] = ser;
            };
        // });

        return o;
        // me.removeUnusedCharts(o);

        // if (me.chartTitle) {
        //     me.writeTitle(o, me.chartTitle);
        // };
    }

    public getColName = (n: number) => {
        var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return abc[n] || abc[(n / 26 - 1) | 0] + abc[n % 26];
    }

}