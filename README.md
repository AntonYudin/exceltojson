# exceltojson
Apache NIFI Excel to JSON and JSON to Excel processors

## ExcelToJson

Converts an excel file in XLSB or XLSX format to a json document.

An example of an output document:

```JSON
{
    "sheets": [
        {
            "name": "Sheet1",
            "content": {
                "rows": [
                    {
                    },
                    {
                        "A": "header A2",
                        "B": "header B2",
                        "C": "header C2"
                    },
                    {
                        "A": "header A3",
                        "B": "header B3",
                        "C": "header C3"
                    },
                    {
                        "A": "cell A4",
                        "B": "cell B4",
                        "C": "cell C4"
                    }
                ]
            }
        }
```

## JsonToExcel

Converts a json document to XLSX format.

An example of an input JSON document:

```JSON
{
    "sheets": [
		{
			"name": "sheet name",
			"color": "ffff00ff",
			"selected": false,
			"active": true,
			"images": [
				{
					"reference": "A8:G10",
					"url": "https://upload.wikimedia.org/wikipedia/commons/thumb/4/47/PNG_transparency_demonstration_1.png/330px-PNG_transparency_demonstration_1.png",
					"type": "png",
					"scale": 1.3
				}
			],
			"rows": [
				{
					"A": {
						"value": "Report",
						"columns": 3,
						"alignment": "center",
						"fontHeight": 32,
						"color": "ff0000ff",
						"fillColor": "ffffde21",
						"fontWeightBold": true
					},
					"D": "Test"
				},
				{
					"A": 5,
					"B": "This is text",
					"C": 123.456
				},
				{
					"A": true,
					"B": 8,
					"C": 789.012
				},
				{
					"A": {
						"value": "SUM(A1:A2)",
						"formula": true
					},
					"B": 9 ,
					"C": 345.678
				},
				{
					"A": {
						"value": "end",
						"alignment": "center",
						"fontHeight": 32
					}
				}
			]
		}
    ]
}
```

