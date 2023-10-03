// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Spreadsheets;

namespace SpreadsheetWriterExample
{
    class Program
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            SpreadsheetWorkbook wb = new SpreadsheetWorkbook
            {
                Worksheets = new SpreadsheetWorksheet[]
                {
                    new SpreadsheetWorksheet
                    {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
                        ColumnHeadings = new SpreadsheetCell[]
                        {
                            new SpreadsheetCell
                            {
                                Value = "Name",
                                Bold = true,
                            },
                            new SpreadsheetCell
                            {
                                Value = "Age",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Left,
                            },
                            new SpreadsheetCell
                            {
                                Value = "Rate",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Left,
                            }
                        },
                        Rows = new SpreadsheetRow[]
                        {
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "Eric",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = 50,
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)45.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "Bob",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = 42,
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)78.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetBuilder.Write(Path.Combine(tempDi.FullName, "Test1.xlsx"), wb);
        }
    }
}
