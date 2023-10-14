// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using OpenXmlPowerTools.Spreadsheets;
using System;
using System.IO;

namespace OpenXmlPowerTools.Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            Example01();
            Example02();
        }

        private static void Example01()
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
        private static void Example02()
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
                        ColumnHeadings = new SpreadsheetCell[]
                        {
                            new SpreadsheetCell
                            {
                                Value = "DataType",
                                Bold = true,
                            },
                            new SpreadsheetCell
                            {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Right,
                            },
                        },
                        Rows = new SpreadsheetRow[]
                        {
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Boolean,
                                        Value = true,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Boolean,
                                        Value = false,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "String",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "int",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (int)100,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "int?",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (int?)100,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "int? (is null)",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = null,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "uint",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (uint)101,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "long",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = Int64.MaxValue,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "float",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (float)123.45,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "double",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (double)123.45,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)123.45,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetBuilder.Write(Path.Combine(tempDi.FullName, "Test2.xlsx"), wb);
        }
    }
}
