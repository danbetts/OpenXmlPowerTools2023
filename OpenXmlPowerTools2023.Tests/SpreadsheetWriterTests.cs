// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Spreadsheets;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class SwTests
    {
        [Fact]
        public void SW001_Simple()
        {
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
            var outXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "SW001-Simple.xlsx"));
            SpreadsheetBuilder.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        [Fact]
        public void SW002_AllDataTypes()
        {
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
                                        CellDataType = CellDataType.String,
                                        Value = "date (t:d, mm-dd-yy)",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff"),
                                        FormatCode = "mm-dd-yy",
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "date (t:d, d-mmm-yy)",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff"),
                                        FormatCode = "d-mmm-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                }
                            },
                            new SpreadsheetRow
                            {
                                Cells = new SpreadsheetCell[]
                                {
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.String,
                                        Value = "date (t:d)",
                                    },
                                    new SpreadsheetCell {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 11).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff"),
                                    },
                                }
                            },
                        }
                    }
                }
            };
            var outXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "SW002-DataTypes.xlsx"));
            SpreadsheetBuilder.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        private void Validate(FileInfo fi)
        {
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(fi.FullName, true))
            {
                OpenXmlValidator v = new OpenXmlValidator();
                var errors = v.Validate(sDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

#if false
                // if a test fails validation post-processing, then can use this code to determine the SDK
                // validation error(s).

                if (errors.Count() != 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var item in errors)
                    {
                        sb.Append(item.Description).Append(Environment.NewLine);
                    }
                    var s = sb.ToString();
                    Console.WriteLine(s);
                }
#endif

                Assert.Empty(errors);
            }
        }

        private static List<string> s_ExpectedErrors = new List<string>()
        {
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        };
    }
}

#endif
