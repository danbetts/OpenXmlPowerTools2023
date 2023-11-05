using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Spreadsheets;
using System;
using System.Collections.Generic;
using System.IO;

namespace OpenXmlPowerTools2023.Tests.Spreadsheets
{
    [TestClass]
    public class SpreadsheetWriterTests : SpreadsheetTestsBase
    {
        [TestMethod]
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
                            Cell(CellDataType.String, "Name", bold: true),
                            Cell(CellDataType.String, "Age", bold: true, hAlign: HorizontalCellAlignment.Left),
                            Cell(CellDataType.String, "Rate", bold: true, hAlign: HorizontalCellAlignment.Left)
                        },
                        Rows = new SpreadsheetRow[]
                        {
                            Row(Cell(CellDataType.String, "Eric"),
                                Cell(CellDataType.Number, 50),
                                Cell(CellDataType.Number, 45.00d, formatCode: "0.00")),
                            Row(Cell(CellDataType.String, "Bob"),
                                Cell(CellDataType.Number, 50),
                                Cell(CellDataType.Number, 78.00d, formatCode: "0.00"))
                        }
                    }
                }
            };

            SpreadsheetBuilder.Write(OutputFile, wb);
            ValidateSpreadsheet(OutputFile);
        }

        [TestMethod]
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
                            new SpreadsheetCell { Value = "DataType", Bold = true },
                            new SpreadsheetCell { Value = "Value", Bold = true, HorizontalCellAlignment = HorizontalCellAlignment.Right },
                        },
                        Rows = new SpreadsheetRow[]
                        {
                            Row(Cell(CellDataType.String, "Boolean"),
                                Cell(CellDataType.Number, true)),
                            Row(Cell(CellDataType.String, "Boolean"),
                                Cell(CellDataType.Boolean, false)),
                            Row(Cell(CellDataType.String, "String"),
                                Cell(CellDataType.String, "A String", hAlign: HorizontalCellAlignment.Right)),
                            Row(Cell(CellDataType.String, "int"),
                                Cell(CellDataType.Number, (int)100)),
                            Row(Cell(CellDataType.String, "int?"),
                                Cell(CellDataType.Number, (int?)100)),
                            Row(Cell(CellDataType.String, "int? (is null)"),
                                Cell(CellDataType.Number, null)),
                            Row(Cell(CellDataType.String, "uint"),
                                Cell(CellDataType.Number, (uint)101)),
                            Row(Cell(CellDataType.String, "long"),
                                Cell(CellDataType.Number, Int64.MaxValue)),
                            Row(Cell(CellDataType.String, "float"),
                                Cell(CellDataType.Number, 123.45d)),
                            Row(Cell(CellDataType.String, "double"),
                                Cell(CellDataType.Number, 123.45d)),
                            Row(Cell(CellDataType.String, "decimal"),
                                Cell(CellDataType.Number, 123.45d)),
                            Row(Cell(CellDataType.String, "date (t:d, mm-dd-yy)"),
                                Cell(CellDataType.Date, new DateTime(2012, 1, 8).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff"), formatCode: "mm-dd-yy" )),
                            Row(Cell(CellDataType.String, "date (t:d, d-mmm-yy)"),
                                Cell(CellDataType.Date, new DateTime(2012, 1, 9).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff"), formatCode: "mm-dd-yy", hAlign: HorizontalCellAlignment.Center)),
                            Row(Cell(CellDataType.String, "date (t:d)"),
                                Cell(CellDataType.Date, new DateTime(2012, 1, 11).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff"))),
                        }
                    }
                }
            };
            string testFile = ToTempPath("SW002-DataTypes.xlsx");
            SpreadsheetBuilder.Write(testFile, wb);
            ValidateSpreadsheet(testFile);
        }
    }
}