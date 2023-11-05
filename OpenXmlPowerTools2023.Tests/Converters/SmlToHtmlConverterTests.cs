using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Spreadsheets;
using OpenXmlPowerTools2023.Tests.Spreadsheets;
using System.IO;

namespace OpenXmlPowerTools2023.Tests.Converters
{
    [TestClass]
    public class SmlToHtmlConverterTests : SpreadsheetTestsBase
    {
        private string SheetName = "Sheet1";

        [TestMethod]
        [DataRow("SH101-SimpleFormats.xlsx")]
        [DataRow("SH102-9-x-9.xlsx")]
        [DataRow("SH103-No-SharedString.xlsx")]
        [DataRow("SH104-With-SharedString.xlsx")]
        [DataRow("SH105-No-SharedString.xlsx")]
        [DataRow("SH106-9-x-9-Formatted.xlsx")]
        [DataRow("SH108-SimpleFormattedCell.xlsx")]
        [DataRow("SH109-CellWithBorder.xlsx")]
        [DataRow("SH110-CellWithMasterStyle.xlsx")]
        [DataRow("SH111-ChangedDefaultColumnWidth.xlsx")]
        [DataRow("SH112-NotVertMergedCell.xlsx")]
        [DataRow("SH113-VertMergedCell.xlsx")]
        [DataRow("SH114-Centered-Cell.xlsx")]
        [DataRow("SH115-DigitsToRight.xlsx")]
        [DataRow("SH116-FmtNumId-1.xlsx")]
        [DataRow("SH117-FmtNumId-2.xlsx")]
        [DataRow("SH118-FmtNumId-3.xlsx")]
        [DataRow("SH119-FmtNumId-4.xlsx")]
        [DataRow("SH120-FmtNumId-9.xlsx")]
        [DataRow("SH121-FmtNumId-11.xlsx")]
        [DataRow("SH122-FmtNumId-12.xlsx")]
        [DataRow("SH123-FmtNumId-14.xlsx")]
        [DataRow("SH124-FmtNumId-15.xlsx")]
        [DataRow("SH125-FmtNumId-16.xlsx")]
        [DataRow("SH126-FmtNumId-17.xlsx")]
        [DataRow("SH127-FmtNumId-18.xlsx")]
        [DataRow("SH128-FmtNumId-19.xlsx")]
        [DataRow("SH129-FmtNumId-20.xlsx")]
        [DataRow("SH130-FmtNumId-21.xlsx")]
        [DataRow("SH131-FmtNumId-22.xlsx")]

        public void SH005_ConvertSheet(string name)
        {
            var source = GetFile(name);

            var dataTemplateFileNameSuffix = "-2-Generated-XmlData-Entire-Sheet.xml";
            var dataXmlFi = source.Replace(".xlsx", dataTemplateFileNameSuffix);
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(source, true))
            {
                var settings = new SmlToHtmlConverterSettings();
                var rangeXml = SmlDataRetriever.RetrieveSheet(sDoc, SheetName);
                rangeXml.Save(dataXmlFi);
            }
        }

        [TestMethod]
        [DataRow("SH101-SimpleFormats.xlsx", "A1:B10")]
        [DataRow("SH101-SimpleFormats.xlsx", "A4:B8")]
        [DataRow("SH102-9-x-9.xlsx", "A1:A1")]
        [DataRow("SH102-9-x-9.xlsx", "C2:C2")]
        [DataRow("SH102-9-x-9.xlsx", "A9:A9")]
        [DataRow("SH102-9-x-9.xlsx", "I1:I1")]
        [DataRow("SH102-9-x-9.xlsx", "I9:I9")]
        [DataRow("SH102-9-x-9.xlsx", "A1:I9")]
        [DataRow("SH102-9-x-9.xlsx", "A2:D4")]
        [DataRow("SH102-9-x-9.xlsx", "C5:G7")]
        [DataRow("SH103-No-SharedString.xlsx", "A1:A1")]
        [DataRow("SH104-With-SharedString.xlsx", "A4:A7")]
        [DataRow("SH105-No-SharedString.xlsx", "A4:A7")]
        [DataRow("SH106-9-x-9-Formatted.xlsx", "A1:I9")]
        [DataRow("SH108-SimpleFormattedCell.xlsx", "A1:A1")]
        [DataRow("SH109-CellWithBorder.xlsx", "A1:A1")]
        [DataRow("SH110-CellWithMasterStyle.xlsx", "A1:A1")]
        [DataRow("SH111-ChangedDefaultColumnWidth.xlsx", "A1:A1")]
        [DataRow("SH112-NotVertMergedCell.xlsx", "A1:A1")]
        [DataRow("SH113-VertMergedCell.xlsx", "A1:A1")]
        [DataRow("SH114-Centered-Cell.xlsx", "A1:A1")]
        [DataRow("SH115-DigitsToRight.xlsx", "A1:A10")]
        [DataRow("SH116-FmtNumId-1.xlsx", "A1:A10")]
        [DataRow("SH117-FmtNumId-2.xlsx", "A1:A10")]
        [DataRow("SH118-FmtNumId-3.xlsx", "A1:A10")]
        [DataRow("SH119-FmtNumId-4.xlsx", "A1:A10")]
        [DataRow("SH120-FmtNumId-9.xlsx", "A1:A10")]
        [DataRow("SH121-FmtNumId-11.xlsx", "A1:A10")]
        [DataRow("SH122-FmtNumId-12.xlsx", "A1:A10")]
        [DataRow("SH123-FmtNumId-14.xlsx", "A1:A10")]
        [DataRow("SH124-FmtNumId-15.xlsx", "A1:A10")]
        [DataRow("SH125-FmtNumId-16.xlsx", "A1:A10")]
        [DataRow("SH126-FmtNumId-17.xlsx", "A1:A10")]
        [DataRow("SH127-FmtNumId-18.xlsx", "A1:A10")]
        [DataRow("SH128-FmtNumId-19.xlsx", "A1:A10")]
        [DataRow("SH129-FmtNumId-20.xlsx", "A1:A10")]
        [DataRow("SH130-FmtNumId-21.xlsx", "A1:A10")]
        [DataRow("SH131-FmtNumId-22.xlsx", "A1:A10")]

        public void SH004_ConvertRange(string name, string range)
        {
            var source = GetFile(name);

            var dataTemplateFileNameSuffix = string.Format("-2-Generated-XmlData-{0}.xml", range.Replace(":", ""));
            var dataXmlFi = source.Replace(".xlsx", dataTemplateFileNameSuffix);
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(source, true))
            {
                var settings = new SmlToHtmlConverterSettings();
                var rangeXml = SmlDataRetriever.RetrieveRange(sDoc, SheetName, range);
                rangeXml.Save(dataXmlFi);
            }
        }


        [TestMethod]
        [DataRow("SH001-Table.xlsx", "MyTable")]
        [DataRow("SH003-TableWithDateInFirstColumn.xlsx", "MyTable")]
        [DataRow("SH004-TableAtOffsetLocation.xlsx", "MyTable")]
        [DataRow("SH005-Table-With-SharedStrings.xlsx", "Table1")]
        [DataRow("SH006-Table-No-SharedStrings.xlsx", "Table1")]
        [DataRow("SH107-9-x-9-Formatted-Table.xlsx", "Table1")]
        [DataRow("SH007-One-Cell-Table.xlsx", "Table1")]
        [DataRow("SH008-Table-With-Tall-Row.xlsx", "Table1")]
        [DataRow("SH009-Table-With-Wide-Column.xlsx", "Table1")]

        public void SH003_ConvertTable(string name, string tableName)
        {
            var source = GetFile(name);

            var dataXmlFi = source.Replace(".xlsx", "-2-Generated-XmlData.xml");
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(source, true))
            {
                var settings = new SmlToHtmlConverterSettings();
                var rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
                rangeXml.Save(dataXmlFi);
            }
        }

        [TestMethod]
        [DataRow("Spreadsheet.xlsx", 2)]
        public void SH002_SheetNames(string name, int numberOfSheets)
        {
            var source = GetFile(name);
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(source, true))
            {
                var SheetNames = SmlDataRetriever.SheetNames(sDoc);
                SheetNames.Length.Should().Be(numberOfSheets);
            }
        }

        [TestMethod]
        [DataRow("SH001-Table.xlsx", 1)]
        [DataRow("SH002-TwoTablesTwoSheets.xlsx", 2)]
        public void SH001_TableNames(string name, int numberOfTables)
        {
            var source = GetFile(name);
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(source, true))
            {
                var table = SmlDataRetriever.TableNames(sDoc);
                table.Length.Should().Be(numberOfTables);
            }
        }
    }
}