using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Spreadsheets;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlPowerTools2023.Tests.Spreadsheets
{
    [TestClass]
    [DeploymentItem(TestResourcePath, OutputPath)]
    public class SpreadsheetTestsBase : TestsBase
    {
        protected override string Extension { get; } = ".xlsx";
        protected override string ModuleFolder { get; } = @".\Spreadsheets";
        protected override string OutputFile { get; set; } = "Output.xlsx";
        protected override string FeatureFolder { get; } = "";


        protected void ValidateDocument(string filepath)
        {
            // Act
            CleanupTest(filepath);
            //Builder.SaveAs(TargetPath);

            IEnumerable<ValidationErrorInfo> validationErrors = new List<ValidationErrorInfo>();
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filepath, true))
            {
                validationErrors = Validator.Validate(wDoc).ToList();
            }

            // Assert
            validationErrors.Count().Should().Be(0);
        }

        protected void ValidateSpreadsheet(string fileName)
        {
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(fileName, true))
            {
                var errors = Validator.Validate(sDoc).Where(ve => ExpectedErrors.Contains(ve.Description));

                errors.Should().BeEmpty();
            }
        }
        protected SpreadsheetCell Cell(CellDataType dataType, object value, string formatCode = null, bool? bold = null, HorizontalCellAlignment? hAlign = null)
        {
            var cell = new SpreadsheetCell { CellDataType = dataType, Value = value };
            if (formatCode != null) cell.FormatCode = formatCode;
            if (bold != null) cell.Bold = bold;
            if (hAlign != null) cell.HorizontalCellAlignment = hAlign;

            return cell;
        }

        protected SpreadsheetRow Row(SpreadsheetCell first, SpreadsheetCell second, SpreadsheetCell third = null)
        {
            var row = new SpreadsheetRow();
            var cells = new List<SpreadsheetCell> { first, second };

            if (third != null) cells.Add(third);

            row.Cells = cells.ToArray();
            return row;
        }
    }
}