using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Converters;
using OpenXmlPowerTools.Documents;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlPowerTools2023.Tests.Converters
{
    [TestClass]
    public class FormattingAssemblerTests : ConverterTestsBase
    {
        protected override string FeatureFolder { get; } = @".\FormattingAssembler";

        [TestMethod]
        [DataRow("001-DeletedRun.docx")]
        [DataRow("002-DeletedNumberedParagraphs.docx")]
        [DataRow("003-DeletedFieldCode.docx")]
        [DataRow("004-InsertedNumberingProperties.docx")]
        [DataRow("005-InsertedNumberedParagraph.docx")]
        [DataRow("006-DeletedTableRow.docx")]
        [DataRow("007-InsertedTableRow.docx")]
        [DataRow("008-InsertedFieldCode.docx")]
        [DataRow("009-InsertedParagraph.docx")]
        [DataRow("010-InsertedRun.docx")]
        [DataRow("011-InsertedMathChar.docx")]
        [DataRow("012-DeletedMathChar.docx")]
        [DataRow("013-DeletedParagraph.docx")]
        [DataRow("014-MovedParagraph.docx")]
        [DataRow("015-InsertedContentControl.docx")]
        [DataRow("016-DeletedContentControl.docx")]
        [DataRow("017-NumberingChange.docx")]
        [DataRow("018-ParagraphPropertiesChange.docx")]
        [DataRow("019-RunPropertiesChange.docx")]
        [DataRow("020-SectionPropertiesChange.docx")]
        [DataRow("021-TableGridChange.docx")]
        [DataRow("022-TablePropertiesChange.docx")]
        [DataRow("023-CellPropertiesChange.docx")]
        [DataRow("024-RowPropertiesChange.docx")]

        public void FA001_DocumentsWithRevTracking(string src)
        {
            var source = GetFile(src);
            WmlDocument wmlSourceDocument = new WmlDocument(source);

            var acceptedOutput = source.ToLower().Replace(".docx", "-accepted.docx");
            var wmlSourceAccepted = RevisionProcessor.AcceptRevisions(wmlSourceDocument);
            wmlSourceAccepted.SaveAs(acceptedOutput);

            var target = GetFile("Output.docx");
            FormattingAssemblerSettings settings = new FormattingAssemblerSettings();
            var assembledWml = FormattingAssembler.AssembleFormatting(wmlSourceDocument, settings);
            assembledWml.SaveAs(target);

            var outAcceptedFi = GetFile("Output-accepted.docx");
            var assembledAcceptedWml = RevisionProcessor.AcceptRevisions(assembledWml);
            assembledAcceptedWml.SaveAs(outAcceptedFi);

            ValidateAgainstExpected(target);
        }
    }
}