using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class DocumentAssemblerTests : DocumentTestsBase
    {
        protected override string FeatureFolder { get; } = @".\DocumentAssembler";

        [TestMethod]
        [DataRow("DA001-TemplateDocument.docx", "DA-Data.xml", false)]
        [DataRow("DA002-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [DataRow("DA003-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [DataRow("DA004-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [DataRow("DA005-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [DataRow("DA006-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [DataRow("DA007-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [DataRow("DA008-TableElementWithNoTable.docx", "DA-Data.xml", true)]
        [DataRow("DA009-InvalidXPath.docx", "DA-Data.xml", true)]
        [DataRow("DA010-InvalidXml.docx", "DA-Data.xml", true)]
        [DataRow("DA011-SchemaError.docx", "DA-Data.xml", true)]
        [DataRow("DA012-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [DataRow("DA013-Runs.docx", "DA-Data.xml", false)]
        [DataRow("DA014-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [DataRow("DA015-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [DataRow("DA016-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [DataRow("DA017-FiveRuns.docx", "DA-Data.xml", true)]
        [DataRow("DA018-SmartQuotes.docx", "DA-Data.xml", false)]
        [DataRow("DA019-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [DataRow("DA020-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [DataRow("DA021-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [DataRow("DA022-InvalidXPath.docx", "DA-Data.xml", true)]
        [DataRow("DA023-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [DataRow("DA026-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [DataRow("DA027-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [DataRow("DA028-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [DataRow("DA029-NoDataForCell.docx", "DA-Data.xml", true)]
        [DataRow("DA030-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", true)]
        [DataRow("DA031-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [DataRow("DA032-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [DataRow("DA033-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [DataRow("DA034-HeaderFooter.docx", "DA-Data.xml", false)]
        [DataRow("DA035-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [DataRow("DA036-SchemaErrorInConditional.docx", "DA-Data.xml", true)]

        [DataRow("DA100-TemplateDocument.docx", "DA-Data.xml", false)]
        [DataRow("DA101-TemplateDocument.docx", "DA-Data.xml", true)]
        [DataRow("DA102-TemplateDocument.docx", "DA-Data.xml", true)]

        [DataRow("DA201-TemplateDocument.docx", "DA-Data.xml", false)]
        [DataRow("DA202-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [DataRow("DA203-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [DataRow("DA204-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [DataRow("DA205-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [DataRow("DA206-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [DataRow("DA207-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [DataRow("DA209-InvalidXPath.docx", "DA-Data.xml", true)]
        [DataRow("DA210-InvalidXml.docx", "DA-Data.xml", true)]
        [DataRow("DA211-SchemaError.docx", "DA-Data.xml", true)]
        [DataRow("DA212-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [DataRow("DA213-Runs.docx", "DA-Data.xml", false)]
        [DataRow("DA214-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [DataRow("DA215-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [DataRow("DA216-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [DataRow("DA217-FiveRuns.docx", "DA-Data.xml", true)]
        [DataRow("DA218-SmartQuotes.docx", "DA-Data.xml", false)]
        [DataRow("DA219-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [DataRow("DA220-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [DataRow("DA221-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [DataRow("DA222-InvalidXPath.docx", "DA-Data.xml", true)]
        [DataRow("DA223-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [DataRow("DA226-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [DataRow("DA227-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [DataRow("DA228-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [DataRow("DA229-NoDataForCell.docx", "DA-Data.xml", true)]
        [DataRow("DA230-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", true)]
        [DataRow("DA231-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [DataRow("DA232-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [DataRow("DA233-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [DataRow("DA234-HeaderFooter.docx", "DA-Data.xml", false)]
        [DataRow("DA235-Crashes.docx", "DA-Content-List.xml", false)]
        [DataRow("DA236-Page-Num-in-Footer.docx", "DA-Content-List.xml", false)]
        [DataRow("DA237-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [DataRow("DA238-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
        [DataRow("DA239-RunLevelCC-Repeat.docx", "DA-Data.xml", false)]

        [DataRow("DA250-ConditionalWithRichXPath.docx", "DA250-Address.xml", false)]
        [DataRow("DA251-EnhancedTables.docx", "DA-Data.xml", false)]
        [DataRow("DA252-Table-With-Sum.docx", "DA-Data.xml", false)]
        [DataRow("DA253-Table-With-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [DataRow("DA254-Table-With-XPath-Sum.docx", "DA-Data.xml", false)]
        [DataRow("DA255-Table-With-XPath-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [DataRow("DA256-NoInvalidDocOnErrorInRun.docx", "DA-Data.xml", true)]
        [DataRow("DA257-OptionalRepeat.docx", "DA-Data.xml", false)]
        [DataRow("DA258-ContentAcceptsCharsAsXPathResult.docx", "DA-Data.xml", false)]
        [DataRow("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        [DataRow("DA260-RunLevelRepeat.docx", "DA-Data.xml", false)]
        [DataRow("DA261-RunLevelConditional.docx", "DA-Data.xml", false)]
        [DataRow("DA262-ConditionalNotMatch.docx", "DA-Data.xml", false)]
        [DataRow("DA263-ConditionalNotMatch.docx", "DA-DataSmallCustomer.xml", false)]
        [DataRow("DA264-InvalidRunLevelRepeat.docx", "DA-Data.xml", true)]
        [DataRow("DA265-RunLevelRepeatWithWhiteSpaceBefore.docx", "DA-Data.xml", false)]
        [DataRow("DA266-RunLevelRepeat-NoData.docx", "DA-Data.xml", true)]
        [DataRow("DA300-TableWithContentInCells.docx", "DA-Data.xml", false)]
        [DataRow("DA267-Repeat-HorizontalAlignType.docx", "DA-Data.xml", false)]
        [DataRow("DA268-Repeat-VerticalAlignType.docx", "DA-Data.xml", false)]
        [DataRow("DA269-Repeat-InvalidAlignType.docx", "DA-Data.xml", true)]
        [DataRow("DA270-ImageSelect.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA270A-ImageSelect.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA271-ImageSelectWithRepeat.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA271A-ImageSelectWithRepeat.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA272-ImageSelectWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA272A-ImageSelectWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA273-ImageSelectInsideTextBoxWithRepeatVerticalAlign.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA273A-ImageSelectInsideTextBoxWithRepeatVerticalAlign.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA274-ImageSelectInsideTextBoxWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA274A-ImageSelectInsideTextBoxWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA275-ImageSelectWithRepeatInvalidAlign.docx", "DA-Data-WithImages.xml", true)]
        [DataRow("DA275A-ImageSelectWithRepeatInvalidAlign.docx", "DA-Data-WithImages.xml", true)]
        [DataRow("DA276-ImageSelectInsideTable.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA276A-ImageSelectInsideTable.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA277-ImageSelectMissingOrInvalidPictureContent.docx", "DA-Data-WithImages.xml", true)]
        [DataRow("DA277A-ImageSelectMissingOrInvalidPictureContent.docx", "DA-Data-WithImages.xml", true)]
        [DataRow("DA278-ImageSelect.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [DataRow("DA278A-ImageSelect.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [DataRow("DA279-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidMIMEType.xml", true)]
        [DataRow("DA279A-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidMIMEType.xml", true)]
        [DataRow("DA280-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidImageDataFormat.xml", true)]
        [DataRow("DA280A-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidImageDataFormat.xml", true)]
        [DataRow("DA281-ImageSelectExtraWhitespaceBeforeImageContent.docx", "DA-Data-WithImages.xml", true)]
        [DataRow("DA281A-ImageSelectExtraWhitespaceBeforeImageContent.docx", "DA-Data-WithImages.xml", true)]
        [DataRow("DA282-ImageSelectWithHeader.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA282A-ImageSelectWithHeader.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA282-ImageSelectWithHeader.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [DataRow("DA282A-ImageSelectWithHeader.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [DataRow("DA283-ImageSelectWithFooter.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA283A-ImageSelectWithFooter.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA284-ImageSelectWithHeaderAndFooter.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA284A-ImageSelectWithHeaderAndFooter.docx", "DA-Data-WithImages.xml", false)]
        [DataRow("DA285-ImageSelectNoParagraphFollowedAfterMetadata.docx", "DA-Data-WithImages.xml", true)]
        [DataRow("DA285A-ImageSelectNoParagraphFollowedAfterMetadata.docx", "DA-Data-WithImages.xml", true)]

        public void DA101(string name, string data, bool err)
        {
            var templateDocx = GetFile(name);
            var dataFile = GetFile(data);

            WmlDocument wmlTemplate = new WmlDocument(templateDocx);
            XElement xmldata = XElement.Load(dataFile);

            bool returnedTemplateError;
            WmlDocument afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out returnedTemplateError);
            var assembledDocx = templateDocx.Replace(".docx", "-processed-by-DocumentAssembler.docx");
            afterAssembling.SaveAs(assembledDocx);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator v = new OpenXmlValidator();
                    var valErrors = v.Validate(wDoc).Where(ve => !ExpectedErrors.Contains(ve.Description));

                    valErrors.Should().BeEmpty();
                }
            }

            err.Equals(returnedTemplateError);
        }

        [TestMethod]
        [DataRow("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        public void DA259(string name, string data, bool err)
        {
            var source = GetFile(name);
            DA101(name, data, err);
            var assembledDocx = source.Replace(".docx", "-processed-by-DocumentAssembler.docx");
            WmlDocument afterAssembling = new WmlDocument(assembledDocx);
            int brCount = afterAssembling.MainDocumentPart
                            .Element(W.body)
                            .Elements(W.p).ElementAt(1)
                            .Elements(W.r)
                            .Elements(W.br).Count();

            brCount.Should().Be(4);
        }

        [TestMethod]
        public void DA240()
        {
            string name = "DA240-Whitespace.docx";
            string source = GetFile(name);
            DA101(name, "DA240-Whitespace.xml", false);
            var assembledDocx = source.Replace(".docx", "-processed-by-DocumentAssembler.docx");
            WmlDocument afterAssembling = new WmlDocument(assembledDocx);

            // when elements are inserted that begin or end with white space, make sure white space is preserved
            string firstParaTextIncorrect = afterAssembling.MainDocumentPart.Element(W.body).Elements(W.p).First().Value;
            firstParaTextIncorrect.Should().Be("Content may or may not have spaces: he/she; he, she; he and she.");
            // warning: XElement.Value returns the string resulting from direct concatenation of all W.t elements. This is fast but ignores
            // proper handling of xml:space="preserve" attributes, which Word honors when rendering content. Below we also check
            // the result of UnicodeMapper.RunToString, which has been enhanced to take xml:space="preserve" into account.
            string firstParaTextCorrect = InnerText(afterAssembling.MainDocumentPart.Element(W.body).Elements(W.p).First());
            firstParaTextCorrect.Should().Be("Content may or may not have spaces: he/she; he, she; he and she.");
        }

        [TestMethod]
        [DataRow("DA024-TrackedRevisions.docx", "DA-Data.xml")]
        [ExpectedException(typeof(OpenXmlPowerToolsException))]
        public void DA102_Throws(string name, string data)
        {
            var templateDocx = GetFile(name);
            var dataFile = GetFile(data);

            WmlDocument wmlTemplate = new WmlDocument(templateDocx);
            XElement xmldata = XElement.Load(dataFile);
            _ = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out bool returnedTemplateError);
        }

        [TestMethod]
        public void DATemplateMaior()
        {
            // this test case was causing incorrect behavior of OpenXmlRegex when replacing fields in paragraphs that contained
            // lastRenderedPageBreak XML elements. Recent fixes relating to UnicodeMapper and OpenXmlRegex addressed it.
            string name = "DA-TemplateMaior.docx";
            string source = GetFile(name);
            DA101(name, "DA-templateMaior.xml", false);
            var assembledDocx = source.Replace(".docx", "-processed-by-DocumentAssembler.docx");
            var afterAssembling = new WmlDocument(assembledDocx);

            var descendants = afterAssembling.MainDocumentPart.Value;
            descendants.Contains(">").Should().BeFalse(because: "Found > on text");
        }

        [TestMethod]
        public void DAXmlError()
        {
            /* The assembly below would originally (prior to bug fixes) cause an exception to be thrown during assembly: 
                 System.ArgumentException : '', hexadecimal value 0x01, is an invalid character.
             */
            string name = "DA-xmlerror.docx";
            string data = "DA-xmlerror.xml";

            var templateDocx = GetFile(name);
            var dataFile = GetFile(data);

            var wmlTemplate = new WmlDocument(templateDocx);
            var xmlData = XElement.Load(dataFile);

            var afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var returnedTemplateError);
            var assembledDocx = templateDocx.Replace(".docx", "-processed-by-DocumentAssembler.docx");
            afterAssembling.SaveAs(assembledDocx);
        }

        [TestMethod]
        [DataRow("DA025-TemplateDocument.docx", "DA-Data.xml", false)]
        public void DA103_UseXmlDocument(string name, string data, bool err)
        {
            var templateDocx = GetFile(name);
            var dataFile = GetFile(data);

            WmlDocument wmlTemplate = new WmlDocument(templateDocx);
            XmlDocument xmldata = new XmlDocument();
            xmldata.Load(dataFile);

            bool returnedTemplateError;
            WmlDocument afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out returnedTemplateError);
            var assembledDocx = templateDocx.Replace(".docx", "-processed-by-DocumentAssembler.docx");
            afterAssembling.SaveAs(assembledDocx);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator v = new OpenXmlValidator();
                    var valErrors = v.Validate(wDoc).Where(ve => !ExpectedErrors.Contains(ve.Description));
                    valErrors.Should().BeEmpty();
                }
            }
            err.Should().Be(returnedTemplateError);
        }
    }
}