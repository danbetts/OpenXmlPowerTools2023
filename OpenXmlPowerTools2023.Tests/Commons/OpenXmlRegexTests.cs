using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools2023.Tests;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Commons
{
    [TestClass]
    public class OpenXmlRegexTests : CommonTestsBase
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        private const string LeftDoubleQuotationMarks = @"[\u0022“„«»”]";
        private const string Words = @"[\w\-&/]+(?:\s[\w\-&/]+)*";
        private const string RightDoubleQuotationMarks = @"[\u0022”‟»«“]";

        private const string QuotationMarksDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">Text can be enclosed in “normal double quotes” and in </w:t>
      </w:r>
      <w:r>
        <w:t>«</w:t>
      </w:r>
      <w:r>
        <w:t>double angle quotation marks</w:t>
      </w:r>
      <w:r>
        <w:t>»</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";
        private const string QuotationMarksAndTrackedChangesDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">Text can be enclosed in “normal </w:t>
      </w:r>
      <w:ins w:id=""8"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:54:00Z"">
        <w:r>
          <w:t xml:space=""preserve"">double </w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t xml:space=""preserve"">quotes” </w:t>
      </w:r>
      <w:del w:id=""9"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:55:00Z"">
        <w:r>
          <w:delText xml:space=""preserve"">or </w:delText>
        </w:r>
      </w:del>
      <w:ins w:id=""10"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:55:00Z"">
        <w:r>
          <w:t xml:space=""preserve"">and </w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t xml:space=""preserve"">in </w:t>
      </w:r>
      <w:r>
        <w:t>«</w:t>
      </w:r>
      <w:r>
        <w:t xml:space=""preserve"">double </w:t>
      </w:r>
      <w:ins w:id=""11"" w:author=""Thomas Barnekow"" w:date=""2016-12-03T15:54:00Z"">
        <w:r>
          <w:t xml:space=""preserve"">angle </w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t>quotation marks</w:t>
      </w:r>
      <w:r>
        <w:t>»</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";
        private const string SymbolsAndTrackedChangesDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">We can also use symbols such as </w:t>
      </w:r>
      <w:del w:id=""4"" w:author=""Thomas Barnekow"" w:date=""2017-04-16T12:31:00Z"">
        <w:r>
          <w:sym w:font=""Wingdings"" w:char=""F028""/>
        </w:r>
        <w:r>
          <w:delText xml:space=""preserve"">, </w:delText>
        </w:r>
      </w:del>
      <w:r>
        <w:sym w:font=""Wingdings"" w:char=""F021""/>
      </w:r>
      <w:r>
        <w:t xml:space=""preserve""> or </w:t>
      </w:r>
      <w:r>
        <w:sym w:font=""Wingdings"" w:char=""F028""/>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";
        private const string FieldsDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val=""Heading1""/>
      </w:pPr>
      <w:bookmarkStart w:id=""0"" w:name=""_Ref491716064""/>
      <w:r>
        <w:t>Article</w:t>
      </w:r>
      <w:bookmarkEnd w:id=""0""/>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val=""Heading2""/>
      </w:pPr>
      <w:bookmarkStart w:id=""1"" w:name=""_Ref491716082""/>
      <w:r>
        <w:t>Section</w:t>
      </w:r>
      <w:bookmarkEnd w:id=""1""/>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val=""HeadingBody2""/>
      </w:pPr>
      <w:r>
        <w:t xml:space=""preserve"">As stated in Article </w:t>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType=""begin""/>
      </w:r>
      <w:r>
        <w:instrText xml:space=""preserve""> REF _Ref491716064 \r \h </w:instrText>
      </w:r>
      <w:r>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType=""separate""/>
      </w:r>
      <w:r>
        <w:t>1</w:t>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType=""end""/>
      </w:r>
      <w:r>
        <w:t xml:space=""preserve""> and this Section </w:t>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType=""begin""/>
      </w:r>
      <w:r>
        <w:instrText xml:space=""preserve""> REF _Ref491716082 \r \h </w:instrText>
      </w:r>
      <w:r>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType=""separate""/>
      </w:r>
      <w:r>
        <w:t>1.1</w:t>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType=""end""/>
      </w:r>
      <w:r>
        <w:t>, this is described in Schedule C (Performance Framework).</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";
        private const string LastRenderedPageBreakXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t>ThisIsAParagraphContainingNoNaturalLi</w:t>
      </w:r>
      <w:r>
        <w:lastRenderedPageBreak/>
        <w:t>neBreaksSoTheLineBreakIsForced.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

        [TestMethod]
        public void CanReplaceTextWithQuotationMarks()
        {
            XDocument partDocument = XDocument.Parse(QuotationMarksDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            innerText.Should().Be("Text can be enclosed in “normal double quotes” and in «double angle quotation marks».");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(string.Format("{0}(?<words>{1}){2}", LeftDoubleQuotationMarks, Words,
                    RightDoubleQuotationMarks));
                int count = PowerToolsRegex.Replace(content, regex, "‘changed ${words}’", null);

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                count.Should().Be(2);
                innerText.Should().Be("Text can be enclosed in ‘changed normal double quotes’ and in ‘changed double angle quotation marks’.");
            }
        }

        [TestMethod]
        public void CanReplaceTextWithQuotationMarksAndAddTrackedChangesWhenReplacing()
        {
            XDocument partDocument = XDocument.Parse(QuotationMarksDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            innerText.Should().Be("Text can be enclosed in “normal double quotes” and in «double angle quotation marks».");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(string.Format("{0}(?<words>{1}){2}", LeftDoubleQuotationMarks, Words,
                    RightDoubleQuotationMarks));
                int count = PowerToolsRegex.Replace(content, regex, "‘changed ${words}’", null, true, "John Doe");

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                count.Should().Be(2);
                innerText.Should().Be("Text can be enclosed in ‘changed normal double quotes’ and in ‘changed double angle quotation marks’.");
                p.Elements(W.ins).Should().Contain(e => InnerText(e) == "‘changed normal double quotes’");
                p.Elements(W.ins).Should().Contain(e => InnerText(e) == "‘changed double angle quotation marks’");
                p.Elements(W.del).Should().Contain(e => InnerDelText(e) == "“normal double quotes”");
                p.Elements(W.del).Should().Contain(e => InnerDelText(e) == "«double angle quotation marks»");
            }
        }

        [TestMethod]
        public void CanReplaceTextWithQuotationMarksAndTrackedChanges()
        {
            XDocument partDocument = XDocument.Parse(QuotationMarksAndTrackedChangesDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            innerText.Should().Be("Text can be enclosed in “normal double quotes” and in «double angle quotation marks».");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(string.Format("{0}(?<words>{1}){2}", LeftDoubleQuotationMarks, Words,
                    RightDoubleQuotationMarks));
                int count = PowerToolsRegex.Replace(content, regex, "‘changed ${words}’", null, true, "John Doe");

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                count.Should().Be(2);
                innerText.Should().Be("Text can be enclosed in ‘changed normal double quotes’ and in ‘changed double angle quotation marks’.");
                p.Elements(W.ins).Should().Contain(e => InnerText(e) == "‘changed normal double quotes’");
                p.Elements(W.ins).Should().Contain(e => InnerText(e) == "‘changed double angle quotation marks’");
            }
        }

        [TestMethod]
        public void CanReplaceTextWithSymbolsAndTrackedChanges()
        {
            XDocument partDocument = XDocument.Parse(SymbolsAndTrackedChangesDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).First();
            string innerText = InnerText(p);

            innerText.Should().Be("We can also use symbols such as \uF021 or \uF028.");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(@"[\uF021]");
                int count = PowerToolsRegex.Replace(content, regex, "\uF028", null, true, "John Doe");

                p = partDocument.Descendants(W.p).First();
                innerText = InnerText(p);

                count.Should().Be(1);
                innerText.Should().Be("We can also use symbols such as \uF028 or \uF028.");
                p.Descendants(W.ins).Should().Contain(ins => ins.Descendants(W.sym).Any(
                        sym => sym.Attribute(W.font).Value == "Wingdings" && 
                               sym.Attribute(W._char).Value == "F028"));
            }
        }

        [TestMethod]
        public void CanReplaceTextWithFields()
        {
            XDocument partDocument = XDocument.Parse(FieldsDocumentXmlString);
            XElement p = partDocument.Descendants(W.p).Last();
            string innerText = InnerText(p);

            innerText.Should().Be("As stated in Article {__1} and this Section {__1.1}, this is described in Schedule C (Performance Framework).");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(@"Schedule C \(Performance Framework\)");
                int count = PowerToolsRegex.Replace(content, regex, "Exhibit 4", null, true, "John Doe");

                p = partDocument.Descendants(W.p).Last();
                innerText = InnerText(p);

                count.Should().Be(1);
                innerText.Should().Be("As stated in Article {__1} and this Section {__1.1}, this is described in Exhibit 4.");
            }
        }

        [TestMethod]
        public void CanMatchDespiteLastRenderedPageBreaks()
        {
            XDocument partDocument = XDocument.Parse(LastRenderedPageBreakXmlString);
            XElement p = partDocument.Descendants(W.p).Last();
            string innerText = InnerText(p);

            innerText.Should().Be("ThisIsAParagraphContainingNoNaturalLineBreaksSoTheLineBreakIsForced.");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                IEnumerable<XElement> content = partDocument.Descendants(W.p);
                var regex = new Regex(@"LineBreak");
                int count = PowerToolsRegex.Replace(content, regex, "LB", null);

                p = partDocument.Descendants(W.p).Last();
                innerText = InnerText(p);

                count.Should().Be(2);
                innerText.Should().Be("ThisIsAParagraphContainingNoNaturalLBsSoTheLBIsForced.");
            }
        }
    }
}