﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class MarkupSimplifierTests
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        private const string SmartTagDocumentTextValue = "The countries include Algeria, Botswana, and Sri Lanka.";
        private const string SmartTagDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p >
      <w:r>
        <w:t xml:space=""preserve"">The countries include </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:r>
          <w:t>Algeria</w:t>
        </w:r>
      </w:smartTag>
      <w:r>
        <w:t xml:space=""preserve"">, </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:r>
          <w:t>Botswana</w:t>
        </w:r>
      </w:smartTag>
      <w:r>
        <w:t xml:space=""preserve"">, and </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""place"">
          <w:r>
            <w:t>Sri Lanka</w:t>
          </w:r>
        </w:smartTag>
      </w:smartTag>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
";
        private const string SdtDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:text/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t>Hello World!</w:t>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
  </w:body>
</w:document>";
        private const string GoBackBookmarkDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id=""0"" w:name=""_GoBack""/>
      <w:bookmarkEnd w:id=""0""/>
    </w:p>
  </w:body>
</w:document>";

        [TestMethod]
        public void CanRemoveSmartTags()
        {
            XDocument partDocument = XDocument.Parse(SmartTagDocumentXmlString);
            partDocument.Descendants(W.smartTag).Any().Should().BeTrue();

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                var settings = new SimplifyMarkupSettings { RemoveSmartTags = true };
                MarkupSimplifier.SimplifyMarkup(wordDocument, settings);

                partDocument = part.GetXDocument();
                XElement t = partDocument.Descendants(W.t).First();

                partDocument.Descendants(W.smartTag).Any().Should().BeFalse();
                SmartTagDocumentTextValue.Should().Be(t.Value);
            }
        }

        [TestMethod]
        public void CanRemoveContentControls()
        {
            XDocument partDocument = XDocument.Parse(SdtDocumentXmlString);

            partDocument.Descendants(W.sdt).Any().Should().BeTrue();

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                var settings = new SimplifyMarkupSettings { RemoveContentControls = true };
                MarkupSimplifier.SimplifyMarkup(wordDocument, settings);

                partDocument = part.GetXDocument();
                XElement element = partDocument
                    .Descendants(W.body)
                    .Descendants()
                    .First();

                partDocument.Descendants(W.sdt).Any().Should().BeFalse();
                element.Name.Should().Be(W.p);
            }
        }

        [TestMethod]
        public void CanRemoveGoBackBookmarks()
        {
            XDocument partDocument = XDocument.Parse(GoBackBookmarkDocumentXmlString);
            partDocument.Descendants(W.bookmarkStart).Should().Contain(e => e.Attribute(W.name).Value == "_GoBack" && e.Attribute(W.id).Value == "0");
            partDocument.Descendants(W.bookmarkEnd).Should().Contain(e => e.Attribute(W.id).Value == "0");

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                var settings = new SimplifyMarkupSettings { RemoveGoBackBookmark = true };
                MarkupSimplifier.SimplifyMarkup(wordDocument, settings);

                partDocument = part.GetXDocument();
                partDocument.Descendants(W.bookmarkStart).Any().Should().BeFalse();
                partDocument.Descendants(W.bookmarkEnd).Any().Should().BeFalse();
            }
        }
    }
}