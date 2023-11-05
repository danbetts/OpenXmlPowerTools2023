using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools2023.Commons;
using OpenXmlPowerTools2023.Tests.Documents;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Commons
{
    [TestClass]
    public class StronglyTypedBlockTests : CommonTestsBase
    {
        [TestMethod]
        public void CanUseStronglyTypedBlockToDemarcateApis()
        {
            using (var stream = new MemoryStream())
            {
                CreateEmptyWordprocessingDocument(stream);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    MainDocumentPart part = wordDocument.MainDocumentPart;

                    // Add a paragraph through the PowerTools.
                    XDocument content = part.GetXDocument();
                    XElement bodyElement = content.Descendants(W.body).First();
                    bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
                    part.PutXDocument();

                    // This demonstrates the use of the StronglyTypedBlock in a using statement to
                    // demarcate the intermittent use of the strongly typed classes.
                    using (new StronglyTypedBlock(wordDocument))
                    {
                        // Assert that we can see the paragraph added through the PowerTools.
                        Body body = part.Document.Body;
                        List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();

                        paragraphs.Count.Should().Be(1);
                        paragraphs[0].InnerText.Should().Be("Added through PowerTools");

                        // Add a paragraph through the SDK.
                        body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));
                    }

                    // Assert that we can see the paragraphs added through the PowerTools and the SDK.
                    content = part.GetXDocument();
                    List<XElement> paragraphElements = content.Descendants(W.p).ToList();
                    paragraphElements.Count.Should().Be(2);
                    paragraphElements[0].Value.Should().Be("Added through PowerTools");
                    paragraphElements[1].Value.Should().Be("Added through SDK");
                }
            }
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorThrowsWhenPassingNull()
        {
            new StronglyTypedBlock(null);
        }
    }
}