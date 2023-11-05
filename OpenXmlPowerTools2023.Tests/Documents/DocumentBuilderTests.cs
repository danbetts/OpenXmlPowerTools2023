using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class DocumentBuilderTests : DocumentTestsBase
    {
        protected override string FeatureFolder { get; } = @".\DocumentBuilder";

        [TestMethod]
        public void DB001_Keep_Sections()
        {
            Builder.AddSource(GetFile("DB001-Sections.docx"));

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB002_Keep_Sections_Discard_Headers()
        {
            Builder.AddSource(GetFile("DB002-Sections-With-Headers.docx"), keepHeadersAndFooters: false);
            Builder.AddSource(GetFile("DB002-Landscape-Section.docx"), keepHeadersAndFooters: false);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB003_Only_Default_Header()
        {
            Builder.AddSource(GetFile("DB003-Only-Default-Header.docx"), keepHeadersAndFooters: true);
            Builder.AddSource(GetFile("DB002-Landscape-Section.docx"), keepHeadersAndFooters: false);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB004_No_Headers()
        {
            Builder.AddSource(GetFile("DB004-No-Headers.docx"), keepHeadersAndFooters: false)
                      .AddSource(GetFile("DB002-Landscape-Section.docx"), keepHeadersAndFooters: false);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB005_Headers_With_Refs_To_Images()
        {
            Builder.AddSource(GetFile("DB005-Headers-With-Images.docx"), keepHeadersAndFooters: true)
                      .AddSource(GetFile("DB002-Landscape-Section.docx"));

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB006_Start_To_Count()
        {
            Builder.AddSource(GetFile("DB006-Source1.docx"), 5, 10);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB006_Remove_Paragraphs()
        {
            var source3 = GetFile("DB006-Source3.docx");

            Builder.AddSource(source3, count: 1)
                      .AddSource(source3, start: 4);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB006_Keep_First_Document_Sections()
        {
            Builder.AddSource(GetFile("DB006-Source1.docx"), keepSections: true)
                      .AddSource(GetFile("DB006-Source2.docx"), keepSections: false);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB006_Keep_Second_Document_Sections()
        {
            Builder.AddSource(GetFile("DB006-Source1.docx"), keepSections: false)
                      .AddSource(GetFile("DB006-Source2.docx"), keepSections: true);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB006_Take_First_Five_Paragraphs_From_Both_Documents()
        {
            Builder.AddSource(GetFile("DB006-Source1.docx"), count: 5, keepSections: false)
                      .AddSource(GetFile("DB006-Source2.docx"));

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB006_White_Paper()
        {
            Builder.AddSource(GetFile("DB007-Spec.docx"), 0, 1, true)
                      .AddSource(GetFile("DB007-Abstract.docx"), keepSections: false)
                      .AddSource(GetFile("DB007-AuthorBiography.docx"), keepSections: false)
                      .AddSource(GetFile("DB007-WhitePaper.docx"), 1, keepSections: false);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB008_Delete_Paragraphs_With_Given_Style()
        {
            // Assemble
            var source = GetFile("DB007-Notes.docx");
            List<WmlSource> sources = new List<WmlSource>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(source, false))
            {
                sources = doc.MainDocumentPart.GetXDocument().Root.Element(W.body).Elements()
                    .Select((p, i) => new { Paragraph = p, Index = i, })
                    .GroupAdjacent(pi => (string)pi.Paragraph.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault() != "Note")
                    .Where(g => g.Key == true).Select(g => new WmlSource(new WmlDocument(source), g.First().Index, g.Last().Index - g.First().Index + 1, true, true, null))
                    .ToList();
            }
            Builder.SetSources(sources);

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        [DataRow("Content-Controls.docx", "Fax.docx")]
        [DataRow("Letterhead.docx", "Fax.docx")]
        [DataRow("Letterhead-with-Watermark.docx", "Fax.docx")]
        [DataRow("Logo.docx", "Fax.docx")]
        [DataRow("Watermark-1.docx", "Fax.docx")]
        [DataRow("Watermark-2.docx", "Fax.docx")]
        [DataRow("Disclaimer.docx", "Fax.docx")]
        [DataRow("Footer.docx", "Fax.docx")]
        //[DataRow("Content-Controls.docx", "Letter.docx")]
        [DataRow("Letterhead.docx", "Letter.docx")]
        [DataRow("Letterhead-with-Watermark.docx", "Letter.docx")]
        //[DataRow("Logo.docx", "Letter.docx")]
        [DataRow("Watermark-1.docx", "Letter.docx")]
        [DataRow("Watermark-2.docx", "Letter.docx")]
        [DataRow("Disclaimer.docx", "Letter.docx")]
        [DataRow("Footer.docx", "Letter.docx")]
        public void DB009_Import_Into_Headers_Footers(string src, string dest)
        {
            var source = GetFile(src, featureFolder: @".\DocumentBuilder\HeadersFooters\Src");
            var destination = GetFile(dest, featureFolder: @".\DocumentBuilder\HeadersFooters\Dest");
            Builder.AddSource(destination).AddSource(source, insertId: "Templafy");

            ValidateAgainstExpected(OutputFile, out string output);
        }

        // Test methods code is too complex.
        //[TestMethod]
        //public void DB009_Shred_Document()
        //{
        //    // Assemble

        //    string spec = GetFile("DB007-Spec.docx");
        //    // Shred a document into multiple parts for each section

        //    using (WordprocessingDocument doc = WordprocessingDocument.Open(spec, false))
        //    {
        //        var sectionCounts = doc.MainDocumentPart.GetBodyElements()
        //            .Rollup(0, (pi, last) => (string)pi.Elements(W.pPr)
        //                                               .Elements(W.pStyle).Attributes(W.val)
        //                                               .FirstOrDefault() == "Heading1" ? last + 1 : last);

        //        var beforeZipped = doc.MainDocumentPart.GetBodyElements().Select((p, i) => new { Paragraph = p, Index = i });

        //        var zipped = beforeZipped.PtZip(sectionCounts, (pi, sc) => new { Paragraph = pi.Paragraph, Index = pi.Index, SectionIndex = sc});
        //        documentList = zipped.GroupAdjacent(p => p.SectionIndex)
        //            .Select(g => new DocumentInfo
        //            {
        //                DocumentNumber = g.Key,
        //                Start = g.First().Index,
        //                Count = g.Last().Index - g.First().Index + 1,
        //            })
        //            .ToList();
        //    }
        //    foreach (var doc in documentList)
        //    {
        //        string fileName = String.Format("DB009-Section{0:000}.docx", doc.DocumentNumber);
        //        var fiSection = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, fileName));
        //        List<WmlSource> documentSource = new List<WmlSource> {
        //            new WmlSource(new WmlDocument(spec.FullName), doc.Start, doc.Count, true)
        //        };
        //        DocumentBuilder.BuildDocument(documentSource, fiSection.FullName);
        //        Validate(fiSection);
        //    }

        //    // Re-assemble the parts into a single document.
        //    List<WmlSource> sources = TestUtil.TempDir
        //        .GetFiles("DB009-Section*.docx")
        //        .Select(d => new WmlSource(new WmlDocument(d.FullName), true))
        //        .ToList();
        //    var fiReassembled = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DB009-Reassembled.docx"));

        //    DocumentBuilder.BuildDocument(sources, fiReassembled.FullName);
        //    using (WordprocessingDocument doc = WordprocessingDocument.Open(fiReassembled.FullName, true))
        //    {
        //        ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]",
        //            @"TOC \o '1-3' \h \z \u", null, null);
        //    }

        //    Validate(OutputFile);
        //}

        [TestMethod]
        public void DB010_Insert_Using_InsertId()
        {

            var front = GetFile("DB010-FrontMatter.docx");
            var insert01 = GetFile("DB010-Insert-01.docx");
            var insert02 = GetFile("DB010-Insert-02.docx");
            var template = GetFile("DB010-Template.docx");

            WmlDocument doc1 = new WmlDocument(template);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    XDocument xDoc = doc.MainDocumentPart.GetXDocument();
                    XElement frontMatterPara = xDoc.Root.Descendants(W.txbxContent).Elements(W.p).FirstOrDefault();
                    frontMatterPara.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Front")));
                    XElement tbl = xDoc.Root.Element(W.body).Elements(W.tbl).FirstOrDefault();
                    XElement firstCell = tbl.Descendants(W.tr).First().Descendants(W.p).First();
                    firstCell.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Liz")));
                    XElement secondCell = tbl.Descendants(W.tr).Skip(1).First().Descendants(W.p).First();
                    secondCell.ReplaceWith(
                        new XElement(PtOpenXml.Insert,
                            new XAttribute("Id", "Eric")));
                    doc.MainDocumentPart.PutXDocument();
                }
                doc1.DocumentByteArray = mem.ToArray();
            }

            Builder.AddSource(doc1)
                      .AddSource(insert01, insertId: "Liz")
                      .AddSource(insert02, insertId: "Eric")
                      .AddSource(front, insertId: "Front");

            ValidateAgainstExpected(OutputFile, out string output);
        }

        [TestMethod]
        public void DB011_Body_And_Header_With_Shapes()
        {
            Builder.AddSource(GetFile("DB011-Header-With-Shape.docx"));
            Builder.AddSource(GetFile("DB011-Body-With-Shape.docx"));

            ValidateAgainstExpected(OutputFile, out string output);
            ValidateUniqueDocPrIds(output);
        }

        [TestMethod]
        public void DB012_Numberings_With_Same_Abstract_Numbering()
        {
            // Assemble
            Builder.AddSource(GetFile("DB012-Lists-With-Different-Numberings.docx"));

            // Act
            var output = BuildAndSave();

            // Assert
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(output, false))
            {
                wDoc.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root.Elements(W.num).Count().Should().Be(9);
            }
        }

        [TestMethod]
        public void DB013a_Localized_StyleIds_Heading()
        {
            // Each of these documents have changed the font color of the Heading 1 style, one to red, the other to green.
            // One of the documents were created with English as the Word display language, the other with Danish as the language.

            // Assemble
            Builder.AddSource(GetFile("DB013a-Red-Heading1-English.docx"));
            Builder.AddSource(GetFile("DB013a-Green-Heading1-Danish.docx"));

            // Act
            var output = BuildAndSave();

            // Assert
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(output, false))
            {
                var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.style).ToArray();
                styles.Count(s => s.Element(W.name).Attribute(W.val).Value == "heading 1").Should().Be(1);

                var styleIds = new HashSet<string>(styles.Select(s => s.Attribute(W.styleId).Value));
                var paragraphStylesIds = new HashSet<string>(wDoc.MainDocumentPart.GetXDocument().Descendants(W.pStyle).Select(p => p.Attribute(W.val).Value));
                styleIds.IsSubsetOf(paragraphStylesIds).Should().BeFalse();
            }
        }

        [TestMethod]
        public void DB013b_Localized_StyleIds_List()
        {
            // Each of these documents have changed the font color of the List Paragraph style, one to orange, the other to blue.
            // One of the documents were created with English as the Word display language, the other with Danish as the language.

            // Arrange
            Builder.AddSource(GetFile("DB013b-Orange-List-Danish.docx"));
            Builder.AddSource(GetFile("DB013b-Blue-List-English.docx"));

            // Act
            var output = BuildAndSave();

            // Assert
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(output, false))
            {
                var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.style).ToArray();
                styles.Count(s => s.Element(W.name).Attribute(W.val).Value == "List Paragraph").Should().Be(1);

                var styleIds = new HashSet<string>(styles.Select(s => s.Attribute(W.styleId).Value));
                var paragraphStylesIds = new HashSet<string>(wDoc.MainDocumentPart.GetXDocument().Descendants(W.pStyle).Select(p => p.Attribute(W.val).Value));
                styleIds.IsSubsetOf(paragraphStylesIds).Should().BeFalse();
            }
        }

        [TestMethod]
        public void DB014_Keep_Web_Extensions()
        {
            Builder.AddSource(GetFile("DB014-WebExtensions.docx"));

            ValidateAgainstExpected(OutputFile, out string output);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(output, false))
            {
                wDoc.WebExTaskpanesPart.Should().NotBeNull();
                wDoc.WebExTaskpanesPart.Taskpanes.ChildElements.Count.Should().Be(2);
                wDoc.WebExTaskpanesPart.WebExtensionParts.Count().Should().Be(2);
            }
        }

        [TestMethod]
        public void DB015_Latent_Styles()
        {
            Builder.AddSource(GetFile("DB015-LatentStyles.docx"));

            ValidateAgainstExpected(OutputFile, out string output);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(output, false))
            {
                wDoc.WebExTaskpanesPart.Should().BeNull();
            }
        }

        [TestMethod]
        public void DB0016_DocDefaultStyles()
        {
            Builder.AddSource(GetFile("DB0016-DocDefaultStyles.docx"));

            ValidateAgainstExpected(OutputFile, out string output);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(output, true))
            {
                var styles = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.docDefaults).ToArray();
                styles.Single();
            }
        }

        [TestMethod]
        [DeploymentItem("CellLevelContentControl-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "InlineContentControl.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("MultilineWithBulletPoints-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "MultilineWithBulletPoints.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("NestedContentControl-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "NestedContentControl.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("RowLevelContentControl-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "RowLevelContentControl.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("ContentControlDanishProofingLanguage-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "ContentControlDanishProofingLanguage.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("ContentControlEnglishProofingLanguage-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "ContentControlEnglishProofingLanguage.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("ContentControlMixedProofingLanguage-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "ContentControlMixedProofingLanguage.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("ContentControlWithContent-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "ContentControlWithContent.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("FooterContent-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "FooterContent.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DeploymentItem("HeaderContent-built.docx")]
        [DataRow("BaseDocument.docx", 0, 4, "HeaderContent.docx", 0, int.MaxValue, "BaseDocument.docx", 4, int.MaxValue)]

        [DataRow("BaseDocument.docx", 0, int.MaxValue, "CellLevelContentControl.docx", 0, int.MaxValue, "NestedContentControl.docx", 0, int.MaxValue)]
        public void WithGlossaryDocuments(string path1, int start1, int count1, string path2, int start2, int count2, string path3, int start3, int count3)
        {
            string featureFolder = @".\DocumentBuilder\GlossaryDocuments";
            // Arrange

            Builder.AddSource(GetFile(path1, featureFolder: featureFolder), start1, count1);
            Builder.AddSource(GetFile(path2, featureFolder: featureFolder), start2, count2);
            Builder.AddSource(GetFile(path3, featureFolder: featureFolder), start3, count3);

            // Act and Assert
            string output = Path.Combine(TestContext.TestResultsDirectory, OutputFile);
            ValidateAgainstExpected(output);
        }

        private void ValidateUniqueDocPrIds(string filepath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, false))
            {
                var docPrIds = new HashSet<string>();
                var main = doc.MainDocumentPart;

                // Doc parts
                foreach (var item in main.GetXDocument().Descendants(WP.docPr))
                {
                    docPrIds.Add(item.Attribute(NoNamespace.id).Value).Should().BeTrue();
                }

                // Header parts
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                {
                    foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                    {
                        docPrIds.Add(item.Attribute(NoNamespace.id).Value).Should().BeTrue();
                    }
                }

                // Footer parts
                foreach (var footer in doc.MainDocumentPart.FooterParts)
                {
                    foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                    {
                        docPrIds.Add(item.Attribute(NoNamespace.id).Value).Should().BeTrue();
                    }
                }

                if (doc.MainDocumentPart.FootnotesPart != null)
                {
                    foreach (var item in doc.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                    {
                        docPrIds.Add(item.Attribute(NoNamespace.id).Value).Should().BeTrue();
                    }
                }

                if (doc.MainDocumentPart.EndnotesPart != null)
                {
                    foreach (var item in doc.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                    {
                        docPrIds.Add(item.Attribute(NoNamespace.id).Value).Should().BeTrue();
                    }
                }
            }
        }
    }
}