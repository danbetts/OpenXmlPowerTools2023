// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Examples
{
    class DocumentBuilderExample
    {
        private class DocumentInfo
        {
            public int DocumentNumber;
            public int Start;
            public int Count;
        }

        static void Main(string[] args)
        {
            Example01();
            Example02();
            Example03();
            Example04();
        }

        private static void Example01()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WmlSourceBuilder sourceBuilder = new WmlSourceBuilder();
            DocumentBuilder docBuilder = new DocumentBuilder();
            string source1 = "../../Examples01/Source1.docx";
            string source2 = "../../Examples01/Source2.docx";
            string source3 = "../../Examples01/Source3.docx";

            // Create new document from 10 paragraphs starting at paragraph 5 of Source1.docx
            docBuilder.AddSource(sourceBuilder.Start(5).Count(10).KeepSections(true).Build(source1));
            docBuilder.SaveAs(Path.Combine(tempDi.FullName, "Out1.docx"));

            // Create new document from paragraph 1, and paragraphs 5 through end of Source3.docx.
            // This effectively 'deletes' paragraphs 2-4
            docBuilder.AddSource(sourceBuilder.Start(0).Count(1).KeepSections(false).Build(source2));
            docBuilder.AddSource(sourceBuilder.Start(5).KeepSections(false).Build(source3));
            docBuilder.SaveAs(Path.Combine(tempDi.FullName, "Out2.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source1.
            docBuilder.Sources.Clear();
            docBuilder.AddSource(sourceBuilder.KeepSections(true).Build(source1));
            docBuilder.AddSource(sourceBuilder.KeepSections(false).Build(source2));
            docBuilder.SaveAs(Path.Combine(tempDi.FullName, "Out3.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source2.
            docBuilder.Sources[0].KeepSections = false;
            docBuilder.Sources[1].KeepSections = true;
            docBuilder.SaveAs(Path.Combine(tempDi.FullName, "Out4.docx"));

            // Create a new document that consists of the first 5 paragraphs of Source1.docx and the first
            // five paragraphs of Source2.docx.  This example returns a new WmlDocument, when you then can
            // serialize to a SharePoint document library, or use in some other interesting scenario.

            docBuilder.Sources.Clear();
            docBuilder.AddSource(sourceBuilder.Start(0).Count(5).KeepSections(false).Build(source1));
            docBuilder.AddSource(sourceBuilder.KeepSections(true).Build(source2)); // note builder already set to start 0 and count 5, so next item will use same values.
            WmlDocument out5 = docBuilder.ToWmlDocument();
            out5.SaveAs(Path.Combine(tempDi.FullName, "Out5.docx"));  // save it to the file system, but we could just as easily done something
        }
        private static void Example02()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WmlSourceBuilder sourceBuilder = new WmlSourceBuilder();
            DocumentBuilder docBuilder = new DocumentBuilder();

            // Create new document from 10 paragraphs starting at paragraph 5 of Source1.docx
            docBuilder.AddSource(sourceBuilder.Start(0).Count(1).KeepSections(true).Build("../../Example02/WhitePaper.docx"));
            docBuilder.AddSource(sourceBuilder.Defaults().KeepSections(false).Build("../../Example02/Abstract.docx"));
            docBuilder.AddSource(sourceBuilder.Build("../../Example02/AuthorBiography.docx"));
            docBuilder.AddSource(sourceBuilder.Start(1).Build("../../Example02/WhitePaper.docx"));
            docBuilder.SaveAs(Path.Combine(tempDi.FullName, "AssembledPaper.docx"));

            List<WmlSource> sources = new List<WmlSource>();

            // Delete all paragraphs with a specific style.
            using (WordprocessingDocument doc = WordprocessingDocument.Open("../../Notes.docx", false))
            {
                sources = doc
                    .MainDocumentPart
                    .GetXDocument()
                    .Root
                    .Element(W.body)
                    .Elements()
                    .Select((p, i) => new
                    {
                        Paragraph = p,
                        Index = i,
                    })
                    .GroupAdjacent(pi => (string)pi.Paragraph
                        .Elements(W.pPr)
                        .Elements(W.pStyle)
                        .Attributes(W.val)
                        .FirstOrDefault() != "Note")
                    .Where(g => g.Key == true)
                    .Select(g => new WmlSource(
                        new WmlDocument("../../Notes.docx"), g.First().Index,
                            g.Last().Index - g.First().Index + 1, true))
                    .ToList();
            }
            docBuilder.SetSources(sources).SaveAs(Path.Combine(tempDi.FullName, "NewNotes.docx"));

            // Shred a document into multiple parts for each section
            List<DocumentInfo> documentList;
            using (WordprocessingDocument doc = WordprocessingDocument.Open("../../Spec.docx", false))
            {
                var sectionCounts = doc
                    .MainDocumentPart
                    .GetXDocument()
                    .Root
                    .Element(W.body)
                    .Elements()
                    .Rollup(0, (pi, last) => (string)pi
                        .Elements(W.pPr)
                        .Elements(W.pStyle)
                        .Attributes(W.val)
                        .FirstOrDefault() == "Heading1" ? last + 1 : last);
                var beforeZipped = doc
                    .MainDocumentPart
                    .GetXDocument()
                    .Root
                    .Element(W.body)
                    .Elements()
                    .Select((p, i) => new
                    {
                        Paragraph = p,
                        Index = i,
                    });
                var zipped = PtExtensions.PtZip(beforeZipped, sectionCounts, (pi, sc) => new
                {
                    Paragraph = pi.Paragraph,
                    Index = pi.Index,
                    SectionIndex = sc,
                });
                documentList = zipped
                    .GroupAdjacent(p => p.SectionIndex)
                    .Select(g => new DocumentInfo
                    {
                        DocumentNumber = g.Key,
                        Start = g.First().Index,
                        Count = g.Last().Index - g.First().Index + 1,
                    })
                    .ToList();
                }
                foreach (var doc in documentList)
                {
                    string fileName = String.Format("Section{0:000}.docx", doc.DocumentNumber);
                    List<WmlSource> documentSource = new List<WmlSource> {
                    new WmlSource(new WmlDocument("../../Spec.docx"), doc.Start, doc.Count, true)
                };

                docBuilder.SetSources(documentSource).SaveAs(Path.Combine(tempDi.FullName, fileName));
            }

            // Re-assemble the parts into a single document.
            sources = tempDi
                .GetFiles("Section*.docx")
                .Select(d => new WmlSource(new WmlDocument(d.FullName), true))
                .ToList();
            docBuilder.SetSources(sources).SaveAs(Path.Combine(tempDi.FullName, "ReassembledSpec.docx"));
        }
        private static void Example03()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WmlDocument doc1 = new WmlDocument(@"..\..\Template.docx");
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

            string outFileName = Path.Combine(tempDi.FullName, "Out.docx");
            List<WmlSource> sources = new List<WmlSource>()
            {
                new WmlSource(doc1, true),
                new WmlSource(new WmlDocument(@"..\..\Insert-01.docx")) { InsertId = "Liz" },
                new WmlSource(new WmlDocument(@"..\..\Insert-02.docx")) { InsertId = "Eric" },
                new WmlSource(new WmlDocument(@"..\..\FrontMatter.docx")) { InsertId = "Front" },
            };
            DocumentBuilder builder = new DocumentBuilder();
            builder.SetSources(sources).SaveAs(outFileName);
        }
        private static void Example04()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WmlDocument solarSystemDoc = new WmlDocument("../../solar-system.docx");
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(solarSystemDoc))
            using (WordprocessingDocument solarSystem = streamDoc.GetWordprocessingDocument())
            {
                // get children elements of the <w:body> element
                var q1 = solarSystem
                    .MainDocumentPart
                    .GetXDocument()
                    .Root
                    .Element(W.body)
                    .Elements();

                // project collection of tuples containing element and type
                var q2 = q1
                    .Select(
                        e =>
                        {
                            string keyForGroupAdjacent = ".NonContentControl";
                            if (e.Name == W.sdt)
                                keyForGroupAdjacent = e.Element(W.sdtPr)
                                    .Element(W.tag)
                                    .Attribute(W.val)
                                    .Value;
                            if (e.Name == W.sectPr)
                                keyForGroupAdjacent = null;
                            return new
                            {
                                Element = e,
                                KeyForGroupAdjacent = keyForGroupAdjacent
                            };
                        }
                    ).Where(e => e.KeyForGroupAdjacent != null);

                // group by type
                var q3 = q2.GroupAdjacent(e => e.KeyForGroupAdjacent);

                // temporary code to dump q3
                foreach (var g in q3)
                    Console.WriteLine("{0}:  {1}", g.Key, g.Count());
                //Environment.Exit(0);


                // validate existence of files referenced in content controls
                foreach (var f in q3.Where(g => g.Key != ".NonContentControl"))
                {
                    string filename = "../../" + f.Key + ".docx";
                    FileInfo fi = new FileInfo(filename);
                    if (!fi.Exists)
                    {
                        Console.WriteLine("{0} doesn't exist.", filename);
                        Environment.Exit(0);
                    }
                }

                // project collection with opened WordProcessingDocument
                var q4 = q3
                    .Select(g => new
                    {
                        Group = g,
                        Document = g.Key != ".NonContentControl" ?
                            new WmlDocument("../../" + g.Key + ".docx") :
                            solarSystemDoc
                    });

                // project collection of OpenXml.PowerTools.Source
                var sources = q4
                    .Select(
                        g =>
                        {
                            if (g.Group.Key == ".NonContentControl")
                                return new WmlSource(
                                    g.Document,
                                    g.Group
                                        .First()
                                        .Element
                                        .ElementsBeforeSelf()
                                        .Count(),
                                    g.Group
                                        .Count(),
                                    false);
                            else
                                return new WmlSource(g.Document, false);
                        }
                    ).ToList();
                DocumentBuilder docBuilder = new DocumentBuilder();
                docBuilder.SetSources(sources).SaveAs(Path.Combine(tempDi.FullName, "solar-system-new.docx"));
            }
        }
    }
}
