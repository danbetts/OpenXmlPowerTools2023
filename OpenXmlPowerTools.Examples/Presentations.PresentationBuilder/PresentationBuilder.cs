// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Presentations;
using OpenXmlPowerTools.Spreadsheets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Examples
{
    class ExamplePresentationBuilder01
    {
        static void Main(string[] args)
        {
            Example01();
            Example02();
        }

        private static void Example01()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            string source1 = "../../Contoso.pptx";
            string source2 = "../../Companies.pptx";
            string source3 = "../../Customer Content.pptx";
            string source4 = "../../Presentation One.pptx";
            string source5 = "../../Presentation Two.pptx";
            string source6 = "../../Presentation Three.pptx";
            string contoso1 = "../../Contoso One.pptx";
            string contoso2 = "../../Contoso Two.pptx";
            string contoso3 = "../../Contoso Three.pptx";
            List<SlideSource> sources = null;

            var sourceDoc = new PmlDocument(source1);
            sources = new List<SlideSource>()
            {
                new SlideSource(sourceDoc, 0, 1, false),  // Title
                new SlideSource(sourceDoc, 1, 1, false),  // First intro (of 3)
                new SlideSource(sourceDoc, 4, 2, false),  // Sales bios
                new SlideSource(sourceDoc, 9, 3, false),  // Content slides
                new SlideSource(sourceDoc, 13, 1, false),  // Closing summary
            };
            Presentation.BuildPresentation(sources, Path.Combine(tempDi.FullName, "Out1.pptx"));

            sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2), 2, 1, true),  // Choose company
                new SlideSource(new PmlDocument(source3), false),       // Content
            };
            Presentation.BuildPresentation(sources, Path.Combine(tempDi.FullName, "Out2.pptx"));

            sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source4), true),
                new SlideSource(new PmlDocument(source5), true),
                new SlideSource(new PmlDocument(source6), true),
            };
            Presentation.BuildPresentation(sources, Path.Combine(tempDi.FullName, "Out3.pptx"));

            sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(contoso1), true),
                new SlideSource(new PmlDocument(contoso2), true),
                new SlideSource(new PmlDocument(contoso3), true),
            };
            Presentation.BuildPresentation(sources, Path.Combine(tempDi.FullName, "Out4.pptx"));
        }

        private static void Example02()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            string presentation = "../../Presentation1.pptx";
            string hiddenPresentation = "../../HiddenPresentation.pptx";

            // First, load both presentations into byte arrays, simulating retrieving presentations from some source
            // such as a SharePoint server
            var baPresentation = File.ReadAllBytes(presentation);
            var baHiddenPresentation = File.ReadAllBytes(hiddenPresentation);

            // Next, replace "thee" with "the" in the main presentation
            var pmlMainPresentation = new PmlDocument("Main.pptx", baPresentation);
            PmlDocument modifiedMainPresentation = null;
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(pmlMainPresentation))
            {
                using (PresentationDocument document = streamDoc.GetPresentationDocument())
                {
                    var pXDoc = document.PresentationPart.GetXDocument();
                    foreach (var slideId in pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        var slidePart = document.PresentationPart.GetPartById(slideRelId);
                        var slideXDoc = slidePart.GetXDocument();
                        var paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("thee"), "the", null);
                        slidePart.PutXDocument();
                    }
                }
                modifiedMainPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // Combine the two presentations into a single presentation
            var slideSources = new List<SlideSource>() {
                new SlideSource(modifiedMainPresentation, 0, 1, true),
                new SlideSource(new PmlDocument("Hidden.pptx", baHiddenPresentation), true),
                new SlideSource(modifiedMainPresentation, 1, true),
            };
            PmlDocument combinedPresentation = Presentation.BuildPresentation(slideSources);

            // Replace <# TRADEMARK #> with AdventureWorks (c)
            PmlDocument modifiedCombinedPresentation = null;
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(combinedPresentation))
            {
                using (PresentationDocument document = streamDoc.GetPresentationDocument())
                {
                    var pXDoc = document.PresentationPart.GetXDocument();
                    foreach (var slideId in pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId).Skip(1).Take(1))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        var slidePart = document.PresentationPart.GetPartById(slideRelId);
                        var slideXDoc = slidePart.GetXDocument();
                        var paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("<# TRADEMARK #>"), "AdventureWorks (c)", null);
                        slidePart.PutXDocument();
                    }
                }
                modifiedCombinedPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // we now have a PmlDocument (which is essentially a byte array) that can be saved as necessary.
            modifiedCombinedPresentation.SaveAs(Path.Combine(tempDi.FullName, "Modified.pptx"));
        }
    }
}
