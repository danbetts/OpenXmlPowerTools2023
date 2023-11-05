using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Presentations;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpenXmlPowerTools2023.Tests.Presentations
{
    [TestClass]
    public class PresentationBuilderTests : PresentationTestsBase
    {
        [TestMethod]
        public void PB001_Formatting()
        {
            var source1 = GetFile("PB001-Input1.pptx");
            var source2 = GetFile("PB001-Input2.pptx");

            List<SlideSource> sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source1), 1, true),
                new SlideSource(new PmlDocument(source2), 0, true),
            };
            var target = GetFile("PB001-Formatting.pptx");
            Presentation.BuildPresentation(sources, target);
        }

        [TestMethod]
        public void PB002_Formatting()
        {
            var source2 = GetFile("PB001-Input2.pptx");

            List<SlideSource> sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2), 0, true),
            };
            var target = GetFile("PB002-Formatting.pptx");
            Presentation.BuildPresentation(sources, target);
        }

        [TestMethod]
        public void PB003_Formatting()
        {
            var source1 = GetFile("PB001-Input1.pptx");
            var source2 = GetFile("PB001-Input3.pptx");

            List<SlideSource> sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source1), 1, true),
                new SlideSource(new PmlDocument(source2), 0, true),
            };
            var target = GetFile("PB003-Formatting.pptx");
            Presentation.BuildPresentation(sources, target);
        }

        [TestMethod]
        public void PB004_Formatting()
        {
            var source1 = GetFile("PB001-Input1.pptx");
            var source2 = GetFile("PB001-Input3.pptx");

            List<SlideSource> sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2), 0, true),
                new SlideSource(new PmlDocument(source1), 1, true),
            };
            var target = GetFile("PB004-Formatting.pptx");
            Presentation.BuildPresentation(sources, target);
        }

        [TestMethod]
        public void PB005_Formatting()
        {
            var source1 = GetFile("PB001-Input1.pptx");
            var source2 = GetFile("PB001-Input3.pptx");

            List<SlideSource> sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2), 0, 0, true),
                new SlideSource(new PmlDocument(source1), 1, true),
                new SlideSource(new PmlDocument(source2), 0, true),
            };
            var target = GetFile("PB005-Formatting.pptx");
            Presentation.BuildPresentation(sources, target);
        }

        [TestMethod]
        public void PB006_VideoFormats()
        {
            // This presentation contains videos with content types video/mp4, video/quicktime, video/unknown, video/x-ms-asf, and video/x-msvideo.
            var source = GetFile("PP006-Videos.pptx");

            var oldMediaDataContentTypes = GetMediaDataContentTypes(source);

            List<SlideSource> sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source), true),
            };
            var target = GetFile("PB006-Videos.pptx");
            Presentation.BuildPresentation(sources, target);

            var newMediaDataContentTypes = GetMediaDataContentTypes(target);

            newMediaDataContentTypes.Should().Equal(oldMediaDataContentTypes);
        }

        private static string[] GetMediaDataContentTypes(string file)
        {
            using (PresentationDocument ptDoc = PresentationDocument.Open(file, false))
            {
                return ptDoc.PresentationPart.SlideParts.SelectMany(
                        p => p.DataPartReferenceRelationships.Select(d => d.DataPart.ContentType))
                    .Distinct()
                    .OrderBy(m => m)
                    .ToArray();
            }
        }
    }
}