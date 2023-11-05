using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;
using System;
using System.IO;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class WmlContentAtomListTests : DocumentTestsBase
    {
        protected override string FeatureFolder { get; } = @".\WmlComparer";

        [TestMethod]
        [DataRow("HC009-Test-04.docx")]
        public void CA002_Annotations(string name)
        {
            var source = GetFile(name, moduleFolder: @".\Converters", featureFolder: @".\HtmlConverter");

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(source, true))
            {
                var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                var settings = new WmlComparerSettings();
                WmlComparer.CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent, settings);
            }
        }

        [TestMethod]
        [DataRow("CA009-altChunk.docx")]
        [ExpectedException(typeof(NotSupportedException))]
        public void CA003_ContentAtoms_Throws(string name)
        {
            var source = GetFile(name, featureFolder: @".\WmlComparer\CA");

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(source, true))
            {
                var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                var settings = new WmlComparerSettings();
                WmlComparer.CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent, settings);
            }
        }
    }
}