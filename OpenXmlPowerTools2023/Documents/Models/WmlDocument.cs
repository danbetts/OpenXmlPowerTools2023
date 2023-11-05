using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System.Linq;

/* Unmerged change from project 'OpenXmlPowerTools (net462)'
Before:
using System.Diagnostics.CodeAnalysis;
After:
using System.Diagnostics.CodeAnalysis;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Documents;
*/
using System.Diagnostics.CodeAnalysis;
using OpenXmlPowerTools.Presentations;
using OpenXmlPowerTools.Converters;

namespace OpenXmlPowerTools.Documents
{
    public class WmlDocument : PowerToolsDocument, IDocument
    {
        public PtMainDocumentPart MainDocumentPart => GetMainDocumentPart();
        #region IDocument
        public WordprocessingDocument Document => GetWordprocessingDocument();
        public MainDocumentPart Main => Document.MainDocumentPart;
        public XDocument MainPart => Document.GetMainPart();
        public XElement Root => MainPart.Root;
        public XElement Body => Root.Element(W.body);
        public IEnumerable<XElement> Children => Main.GetBodyElements();
        private string[] extensions { get; set; }
        public string[] Extensions
        {
            get => extensions = extensions ?? Wordprocessing.Extensions;
            private set => extensions = value;
        }
        public IEnumerable<FooterPart> FooterParts => throw new NotImplementedException();
        public IEnumerable<HeaderPart> HeaderParts => throw new NotImplementedException();
        public FontTablePart FontTablePart => throw new NotImplementedException();
        public XDocument FontFamilyTablePart => throw new NotImplementedException();
        #endregion
        public WmlDocument(PowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }
        public WmlDocument(PowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }
        public WmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }
        public WmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }
        public WmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }
        public WmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }
        public WmlDocument(string fileName, MemoryStream memStream)
            : base(fileName, memStream)
        {
        }
        public WmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
            : base(fileName, memStream, convertToTransitional)
        {
        }
        public WmlDocument(WmlDocument other, params XElement[] replacementParts) : base(other)
        {
            using (MemoryStreamDocument streamDoc = new MemoryStreamDocument(this))
            {
                using (Package package = streamDoc.GetPackage())
                {
                    foreach (var replacementPart in replacementParts)
                    {
                        XAttribute uriAttribute = replacementPart.Attribute(PtOpenXml.Uri);
                        if (uriAttribute == null)
                            throw new OpenXmlPowerToolsException("Replacement part does not contain a Uri as an attribute");
                        string uri = uriAttribute.Value;
                        var part = package.GetParts().FirstOrDefault(p => p.Uri.ToString() == uri);
                        using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                        using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                            replacementPart.Save(partXmlWriter);
                    }
                }
                DocumentByteArray = streamDoc.GetModifiedDocument().DocumentByteArray;
            }
        }
        public XElement ConvertToHtml(WmlToHtmlConverterSettings htmlConverterSettings)
        {
            return WmlToHtmlConverter.ConvertToHtml(this, htmlConverterSettings);
        }
        public XElement ConvertToHtml(HtmlConverterSettings htmlConverterSettings)
        {
            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(this, settings);
        }
        public WmlDocument AddToc(string xPath, string switches, string title, int? rightTabPos)
        {
            return ReferenceAdder.AddToc(this, xPath, switches, title, rightTabPos);
        }
        public WmlDocument AddTof(string xPath, string switches, int? rightTabPos)
        {
            return ReferenceAdder.AddTof(this, xPath, switches, rightTabPos);
        }
        public WmlDocument AddToa(string xPath, string switches, int? rightTabPos)
        {
            return ReferenceAdder.AddToa(this, xPath, switches, rightTabPos);
        }
        public WordprocessingDocument GetWordprocessingDocument()
        {
            using (MemoryStream ms = new MemoryStream(DocumentByteArray))
            {
                return WordprocessingDocument.Open(ms, false);
            }
        }
        public PtMainDocumentPart GetMainDocumentPart()
        {
            using (MemoryStream ms = new MemoryStream(DocumentByteArray))
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
            {
                XElement partElement = wDoc.MainDocumentPart.GetXDocument().Root;
                var childNodes = partElement.Nodes().ToList();
                foreach (var item in childNodes)
                {
                    item.Remove();
                }
                return new PtMainDocumentPart(this, wDoc.MainDocumentPart.Uri, partElement.Name, partElement.Attributes(), childNodes);
            }
        }
        public WmlDocument SearchAndReplace(string search, string replace, bool matchCase)
        {
            return TextReplacer.SearchAndReplace(this, search, replace, matchCase);
        }
        public WmlDocument AcceptRevisions(WmlDocument document)
        {
            return RevisionAccepter.AcceptRevisions(document);
        }
        public bool HasTrackedRevisions(WmlDocument document)
        {
            return RevisionAccepter.HasTrackedRevisions(document);
        }
        public WmlDocument SimplifyMarkup(SimplifyMarkupSettings settings)
        {
            return MarkupSimplifier.SimplifyMarkup(this, settings);
        }
    }
}