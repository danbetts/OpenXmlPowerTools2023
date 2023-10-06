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
    public class WmlDocument : OpenXmlPowerToolsDocument
    {
        public PtMainDocumentPart MainDocumentPart
        {
            get
            {
                using (MemoryStream ms = new MemoryStream(DocumentByteArray))
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    XElement partElement = wDoc.MainDocumentPart.GetXDocument().Root;
                    var childNodes = partElement.Nodes().ToList();
                    foreach (var item in childNodes)
                        item.Remove();
                    return new PtMainDocumentPart(this, wDoc.MainDocumentPart.Uri, partElement.Name, partElement.Attributes(), childNodes);
                }
            }
        }

        public WmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }

        public WmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
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
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(this))
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
        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public XElement ConvertToHtml(WmlToHtmlConverterSettings htmlConverterSettings)
        {
            return WmlToHtmlConverter.ConvertToHtml(this, htmlConverterSettings);
        }

        [SuppressMessage("ReSharper", "UnusedMember.Global")]
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