using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Presentations
{
    public class PtMainDocumentPart : XElement
    {
        private WmlDocument ParentWmlDocument;

        public PtWordprocessingCommentsPart WordprocessingCommentsPart
        {
            get
            {
                using (MemoryStream ms = new MemoryStream(ParentWmlDocument.DocumentByteArray))
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    WordprocessingCommentsPart commentsPart = wDoc.MainDocumentPart.WordprocessingCommentsPart;
                    if (commentsPart == null)
                        return null;
                    XElement partElement = commentsPart.GetXDocument().Root;
                    var childNodes = partElement.Nodes().ToList();
                    foreach (var item in childNodes)
                        item.Remove();
                    return new PtWordprocessingCommentsPart(ParentWmlDocument, commentsPart.Uri, partElement.Name, partElement.Attributes(), childNodes);
                }
            }
        }

        public PtMainDocumentPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
            : base(name, values)
        {
            ParentWmlDocument = wmlDocument;
            Add(
                new XAttribute(PtOpenXml.Uri, uri),
                new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
            );
        }
    }
}