using System;
using System.Xml.Linq;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;

namespace OpenXmlPowerTools.Presentations
{
    public class PtWordprocessingCommentsPart : XElement
    {
        private WmlDocument ParentWmlDocument;

        public PtWordprocessingCommentsPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
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