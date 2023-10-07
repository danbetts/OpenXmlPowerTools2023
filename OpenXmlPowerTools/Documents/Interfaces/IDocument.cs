using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    public interface IDocument
    {
        XElement Body { get; }
        string[] Extensions { get; }
        XDocument Main { get; }
        MainDocumentPart MainPart { get; }
        XElement Root { get; }
        IEnumerable<FooterPart> FooterParts { get; }
        IEnumerable<HeaderPart> HeaderParts { get; }
        FontTablePart FontTablePart { get; }
        XDocument FontFamilyTablePart { get; }
    }
}