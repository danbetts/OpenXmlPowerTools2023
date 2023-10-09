using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    public interface IDocument
    {
        WordprocessingDocument Document { get; }
        XElement Body { get; }
        string[] Extensions { get; }
        XDocument MainPart { get; }
        MainDocumentPart Main { get; }
        XElement Root { get; }
        IEnumerable<FooterPart> FooterParts { get; }
        IEnumerable<HeaderPart> HeaderParts { get; }
        FontTablePart FontTablePart { get; }
        XDocument FontFamilyTablePart { get; }
    }
}