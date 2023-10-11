using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Documents;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Commons
{
    public interface IPackage
    {
        IList<ISource> Sources { get; set; }
        IList<ImageData> Images { get; set; }
        IEnumerable<XElement> Contents { get; set; }
        XElement Section { get; set; }
        IDictionary<XName, XName[]> RelationshipMarkup { get; }
        bool HasSources { get; }
        bool KeepNoSections { get; }
        bool KeepAllSections { get; }
        bool KeepNoHeadersAndFooters { get; }
        bool KeepAllHeadersAndFooters { get; }
        IPackage SetSource(TypedOpenXmlPackage target);

    }
}