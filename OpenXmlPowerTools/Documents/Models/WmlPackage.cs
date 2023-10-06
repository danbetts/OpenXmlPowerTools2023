using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// Wordprocessing package used during building to simplify method calls and data access
    /// </summary>
    public  class WmlPackage
    {
        public WordprocessingDocument Document { get; set; }
        public MainDocumentPart MainPart => Document.MainDocumentPart;
        public XDocument Main => Document.GetMainPart();
        public XElement Root => Main.Root;
        public XElement Body => Root.Element(W.body);
        public IEnumerable<XElement> Children => Main.GetBodyElements();
        public IList<WmlSource> Sources { get; set; } = new List<WmlSource>();
        public IList<ImageData> Images { get; set; } = new List<ImageData>();
        private Dictionary<XName, XName[]> relationshipMarkup { get; set; }
        public Dictionary<XName, XName[]> RelationshipMarkup
        {
            get => relationshipMarkup = relationshipMarkup ?? Wordprocessing.RelationshipMarkup;
            private set => relationshipMarkup = value;
        }
        public string[] extensions { get; set; }
        public string[] Extensions
        {
            get => extensions = extensions ?? Wordprocessing.Extensions;
            private set => extensions = value;
        }
        public bool HasSources() => Sources?.Any() == true;
        public bool KeepNoSections() => Sources.All(p => p.KeepSections == false);
        public bool KeepAllSections() => Sources.All(p => p.KeepSections == true);
        public bool KeepNoHeadersAndFooters() => Sources.All(p => p.KeepHeadersAndFooters == false);
        public bool KeepAllHeadersAndFooters() => Sources.All(p => p.KeepHeadersAndFooters == true);
        public WmlPackage SetDocument(WordprocessingDocument document)
        {
            Document = document;
            return this;
        }
    }
}