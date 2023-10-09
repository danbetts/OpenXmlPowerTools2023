using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools.Commons;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// Wordprocessing package used during building to simplify method calls and data access
    /// </summary>
    public class WmlPackage : IPackage
    {
        #region IPackage
        public WordprocessingDocument Target { get; set; }
        public MainDocumentPart Main => Target.MainDocumentPart;
        public XDocument MainPart => Target.GetMainPart();
        public T GetPart<T>() where T : OpenXmlPart, IFixedContentTypePart
        {
            return Target.GetPart<T>();
        }
        public XElement Body => Main.GetBody();
        public IEnumerable<XElement> Contents { get; set; }
        public XElement Section { get; set; }
        public IList<WmlSource> Sources { get; set; } = new List<WmlSource>();
        IList<ISource> IPackage.Sources
        {
            get => Sources.Cast<ISource>().ToList();
            set => Sources = value.Cast<WmlSource>().ToList();
        }
        public IList<ImageData> Images { get; set; } = new List<ImageData>();
        private IDictionary<XName, XName[]> relationshipMarkup { get; set; }
        public IDictionary<XName, XName[]> RelationshipMarkup
        {
            get => relationshipMarkup = relationshipMarkup ?? Wordprocessing.RelationshipMarkup;
            private set => relationshipMarkup = value;
        }
        public bool HasSources { get => Sources?.Any() == true; }
        public bool KeepNoSections { get => Sources.All(p => p.KeepSections == false); }
        public bool KeepAllSections { get => Sources.All(p => p.KeepSections == true); }
        public bool KeepNoHeadersAndFooters { get => Sources.All(p => p.KeepHeadersAndFooters == false); }
        public bool KeepAllHeadersAndFooters { get => Sources.All(p => p.KeepHeadersAndFooters == true); }

        public IEnumerable<XElement> GetContents(int start, int count) => Main.GetContents(start, count);
        public IPackage SetSource(TypedOpenXmlPackage source)
        {
            if (source.GetType() != typeof(WordprocessingDocument)) throw new InvalidCastException($"{source.GetType().Name} is not a word processing document.");

            Target = source as WordprocessingDocument;
            return this;
        }
        #endregion
    }
}