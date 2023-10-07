using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System;
using System.Collections.Generic;
using System.Linq;
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
        public WordprocessingDocument Source { get; set; }
        public IList<WmlSource> Sources { get; set; } = new List<WmlSource>();
        IList<ISource> IPackage.Sources
        {
            get => Sources.Cast<ISource>().ToList();
            set => Sources = value.Cast<WmlSource>().ToList();
        }
        public IList<ImageData> Images { get; set; } = new List<ImageData>();
        public IEnumerable<XElement> Contents { get; set; }
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

        public IPackage SetSource(TypedOpenXmlPackage source)
        {
            if (source.GetType() != typeof(WordprocessingDocument)) throw new InvalidCastException($"{source.GetType().Name} is not a word processing document.");

            Target = source as WordprocessingDocument;
            return this;
        }
        #endregion

    }
}