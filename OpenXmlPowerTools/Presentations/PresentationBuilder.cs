using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Spreadsheets;

namespace OpenXmlPowerTools.Presentations
{
    public class PresentationBuilder
    {
        private string _fileName { get; set; } = string.Empty;
        public List<SlideSource> Sources { get; set; } = new List<SlideSource>();
        private int _start { get; set; } = 0;
        private int _count { get; set; } = 0;
        private HashSet<string> _customXmlGuidList { get; set; } = null;
        private bool _normalizeStyleIds { get; set; } = false;
        private static Dictionary<XName, XName[]> _relationshipMarkup { get; set; } = null;

        /// <summary>
        /// Set file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public PresentationBuilder FileName(string fileName)
        {
            _fileName = fileName;
            return this;
        }

        /// <summary>
        /// Set sources to a collection of sources
        /// </summary>
        /// <param name="sources"></param>
        /// <returns></returns>
        public PresentationBuilder SetSources(IEnumerable<SlideSource> sources)
        {
            Sources = sources.ToList();
            return this;
        }

        /// <summary>
        /// Add a new source
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public PresentationBuilder AddSource(SlideSource source)
        {
            Sources.Add(source);
            return this;
        }

        /// <summary>
        /// Add a range of new sources to existing sources
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public PresentationBuilder AppendSources(IEnumerable<SlideSource> source)
        {
            Sources.AddRange(source);
            return this;
        }

        /// <summary>
        /// Build and saves presentation document
        /// </summary>
        public static void Save()
        {

        }

        /// <summary>
        /// Build PresentationDocument
        /// </summary>
        /// <returns></returns>
        //public static PresentationDocument Build()
        //{
        //    return Extensions.CreatePresentationDocument();
        //}
    }
}
