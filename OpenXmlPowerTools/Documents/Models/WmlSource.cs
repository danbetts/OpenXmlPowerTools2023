using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System;

namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// Source for documents used in a wordprocessingdocument merge
    /// </summary>
    public class WmlSource
    {
        public WmlDocument WmlDocument { get; set; }
        public int Start { get; set; } = 0;
        public int Count { get; set; } = int.MaxValue;
        public bool KeepSections { get; set; } = true;
        public bool KeepHeadersAndFooters { get; set; } = true;
        public string InsertId { get; set; } = null;

        public WmlSource() { }
        public WmlSource(string filename)
            : this(new WmlDocument(filename), 0, int.MaxValue, true, true, null) { }
        public WmlSource(string filename, int start)
            : this(new WmlDocument(filename), start, int.MaxValue, true, true, null) { }
        public WmlSource(string filename, int start, int count)
            : this(new WmlDocument(filename), start, count, true, true, null) { }
        public WmlSource(string filename, int start, int count, bool keepSections)
            : this(new WmlDocument(filename), start, count, keepSections, true, null) { }
        public WmlSource(WmlDocument document)
            : this(document, 0, int.MaxValue, true, true, null) { }
        public WmlSource(WmlDocument document, int start)
            : this(document, start, int.MaxValue, true, true, null) { }
        public WmlSource(WmlDocument document, int start, int count)
            : this(document, start, count, true, true, null) { }
        public WmlSource(WmlDocument document, int start, int count, bool keepSections)
            : this(document, start, count, keepSections, true, null) { }
        public WmlSource(WmlDocument document, int start, int count, bool keepSections, bool keepHeadersAndFooters, string insertId)
        {
            this.WmlDocument = document;
            this.Start = start;
            this.Count = count;
            this.KeepSections = keepSections;
            this.KeepHeadersAndFooters = keepHeadersAndFooters;
            this.InsertId = insertId;
        }
    }
}