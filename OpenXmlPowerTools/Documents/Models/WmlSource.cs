using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System;

namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// Source for documents used in a wordprocessingdocument merge
    /// </summary>
    public class WmlSource : ISource
    {
        #region ISource
        public WmlDocument WmlDocument { get; set; }
        OpenXmlPowerToolsDocument ISource.WmlDocument
        {
            get => WmlDocument;
            set => WmlDocument = (WmlDocument)Convert.ChangeType(value, typeof(WmlDocument));
        }
        public WordprocessingDocument Document { get; set; }
        public int Start { get; set; } = 0;
        public int Count { get; set; } = int.MaxValue;
        public bool KeepHeadersAndFooters { get; set; } = true;
        public bool KeepSections { get; set; } = true;
        public string InsertId { get; set; } = null;
        #endregion

        public WmlSource() { }
        public WmlSource(string filename) : this(new WmlDocument(filename), 0, int.MaxValue) { }
        public WmlSource(string filename, int start) : this(new WmlDocument(filename), start, int.MaxValue) { }
        public WmlSource(string filename, bool keepSections) : this(new WmlDocument(filename), keepSections) { }
        public WmlSource(string filename, int start, int count) : this(new WmlDocument(filename), start, count) { }
        public WmlSource(string filename, int start, bool keepSections) : this(new WmlDocument(filename), start, keepSections) { }
        public WmlSource(string filename, int start, int count, bool keepSections) : this(new WmlDocument(filename), start, count, keepSections) { }
        public WmlSource(WmlDocument document) : this(document, 0) { }
        public WmlSource(WmlDocument document, int start) : this(document, start, int.MaxValue) { }
        public WmlSource(WmlDocument doucment, bool keepSections) : this(doucment, 0, int.MaxValue, keepSections) { }
        public WmlSource(WmlDocument document, int start, int count) : this(document, start, count, true) { }
        public WmlSource(WmlDocument document, int start, bool keepSections) : this(document, start, int.MaxValue, keepSections) { }
        public WmlSource(WmlDocument document, int start, int count, bool keepSections)
        { 
            this.WmlDocument = document;
            this.Start = start;
            this.Count = count;
            this.KeepSections = KeepSections;
        }
    }
}