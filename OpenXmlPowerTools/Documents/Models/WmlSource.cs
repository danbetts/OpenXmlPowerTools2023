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
        public bool ContentOnly { get; set; } = false;
        public bool InheritLayout { get; set; } = true;
        public bool KeepHeadersAndFooters { get; set; } = true;
        public bool KeepSections { get; set; } = true;
        public string InsertId { get; set; } = null;
        #endregion

        public WmlSource() { }
        public WmlSource(string filename) : this(new WmlDocument(filename), 0, int.MaxValue) { }
        public WmlSource(string filename, int start) : this(new WmlDocument(filename), start, int.MaxValue) { }
        public WmlSource(string filename, int start, int count) : this(new WmlDocument(filename), start, count) { }
        public WmlSource(WmlDocument document) : this(document, 0, int.MaxValue) { }
        public WmlSource(WmlDocument document, int start) : this(document, start, int.MaxValue) { }
        public WmlSource(WmlDocument document, int start, int count)
        {
            this.WmlDocument = document;
            this.Start = start;
            this.Count = count;
        }
    }
}