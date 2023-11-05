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
        PowerToolsDocument ISource.WmlDocument
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
        public WmlSource(string filename) : this(new WmlDocument(filename)) { }
        public WmlSource(string filename, int start, int count) : this(new WmlDocument(filename), start, count) { }
        public WmlSource(string filename, int start, int count, bool keepSections, bool keepHeaderAndFooters, string insertId)
            : this(new WmlDocument(filename), start, count, keepSections, keepHeaderAndFooters, insertId) { }
        public WmlSource(WmlDocument document) 
        {
            WmlDocument = document;
        }
        public WmlSource(WmlDocument doucment, int start, int count)
        {
            WmlDocument = doucment;
            Start = start;
            Count = count;
        }

        public WmlSource(WmlDocument document, int start, int count, bool keepSections, bool keepHeaderAndFooters, string insertId)
        { 
            WmlDocument = document;
            Start = start;
            Count = count;
            KeepSections = keepSections;
            KeepHeadersAndFooters = keepHeaderAndFooters;
            InsertId = insertId;
        }
    }
}