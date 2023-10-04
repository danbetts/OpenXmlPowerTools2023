namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// Source for documents used in a wordprocessingdocument merge
    /// </summary>
    public class WmlSource
    {
        public WmlDocument WmlDocument { get; set; }
        public int Start { get; set; }
        public int Count { get; set; }
        public bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }
        public string InsertId { get; set; }

        public WmlSource(string fileName)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public WmlSource(WmlDocument source)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public WmlSource(string fileName, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public WmlSource(WmlDocument source, bool keepSections)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public WmlSource(string fileName, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public WmlSource(WmlDocument source, string insertId)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public WmlSource(string fileName, int start, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public WmlSource(WmlDocument source, int start, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public WmlSource(string fileName, int start, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public WmlSource(WmlDocument source, int start, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public WmlSource(string fileName, int start, int count, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public WmlSource(WmlDocument source, int start, int count, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public WmlSource(string fileName, int start, int count, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }

        public WmlSource(WmlDocument source, int start, int count, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }
    }
}