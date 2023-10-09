using System;

namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// WmlSource builder
    /// </summary>
    public class WmlSourceBuilder
    {
        private WmlDocument _document { get; set; } = null;
        private int _start { get; set; } = 0;
        private int _count { get; set; } = int.MaxValue;
        private bool _contentOnly { get; set; } = false;
        private bool _inheritLayout { get; set; } = false;
        private bool _keepHeadersAndFooters { get; set; } = true;
        private bool _keepSections { get; set; } = true;
        private string _insertId { get; set; } = null;

        public WmlSourceBuilder() { }
        public WmlSourceBuilder(string fileName) : this(new WmlDocument(fileName)) { }

        public WmlSourceBuilder(WmlDocument document)
        {
            _document = document;
        }

        /// <summary>
        /// Set file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public WmlSourceBuilder FileName(string fileName)
        {
            _document = new WmlDocument(fileName);
            return this;
        }

        /// <summary>
        /// Set OpenXml Wordprocessing document 
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public WmlSourceBuilder Document(WmlDocument document)
        {
            _document = document;
            return this;
        }

        /// <summary>
        /// Set starting body part
        /// </summary>
        /// <param name="start"></param>
        /// <returns></returns>
        public WmlSourceBuilder Start(int start)
        {
            _start = start;
            return this;
        }

        /// <summary>
        /// Set number of body parts to include
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public WmlSourceBuilder Count(int count)
        {
            _count = count;
            return this;
        }

        /// <summary>
        /// Set content only
        /// </summary>
        /// <param name="contentOnly"></param>
        /// <returns></returns>
        public WmlSourceBuilder ContentOnly(bool contentOnly = true)
        {
            _contentOnly = contentOnly;
            _keepSections = false;
            return this;
        }

        /// <summary>
        /// Set inherit layout
        /// </summary>
        /// <param name="inheritLayout"></param>
        /// <returns></returns>
        public WmlSourceBuilder InheritLayout(bool inheritLayout = true)
        {
            _inheritLayout = inheritLayout;
            return this;
        }

        /// <summary>
        /// Set keep headers and footers
        /// </summary>
        /// <param name="keepHeadersAndFooters"></param>
        /// <returns></returns>
        public WmlSourceBuilder KeepHeadersAndFooters(bool keepHeadersAndFooters = true)
        {
            _keepHeadersAndFooters = keepHeadersAndFooters;
            return this;
        }

        /// <summary>
        /// Set keep sections
        /// </summary>
        /// <param name="keepSections"></param>
        /// <returns></returns>
        public WmlSourceBuilder KeepSections(bool keepSections = true)
        {
            _keepSections = keepSections;
            return this;
        }

        /// <summary>
        /// Processes and saves document
        /// </summary>
        public WmlSource Build()
        {
            if (_document == null) throw new ArgumentException("Document cannot be null.");

            return new WmlSource()
            {
                WmlDocument = _document,
                Start = _start,
                Count = _count,
                ContentOnly = _contentOnly,
                InheritLayout = _inheritLayout,
                KeepHeadersAndFooters = _keepHeadersAndFooters,
                KeepSections = _keepSections,
                InsertId = _insertId,
            };
        }
    }
}