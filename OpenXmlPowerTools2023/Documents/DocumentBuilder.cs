using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    public class DocumentBuilder
    {
        private string output { get; set; } = string.Empty;
        public List<WmlSource> Sources { get; set; } = new List<WmlSource>();
        private HashSet<string> customXmlGuidList { get; set; } = null;
        private bool normalizeStyleIds { get; set; } = true;

        #region Settings
        /// <summary>
        /// Set output path
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public DocumentBuilder Output(string output)
        {
            this.output = output;
            return this;
        }

        /// <summary>
        /// Set sources to a collection of sources
        /// </summary>
        /// <param name="sources"></param>
        /// <returns></returns>
        public DocumentBuilder SetSources(IList<WmlSource> sources)
        {
            Sources = sources.ToList();
            return this;
        }

        /// <summary>
        /// Add a new srouce by path, and every other optional field
        /// </summary>
        /// <param name="source"></param>
        /// <param name="start"></param>
        /// <param name="count"></param>
        /// <param name="keepSections"></param>
        /// <param name="keepHeadersAndFooters"></param>
        /// <param name="insertId"></param>
        /// <returns></returns>
        public DocumentBuilder AddSource(string source, int start = 0, int count = int.MaxValue, bool keepSections = true, bool keepHeadersAndFooters = true, string insertId = null) 
        {
            AddSource(new WmlDocument(source), 0, count, keepSections, keepHeadersAndFooters, insertId);
            return this;
        }

        /// <summary>
        /// Add a new source
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public DocumentBuilder AddSource(WmlDocument document, int start = 0, int count = int.MaxValue, bool keepSections = true, bool keepHeadersAndFooters = true, string insertId = null)
        {
            AddSource(new WmlSource(document, start, count, keepSections, keepHeadersAndFooters, insertId));
            return this;
        }

        /// <summary>
        /// Add a new source
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public DocumentBuilder AddSource(WmlSource source)
        {
            Sources.Add(source); 
            return this;
        }

        /// <summary>
        /// Add a range of new sources to existing sources
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public DocumentBuilder AppendSources(IList<WmlSource> source)
        {
            Sources.AddRange(source); 
            return this;
        }

        /// <summary>
        /// Normalise Style Ids
        /// </summary>
        /// <param name="normalizeStyleIds"></param>
        /// <returns></returns>
        public DocumentBuilder NormalStyleIds(bool normalizeStyleIds)
        {
            this.normalizeStyleIds = normalizeStyleIds;
            return this;
        }

        /// <summary>
        /// Use custom Xml GUIDs
        /// </summary>
        /// <param name="customXmlGuidList"></param>
        /// <returns></returns>
        public DocumentBuilder CustomXmlGuidList(HashSet<string> customXmlGuidList)
        {
            this.customXmlGuidList = customXmlGuidList;
            return this;
        }

        /// <summary>
        /// Reset document builder to defaults
        /// </summary>
        /// <returns></returns>
        public DocumentBuilder Reset() => new DocumentBuilder();
        #endregion

        #region Process
        /// Build document
        /// </summary>
        public void Save()
        {
            using (MemoryStreamDocument streamDoc = Wordprocessing.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument target = streamDoc.GetWordprocessingDocument())
                {
                    Process(target);
                    target.Close();
                }
                streamDoc.GetModifiedDocument().SaveAs(output);
            }
        }

        public void SaveAs(string filename)
        {
            output = filename;
            Save();
        }

        public WmlDocument ToWmlDocument()
        {
            using (MemoryStreamDocument streamDoc = Wordprocessing.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument target = streamDoc.GetWordprocessingDocument())
                {
                    Process(target);
                    target.Close();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public WordprocessingDocument Build()
        {
            WordprocessingDocument result;
            using (MemoryStreamDocument streamDoc = Wordprocessing.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument target = streamDoc.GetWordprocessingDocument())
                {
                    Process(target);
                    result = target.Clone(output, isEditable: true) as WordprocessingDocument;
                    target.Close();
                }
            }
            return result;
        }

        private void Process(WordprocessingDocument target)
        {
            WmlPackage package = new WmlPackage();
            package.Target = target;
            package.Sources = Sources;
            var targetMain = target.GetMainPart();
            targetMain.Declaration.SetDeclaration();
            targetMain.Root.ReplaceWith(Wordprocessing.CreateRoot());

            if (package.HasSources)
            {
                if (normalizeStyleIds) package.Sources.NormaliseStyleNamesAndIds();
                package.CopyFirstSourceCoreParts();
                HandleInsertId();

                if (package.KeepNoSections)
                {
                    package.RemoveAllSectionsExceptLastKept();
                }
                else HandleKeepSections();

                HandleHeadersAndFooters();

                if (package.Sources.Any(s => s.KeepSections == true))
                {
                    package.CopyAllSections();
                }

                target.AdjustDocPrIds();
                //target.NormalisePageLayout(package);
            }
            HandleGlossary();
            PutAllChanges();

            void HandleInsertId()
            {
                for (int index = 0; index < package.Sources.Count; index++)
                {
                    var source = Sources[index];
                    var sourceId = source.InsertId;
                    var idMatch = targetMain.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == sourceId);

                    using (MemoryStreamDocument streamDoc = new MemoryStreamDocument(source.WmlDocument))
                    using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                    {
                        
                        // Prepare source for merging
                        doc.TestForUnsupportedDocument(package.Sources.IndexOf(source));

                        // Case: Keep section
                        if (source.KeepSections)
                        {
                            // Case: but remove header and footer
                            if (!source.KeepHeadersAndFooters) doc.RemoveHeadersAndFootersFromSections();
                            
                            doc.LinkToPreviousHeadersAndFooters();
                        }
                        else
                        {
                            // Case: Remove section. Since header and footers are imbedded into section they are removed with the section.
                            doc.RemoveSections();

                            // Case: maybe a case to handle here where header and footers are kept by cloning previous section and updating header and footer
                        }

                        if (doc.MainDocumentPart.GetBody() == null) throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains no body element in the correct namespace", index));

                        // Merge in source
                        var contents = doc.MainDocumentPart.GetContents(source.Start, source.Count);
                        try
                        {
                            doc.AppendDocument(package, source, contents);
                        }
                        catch (DocumentBuilderInternalException dbie)
                        {
                            if (dbie.Message.Contains("{0}")) throw new DocumentBuilderException(string.Format(dbie.Message, index));
                            else throw dbie;
                        }
                    }
                }
            }
            void HandleKeepSections()
            {
                //target.FixSectionProperties(); // TODO this code breaks document layout

                var sections = targetMain.Descendants(W.sectPr).ToList();

                CachedHeaderFooter[] cachedHeaderFooter = Wordprocessing.CachedHeadersAndFooters;

                for (int index = 0; index < sections.Count; index++)
                {
                    var sect = sections[index];
                    foreach (var item in cachedHeaderFooter)
                    {
                        if (index == 0)
                        {
                            var referenceElement = sect.Elements(item.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == item.Type);
                            if (referenceElement != null) item.CachedPartRid = (string)referenceElement.Attribute(R.id);
                        }
                        else target.CopyOrCacheHeaderOrFooter(cachedHeaderFooter, sect, item.Ref, item.Type);
                    }
                }
            }
            void HandleHeadersAndFooters()
            {
                for (int index = 0; index < Sources.Count; index++)
                {
                    var source = Sources[index];
                    var sourceId = source.InsertId;

                    if (sourceId != null && (target.MainDocumentPart.HeaderParts.Any(hp => hp.GetXDocument().Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == sourceId)) ||
                        target.MainDocumentPart.FooterParts.Any(fp => fp.GetXDocument().Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == sourceId))))
                    {
                        using (MemoryStreamDocument streamDoc = new MemoryStreamDocument(source.WmlDocument))
                        using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                        {
                            var body = doc.MainDocumentPart.GetBody();
                            doc.TestForUnsupportedDocument(index);

                            var partList = target.MainDocumentPart.HeaderParts.Cast<OpenXmlPart>().Concat(target.MainDocumentPart.FooterParts.Cast<OpenXmlPart>()).ToList();
                            foreach (var part in partList)
                            {
                                if (!part.GetXDocument().Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == sourceId)) continue;
                                List<XElement> contents = doc.MainDocumentPart.GetContents(source.Start, source.Count).ToList();

                                try
                                {
                                    doc.AppendPart(package, contents, source, part);
                                }
                                catch (DocumentBuilderInternalException dbie)
                                {
                                    if (dbie.Message.Contains("{0}")) throw new DocumentBuilderException(string.Format(dbie.Message, index));
                                    else throw dbie;
                                }
                            }
                        }
                    }
                }
            }
            void HandleGlossary()
            {
                WmlDocument wmlGlossaryDocument = Sources.CoalesceGlossaryDocumentParts(package);
                if (wmlGlossaryDocument != null) wmlGlossaryDocument.CopyGlossaryDocumentPart(package);
            }
            void PutAllChanges()
            {
                foreach (var part in target.GetAllParts())
                {
                    if (part.Annotation<XDocument>() != null)
                    {
                        part.PutXDocument();
                    }
                }
            }
        }
        #endregion
    }
}