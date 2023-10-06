using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        /// Build document
        /// </summary>
        public void Save()
        {
            using (OpenXmlMemoryStreamDocument streamDoc = Wordprocessing.CreateWordprocessingDocument())
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

        public WmlDocument Build()
        {
            using (OpenXmlMemoryStreamDocument streamDoc = Wordprocessing.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument target = streamDoc.GetWordprocessingDocument())
                {
                    Process(target);
                    target.Close();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        private void Process(WordprocessingDocument target)
        {
            WmlPackage package = new WmlPackage();
            package.Document = target;
            package.Sources = Sources;
            package.Main.Declaration.SetDeclaration();
            package.Main.Root.ReplaceWith(Wordprocessing.CreateRoot());

            if (package.HasSources())
            {
                if (normalizeStyleIds)
                {
                    package.Sources.NormalizeStyleNamesAndIds();
                }

                package.CopyFirstSourceCoreParts();
                HandleInsertId();

                if (package.KeepNoSections())
                {
                    package.RemoveAllSectionsExceptLast();
                }
                else HandleKeepSections();
                        
                HandleHeadersAndFooters();

                if (Sources.Any(s => s.KeepSections == true))
                {
                    package.CopyAllSections();
                }

                target.AdjustDocPrIds();
            }
            HandleGlossary();
            PutAllChanges();

            void HandleInsertId()
            {
                for (int index = 0; index < Sources.Count; index++)
                {
                    var source = Sources[index];
                    if (string.IsNullOrEmpty(source.InsertId))
                    {
                        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                        using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                        {
                            // throws exceptions if a document contains unsupported content
                            //TestForUnsupportedDocument(doc, sources.IndexOf(source));

                            if (source.KeepSections && !source.KeepHeadersAndFooters)
                                doc.RemoveHeadersAndFootersFromSections();
                            else if (source.KeepSections)
                                doc.LinkToPreviousHeadersAndFooters();

                            var body = doc.GetBody();

                            if (body == null)
                                throw new DocumentBuilderException(
                                    String.Format("Source {0} is unsupported document - contains no body element in the correct namespace", index));

                            var contents = doc.GetContents(source.Start, source.Count);
                            try
                            {
                                doc.AppendDocument(package, contents, source.KeepSections, null);
                            }
                            catch (DocumentBuilderInternalException dbie)
                            {
                                if (dbie.Message.Contains("{0}"))
                                    throw new DocumentBuilderException(string.Format(dbie.Message, index));
                                else
                                    throw dbie;
                            }
                        }
                    }
                    else if (package.Main.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                    {
                        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                        using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                        {
                            // throws exceptions if a document contains unsupported content
                            //TestForUnsupportedDocument(doc, sources.IndexOf(source));

                            if (source.KeepSections && !source.KeepHeadersAndFooters)
                                doc.RemoveHeadersAndFootersFromSections();
                            else if (source.KeepSections)
                                doc.LinkToPreviousHeadersAndFooters();

                            var contents = doc.GetContents(source.Start, source.Count);

                            try
                            {
                                doc.AppendDocument(package, contents, source.KeepSections, source.InsertId);
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
            void HandleKeepSections()
            {
                target.FixSectionProperties();

                var targetMain = package.Main;
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
                int sourceNum = 0;
                for (int index = 0; index < Sources.Count; index++)
                {
                    var source = Sources[index];
                    if (source.InsertId != null)
                    {
                        while (true)
                        {
                            //this uses an overload of AppendDocument that takes a part.
                            //for each in main document part, header parts, footer parts
                            //    are there any PtOpenXml.Insert elements in any of them?
                            //if so, then open and process all.
                            bool foundInHeadersFooters = false;
                            if (target.MainDocumentPart.HeaderParts.Any(hp =>
                            {
                                var hpXDoc = hp.GetXDocument();
                                return hpXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId);
                            }))
                                foundInHeadersFooters = true;
                            if (target.MainDocumentPart.FooterParts.Any(fp =>
                            {
                                var hpXDoc = fp.GetXDocument();
                                return hpXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId);
                            }))
                                foundInHeadersFooters = true;

                            if (foundInHeadersFooters)
                            {
                                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                                {
                                    //// throws exceptions if a document contains unsupported content
                                    //TestForUnsupportedDocument(doc, sources.IndexOf(source));

                                    var partList = target.MainDocumentPart.HeaderParts.Cast<OpenXmlPart>().Concat(target.MainDocumentPart.FooterParts.Cast<OpenXmlPart>()).ToList();
                                    foreach (var part in partList)
                                    {
                                        var partXDoc = part.GetXDocument();
                                        if (!partXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                                            continue;
                                        List<XElement> contents = doc.GetContents(source.Start, source.Count).ToList();

                                        try
                                        {
                                            // append with a part
                                            doc.AppendPart(package, contents, source.KeepSections, source.InsertId, part);
                                        }
                                        catch (DocumentBuilderInternalException dbie)
                                        {
                                            if (dbie.Message.Contains("{0}"))
                                                throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum));
                                            else
                                                throw dbie;
                                        }
                                    }
                                }
                            }
                            else
                                break;
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
    }
}