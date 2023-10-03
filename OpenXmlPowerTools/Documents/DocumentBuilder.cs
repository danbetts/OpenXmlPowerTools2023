using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    public class DocumentBuilder
    {
        private string _fileName { get; set; } = string.Empty;
        public List<Source> Sources { get; set; } = new List<Source>();
        private int _start { get; set; } = 0;
        private int _count { get; set; } = 0;
        private HashSet<string> _customXmlGuidList { get; set; } = null;
        private bool _normalizeStyleIds { get; set; } = false;
        private static Dictionary<XName, XName[]> _relationshipMarkup = Constants.WordprocessingRelationshipMarkup;

        /// <summary>
        /// Set file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public DocumentBuilder FileName(string fileName)
        {
            _fileName = fileName;
            return this;
        }

        /// <summary>
        /// Set sources to a collection of sources
        /// </summary>
        /// <param name="sources"></param>
        /// <returns></returns>
        public DocumentBuilder SetSources(IEnumerable<Source> sources)
        {
            Sources = sources.ToList();
            return this;
        }

        /// <summary>
        /// Add a new source
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public DocumentBuilder AddSource(Source source)
        {
            Sources.Add(source);
            return this;
        }

        /// <summary>
        /// Add a range of new sources to existing sources
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public DocumentBuilder AppendSources(IEnumerable<Source> source)
        {
            Sources.AddRange(source); 
            return this;
        }

        /// <summary>
        /// Set starting body part
        /// </summary>
        /// <param name="start"></param>
        /// <returns></returns>
        public DocumentBuilder Start(int start)
        {
            _start = start;
            return this;
        }

        /// <summary>
        /// Set number of body parts to include
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public DocumentBuilder Count(int count)
        {
            _count = count;
            return this;
        }

        /// <summary>
        /// Normalise Style Ids
        /// </summary>
        /// <param name="normalizeStyleIds"></param>
        /// <returns></returns>
        public DocumentBuilder NormalStyleIds(bool normalizeStyleIds)
        {
            _normalizeStyleIds = normalizeStyleIds;
            return this;
        }

        /// <summary>
        /// Use custom Xml GUIDs
        /// </summary>
        /// <param name="customXmlGuidList"></param>
        /// <returns></returns>
        public DocumentBuilder CustomXmlGuidList(HashSet<string> customXmlGuidList)
        {
            _customXmlGuidList = customXmlGuidList;
            return this;
        }

        /// <summary>
        /// Processes and saves document
        /// </summary>
        public static void Create()
        {

        }

        /// <summary>
        /// Processes and returns processed WordprocessingDocument
        /// </summary>
        /// <returns></returns>
        //public static WordprocessingDocument Docx()
        //{
        //    return Extensions.CreateWordprocessingDocument();
        //}



        public static void BuildDocument(List<Source> sources, string fileName)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = Wordprocessing.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                {
                    BuildDocument(sources, output);
                    output.Close();
                }
                streamDoc.GetModifiedDocument().SaveAs(fileName);
            }
        }

        public static WmlDocument BuildDocument(List<Source> sources)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = Wordprocessing.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                {
                    BuildDocument(sources, output);
                    output.Close();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        private static void BuildDocument(List<Source> sources, WordprocessingDocument output)
        {
            WmlDocument wmlGlossaryDocument = sources.CoalesceGlossaryDocumentParts();

            // This list is used to eliminate duplicate images
            List<ImageData> images = new List<ImageData>();
            XDocument mainPart = output.MainDocumentPart.GetXDocument();
            mainPart.Declaration.SetDeclaration();
            mainPart.Root.ReplaceWith(
                new XElement(W.document, Constants.NamespaceAttributes,
                    new XElement(W.body)));
            if (sources.Count > 0)
            {
                // the following function makes sure that for a given style name, the same style ID is used for all documents.
                //if (settings != null && settings.NormalizeStyleIds)
                //    sources = sources.NormalizeStyleNamesAndIds();

                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument))
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    doc.CopyStartingParts(output, images);
                    doc.CopySpecifiedCustomXmlParts(output);
                }

                int sourceNum2 = 0;
                foreach (Source source in sources)
                {
                    if (source.InsertId != null)
                    {
                        while (true)
                        {
#if false
                            modify AppendDocument so that it can take a part.
                            for each in main document part, header parts, footer parts
                                are there any PtOpenXml.Insert elements in any of them?
                            if so, then open and process all.
#endif
                            bool foundInMainDocPart = false;
                            XDocument mainXDoc = output.MainDocumentPart.GetXDocument();
                            if (mainXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                                foundInMainDocPart = true;
                            if (foundInMainDocPart)
                            {
                                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                                {
#if TestForUnsupportedDocuments
                                    // throws exceptions if a document contains unsupported content
                                    TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                                    if (foundInMainDocPart)
                                    {
                                        if (source.KeepSections && source.DiscardHeadersAndFootersInKeptSections)
                                            doc.RemoveHeadersAndFootersFromSections();
                                        else if (source.KeepSections)
                                            doc.ProcessSectionsForLinkToPreviousHeadersAndFooters();

                                        List<XElement> contents = doc.MainDocumentPart.GetXDocument()
                                            .Root
                                            .Element(W.body)
                                            .Elements()
                                            .Skip(source.Start)
                                            .Take(source.Count)
                                            .ToList();

                                        try
                                        {
                                            doc.AppendDocument(output, contents, source.KeepSections, source.InsertId, images);
                                        }
                                        catch (DocumentBuilderInternalException dbie)
                                        {
                                            if (dbie.Message.Contains("{0}"))
                                                throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum2));
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
                    else
                    {
                        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                        using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                        {
#if TestForUnsupportedDocuments
                            // throws exceptions if a document contains unsupported content
                            TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                            if (source.KeepSections && source.DiscardHeadersAndFootersInKeptSections)
                                doc.RemoveHeadersAndFootersFromSections();
                            else if (source.KeepSections)
                                doc.ProcessSectionsForLinkToPreviousHeadersAndFooters();

                            var body = doc.MainDocumentPart.GetXDocument()
                                .Root
                                .Element(W.body);

                            if (body == null)
                                throw new DocumentBuilderException(
                                    String.Format("Source {0} is unsupported document - contains no body element in the correct namespace", sourceNum2));

                            List<XElement> contents = body
                                .Elements()
                                .Skip(source.Start)
                                .Take(source.Count)
                                .ToList();
                            try
                            {
                                doc.AppendDocument(output, contents, source.KeepSections, null, images);
                            }
                            catch (DocumentBuilderInternalException dbie)
                            {
                                if (dbie.Message.Contains("{0}"))
                                    throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum2));
                                else
                                    throw dbie;
                            }
                        }
                    }
                    ++sourceNum2;
                }
                if (!sources.Any(s => s.KeepSections))
                {
                    using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument))
                    using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                    {
                        var body = doc.MainDocumentPart.GetXDocument().Root.Element(W.body);

                        if (body != null && body.Elements().Any())
                        {
                            var sectPr = doc.MainDocumentPart.GetXDocument().Root.Elements(W.body)
                                .Elements().LastOrDefault();
                            if (sectPr != null && sectPr.Name == W.sectPr)
                            {
                                doc.AddSectionAndDependencies(output, sectPr, images);
                                output.MainDocumentPart.GetXDocument().Root.Element(W.body).Add(sectPr);
                            }
                        }
                    }
                }
                else
                {
                    output.FixUpSectionProperties();

                    // Any sectPr elements that do not have headers and footers should take their headers and footers from the *next* section,
                    // i.e. from the running section.
                    var mxd = output.MainDocumentPart.GetXDocument();
                    var sections = mxd.Descendants(W.sectPr).Reverse().ToList();

                    CachedHeaderFooter[] cachedHeaderFooter = new[]
                    {
                        new CachedHeaderFooter() { Ref = W.headerReference, Type = "first" },
                        new CachedHeaderFooter() { Ref = W.headerReference, Type = "even" },
                        new CachedHeaderFooter() { Ref = W.headerReference, Type = "default" },
                        new CachedHeaderFooter() { Ref = W.footerReference, Type = "first" },
                        new CachedHeaderFooter() { Ref = W.footerReference, Type = "even" },
                        new CachedHeaderFooter() { Ref = W.footerReference, Type = "default" },
                    };

                    bool firstSection = true;
                    foreach (var sect in sections)
                    {
                        if (firstSection)
                        {
                            foreach (var hf in cachedHeaderFooter)
                            {
                                var referenceElement = sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type);
                                if (referenceElement != null)
                                    hf.CachedPartRid = (string)referenceElement.Attribute(R.id);
                            }
                            firstSection = false;
                            continue;
                        }
                        else
                        {
                            output.CopyOrCacheHeaderOrFooter(cachedHeaderFooter, sect, W.headerReference, "first");
                            output.CopyOrCacheHeaderOrFooter(cachedHeaderFooter, sect, W.headerReference, "even");
                            output.CopyOrCacheHeaderOrFooter(cachedHeaderFooter, sect, W.headerReference, "default");
                            output.CopyOrCacheHeaderOrFooter(cachedHeaderFooter, sect, W.footerReference, "first");
                            output.CopyOrCacheHeaderOrFooter(cachedHeaderFooter, sect, W.footerReference, "even");
                            output.CopyOrCacheHeaderOrFooter(cachedHeaderFooter, sect, W.footerReference, "default");
                        }

                    }
                }

                // Now can process PtOpenXml:Insert elements in headers / footers
                int sourceNum = 0;
                foreach (Source source in sources)
                {
                    if (source.InsertId != null)
                    {
                        while (true)
                        {
                            //this uses an overload of AppendDocument that takes a part.
                            //for each in main document part, header parts, footer parts
                            //    are there any PtOpenXml.Insert elements in any of them?
                            //if so, then open and process all.
                            bool foundInHeadersFooters = false;
                            if (output.MainDocumentPart.HeaderParts.Any(hp =>
                            {
                                var hpXDoc = hp.GetXDocument();
                                return hpXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId);
                            }))
                                foundInHeadersFooters = true;
                            if (output.MainDocumentPart.FooterParts.Any(fp =>
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

                                    var partList = output.MainDocumentPart.HeaderParts.Cast<OpenXmlPart>().Concat(output.MainDocumentPart.FooterParts.Cast<OpenXmlPart>()).ToList();
                                    foreach (var part in partList)
                                    {
                                        var partXDoc = part.GetXDocument();
                                        if (!partXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                                            continue;
                                        List<XElement> contents = doc.MainDocumentPart.GetXDocument()
                                            .Root
                                            .Element(W.body)
                                            .Elements()
                                            .Skip(source.Start)
                                            .Take(source.Count)
                                            .ToList();

                                        try
                                        {
                                            doc.AppendDocument(output, part, contents, source.KeepSections, source.InsertId, images);
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
                    ++sourceNum;
                }
                if (sources.Any(s => s.KeepSections) && !output.MainDocumentPart.GetXDocument().Root.Descendants(W.sectPr).Any())
                {
                    using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument))
                    using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                    {
                        var sectPr = doc.MainDocumentPart.GetXDocument().Root.Element(W.body)
                            .Elements().LastOrDefault();
                        if (sectPr != null && sectPr.Name == W.sectPr)
                        {
                            doc.AddSectionAndDependencies(output, sectPr, images);
                            output.MainDocumentPart.GetXDocument().Root.Element(W.body).Add(sectPr);
                        }
                    }
                }
                output.AdjustDocPrIds();
            }

            if (wmlGlossaryDocument != null)
                wmlGlossaryDocument.WriteGlossaryDocumentPart(output, images);

            foreach (var part in output.GetAllParts())
                if (part.Annotation<XDocument>() != null)
                    part.PutXDocument();
        }

    }
}