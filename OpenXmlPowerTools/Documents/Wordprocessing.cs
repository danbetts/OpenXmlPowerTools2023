using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using FontFamily = System.Drawing.FontFamily;

namespace OpenXmlPowerTools.Documents
{
    public static class Wordprocessing
    {
        public static bool IsWordprocessing(string ext) => Extensions.Contains(ext.ToLower());
        public static int StringToTwips(string twipsOrPoints)
        {
            // if the pos value is in points, not twips
            if (twipsOrPoints.EndsWith("pt"))
            {
                decimal decimalValue = decimal.Parse(twipsOrPoints.Substring(0, twipsOrPoints.Length - 2));
                return (int)(decimalValue * 20);
            }
            return int.Parse(twipsOrPoints);
        }

        #region XDocument / XElement / XAttribute / XNode
        public static void AddToIgnorable(this XElement root, string v)
        {
            var ignorable = root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                var val = (string)ignorable;
                val = val + " " + v;
                ignorable.Remove();
                root.SetAttributeValue(MC.Ignorable, val);
            }
        }
        public static void AddReferenceToExistingHeaderOrFooter(this XElement sect, string rId, XName reference, string toType)
        {
            if (reference == W.headerReference)
            {
                var referenceToAdd = new XElement(W.headerReference, new XAttribute(W.type, toType), new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
            }
            else
            {
                var referenceToAdd = new XElement(W.footerReference, new XAttribute(W.type, toType), new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
            }
        }
        public static int CalcWidthOfRunInTwips(XElement r)
        {
            HashSet<string> KnownFamilies = null;
            HashSet<string> UnknownFonts = new HashSet<string>();

            if (KnownFamilies == null)
            {
                KnownFamilies = new HashSet<string>();
                var families = FontFamily.Families;
                foreach (var fam in families)
                    KnownFamilies.Add(fam.Name);
            }

            var fontName = (string)r.Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
                fontName = (string)r.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
                throw new OpenXmlPowerToolsException("Internal Error, should have FontName attribute");
            if (UnknownFonts.Contains(fontName))
                return 0;

            var rPr = r.Element(W.rPr);
            if (rPr == null)
                throw new OpenXmlPowerToolsException("Internal Error, should have run properties");
            var languageType = (string)r.Attribute(PtOpenXml.LanguageType);
            decimal? szn = null;
            if (languageType == "bidi")
                szn = (decimal?)rPr.Elements(W.szCs).Attributes(W.val).FirstOrDefault();
            else
                szn = (decimal?)rPr.Elements(W.sz).Attributes(W.val).FirstOrDefault();
            if (szn == null)
                szn = 22m;

            var sz = szn.GetValueOrDefault();

            // unknown font families will throw ArgumentException, in which case just return 0
            if (!KnownFamilies.Contains(fontName))
                return 0;
            // in theory, all unknown fonts are found by the above test, but if not...
            FontFamily ff;
            try
            {
                ff = new FontFamily(fontName);
            }
            catch (ArgumentException)
            {
                UnknownFonts.Add(fontName);

                return 0;
            }
            FontStyle fs = FontStyle.Regular;
            var bold = GetBoolProp(rPr, W.b) || GetBoolProp(rPr, W.bCs);
            var italic = GetBoolProp(rPr, W.i) || GetBoolProp(rPr, W.iCs);
            if (bold && !italic)
                fs = FontStyle.Bold;
            if (italic && !bold)
                fs = FontStyle.Italic;
            if (bold && italic)
                fs = FontStyle.Bold | FontStyle.Italic;

            var runText = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.t)
                .Select(t => (string)t)
                .StringConcatenate();

            var tabLength = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.tab)
                .Select(t => (decimal)t.Attribute(PtOpenXml.TabWidth))
                .Sum();

            if (runText.Length == 0 && tabLength == 0)
                return 0;

            int multiplier = 1;
            if (runText.Length <= 2) multiplier = 100;
            else if (runText.Length <= 4) multiplier = 50;
            else if (runText.Length <= 8) multiplier = 25;
            else if (runText.Length <= 16) multiplier = 12;
            else if (runText.Length <= 32) multiplier = 6;
            if (multiplier != 1)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < multiplier; i++)
                    sb.Append(runText);
                runText = sb.ToString();
            }

            var w = MetricsGetter.GetTextWidth(ff, fs, sz, runText);

            return (int)(w / 96m * 1440m / multiplier + tabLength * 1440m);
        }
        public static XElement FindReference(this XElement sect, XName reference, string type)
        {
            return sect.Elements(reference).FirstOrDefault(z =>
            {
                return (string)z.Attribute(W.type) == type;
            });
        }

        public static int? AttributeToTwips(XAttribute attribute)
        {
            if (attribute == null) return null;

            string twipsOrPoints = (string)attribute;

            if (twipsOrPoints.EndsWith("pt")) return (int)(decimal.Parse(twipsOrPoints.Substring(0, twipsOrPoints.Length - 2)) * 20);
            else if (twipsOrPoints.Contains('.')) return (int)decimal.Parse(twipsOrPoints);
            else return int.Parse(twipsOrPoints);
        }
        // This method is a mess
        public static XElement CoalesceAdjacentRunsWithIdenticalFormatting(XElement runContainer)
        {
            const string dontConsolidate = "DontConsolidate";

            IEnumerable<IGrouping<string, XElement>> groupedAdjacentRunsWithIdenticalFormatting =
                runContainer.Elements().GroupAdjacent(ce =>
                {
                    if (ce.Name == W.r)
                    {
                        if (ce.Elements().Count(e => e.Name != W.rPr) != 1)
                            return dontConsolidate;

                        if (ce.Attribute(PtOpenXml.AbstractNumId) != null)
                            return dontConsolidate;

                        XElement rPr = ce.Element(W.rPr);
                        string rPrString = rPr != null ? rPr.ToString(SaveOptions.None) : string.Empty;

                        if (ce.Element(W.t) != null)
                            return "Wt" + rPrString;

                        if (ce.Element(W.instrText) != null)
                            return "WinstrText" + rPrString;

                        return dontConsolidate;
                    }

                    if (ce.Name == W.ins)
                    {
                        if (ce.Elements(W.del).Any())
                        {
                            return dontConsolidate;
#if false
                                // for w:ins/w:del/w:r/w:delText
                                if ((ce.Elements(W.del).Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1) ||
                                    !ce.Elements().Elements().Elements(W.delText).Any())
                                    return dontConsolidate;

                                XAttribute dateIns = ce.Attribute(W.date);
                                XElement del = ce.Element(W.del);
                                XAttribute dateDel = del.Attribute(W.date);

                                string authorIns = (string) ce.Attribute(W.author) ?? string.Empty;
                                string dateInsString = dateIns != null
                                    ? ((DateTime) dateIns).ToString("s")
                                    : string.Empty;
                                string authorDel = (string) del.Attribute(W.author) ?? string.Empty;
                                string dateDelString = dateDel != null
                                    ? ((DateTime) dateDel).ToString("s")
                                    : string.Empty;

                                return "Wins" +
                                       authorIns +
                                       dateInsString +
                                       authorDel +
                                       dateDelString +
                                       ce.Elements(W.del)
                                           .Elements(W.r)
                                           .Elements(W.rPr)
                                           .Select(rPr => rPr.ToString(SaveOptions.None))
                                           .StringConcatenate();
#endif
                        }

                        // w:ins/w:r/w:t
                        if ((ce.Elements().Elements().Count(e => e.Name != W.rPr) != 1) ||
                            !ce.Elements().Elements(W.t).Any())
                            return dontConsolidate;

                        XAttribute dateIns2 = ce.Attribute(W.date);

                        string authorIns2 = (string)ce.Attribute(W.author) ?? string.Empty;
                        string dateInsString2 = dateIns2 != null
                            ? ((DateTime)dateIns2).ToString("s")
                            : string.Empty;

                        string idIns2 = (string)ce.Attribute(W.id);

                        return "Wins2" +
                               authorIns2 +
                               dateInsString2 +
                               idIns2 +
                               ce.Elements()
                                   .Elements(W.rPr)
                                   .Select(rPr => rPr.ToString(SaveOptions.None))
                                   .StringConcatenate();
                    }

                    if (ce.Name == W.del)
                    {
                        if ((ce.Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1) ||
                            !ce.Elements().Elements(W.delText).Any())
                            return dontConsolidate;

                        XAttribute dateDel2 = ce.Attribute(W.date);

                        string authorDel2 = (string)ce.Attribute(W.author) ?? string.Empty;
                        string dateDelString2 = dateDel2 != null ? ((DateTime)dateDel2).ToString("s") : string.Empty;

                        return "Wdel" +
                               authorDel2 +
                               dateDelString2 +
                               ce.Elements(W.r)
                                   .Elements(W.rPr)
                                   .Select(rPr => rPr.ToString(SaveOptions.None))
                                   .StringConcatenate();
                    }

                    return dontConsolidate;
                });

            var runContainerWithConsolidatedRuns = new XElement(runContainer.Name,
                runContainer.Attributes(),
                groupedAdjacentRunsWithIdenticalFormatting.Select(g =>
                {
                    if (g.Key == dontConsolidate)
                        return (object)g;

                    string textValue = g
                        .Select(r =>
                            r.Descendants()
                                .Where(d => (d.Name == W.t) || (d.Name == W.delText) || (d.Name == W.instrText))
                                .Select(d => d.Value)
                                .StringConcatenate())
                        .StringConcatenate();
                    XAttribute xs = Common.GetXmlSpaceAttribute(textValue);

                    if (g.First().Name == W.r)
                    {
                        if (g.First().Element(W.t) != null)
                        {
                            IEnumerable<IEnumerable<XAttribute>> statusAtt =
                                g.Select(r => r.Descendants(W.t).Take(1).Attributes(PtOpenXml.Status));
                            return new XElement(W.r,
                                g.First().Attributes(),
                                g.First().Elements(W.rPr),
                                new XElement(W.t, statusAtt, xs, textValue));
                        }

                        if (g.First().Element(W.instrText) != null)
                            return new XElement(W.r,
                                g.First().Attributes(),
                                g.First().Elements(W.rPr),
                                new XElement(W.instrText, xs, textValue));
                    }

                    if (g.First().Name == W.ins)
                    {
                        XElement firstR = g.First().Element(W.r);
                        return new XElement(W.ins,
                            g.First().Attributes(),
                            new XElement(W.r,
                                firstR?.Attributes(),
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.t, xs, textValue)));
                    }

                    if (g.First().Name == W.del)
                    {
                        XElement firstR = g.First().Element(W.r);
                        return new XElement(W.del,
                            g.First().Attributes(),
                            new XElement(W.r,
                                firstR?.Attributes(),
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.delText, xs, textValue)));
                    }
                    return g;
                }));

            // Process w:txbxContent//w:p
            foreach (XElement txbx in runContainerWithConsolidatedRuns.Descendants(W.txbxContent))
                foreach (XElement txbxPara in txbx.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p))
                {
                    XElement newPara = CoalesceAdjacentRunsWithIdenticalFormatting(txbxPara);
                    txbxPara.ReplaceWith(newPara);
                }

            // Process additional run containers.
            List<XElement> runContainers = runContainerWithConsolidatedRuns
                .Descendants()
                .Where(d => Constants.AdditionalRunContainerNames.Contains(d.Name))
                .ToList();
            foreach (XElement container in runContainers)
            {
                XElement newContainer = CoalesceAdjacentRunsWithIdenticalFormatting(container);
                container.ReplaceWith(newContainer);
            }

            return runContainerWithConsolidatedRuns;
        }
        public static bool GetBoolProp(XElement runProps, XName xName)
        {
            var value = runProps?.Element(xName)?.Attribute(W.val) ?? null;

            if (value == null) return true;
            else
            {
                var lower = value.Value.ToLower();
                return (lower == "1" || lower == "true");
            }
        }
        public static void RemoveGfxdata(this IEnumerable<XElement> newContent)
        {
            newContent.DescendantsAndSelf().Attributes(O.gfxdata).Remove();
        }
        /// <summary>
        /// Fix paragraphs if they do not have the correct start/end for parts
        /// </summary>
        /// <param name="sourceDocument"></param>
        /// <param name="newContent"></param>
        public static void FixRanges(this XDocument sourceDocument, IEnumerable<XElement> newContent)
        {
            FixRange(W.commentRangeStart, W.commentRangeEnd, W.id, W.commentReference);
            FixRange(W.bookmarkStart, W.bookmarkEnd, W.id, null);
            FixRange(W.permStart, W.permEnd, W.id, null);
            FixRange(W.moveFromRangeStart, W.moveFromRangeEnd, W.id, null);
            FixRange(W.moveToRangeStart, W.moveToRangeEnd, W.id, null);
            DeleteUnmatchedRange(W.moveFromRangeStart, W.moveFromRangeEnd, W.moveToRangeStart, W.name, W.id);
            DeleteUnmatchedRange(W.moveToRangeStart, W.moveToRangeEnd, W.moveFromRangeStart, W.name, W.id);

            void DeleteUnmatchedRange(XName startElement, XName endElement, XName matchTo, XName matchAttr, XName idAttr)
            {
                List<string> deleteList = new List<string>();
                foreach (XElement start in newContent.Elements(startElement))
                {
                    string id = start.Attribute(matchAttr).Value;
                    if (!newContent.Elements(matchTo).Where(n => n.Attribute(matchAttr).Value == id).Any())
                        deleteList.Add(start.Attribute(idAttr).Value);
                }
                foreach (string item in deleteList)
                {
                    newContent.Elements(startElement).Where(n => n.Attribute(idAttr).Value == item).Remove();
                    newContent.Elements(endElement).Where(n => n.Attribute(idAttr).Value == item).Remove();
                    newContent.Where(p => p.Name == startElement && p.Attribute(idAttr).Value == item).Remove();
                    newContent.Where(p => p.Name == endElement && p.Attribute(idAttr).Value == item).Remove();
                }
            }
            void FixRange(XName startElement, XName endElement, XName idAttribute, XName refElement)
            {
                foreach (XElement start in newContent.DescendantsAndSelf(startElement))
                {
                    string rangeId = start.Attribute(idAttribute).Value;
                    if (newContent
                        .DescendantsAndSelf(endElement)
                        .Where(e => e.Attribute(idAttribute).Value == rangeId)
                        .Count() == 0)
                    {
                        XElement end = sourceDocument
                            .Descendants(endElement)
                            .Where(o => o.Attribute(idAttribute).Value == rangeId)
                            .FirstOrDefault();
                        if (end != null)
                        {
                            AddAtEnd(new XElement(end));
                            if (refElement != null)
                            {
                                XElement newRef = new XElement(refElement, new XAttribute(idAttribute, rangeId));
                                AddAtEnd(new XElement(newRef));
                            }
                        }
                    }
                }
                foreach (XElement end in newContent.Elements(endElement))
                {
                    string rangeId = end.Attribute(idAttribute).Value;
                    if (newContent
                        .DescendantsAndSelf(startElement)
                        .Where(s => s.Attribute(idAttribute).Value == rangeId)
                        .Count() == 0)
                    {
                        XElement start = sourceDocument
                            .Descendants(startElement)
                            .Where(o => o.Attribute(idAttribute).Value == rangeId)
                            .FirstOrDefault();
                        if (start != null)
                            AddAtBeginning(new XElement(start));
                    }
                }

                void AddAtBeginning(XElement contentToAdd)
                {
                    if (newContent.First().Element(W.pPr) != null)
                        newContent.First().Element(W.pPr).AddAfterSelf(contentToAdd);
                    else
                        newContent.First().AddFirst(new XElement(contentToAdd));
                }
                void AddAtEnd(XElement contentToAdd)
                {
                    if (newContent.Last().Element(W.pPr) != null)
                        newContent.Last().Element(W.pPr).AddAfterSelf(new XElement(contentToAdd));
                    else
                        newContent.Last().Add(new XElement(contentToAdd));
                }
            }
        }
        // This method is a mess
        public static void MergeDocDefaultStyles(this XDocument xDocument, XDocument newXDoc)
        {
            var docDefaultStyles = xDocument.Descendants(W.docDefaults);
            foreach (var docDefaultStyle in docDefaultStyles)
            {
                newXDoc.Root.Add(docDefaultStyle);
            }
        }
        public static void MergeFontTables(this XDocument source, XDocument target)
        {
            foreach (XElement font in source.Root.Elements(W.font))
            {
                if (!target.Root.Elements(W.font).Any(o => o.Attribute(W.name).Value == font.Attribute(W.name).Value))
                {
                    target.Root.Add(new XElement(font));
                }
            }
        }
        public static void MergeLatentStyles(this XDocument source, XDocument target)
        {
            var fromLatentStyles = source.Descendants(W.latentStyles).FirstOrDefault();
            if (fromLatentStyles == null)
                return;

            var toLatentStyles = target.Descendants(W.latentStyles).FirstOrDefault();
            if (toLatentStyles == null)
            {
                var newLatentStylesElement = new XElement(W.latentStyles,
                    fromLatentStyles.Attributes());
                var globalDefaults = target
                    .Descendants(W.docDefaults)
                    .FirstOrDefault();
                if (globalDefaults == null)
                {
                    var firstStyle = target
                        .Root
                        .Elements(W.style)
                        .FirstOrDefault();
                    if (firstStyle == null)
                        target.Root.Add(newLatentStylesElement);
                    else
                        firstStyle.AddBeforeSelf(newLatentStylesElement);
                }
                else
                    globalDefaults.AddAfterSelf(newLatentStylesElement);
            }
            toLatentStyles = target.Descendants(W.latentStyles).FirstOrDefault();
            if (toLatentStyles == null)
                throw new OpenXmlPowerToolsException("Internal error");

            var toStylesHash = new HashSet<string>();
            foreach (var lse in toLatentStyles.Elements(W.lsdException))
                toStylesHash.Add((string)lse.Attribute(W.name));

            foreach (var fls in fromLatentStyles.Elements(W.lsdException))
            {
                var name = (string)fls.Attribute(W.name);
                if (toStylesHash.Contains(name))
                    continue;
                toLatentStyles.Add(fls);
                toStylesHash.Add(name);
            }

            var count = toLatentStyles
                .Elements(W.lsdException)
                .Count();

            toLatentStyles.SetAttributeValue(W.count, count);
        }

        public static object OrderElementsPerStandard(XNode node)
        {
            if (node == null) return node;
            var element = node as XElement;

            if (element.Name == W.p) return ProcessRoot(W.pPr);
            else if (element.Name == W.r) return ProcessRoot(W.rPr);
            else if (element.Name == W.pPr) return ProcessBranch(ParagraphPropertiesOrder);
            else if (element.Name == W.rPr) return ProcessBranch(RunPropertiesOrder);
            else if (element.Name == W.tblPr) return ProcessBranch(TablePropertiesOrder);
            else if (element.Name == W.tcPr) return ProcessBranch(TableCellPropertiesOrder);
            else if (element.Name == W.tcBorders) return ProcessBranch(TableCellBordersOrder);
            else if (element.Name == W.tblBorders) return ProcessBranch(TableBordersOrder);
            else if (element.Name == W.pBdr) return ProcessBranch(ParagraphBordersOrder);
            else if (element.Name == W.settings) return ProcessBranch(SettingsOrder);
            else return new XElement(element.Name, element.Attributes(), element.Nodes().Select(n => OrderElementsPerStandard(n)));

            object ProcessRoot(XName xName)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements(xName).Select(e => (XElement)OrderElementsPerStandard(e)),
                    element.Elements().Where(e => e.Name != xName).Select(e => (XElement)OrderElementsPerStandard(e)));
            }
            object ProcessBranch(Dictionary<XName, int> order)
            {
                return new XElement(element.Name, element.Attributes(), element.Elements()
                   .Select(e => (XElement)OrderElementsPerStandard(e))
                   .OrderBy(e =>
                   {
                       return order.ContainsKey(e.Name)
                       ? order[e.Name]
                       : 999;
                   }));
            }
        }
        #endregion

        #region Creators
        /// <summary>
        /// Create a new WordprocessingDocument as an OpenXml memory stream document
        /// </summary>
        /// <returns></returns>
        public static OpenXmlMemoryStreamDocument CreateWordprocessingDocument()
        {
            MemoryStream stream = new MemoryStream();
            using (WordprocessingDocument doc = stream.CreateWordprocessingDocument())
            {
                doc.Close();
                return new OpenXmlMemoryStreamDocument(stream);
            }
        }

        /// <summary>
        /// Create a new WordprocessingDocument for a specified file path
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static WordprocessingDocument CreateWordprocessingDocument(this string filePath)
        {
            WordprocessingDocument doc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();
            doc.MainDocumentPart.PutXDocument(CreateMainDocumentPart());
            return doc;
        }

        /// <summary>
        /// Create a new WordprocessingDocument from a stream
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static WordprocessingDocument CreateWordprocessingDocument(this Stream stream)
        {
            WordprocessingDocument doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();
            doc.MainDocumentPart.PutXDocument(CreateMainDocumentPart());
            return doc;
        }


        /// <summary>
        /// Create a new main document part
        /// </summary>
        /// <returns></returns>
        public static XDocument CreateMainDocumentPart()
        {
            return new XDocument(CreateDocument());
        }

        /// <summary>
        /// Create new word processing document root
        /// </summary>
        /// <returns></returns>
        public static XElement CreateRoot()
        {
            return new XElement(W.document, Constants.NamespaceAttributes, new XElement(W.body));
        }

        /// <summary>
        /// Create new wordprocessing document part with body
        /// </summary>
        /// <returns></returns>
        public static XElement CreateDocument()
        {
            return new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XElement(W.body));
        }

        /// <summary>
        /// Create new word processing document body part
        /// </summary>
        /// <param name="mainPart"></param>
        /// <returns></returns>
        public static XElement CreateBody(this XDocument mainPart)
        {
            return new XElement(W.body, new XElement(W.docParts, mainPart.Root.Element(W.body).Elements(W.docParts).Elements(W.docPart)));
        }

        /// <summary>
        /// Create new word processing document glossary
        /// </summary>
        /// <param name="mainPart"></param>
        /// <returns></returns>
        public static XDocument CreateGlossary(this XDocument mainPart)
        {
            return new XDocument(Common.CreateDeclaration(), new XElement(W.glossaryDocument, Constants.NamespaceAttributes,
                new XElement(W.docParts, mainPart.Descendants(W.docPart))));
        }
        #endregion

        #region WmlSource
        public static WmlDocument BreakLinkToTemplate(WmlDocument source)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var efpp = wDoc.ExtendedFilePropertiesPart;
                    if (efpp != null)
                    {
                        var xd = efpp.GetXDocument();
                        var template = xd.Descendants(EP.Template).FirstOrDefault();
                        if (template != null)
                            template.Value = "";
                        efpp.PutXDocument();
                    }
                }
                var result = new WmlDocument(source.FileName, ms.ToArray());
                return result;
            }
        }
        // This method is massive and complex (recusive solution?) 
        public static WmlDocument CoalesceGlossaryDocumentParts(this IList<WmlSource> sources, WmlPackage package)
        {
            List<WmlSource> allGlossaryDocuments = sources
                .Select(s => ExtractGlossaryDocument(s.WmlDocument))
                .Where(s => s != null)
                .Select(s => new WmlSource(s)).ToList();

            if (!allGlossaryDocuments.Any()) return null;

            WmlDocument coalescedRaw = new DocumentBuilder().SetSources(allGlossaryDocuments).Modify();

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(coalescedRaw.DocumentByteArray, 0, coalescedRaw.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var mainXDoc = wDoc.MainDocumentPart.GetXDocument();
                    var newBody = CreateBody(mainXDoc);
                    mainXDoc.Root.Element(W.body).ReplaceWith(newBody);
                    wDoc.MainDocumentPart.PutXDocument();
                }
                WmlDocument coalescedGlossaryDocument = new WmlDocument("Coalesced.docx", ms.ToArray());
                return coalescedGlossaryDocument;
            }

            WmlDocument ExtractGlossaryDocument(WmlDocument wmlGlossaryDocument)
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    memoryStream.Write(wmlGlossaryDocument.DocumentByteArray, 0, wmlGlossaryDocument.DocumentByteArray.Length);
                    using (WordprocessingDocument source = WordprocessingDocument.Open(memoryStream, false))
                    {
                        if (source.MainDocumentPart.GlossaryDocumentPart == null)
                            return null;

                        var sourceGlossary = source.MainDocumentPart.GlossaryDocumentPart.GetXDocument();
                        if (sourceGlossary.Root == null)
                            return null;

                        using (MemoryStream outMs = new MemoryStream())
                        {
                            using (WordprocessingDocument target = WordprocessingDocument.Create(outMs, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                            {
                                List<ImageData> images = new List<ImageData>();
                                MainDocumentPart mdp = target.AddMainDocumentPart();
                                var mdpXd = mdp.GetXDocument();
                                XElement root = new XElement(W.document);
                                if (mdpXd.Root == null)
                                    mdpXd.Add(root);
                                else
                                    mdpXd.Root.ReplaceWith(root);
                                root.Add(new XElement(W.body,
                                    sourceGlossary.Root.Elements(W.docParts)));
                                mdp.PutXDocument();

                                var newContent = sourceGlossary.Root.Elements(W.docParts);
                                CopyGlossaryDocumentPartsFromGD(source, target, newContent);
                                source.MainDocumentPart.GlossaryDocumentPart.CopyRelatedPartsForContentParts(mdp, package.RelationshipMarkup, newContent, package.Images);
                            }
                            return new WmlDocument("Glossary.docx", outMs.ToArray());
                        }
                    }
                }
                void CopyGlossaryDocumentPartsFromGD(WordprocessingDocument source, WordprocessingDocument target, IEnumerable<XElement> newContent)
                {
                    // Copy all styles to the new document
                    if (source.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart != null)
                    {
                        XDocument oldStyles = source.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                        if (target.MainDocumentPart.StyleDefinitionsPart == null)
                        {
                            target.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                            XDocument newStyles = target.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                            newStyles.Declaration.SetDeclaration();
                            newStyles.Add(oldStyles.Root);
                            target.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                        }
                        else
                        {
                            XDocument newStyles = target.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                            MergeStyles(source, target, oldStyles, newStyles, newContent);
                            target.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                        }
                    }

                    // Copy fontTable to the new document
                    if (source.MainDocumentPart.GlossaryDocumentPart.FontTablePart != null)
                    {
                        XDocument oldFontTable = source.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                        if (target.MainDocumentPart.FontTablePart == null)
                        {
                            target.MainDocumentPart.AddNewPart<FontTablePart>();
                            XDocument newFontTable = target.MainDocumentPart.FontTablePart.GetXDocument();
                            newFontTable.Declaration.SetDeclaration();
                            newFontTable.Add(oldFontTable.Root);
                            target.MainDocumentPart.FontTablePart.PutXDocument();
                        }
                        else
                        {
                            XDocument newFontTable = target.MainDocumentPart.FontTablePart.GetXDocument();
                            oldFontTable.MergeFontTables(newFontTable);
                            target.MainDocumentPart.FontTablePart.PutXDocument();
                        }
                    }

                    DocumentSettingsPart oldSettingsPart = source.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart;
                    if (oldSettingsPart != null)
                    {
                        DocumentSettingsPart newSettingsPart = target.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                        XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                        oldSettingsPart.AddRelationships(newSettingsPart, package.RelationshipMarkup, new[] { settingsXDoc.Root });
                        source.CopyFootnotesPart(package, settingsXDoc);
                        source.CopyEndnotesPart(package, settingsXDoc);
                        XDocument newXDoc = target.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                        newXDoc.Declaration.SetDeclaration();
                        newXDoc.Add(settingsXDoc.Root);
                        oldSettingsPart.CopyRelatedPartsForContentParts(newSettingsPart, package.RelationshipMarkup, new[] { newXDoc.Root }, package.Images);
                        newSettingsPart.PutXDocument(newXDoc);
                    }

                    WebSettingsPart oldWebSettingsPart = source.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart;
                    if (oldWebSettingsPart != null)
                    {
                        WebSettingsPart newWebSettingsPart = target.MainDocumentPart.AddNewPart<WebSettingsPart>();
                        XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                        oldWebSettingsPart.AddRelationships(newWebSettingsPart, package.RelationshipMarkup, new[] { settingsXDoc.Root });
                        XDocument newXDoc = target.MainDocumentPart.WebSettingsPart.GetXDocument();
                        newXDoc.Declaration.SetDeclaration();
                        newXDoc.Add(settingsXDoc.Root);
                        newWebSettingsPart.PutXDocument(newXDoc);
                    }

                    NumberingDefinitionsPart oldNumberingDefinitionsPart = source.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart;
                    if (oldNumberingDefinitionsPart != null)
                    {
                        CopyNumberingForGlossaryDocumentPartFromGD(oldNumberingDefinitionsPart, newContent);
                    }

                    void CopyNumberingForGlossaryDocumentPartFromGD(NumberingDefinitionsPart sourceNumberingPart, IEnumerable<XElement> targetContent)
                    {
                        Dictionary<int, int> numIdMap = new Dictionary<int, int>();
                        int number = 1;
                        int abstractNumber = 0;
                        XDocument oldNumbering = null;
                        XDocument newNumbering = null;

                        foreach (XElement numReference in targetContent.DescendantsAndSelf(W.numPr))
                        {
                            XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                            if (idElement != null)
                            {
                                if (oldNumbering == null)
                                    oldNumbering = sourceNumberingPart.GetXDocument();
                                if (newNumbering == null)
                                {
                                    if (target.MainDocumentPart.NumberingDefinitionsPart != null)
                                    {
                                        newNumbering = target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                        var numIds = newNumbering
                                            .Root
                                            .Elements(W.num)
                                            .Select(f => (int)f.Attribute(W.numId));
                                        if (numIds.Any())
                                            number = numIds.Max() + 1;
                                        numIds = newNumbering
                                            .Root
                                            .Elements(W.abstractNum)
                                            .Select(f => (int)f.Attribute(W.abstractNumId));
                                        if (numIds.Any())
                                            abstractNumber = numIds.Max() + 1;
                                    }
                                    else
                                    {
                                        target.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                                        newNumbering = target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                        newNumbering.Declaration.SetDeclaration();
                                        newNumbering.Add(new XElement(W.numbering, Constants.NamespaceAttributes));
                                    }
                                }
                                int numId = (int)idElement.Attribute(W.val);
                                if (numId != 0)
                                {
                                    XElement element = oldNumbering
                                        .Descendants(W.num)
                                        .Where(p => ((int)p.Attribute(W.numId)) == numId)
                                        .FirstOrDefault();
                                    if (element == null)
                                        continue;

                                    // Copy abstract numbering element, if necessary (use matching NSID)
                                    string abstractNumIdStr = (string)element
                                        .Elements(W.abstractNumId)
                                        .First()
                                        .Attribute(W.val);
                                    int abstractNumId;
                                    if (!int.TryParse(abstractNumIdStr, out abstractNumId))
                                        throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");
                                    XElement abstractElement = oldNumbering
                                        .Descendants()
                                        .Elements(W.abstractNum)
                                        .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                                        .First();
                                    XElement nsidElement = abstractElement
                                        .Element(W.nsid);
                                    string abstractNSID = null;
                                    if (nsidElement != null)
                                        abstractNSID = (string)nsidElement
                                            .Attribute(W.val);
                                    XElement newAbstractElement = newNumbering
                                        .Descendants()
                                        .Elements(W.abstractNum)
                                        .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                        .Where(p =>
                                        {
                                            var thisNsidElement = p.Element(W.nsid);
                                            if (thisNsidElement == null)
                                                return false;
                                            return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                                        })
                                        .FirstOrDefault();
                                    if (newAbstractElement == null)
                                    {
                                        newAbstractElement = new XElement(abstractElement);
                                        newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                                        abstractNumber++;
                                        if (newNumbering.Root.Elements(W.abstractNum).Any())
                                            newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                                        else
                                            newNumbering.Root.Add(newAbstractElement);

                                        foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                                        {
                                            string bulletId = (string)pictId.Attribute(W.val);
                                            XElement numPicBullet = oldNumbering
                                                .Descendants(W.numPicBullet)
                                                .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                            int maxNumPicBulletId = new int[] { -1 }.Concat(
                                                newNumbering.Descendants(W.numPicBullet)
                                                .Attributes(W.numPicBulletId)
                                                .Select(a => (int)a))
                                                .Max() + 1;
                                            XElement newNumPicBullet = new XElement(numPicBullet);
                                            newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                            pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                            newNumbering.Root.AddFirst(newNumPicBullet);
                                        }
                                    }
                                    string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                                    // Copy numbering element, if necessary (use matching element with no overrides)
                                    XElement newElement;
                                    if (numIdMap.ContainsKey(numId))
                                    {
                                        newElement = newNumbering
                                            .Descendants()
                                            .Elements(W.num)
                                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                            .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                            .First();
                                    }
                                    else
                                    {
                                        newElement = new XElement(element);
                                        newElement
                                            .Elements(W.abstractNumId)
                                            .First()
                                            .Attribute(W.val).Value = newAbstractId;
                                        newElement.Attribute(W.numId).Value = number.ToString();
                                        numIdMap.Add(numId, number);
                                        number++;
                                        newNumbering.Root.Add(newElement);
                                    }
                                    idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                                }
                            }
                        }
                        if (newNumbering != null)
                        {
                            foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                                abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                            foreach (var num in newNumbering.Descendants(W.num))
                                num.AddAnnotation(new FromPreviousSourceSemaphore());
                        }

                        if (target.MainDocumentPart.NumberingDefinitionsPart != null &&
                            sourceNumberingPart != null)
                        {
                            sourceNumberingPart.AddRelationships(target.MainDocumentPart.NumberingDefinitionsPart,
                                package.RelationshipMarkup, new[] { target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                            sourceNumberingPart.CopyRelatedPartsForContentParts(target.MainDocumentPart.NumberingDefinitionsPart,
                                package.RelationshipMarkup, new[] { target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, package.Images);
                        }
                        if (target.MainDocumentPart.NumberingDefinitionsPart != null)
                            target.MainDocumentPart.NumberingDefinitionsPart.PutXDocument();
                    }
                }
            }
        }
        public static void CopyAllSections(this WmlPackage package)
        {
            var outputMain = package.Target.MainDocumentPart;
            if (!outputMain.AnySections())
            {
                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(package.Sources.ElementAt(0).WmlDocument))
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    var sectPr = doc.MainDocumentPart.GetLastElement();
                    if (sectPr?.Name == W.sectPr)
                    {
                        doc.AddSectionAndDependencies(package, sectPr);
                        outputMain.GetBody().Add(sectPr);
                    }
                }
            }
        }
        public static void CopyFirstSourceCoreParts(this WmlPackage package)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(package.Sources.ElementAt(0).WmlDocument))
            {
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    doc.CopyStartingParts(package);
                    doc.CopySpecifiedCustomXmlParts(package.Target);
                }
            }
        }
        // This method is a mess
        public static IList<WmlSource> NormaliseStyleNamesAndIds(this IList<WmlSource> sources)
        {
            Dictionary<string, string> styleNameMap = new Dictionary<string, string>();
            HashSet<string> styleIds = new HashSet<string>();
            List<WmlSource> newSources = new List<WmlSource>(sources);

            foreach (var src in newSources)
            {
                AddAndRectify(src);
            }
            return newSources;


            WmlSource AddAndRectify(WmlSource src)
            {
                Dictionary<string, string> correctionList = new Dictionary<string, string>();

                bool modified = false;
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(src.WmlDocument.DocumentByteArray, 0, src.WmlDocument.DocumentByteArray.Length);
                    using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                    {

                        foreach (var pair in wDoc.MainDocumentPart.GetStyleNameMap())
                        {
                            var styleName = pair.Key;
                            var styleId = pair.Value;

                            if (styleNameMap.ContainsKey(styleName))
                            {
                                if (styleNameMap[styleName] != styleId) correctionList.Add(styleId, styleNameMap[styleName]);
                                continue;
                            }
                            else
                            {
                                // if the id is already used
                                if (styleIds.Contains(styleId))
                                {
                                    // this style uses a styleId that is used for another style.
                                    // randomly generate new styleId
                                    while (true)
                                    {
                                        var newStyleId = GenStyleIdFromStyleName(styleName);
                                        if (!styleIds.Contains(newStyleId))
                                        {
                                            correctionList.Add(styleId, newStyleId);
                                            styleNameMap.Add(styleName, newStyleId);
                                            styleIds.Add(newStyleId);
                                            break;
                                        }
                                    }
                                }
                                // otherwise we just add to the styleNameMap
                                else
                                {
                                    styleNameMap.Add(styleName, styleId);
                                    styleIds.Add(styleId);
                                }

                            }
                        }

                        if (correctionList.Any())
                        {
                            modified = true;
                            AdjustStyleIdsForDocument(wDoc);
                        }
                    }
                    if (modified)
                    {
                        var newWmlDocument = new WmlDocument(src.WmlDocument.FileName, ms.ToArray());
                        var newSrc = new WmlSource(newWmlDocument, src.Start, src.Count) { KeepSections = true };
                        newSrc.KeepHeadersAndFooters = src.KeepHeadersAndFooters;
                        newSrc.InsertId = src.InsertId;
                        return newSrc;
                    }
                }
                return src;

                void AdjustStyleIdsForDocument(WordprocessingDocument wDoc)
                {
                    var main = wDoc.MainDocumentPart;
                    // update styles part
                    UpdateStyleIdsForStylePart(main.StyleDefinitionsPart);
                    UpdateStyleIdsForStylePart(main.StylesWithEffectsPart);

                    // Update all content parts
                    UpdateStyleIdsForContentPart(main);
                    UpdateStyleIdsForContentPart(main.EndnotesPart);
                    UpdateStyleIdsForContentPart(main.FootnotesPart);
                    UpdateStyleIdsForContentPart(main.NumberingDefinitionsPart);
                    UpdateStyleIdsForContentPart(main.WordprocessingCommentsExPart);
                    UpdateStyleIdsForContentPart(main.WordprocessingCommentsPart);

                    foreach (var part in main.FooterParts)
                    {
                        UpdateStyleIdsForContentPart(part);
                    }

                    foreach (var part in main.HeaderParts)
                    {
                        UpdateStyleIdsForContentPart(part);
                    }

                    void UpdateStyleIdsForStylePart(StylesPart part)
                    {
                        if (part == null) return;

                        var styleXDoc = part.GetXDocument();
                        var styleAttributeChangeList = correctionList
                            .Select(cor => new
                            {
                                NewId = cor.Value,
                                StyleIdAttributesToChange = styleXDoc.Root.Elements(W.style).Attributes(W.styleId).Where(a => a.Value == cor.Key).ToList(),
                                BasedOnAttributesToChange = styleXDoc.Root.Elements(W.style).Elements(W.basedOn).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                                NextAttributesToChange = styleXDoc.Root.Elements(W.style).Elements(W.next).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                                LinkAttributesToChange = styleXDoc.Root.Elements(W.style).Elements(W.link).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                            }).ToList();

                        foreach (var item in styleAttributeChangeList)
                        {
                            foreach (var att in item.StyleIdAttributesToChange)
                                att.Value = item.NewId;
                            foreach (var att in item.BasedOnAttributesToChange)
                                att.Value = item.NewId;
                            foreach (var att in item.NextAttributesToChange)
                                att.Value = item.NewId;
                            foreach (var att in item.LinkAttributesToChange)
                                att.Value = item.NewId;
                        }
                        part.PutXDocument();
                    }
                    void UpdateStyleIdsForContentPart(OpenXmlPart part)
                    {
                        if (part == null) return;

                        var xDoc = part.GetXDocument();
                        var mainAttributeChangeList = correctionList
                            .Select(cor => new
                            {
                                NewId = cor.Value,
                                PStyleAttributesToChange = xDoc.Descendants(W.pStyle).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                                RStyleAttributesToChange = xDoc.Descendants(W.rStyle).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                                TblStyleAttributesToChange = xDoc.Descendants(W.tblStyle).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                            }).ToList();

                        foreach (var item in mainAttributeChangeList)
                        {
                            foreach (var att in item.PStyleAttributesToChange)
                                att.Value = item.NewId;
                            foreach (var att in item.RStyleAttributesToChange)
                                att.Value = item.NewId;
                            foreach (var att in item.TblStyleAttributesToChange)
                                att.Value = item.NewId;
                        }
                        part.PutXDocument();
                    }
                }
            }
            string GenStyleIdFromStyleName(string styleName)
            {
                var newStyleId = styleName
                    .Replace("_", "")
                    .Replace("#", "")
                    .Replace(".", "") + ((new Random()).Next(990) + 9).ToString();
                return newStyleId;
            }
        }
        public static void RemoveAllSectionsExceptLastKept(this WmlPackage package)
        {
            var outputMain = package.Target.MainDocumentPart;
            if (!outputMain.AnySections())
            {
                var source = package.Sources.Reverse().First(p => p.KeepSections);
                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    var body = doc.GetMainPart().Root.Element(W.body);
                    if (body?.Elements().Any() == true)
                    {
                        var sectPr = doc.MainDocumentPart.GetLastElement();
                        if (sectPr?.Name == W.sectPr)
                        {
                            doc.AddSectionAndDependencies(package, sectPr);
                            outputMain.GetBody().Add(sectPr);
                        }
                    }
                }
            }
        }
        #endregion

        #region WmlDocument
        public static IEnumerable<WmlDocument> SplitOnSections(this WmlDocument doc)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    IEnumerable<Atbid> divs = document.MainDocumentPart.GetDivs();
                    var groups = divs.GroupAdjacent(b => b.Div);
                    var tempSourceList = groups.Select(g => (Start: g.First().Index, Count: g.Count())).ToList();
                    foreach (var ts in tempSourceList)
                    {
                        var sources = new List<WmlSource>() { new WmlSource(doc, ts.Start, ts.Count) };
                        WmlDocument newDoc = new DocumentBuilder().SetSources(sources).Modify();
                        newDoc = AdjustSectionBreak(newDoc);
                        yield return newDoc;
                    }
                }
            }
        }
        public static WmlDocument AdjustSectionBreak(this WmlDocument doc)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    XDocument mainXDoc = document.MainDocumentPart.GetXDocument();
                    XElement lastElement = mainXDoc.Root.Element(W.body).Elements().LastOrDefault();
                    if (lastElement != null)
                    {
                        if (lastElement.Name != W.sectPr && lastElement.Descendants(W.sectPr).Any())
                        {
                            mainXDoc.Root.Element(W.body).Add(lastElement.Descendants(W.sectPr).First());
                            lastElement.Descendants(W.sectPr).Remove();
                            if (!lastElement.Elements().Where(e => e.Name != W.pPr).Any())
                            {
                                lastElement.Remove();
                            }
                            document.MainDocumentPart.PutXDocument();
                        }
                    }
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }
        public static WmlDocument SimplifyMarkup(this WmlDocument doc, SimplifyMarkupSettings settings)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    //SimplifyMarkup(document, settings);
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }
        public static string GetBackgroundColor(this WmlDocument doc)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
            {
                XDocument mainDocument = document.MainDocumentPart.GetXDocument();
                XElement backgroundElement = mainDocument.Descendants(W.background).FirstOrDefault();
                return (backgroundElement == null) ? string.Empty : backgroundElement.Attribute(W.color).Value;
            }
        }
        #endregion

        #region WordprocessingDocument
        public static void AddSectionAndDependencies(this WordprocessingDocument source, WmlPackage package, XElement sectionMarkup)
        {
            WordprocessingDocument target = package.Target;
            var headerReferences = sectionMarkup.Elements(W.headerReference);
            foreach (var headerReference in headerReferences)
            {
                string oldRid = headerReference.Attribute(R.id).Value;
                HeaderPart oldHeaderPart = null;
                try
                {
                    oldHeaderPart = (HeaderPart)source.MainDocumentPart.GetPartById(oldRid);
                }
                catch (ArgumentOutOfRangeException)
                {
                    var message = string.Format("ArgumentOutOfRangeException, attempting to get header rId={0}", oldRid);
                    throw new OpenXmlPowerToolsException(message);
                }
                XDocument oldHeaderXDoc = oldHeaderPart.GetXDocument();
                if (oldHeaderXDoc != null && oldHeaderXDoc.Root != null)
                {
                    source.CopyNumbering(package, new[] { oldHeaderXDoc.Root });
                }
                HeaderPart targetHeaderPart = target.MainDocumentPart.AddNewPart<HeaderPart>();
                XDocument targetHeader = targetHeaderPart.GetXDocument();
                targetHeader.Declaration.SetDeclaration();
                targetHeader.Add(oldHeaderXDoc.Root);
                headerReference.Attribute(R.id).Value = target.MainDocumentPart.GetIdOfPart(targetHeaderPart);
                oldHeaderPart.AddRelationships(targetHeaderPart, package.RelationshipMarkup, new[] { targetHeader.Root });
                oldHeaderPart.CopyRelatedPartsForContentParts(targetHeaderPart, package.RelationshipMarkup, new[] { targetHeader.Root }, package.Images);
            }

            var footerReferences = sectionMarkup.Elements(W.footerReference);
            foreach (var footerReference in footerReferences)
            {
                string oldRid = footerReference.Attribute(R.id).Value;

                if (source.MainDocumentPart.GetPartById(oldRid) is FooterPart sourceFooterPart)
                {
                    XDocument sourceFooter = sourceFooterPart.GetXDocument();
                    if (sourceFooter != null && sourceFooter.Root != null)
                    {
                        source.CopyNumbering(package, new[] { sourceFooter.Root });
                    }
                    FooterPart targetFooterPart = package.Target.MainDocumentPart.AddNewPart<FooterPart>();
                    XDocument targetFooter = targetFooterPart.GetXDocument();
                    targetFooter.Declaration.SetDeclaration();
                    targetFooter.Add(sourceFooter.Root);
                    footerReference.Attribute(R.id).Value = package.Target.MainDocumentPart.GetIdOfPart(targetFooterPart);
                    sourceFooterPart.AddRelationships(targetFooterPart, package.RelationshipMarkup, new[] { targetFooter.Root });
                    sourceFooterPart.CopyRelatedPartsForContentParts(targetFooterPart, package.RelationshipMarkup, new[] { targetFooter.Root }, package.Images);
                }
                else throw new DocumentBuilderException("Invalid document - invalid footer part.");

            }
        }
        public static void AppendDocument(this WordprocessingDocument source, WmlPackage package, WmlSource wmlSource, IEnumerable<XElement> targetContent)
        {
            XDocument sourceMain = source.GetMainPart();
            MainDocumentPart sourceMainPart = source.MainDocumentPart;
            WordprocessingDocument target = package.Target;
            MainDocumentPart targetMainPart = target.MainDocumentPart;

            sourceMain.FixRanges(targetContent);
            sourceMainPart.AddRelationships(targetMainPart, package.RelationshipMarkup, targetContent);
            sourceMainPart.CopyRelatedPartsForContentParts(targetMainPart, package.RelationshipMarkup, targetContent, package.Images);

            XDocument targetMain = target.GetMainPart();
            targetMain.Declaration.SetDeclaration();

            if (!wmlSource.KeepSections)
            {
                List<XElement> adjustedContents = targetContent.Where(e => e.Name != W.sectPr).ToList();
                adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
                targetContent = adjustedContents;
            }

            var listOfSectionProps = targetContent.DescendantsAndSelf(W.sectPr).ToList();

            foreach (var sectPr in listOfSectionProps)
            {
                source.AddSectionAndDependencies(package, sectPr);
            }
            
            source.CopyStylesAndFonts(package, targetContent);
            source.CopyNumbering(package, targetContent);
            source.CopyComments(package, targetContent);
            source.CopyFootnotes(package, targetContent);
            source.CopyEndnotes(package, targetContent);
            source.AdjustUniqueIds(target, targetContent);
            targetContent.RemoveGfxdata();
            source.CopyCustomXmlParts(package, targetContent);
            CopyWebExtensions(source, target);

            if (wmlSource.InsertId != null)
            {
                XElement insertElementToReplace = targetMain.Descendants(PtOpenXml.Insert).FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == wmlSource.InsertId);
                if (insertElementToReplace != null)
                {
                    insertElementToReplace.AddAnnotation(new ReplaceSemaphore());
                }

                targetMain.Element(W.document).ReplaceWith((XElement)targetMain.Root.InsertTransform(targetContent));
            }
            else targetMain.Root.Element(W.body).Add(targetContent);

            if (targetMain.Descendants().Any(d => (d.Name.Namespace == PtOpenXml.pt || d.Name.Namespace == PtOpenXml.ptOpenXml) ||
                    (d.Attributes().Any(att => att.Name.Namespace == PtOpenXml.pt || att.Name.Namespace == PtOpenXml.ptOpenXml))))
            {
                var root = targetMain.Root;
                if (!root.Attributes().Any(na => na.Value == PtOpenXml.pt.NamespaceName))
                {
                    root.Add(new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt.NamespaceName));
                    root.AddToIgnorable("pt");
                }
                if (!root.Attributes().Any(na => na.Value == PtOpenXml.ptOpenXml.NamespaceName))
                {
                    root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.ptOpenXml.NamespaceName));
                    root.AddToIgnorable("pt14");
                }
            }
        }
        public static void AppendPart(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent, WmlSource wmlSource, OpenXmlPart part)
        {
            var target = package.Target;
            XDocument partXDoc = part.GetXDocument();
            partXDoc.Declaration.SetDeclaration();

            partXDoc.FixRanges(targetContent);
            source.MainDocumentPart.AddRelationships(part, package.RelationshipMarkup, targetContent);
            source.MainDocumentPart.CopyRelatedPartsForContentParts(part, package.RelationshipMarkup, targetContent, package.Images);

            // never keep sections for content to be inserted into a header/footer
            List<XElement> adjustedContents = targetContent.Where(e => e.Name != W.sectPr).ToList();
            adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
            targetContent = adjustedContents;

            source.CopyNumbering(package, targetContent);
            source.CopyComments(package, targetContent);
            source.AdjustUniqueIds(target, targetContent);
            targetContent.RemoveGfxdata();

            if (wmlSource.InsertId == null) throw new OpenXmlPowerToolsException("Internal error");

            XElement insertElementToReplace = partXDoc.Descendants(PtOpenXml.Insert)
                .FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == wmlSource.InsertId);

            if (insertElementToReplace != null) insertElementToReplace.AddAnnotation(new ReplaceSemaphore());

            partXDoc.Elements().First().ReplaceWith((XElement)partXDoc.Root.InsertTransform(targetContent));
        }
        public static XDocument GetMainPart(this WordprocessingDocument doc) => doc.MainDocumentPart.GetXDocument();
        public static OpenXmlPowerToolsDocument SplitDocument(this WordprocessingDocument source, IEnumerable<XElement> contents, string newFileName)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = CreateWordprocessingDocument())
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    source.MainDocumentPart.GetXDocument().FixRanges(contents);
                    //Open.SetContent(document, contents);
                }
                OpenXmlPowerToolsDocument newDoc = streamDoc.GetModifiedDocument();
                newDoc.FileName = newFileName;
                return newDoc;
            }
        }

        #region Copiers
        // Messy
        public static void CopyComments(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent)
        {
            var target = package.Target;
            Dictionary<int, int> commentIdMap = new Dictionary<int, int>();
            int number = 0;
            XDocument oldComments = null;
            XDocument newComments = null;
            foreach (XElement comment in targetContent.DescendantsAndSelf(W.commentReference))
            {
                if (oldComments == null)
                    oldComments = source.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                if (newComments == null)
                {
                    if (target.MainDocumentPart.WordprocessingCommentsPart != null)
                    {
                        newComments = target.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.SetDeclaration();
                        var ids = newComments.Root.Elements(W.comment).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        target.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                        newComments = target.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.SetDeclaration();
                        newComments.Add(new XElement(W.comments, Constants.NamespaceAttributes));
                    }
                }
                int id;
                if (!int.TryParse((string)comment.Attribute(W.id), out id))
                    throw new DocumentBuilderException("Invalid document - invalid comment id");
                XElement element = oldComments
                    .Descendants()
                    .Elements(W.comment)
                    .Where(p =>
                    {
                        int thisId;
                        if (!int.TryParse((string)p.Attribute(W.id), out thisId))
                            throw new DocumentBuilderException("Invalid document - invalid comment id");
                        return thisId == id;
                    })
                    .FirstOrDefault();
                if (element == null)
                    throw new DocumentBuilderException("Invalid document - comment reference without associated comment in comments part");
                XElement newElement = new XElement(element);
                newElement.Attribute(W.id).Value = number.ToString();
                newComments.Root.Add(newElement);
                if (!commentIdMap.ContainsKey(id))
                    commentIdMap.Add(id, number);
                number++;
            }
            foreach (var item in targetContent.DescendantsAndSelf()
                .Where(d => d.Name == W.commentReference ||
                            d.Name == W.commentRangeStart ||
                            d.Name == W.commentRangeEnd)
                .ToList())
            {
                if (commentIdMap.ContainsKey((int)item.Attribute(W.id)))
                    item.Attribute(W.id).Value = commentIdMap[(int)item.Attribute(W.id)].ToString();
            }
            if (source.MainDocumentPart.WordprocessingCommentsPart != null &&
                target.MainDocumentPart.WordprocessingCommentsPart != null)
            {
                source.MainDocumentPart.WordprocessingCommentsPart.AddRelationships(target.MainDocumentPart.WordprocessingCommentsPart,
                    package.RelationshipMarkup, new[] { target.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root });
                source.MainDocumentPart.WordprocessingCommentsPart.CopyRelatedPartsForContentParts(target.MainDocumentPart.WordprocessingCommentsPart,
                    package.RelationshipMarkup, new[] { target.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root }, package.Images);
            }
        }
        public static void CopyCustomXmlParts(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent)
        {
            List<string> itemList = new List<string>();
            var target = package.Target;

            var itemIds = targetContent.Descendants(W.dataBinding).Select(e => (string)e.Attribute(W.storeItemID));
            foreach (string itemId in itemIds)
            {
                if (!itemList.Contains(itemId)) itemList.Add(itemId);
            }

            foreach (CustomXmlPart customXmlPart in source.MainDocumentPart.CustomXmlParts)
            {
                OpenXmlPart propertyPart = customXmlPart.Parts.Select(p => p.OpenXmlPart)
                    .Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                    .FirstOrDefault();

                if (propertyPart != null)
                {
                    XDocument propertyPartDoc = propertyPart.GetXDocument();
                    if (itemList.Contains(propertyPartDoc.Root.Attribute(DS.itemID).Value))
                    {
                        CustomXmlPart newPart = target.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
                        newPart.GetXDocument().Add(customXmlPart.GetXDocument().Root);

                        foreach (OpenXmlPart propPart in customXmlPart.Parts.Select(p => p.OpenXmlPart))
                        {
                            CustomXmlPropertiesPart newPropPart = newPart.AddNewPart<CustomXmlPropertiesPart>();
                            newPropPart.GetXDocument().Add(propPart.GetXDocument().Root);
                        }
                    }
                }
            }
        }
        public static void CopyEndnotes(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent)
        {
            source.CopyContainerPart<EndnotesPart>(package, targetContent, W.endnote, W.endnotes, W.endnoteReference);
        }
        public static void CopyFootnotes(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent)
        {
            source.CopyContainerPart<FootnotesPart>(package, targetContent, W.footnote, W.footnotes, W.endnotePr);
        }
        private static void CopyContainerPart<T>(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent, XName partName, XName partContainer, XName partRef)
            where T : OpenXmlPart, IFixedContentTypePart
        {
            var target = package.Target;
            T sourcePart = source.MainDocumentPart.GetPart<T>();
            T targetPart = target.MainDocumentPart.GetPart<T>();
            XDocument sourcePartXDoc = null;
            XDocument targetPartXDoc = null;
            var parts = targetContent.DescendantsAndSelf(partRef).ToList();
            for (int index = 0; index < parts.Count(); index++)
            {
                var part = parts[index];
                if (sourcePartXDoc == null) sourcePartXDoc = sourcePart.GetXDocument();
                if (targetPartXDoc == null)
                {
                    if (targetPart != null)
                    {
                        targetPartXDoc = targetPart.GetXDocument();
                        var ids = targetPartXDoc.Root.Elements(partName).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any()) index = ids.Max() + 1;
                    }
                    else
                    {
                        target.MainDocumentPart.AddNewPart<T>();
                        targetPartXDoc = targetPart.GetXDocument();
                        targetPartXDoc.Declaration.SetDeclaration();
                        targetPartXDoc.Add(new XElement(partContainer, Constants.NamespaceAttributes));
                    }
                }
                string id = (string)part.Attribute(W.id);
                XElement element = sourcePartXDoc.Descendants().Elements(partName).Where(p => ((string)p.Attribute(W.id)) == id).FirstOrDefault();
                if (element != null)
                {
                    XElement newElement = new XElement(element);
                    newElement.Attribute(W.id).Value = index.ToString();
                    targetPartXDoc.Root.Add(newElement);
                    part.Attribute(W.id).Value = index.ToString();
                }
            }
            if (sourcePart != null && targetPart != null)
            {
                sourcePart.AddRelationships(targetPart, RelationshipMarkup, new[] { targetPart.GetXDocument().Root });
                sourcePart.CopyRelatedPartsForContentParts(targetPart, RelationshipMarkup, new[] { targetPart.GetXDocument().Root }, package.Images);
            }
        }
        // messy
        public static void CopyEndnotesPart(this WordprocessingDocument source, WmlPackage package, XDocument settingsXDoc)
        {
            var target = package.Target;
            int number = 0;
            XDocument oldEndnotes = null;
            XDocument newEndnotes = null;
            XElement endnotePr = settingsXDoc.Root.Element(W.endnotePr);
            if (endnotePr == null)
                return;
            if (source.MainDocumentPart.EndnotesPart == null)
                return;
            foreach (XElement endnote in endnotePr.Elements(W.endnote))
            {
                if (oldEndnotes == null)
                    oldEndnotes = source.MainDocumentPart.EndnotesPart.GetXDocument();
                if (newEndnotes == null)
                {
                    if (target.MainDocumentPart.EndnotesPart != null)
                    {
                        newEndnotes = target.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.SetDeclaration();
                        var ids = newEndnotes.Root
                            .Elements(W.endnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        target.MainDocumentPart.AddNewPart<EndnotesPart>();
                        newEndnotes = target.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.SetDeclaration();
                        newEndnotes.Add(new XElement(W.endnotes, Constants.NamespaceAttributes));
                    }
                }
                string id = (string)endnote.Attribute(W.id);
                XElement element = oldEndnotes.Descendants()
                    .Elements(W.endnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    XElement newElement = new XElement(element);
                    newElement.Attribute(W.id).Value = number.ToString();
                    newEndnotes.Root.Add(newElement);
                    endnote.Attribute(W.id).Value = number.ToString();
                    number++;
                }
            }
        }
        public static void CopyFootnotesPart(this WordprocessingDocument source, WmlPackage package, XDocument settingsXDoc)
        {
            var target = package.Target;
            int number = 0;
            XDocument oldFootnotes = null;
            XDocument newFootnotes = null;
            XElement footnotePr = settingsXDoc.Root.Element(W.footnotePr);
            if (footnotePr == null)
                return;
            if (source.MainDocumentPart.FootnotesPart == null)
                return;
            foreach (XElement footnote in footnotePr.Elements(W.footnote))
            {
                if (oldFootnotes == null)
                    oldFootnotes = source.MainDocumentPart.FootnotesPart.GetXDocument();
                if (newFootnotes == null)
                {
                    if (target.MainDocumentPart.FootnotesPart != null)
                    {
                        newFootnotes = target.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.SetDeclaration();
                        var ids = newFootnotes.Root.Elements(W.footnote).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        target.MainDocumentPart.AddNewPart<FootnotesPart>();
                        newFootnotes = target.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.SetDeclaration();
                        newFootnotes.Add(new XElement(W.footnotes, Constants.NamespaceAttributes));
                    }
                }
                string id = (string)footnote.Attribute(W.id);
                XElement element = oldFootnotes.Descendants()
                    .Elements(W.footnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    XElement newElement = new XElement(element);
                    // the following adds the footnote into the new settings part
                    newElement.Attribute(W.id).Value = number.ToString();
                    newFootnotes.Root.Add(newElement);
                    footnote.Attribute(W.id).Value = number.ToString();
                    number++;
                }
            }
        }

        // messy
        public static void CopyGlossaryDocumentPart(this WmlDocument sourceWmlDoc, WmlPackage package)
        {
            var target = package.Target;
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(sourceWmlDoc.DocumentByteArray, 0, sourceWmlDoc.DocumentByteArray.Length);
                using (WordprocessingDocument source = WordprocessingDocument.Open(ms, true))
                {
                    var sourceMain = source.GetMainPart();
                    var sourceGlossary = sourceMain.CreateGlossary();
                    var outputGlossaryDocumentPart = target.MainDocumentPart.AddNewPart<GlossaryDocumentPart>();
                    outputGlossaryDocumentPart.PutXDocument(sourceGlossary);

                    CopyGlossaryDocumentPartsToGD(source, sourceMain.Root.Descendants(W.docPart));
                    source.MainDocumentPart.CopyRelatedPartsForContentParts(outputGlossaryDocumentPart, package.RelationshipMarkup, new[] { sourceMain.Root }, package.Images);
                }
            }

            void CopyGlossaryDocumentPartsToGD(WordprocessingDocument source, IEnumerable<XElement> newContent)
            {
                // Copy all styles to the new document
                if (source.MainDocumentPart.StyleDefinitionsPart != null)
                {
                    XDocument oldStyles = source.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    target.MainDocumentPart.GlossaryDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    XDocument newStyles = target.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newStyles.Declaration.SetDeclaration();
                    newStyles.Add(oldStyles.Root);
                    target.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.PutXDocument();
                }

                // Copy fontTable to the new document
                if (source.MainDocumentPart.FontTablePart != null)
                {
                    XDocument oldFontTable = source.MainDocumentPart.FontTablePart.GetXDocument();
                    target.MainDocumentPart.GlossaryDocumentPart.AddNewPart<FontTablePart>();
                    XDocument newFontTable = target.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                    newFontTable.Declaration.SetDeclaration();
                    newFontTable.Add(oldFontTable.Root);
                    target.MainDocumentPart.FontTablePart.PutXDocument();
                }

                DocumentSettingsPart oldSettingsPart = source.MainDocumentPart.DocumentSettingsPart;
                if (oldSettingsPart != null)
                {
                    DocumentSettingsPart newSettingsPart = target.MainDocumentPart.GlossaryDocumentPart.AddNewPart<DocumentSettingsPart>();
                    XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                    oldSettingsPart.AddRelationships(newSettingsPart, package.RelationshipMarkup, new[] { settingsXDoc.Root });
                    //CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                    //CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                    XDocument newXDoc = target.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart.GetXDocument();
                    newXDoc.Declaration.SetDeclaration();
                    newXDoc.Add(settingsXDoc.Root);
                    oldSettingsPart.CopyRelatedPartsForContentParts(newSettingsPart, package.RelationshipMarkup, new[] { newXDoc.Root }, package.Images);
                    newSettingsPart.PutXDocument(newXDoc);
                }

                WebSettingsPart oldWebSettingsPart = source.MainDocumentPart.WebSettingsPart;
                if (oldWebSettingsPart != null)
                {
                    WebSettingsPart newWebSettingsPart = target.MainDocumentPart.GlossaryDocumentPart.AddNewPart<WebSettingsPart>();
                    XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                    oldWebSettingsPart.AddRelationships(newWebSettingsPart, package.RelationshipMarkup, new[] { settingsXDoc.Root });
                    XDocument newXDoc = target.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart.GetXDocument();
                    newXDoc.Declaration.SetDeclaration();
                    newXDoc.Add(settingsXDoc.Root);
                    newWebSettingsPart.PutXDocument(newXDoc);
                }

                NumberingDefinitionsPart oldNumberingDefinitionsPart = source.MainDocumentPart.NumberingDefinitionsPart;
                if (oldNumberingDefinitionsPart != null)
                {
                    CopyNumberingForGlossaryDocumentPartToGD(oldNumberingDefinitionsPart, newContent);
                }

                void CopyNumberingForGlossaryDocumentPartToGD(NumberingDefinitionsPart sourceNumberingPart, IEnumerable<XElement> content)
                {
                    Dictionary<int, int> numIdMap = new Dictionary<int, int>();
                    int number = 1;
                    int abstractNumber = 0;
                    XDocument oldNumbering = null;
                    XDocument newNumbering = null;

                    foreach (XElement numReference in content.DescendantsAndSelf(W.numPr))
                    {
                        XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                        if (idElement != null)
                        {
                            if (oldNumbering == null)
                                oldNumbering = sourceNumberingPart.GetXDocument();
                            if (newNumbering == null)
                            {
                                if (target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                                {
                                    newNumbering = target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    var numIds = newNumbering
                                        .Root
                                        .Elements(W.num)
                                        .Select(f => (int)f.Attribute(W.numId));
                                    if (numIds.Any())
                                        number = numIds.Max() + 1;
                                    numIds = newNumbering
                                        .Root
                                        .Elements(W.abstractNum)
                                        .Select(f => (int)f.Attribute(W.abstractNumId));
                                    if (numIds.Any())
                                        abstractNumber = numIds.Max() + 1;
                                }
                                else
                                {
                                    target.MainDocumentPart.GlossaryDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                                    newNumbering = target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    newNumbering.Declaration.SetDeclaration();
                                    newNumbering.Add(new XElement(W.numbering, Constants.NamespaceAttributes));
                                }
                            }
                            int numId = (int)idElement.Attribute(W.val);
                            if (numId != 0)
                            {
                                XElement element = oldNumbering
                                    .Descendants(W.num)
                                    .Where(p => ((int)p.Attribute(W.numId)) == numId)
                                    .FirstOrDefault();
                                if (element == null)
                                    continue;

                                // Copy abstract numbering element, if necessary (use matching NSID)
                                string abstractNumIdStr = (string)element
                                    .Elements(W.abstractNumId)
                                    .First()
                                    .Attribute(W.val);
                                int abstractNumId;
                                if (!int.TryParse(abstractNumIdStr, out abstractNumId))
                                    throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");
                                XElement abstractElement = oldNumbering
                                    .Descendants()
                                    .Elements(W.abstractNum)
                                    .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                                    .First();
                                XElement nsidElement = abstractElement
                                    .Element(W.nsid);
                                string abstractNSID = null;
                                if (nsidElement != null)
                                    abstractNSID = (string)nsidElement
                                        .Attribute(W.val);
                                XElement newAbstractElement = newNumbering
                                    .Descendants()
                                    .Elements(W.abstractNum)
                                    .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                    .Where(p =>
                                    {
                                        var thisNsidElement = p.Element(W.nsid);
                                        if (thisNsidElement == null)
                                            return false;
                                        return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                                    })
                                    .FirstOrDefault();
                                if (newAbstractElement == null)
                                {
                                    newAbstractElement = new XElement(abstractElement);
                                    newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                                    abstractNumber++;
                                    if (newNumbering.Root.Elements(W.abstractNum).Any())
                                        newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                                    else
                                        newNumbering.Root.Add(newAbstractElement);

                                    foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                                    {
                                        string bulletId = (string)pictId.Attribute(W.val);
                                        XElement numPicBullet = oldNumbering
                                            .Descendants(W.numPicBullet)
                                            .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                        int maxNumPicBulletId = new int[] { -1 }.Concat(
                                            newNumbering.Descendants(W.numPicBullet)
                                            .Attributes(W.numPicBulletId)
                                            .Select(a => (int)a))
                                            .Max() + 1;
                                        XElement newNumPicBullet = new XElement(numPicBullet);
                                        newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                        pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                        newNumbering.Root.AddFirst(newNumPicBullet);
                                    }
                                }
                                string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                                // Copy numbering element, if necessary (use matching element with no overrides)
                                XElement newElement;
                                if (numIdMap.ContainsKey(numId))
                                {
                                    newElement = newNumbering
                                        .Descendants()
                                        .Elements(W.num)
                                        .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                        .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                        .First();
                                }
                                else
                                {
                                    newElement = new XElement(element);
                                    newElement
                                        .Elements(W.abstractNumId)
                                        .First()
                                        .Attribute(W.val).Value = newAbstractId;
                                    newElement.Attribute(W.numId).Value = number.ToString();
                                    numIdMap.Add(numId, number);
                                    number++;
                                    newNumbering.Root.Add(newElement);
                                }
                                idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                            }
                        }
                    }
                    if (newNumbering != null)
                    {
                        foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                            abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                        foreach (var num in newNumbering.Descendants(W.num))
                            num.AddAnnotation(new FromPreviousSourceSemaphore());
                    }

                    if (target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null &&
                        sourceNumberingPart != null)
                    {
                        sourceNumberingPart.AddRelationships(target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                            package.RelationshipMarkup, new[] { target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                        sourceNumberingPart.CopyRelatedPartsForContentParts(target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                            package.RelationshipMarkup, new[] { target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, package.Images);
                    }
                    if (target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                        target.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.PutXDocument();
                }
            }
        }
        // messy
        public static void CopyNumbering(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent)
        {
            var target = package.Target;
            var sourceMainPart = source.MainDocumentPart;
            var targetMainPart = target.MainDocumentPart;
            var sourceMain = source.GetMainPart();
            var targetMain = target.GetMainPart();
            int number = 1;
            int abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (XElement numReference in targetContent.DescendantsAndSelf(W.numPr))
            {
                XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    if (oldNumbering == null)
                        oldNumbering = source.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (target.MainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            var numIds = newNumbering
                                .Root
                                .Elements(W.num)
                                .Select(f => (int)f.Attribute(W.numId));
                            if (numIds.Any())
                                number = numIds.Max() + 1;
                            numIds = newNumbering
                                .Root
                                .Elements(W.abstractNum)
                                .Select(f => (int)f.Attribute(W.abstractNumId));
                            if (numIds.Any())
                                abstractNumber = numIds.Max() + 1;
                        }
                        else
                        {
                            target.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            newNumbering.Declaration.SetDeclaration();
                            newNumbering.Add(new XElement(W.numbering, Constants.NamespaceAttributes));
                        }
                    }
                    int numId = (int)idElement.Attribute(W.val);
                    if (numId != 0)
                    {
                        XElement element = oldNumbering
                            .Descendants(W.num)
                            .Where(p => ((int)p.Attribute(W.numId)) == numId)
                            .FirstOrDefault();
                        if (element == null)
                            continue;

                        // Copy abstract numbering element, if necessary (use matching NSID)
                        string abstractNumIdStr = (string)element
                            .Elements(W.abstractNumId)
                            .First()
                            .Attribute(W.val);
                        int abstractNumId;
                        if (!int.TryParse(abstractNumIdStr, out abstractNumId))
                            throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");

                        XElement abstractElement = oldNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                            .First();
                        XElement nsidElement = abstractElement
                            .Element(W.nsid);
                        string abstractNSID = null;
                        if (nsidElement != null)
                            abstractNSID = (string)nsidElement
                                .Attribute(W.val);
                        XElement newAbstractElement = newNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                            .Where(p =>
                            {
                                var thisNsidElement = p.Element(W.nsid);
                                if (thisNsidElement == null)
                                    return false;
                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                            })
                            .FirstOrDefault();
                        if (newAbstractElement == null)
                        {
                            newAbstractElement = new XElement(abstractElement);
                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                            abstractNumber++;
                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                            else
                                newNumbering.Root.Add(newAbstractElement);

                            foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                            {
                                string bulletId = (string)pictId.Attribute(W.val);
                                XElement numPicBullet = oldNumbering
                                    .Descendants(W.numPicBullet)
                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                int maxNumPicBulletId = new int[] { -1 }.Concat(
                                    newNumbering.Descendants(W.numPicBullet)
                                    .Attributes(W.numPicBulletId)
                                    .Select(a => (int)a))
                                    .Max() + 1;
                                XElement newNumPicBullet = new XElement(numPicBullet);
                                newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                newNumbering.Root.AddFirst(newNumPicBullet);
                            }
                        }
                        string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                        // Copy numbering element, if necessary (use matching element with no overrides)
                        XElement newElement;
                        Dictionary<int, int> numIdMap = new Dictionary<int, int>();
                        if (numIdMap.ContainsKey(numId))
                        {
                            newElement = newNumbering
                                .Descendants()
                                .Elements(W.num)
                                .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                .First();
                        }
                        else
                        {
                            newElement = new XElement(element);
                            newElement
                                .Elements(W.abstractNumId)
                                .First()
                                .Attribute(W.val).Value = newAbstractId;
                            newElement.Attribute(W.numId).Value = number.ToString();
                            numIdMap.Add(numId, number);
                            number++;
                            newNumbering.Root.Add(newElement);
                        }
                        idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                    }
                }
            }
            if (newNumbering != null)
            {
                foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                    abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                foreach (var num in newNumbering.Descendants(W.num))
                    num.AddAnnotation(new FromPreviousSourceSemaphore());
            }

            if (target.MainDocumentPart.NumberingDefinitionsPart != null &&
                source.MainDocumentPart.NumberingDefinitionsPart != null)
            {
                source.MainDocumentPart.NumberingDefinitionsPart.AddRelationships(target.MainDocumentPart.NumberingDefinitionsPart,
                    package.RelationshipMarkup, new[] { target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                source.MainDocumentPart.NumberingDefinitionsPart.CopyRelatedPartsForContentParts(target.MainDocumentPart.NumberingDefinitionsPart,
                    package.RelationshipMarkup, new[] { target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, package.Images);
            }
        }
        public static void CopyOrCacheHeaderOrFooter(this WordprocessingDocument doc, CachedHeaderFooter[] cachedHeaderFooter, XElement sect, XName referenceXName, string type)
        {
            var referenceElement = sect.FindReference(referenceXName, type);
            if (referenceElement == null)
            {
                var cachedPartRid = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type).CachedPartRid;
                sect.AddReferenceToExistingHeaderOrFooter(cachedPartRid, referenceXName, type);
            }
            else
            {
                var cachedPart = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type);
                cachedPart.CachedPartRid = (string)referenceElement.Attribute(R.id);
            }
        }
        public static void CopySpecifiedCustomXmlParts(this WordprocessingDocument source, WordprocessingDocument target)
        {
            //if (settings.CustomXmlGuidList == null || !settings.CustomXmlGuidList.Any())
            //    return;

            foreach (CustomXmlPart customXmlPart in source.MainDocumentPart.CustomXmlParts)
            {
                OpenXmlPart propertyPart = customXmlPart
                    .Parts
                    .Select(p => p.OpenXmlPart)
                    .Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                    .FirstOrDefault();
                if (propertyPart != null)
                {
                    XDocument propertyPartDoc = propertyPart.GetXDocument();

                    var itemID = (string)propertyPartDoc.Root.Attribute(DS.itemID);
                    if (itemID != null)
                    {
                        itemID = itemID.Trim('{', '}');
                        if (true)//settings.CustomXmlGuidList.Contains(itemID))
                        {
                            CustomXmlPart newPart = target.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
                            newPart.GetXDocument().Add(customXmlPart.GetXDocument().Root);
                            foreach (OpenXmlPart propPart in customXmlPart.Parts.Select(p => p.OpenXmlPart))
                            {
                                CustomXmlPropertiesPart newPropPart = newPart.AddNewPart<CustomXmlPropertiesPart>();
                                newPropPart.GetXDocument().Add(propPart.GetXDocument().Root);
                            }
                        }
                    }
                }
            }
        }
        public static void CopyStartingParts(this WordprocessingDocument source, WmlPackage package)
        {
            var target = package.Target;

            ProcessCoreFilePropertiesPart();
            ProcessExtendedFilePropertiesParty();
            ProcessCustomFilePropertiesParty();
            AddDocumentSettingsPart();
            AddWebSettingsPart();
            AddThemePart();
            // If needed to handle GlossaryDocumentPart in the future, then
            // this code should handle the following parts:
            //   MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart
            //   MainDocumentPart.GlossaryDocumentPart.StylesWithEffectsPart
            //GlossaryDocumentPart
            AddStyleDefinitionsPart();
            ProcessFontTablePart();

            void ProcessCoreFilePropertiesPart()
            {
                CoreFilePropertiesPart corePart = source.CoreFilePropertiesPart;
                if (corePart?.GetXDocument()?.Root != null)
                {
                    target.AddCoreFilePropertiesPart();
                    XDocument targetPart = target.CoreFilePropertiesPart.GetXDocument();
                    targetPart.Declaration.SetDeclaration();
                    targetPart.Add(corePart.GetXDocument().Root);
                }
            }
            void ProcessExtendedFilePropertiesParty()
            {
                ExtendedFilePropertiesPart sourcePart = source.ExtendedFilePropertiesPart;
                if (sourcePart != null)
                {
                    target.AddExtendedFilePropertiesPart();
                    XDocument targetPart = target.ExtendedFilePropertiesPart.GetXDocument();
                    targetPart.Declaration.SetDeclaration();
                    targetPart.Add(sourcePart.GetXDocument().Root);
                }
            }
            void ProcessCustomFilePropertiesParty()
            {
                CustomFilePropertiesPart sourcePart = source.CustomFilePropertiesPart;
                if (source.CustomFilePropertiesPart != null)
                {
                    target.AddCustomFilePropertiesPart();
                    XDocument targetPart = target.CustomFilePropertiesPart.GetXDocument();
                    targetPart.Declaration.SetDeclaration();
                    targetPart.Add(sourcePart.GetXDocument().Root);
                }
            }
            void AddDocumentSettingsPart()
            {
                DocumentSettingsPart sourcePart = source.MainDocumentPart.DocumentSettingsPart;
                if (sourcePart != null)
                {
                    DocumentSettingsPart targetPart = target.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    XDocument sourceXDoc = sourcePart.GetXDocument();
                    sourcePart.AddRelationships(targetPart, package.RelationshipMarkup, new[] { sourceXDoc.Root });
                    source.CopyFootnotesPart(package, sourceXDoc);
                    source.CopyEndnotesPart(package, sourceXDoc);
                    XDocument targetXDoc = target.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                    targetXDoc.Declaration.SetDeclaration();
                    targetXDoc.Add(sourceXDoc.Root);
                    sourcePart.CopyRelatedPartsForContentParts(targetPart, package.RelationshipMarkup, new[] { targetXDoc.Root }, package.Images);
                }
            }
            void AddWebSettingsPart()
            {
                WebSettingsPart sourcePart = source.MainDocumentPart.WebSettingsPart;
                if (sourcePart != null)
                {
                    WebSettingsPart targetPart = target.MainDocumentPart.AddNewPart<WebSettingsPart>();
                    XDocument sourceXDoc = sourcePart.GetXDocument();
                    sourcePart.AddRelationships(targetPart, package.RelationshipMarkup, new[] { sourceXDoc.Root });
                    XDocument targetXDoc = target.MainDocumentPart.WebSettingsPart.GetXDocument();
                    targetXDoc.Declaration.SetDeclaration();
                    targetXDoc.Add(sourceXDoc.Root);
                }
            }
            void AddThemePart()
            {
                ThemePart sourcePart = source.MainDocumentPart.ThemePart;
                if (sourcePart != null)
                {
                    ThemePart targetPart = target.MainDocumentPart.AddNewPart<ThemePart>();
                    XDocument targetXDoc = targetPart.GetXDocument();
                    targetXDoc.Declaration.SetDeclaration();
                    targetXDoc.Add(sourcePart.GetXDocument().Root);
                    sourcePart.CopyRelatedPartsForContentParts(targetPart, package.RelationshipMarkup, new[] { targetPart.GetXDocument().Root }, package.Images);
                }
            }

            void AddStyleDefinitionsPart()
            {
                StyleDefinitionsPart stylesPart = source.MainDocumentPart.StyleDefinitionsPart;
                if (stylesPart != null)
                {
                    target.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    XDocument newXDoc = target.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newXDoc.Declaration.SetDeclaration();
                    newXDoc.Add(new XElement(W.styles,new XAttribute(XNamespace.Xmlns + "w", W.w)

                        //,
                        //stylesPart.GetXDocument().Descendants(W.docDefaults)

                        //,
                        //new XElement(W.latentStyles, stylesPart.GetXDocument().Descendants(W.latentStyles).Attributes())

                        ));
                    stylesPart.GetXDocument().MergeDocDefaultStyles(newXDoc);
                    MergeStyles(source, target, stylesPart.GetXDocument(), newXDoc, Enumerable.Empty<XElement>());
                    stylesPart.GetXDocument().MergeLatentStyles(newXDoc);
                }
            }
            void ProcessFontTablePart()
            {
                FontTablePart sourceFontTablePart = source.MainDocumentPart.FontTablePart;
                if (sourceFontTablePart != null)
                {
                    target.MainDocumentPart.AddNewPart<FontTablePart>();
                    FontTablePart targetFontTablePart = target.MainDocumentPart.FontTablePart;
                    XDocument newXDoc = targetFontTablePart.GetXDocument();
                    newXDoc.Declaration.SetDeclaration();
                    CopyFontTable(targetFontTablePart);
                    newXDoc.Add(sourceFontTablePart.GetXDocument().Root);
                }
                void CopyFontTable(FontTablePart targetPart)
                {
                    var relevantElements = sourceFontTablePart.GetXDocument().Descendants().Where(d => d.Name == W.embedRegular ||
                        d.Name == W.embedBold || d.Name == W.embedItalic || d.Name == W.embedBoldItalic).ToList();
                    foreach (XElement fontReference in relevantElements)
                    {
                        string relId = (string)fontReference.Attribute(R.id);
                        if (string.IsNullOrEmpty(relId))
                            continue;

                        var ipp1 = targetPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                        if (ipp1 != null)
                        {
                            OpenXmlPart tempPart = ipp1.OpenXmlPart;
                            continue;
                        }

                        ExternalRelationship tempEr1 = targetPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                        if (tempEr1 != null)
                            continue;

                        var oldPart2 = sourceFontTablePart.GetPartById(relId);
                        if (oldPart2 == null || (!(oldPart2 is FontPart)))
                            throw new DocumentBuilderException("Invalid document - FontTablePart contains invalid relationship");

                        FontPart oldPart = (FontPart)oldPart2;
                        FontPart newPart = targetPart.AddFontPart(oldPart.ContentType);
                        var ResourceID = targetPart.GetIdOfPart(newPart);
                        using (Stream oldFont = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                        using (Stream newFont = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                        {
                            int byteCount;
                            byte[] buffer = new byte[65536];
                            while ((byteCount = oldFont.Read(buffer, 0, 65536)) != 0)
                                newFont.Write(buffer, 0, byteCount);
                        }
                        fontReference.Attribute(R.id).Value = ResourceID;
                    }
                }
            }
        }

        // Generic candidate method (not complete)
        private static void CopyPart<T>(this WordprocessingDocument source, WmlPackage package, Action addStrategy = null, bool addRelationship = false, Action copyStrategy = null, bool copyRelated = false)
            where T : OpenXmlPart, IFixedContentTypePart
        {
            var sourceMain = source.MainDocumentPart;
            var targetMain = package.Target.MainDocumentPart;
            T sourcePart = sourceMain.GetPart<T>();

            if (sourcePart != null)
            {
                XDocument sourceXDoc = sourcePart.GetXDocument();
                
                if (addStrategy != null) addStrategy();
                else targetMain.AddNewPart<T>();

                T targetPart = targetMain.GetPart<T>();
                XDocument targetXDoc = targetPart.GetXDocument();
                targetXDoc.Declaration.SetDeclaration();

                if (addRelationship) sourcePart.AddRelationships(targetPart, package.RelationshipMarkup, new[] { sourceXDoc.Root });
                if (copyStrategy != null) copyStrategy();
                targetXDoc.Add(sourceXDoc.Root);
                if (copyRelated) sourcePart.CopyRelatedPartsForContentParts(targetPart, package.RelationshipMarkup, new[] { targetXDoc.Root }, package.Images);
            }
        }

        public static void CopyStylesAndFonts(this WordprocessingDocument source, WmlPackage package, IEnumerable<XElement> targetContent)
        {
            var target = package.Target;
            var sourceMainPart = source.MainDocumentPart;
            var targetMainPart = target.MainDocumentPart;
            // Copy all styles to the new document
            if (source.MainDocumentPart.StyleDefinitionsPart != null)
            {
                XDocument sourceStyle = sourceMainPart.GetStylePart();
                if (targetMainPart.StyleDefinitionsPart == null)
                {
                    targetMainPart.AddNewPart<StyleDefinitionsPart>();
                    XDocument targetStyle = targetMainPart.GetStylePart();
                    targetStyle.Declaration.SetDeclaration();
                    targetStyle.Add(sourceStyle.Root);
                }
                else
                {
                    XDocument targetStylePart = targetMainPart.GetStylePart();
                    MergeStyles(source, target, sourceStyle, targetStylePart, targetContent);
                    sourceStyle.MergeLatentStyles(targetStylePart);
                }
            }

            // Copy all styles with effects to the new document
            if (source.MainDocumentPart.StylesWithEffectsPart != null)
            {
                XDocument sourceStyle = source.MainDocumentPart.GetStylesWithEffectsPart();
                if (targetMainPart.StylesWithEffectsPart == null)
                {
                    targetMainPart.AddNewPart<StylesWithEffectsPart>();
                    XDocument targetStyle = targetMainPart.GetStylesWithEffectsPart();
                    targetStyle.Declaration.SetDeclaration();
                    targetStyle.Add(sourceStyle.Root);
                }
                else
                {
                    XDocument targetStyle = targetMainPart.GetStylesWithEffectsPart();
                    MergeStyles(source, target, sourceStyle, targetStyle, targetContent);
                    sourceStyle.MergeLatentStyles(targetStyle);
                }
            }

            // Copy fontTable to the new document
            if (source.MainDocumentPart.FontTablePart != null)
            {
                XDocument oldFontTable = sourceMainPart.GetFontTablePart();
                if (targetMainPart.FontTablePart == null)
                {
                    targetMainPart.AddNewPart<FontTablePart>();
                    XDocument newFontTable = targetMainPart.GetFontTablePart();
                    newFontTable.Declaration.SetDeclaration();
                    newFontTable.Add(oldFontTable.Root);
                }
                else
                {
                    oldFontTable.MergeFontTables(targetMainPart.GetFontTablePart());
                }
            }
        }
        public static void CopyWebExtensions(this WordprocessingDocument source, WordprocessingDocument target)
        {
            if (source.WebExTaskpanesPart != null && target.WebExTaskpanesPart == null)
            {
                target.AddWebExTaskpanesPart();
                target.WebExTaskpanesPart.GetXDocument().Add(source.WebExTaskpanesPart.GetXDocument().Root);

                foreach (var sourceWebExtensionPart in source.WebExTaskpanesPart.WebExtensionParts)
                {
                    var newWebExtensionpart = target.WebExTaskpanesPart.AddNewPart<WebExtensionPart>(
                        source.WebExTaskpanesPart.GetIdOfPart(sourceWebExtensionPart));
                    newWebExtensionpart.GetXDocument().Add(sourceWebExtensionPart.GetXDocument().Root);
                }
            }
        }
        #endregion

        #region Modifiers
        public static void AdjustUniqueIds(this WordprocessingDocument source, WordprocessingDocument target, IEnumerable<XElement> newContent)
        {
            // adjust bookmark unique ids
            int maxId = 0;
            if (target.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart).Any())
                maxId = target.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart)
                    .Select(d => (int)d.Attribute(W.id)).Max();
            Dictionary<int, int> bookmarkIdMap = new Dictionary<int, int>();
            foreach (var item in newContent.DescendantsAndSelf().Where(bm => bm.Name == W.bookmarkStart ||
                bm.Name == W.bookmarkEnd))
            {
                int id;
                if (!int.TryParse((string)item.Attribute(W.id), out id))
                    throw new DocumentBuilderException("Invalid document - invalid value for bookmark ID");
                if (!bookmarkIdMap.ContainsKey(id))
                    bookmarkIdMap.Add(id, ++maxId);
            }
            foreach (var bookmarkElement in newContent.DescendantsAndSelf().Where(e => e.Name == W.bookmarkStart ||
                e.Name == W.bookmarkEnd))
                bookmarkElement.Attribute(W.id).Value = bookmarkIdMap[(int)bookmarkElement.Attribute(W.id)].ToString();

            // adjust shape unique ids
            // This doesn't work because OLEObjects refer to shapes by ID.
            // Punting on this, because sooner or later, this will be a non-issue.
            //foreach (var item in newContent.DescendantsAndSelf(VML.shape))
            //{
            //    Guid g = Guid.NewGuid();
            //    string s = "R" + g.ToString().Replace("-", "");
            //    item.Attribute(NoNamespace.id).Value = s;
            //}
        }
        // messy
        public static void AdjustDocPrIds(this WordprocessingDocument doc)
        {
            int docPrId = 0;
            foreach (var item in doc.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
                item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            foreach (var header in doc.MainDocumentPart.HeaderParts)
                foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            foreach (var footer in doc.MainDocumentPart.FooterParts)
                foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            if (doc.MainDocumentPart.FootnotesPart != null)
                foreach (var item in doc.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            if (doc.MainDocumentPart.EndnotesPart != null)
                foreach (var item in doc.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
        }
        public static void FixSectionProperties(this WordprocessingDocument doc, WmlPackage package)
        {
            XDocument mainDocumentXDoc = doc.MainDocumentPart.GetXDocument();
            mainDocumentXDoc.Declaration.SetDeclaration();
            XElement body = mainDocumentXDoc.Root.Element(W.body);
            var sectionPropQueue = new Queue<XElement>(body.Elements().Take(body.Elements().Count() - 1).Where(e => e.Name == W.sectPr));

            while (sectionPropQueue.Count > 0)
            {
                var sectionProp = sectionPropQueue.Dequeue();
                var p = sectionProp.SiblingsBeforeSelfReverseDocumentOrder().First();
                if (p.Element(W.pPr) == null) p.AddFirst(new XElement(W.pPr));
                p.Element(W.pPr).Add(sectionProp); // TODO this adds an additional sections
                sectionProp.Remove();
            }
        }
        public static void InitEmptyHeaderOrFooter(this MainDocumentPart mainDocPart, XElement sect, XName referenceXName, string toType)
        {
            XDocument xDoc = null;
            if (referenceXName == W.headerReference)
            {
                xDoc = XDocument.Parse(
                    @"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                    <w:hdr xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
                           xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
                           xmlns:o='urn:schemas-microsoft-com:office:office'
                           xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                           xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
                           xmlns:v='urn:schemas-microsoft-com:vml'
                           xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
                           xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                           xmlns:w10='urn:schemas-microsoft-com:office:word'
                           xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                           xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
                           xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'
                           xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
                           xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
                           xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
                           xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
                           mc:Ignorable='w14 w15 wp14'>
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                        </w:pPr>
                        <w:r>
                          <w:t></w:t>
                        </w:r>
                      </w:p>
                    </w:hdr>");
                var newHeaderPart = mainDocPart.AddNewPart<HeaderPart>();
                newHeaderPart.PutXDocument(xDoc);
                var referenceToAdd = new XElement(W.headerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, mainDocPart.GetIdOfPart(newHeaderPart)));
                sect.AddFirst(referenceToAdd);
            }
            else
            {
                xDoc = XDocument.Parse(@"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                    <w:ftr xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
                           xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
                           xmlns:o='urn:schemas-microsoft-com:office:office'
                           xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                           xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
                           xmlns:v='urn:schemas-microsoft-com:vml'
                           xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
                           xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                           xmlns:w10='urn:schemas-microsoft-com:office:word'
                           xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                           xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
                           xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'
                           xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
                           xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
                           xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
                           xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
                           mc:Ignorable='w14 w15 wp14'>
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val='Footer' />
                        </w:pPr>
                        <w:r>
                          <w:t></w:t>
                        </w:r>
                      </w:p>
                    </w:ftr>");
                var newFooterPart = mainDocPart.AddNewPart<FooterPart>();
                newFooterPart.PutXDocument(xDoc);
                var referenceToAdd = new XElement(W.footerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, mainDocPart.GetIdOfPart(newFooterPart)));
                sect.AddFirst(referenceToAdd);
            }
        }
        public static void LinkToPreviousHeadersAndFooters(this WordprocessingDocument doc)
        {
            CachedHeaderFooter[] cachedHeaderFooter = new[]
            {
                new CachedHeaderFooter() { Ref = W.headerReference, Type = "first" },
                new CachedHeaderFooter() { Ref = W.headerReference, Type = "even" },
                new CachedHeaderFooter() { Ref = W.headerReference, Type = "default" },
                new CachedHeaderFooter() { Ref = W.footerReference, Type = "first" },
                new CachedHeaderFooter() { Ref = W.footerReference, Type = "even" },
                new CachedHeaderFooter() { Ref = W.footerReference, Type = "default" },
            };

            var mdXDoc = doc.MainDocumentPart.GetXDocument();
            var sections = mdXDoc.Descendants(W.sectPr).ToList();
            var firstSection = true;
            foreach (var sect in sections)
            {
                if (firstSection)
                {
                    var headerFirst = sect.FindReference(W.headerReference, "first");
                    var headerDefault = sect.FindReference(W.headerReference, "default");
                    var headerEven = sect.FindReference(W.headerReference, "even");
                    var footerFirst = sect.FindReference(W.footerReference, "first");
                    var footerDefault = sect.FindReference(W.footerReference, "default");
                    var footerEven = sect.FindReference(W.footerReference, "even");

                    if (headerEven == null)
                    {
                        if (headerDefault != null)
                            sect.AddReferenceToExistingHeaderOrFooter((string)headerDefault.Attribute(R.id), W.headerReference, "even");
                        else
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, W.headerReference, "even");
                    }

                    if (headerFirst == null)
                    {
                        if (headerDefault != null)
                            sect.AddReferenceToExistingHeaderOrFooter((string)headerDefault.Attribute(R.id), W.headerReference, "first");
                        else
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, W.headerReference, "first");
                    }

                    if (footerEven == null)
                    {
                        if (footerDefault != null)
                            sect.AddReferenceToExistingHeaderOrFooter((string)footerDefault.Attribute(R.id), W.footerReference, "even");
                        else
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, W.footerReference, "even");
                    }

                    if (footerFirst == null)
                    {
                        if (footerDefault != null)
                            sect.AddReferenceToExistingHeaderOrFooter((string)footerDefault.Attribute(R.id), W.footerReference, "first");
                        else
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, W.footerReference, "first");
                    }

                    foreach (var hf in cachedHeaderFooter)
                    {
                        if (sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type) == null)
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, hf.Ref, hf.Type);
                        var reference = sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type);
                        if (reference == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        hf.CachedPartRid = (string)reference.Attribute(R.id);
                    }
                    firstSection = false;
                    continue;
                }
                else
                {
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "first");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "even");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "default");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "first");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "even");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "default");
                }
            }
            doc.MainDocumentPart.PutXDocument();
        }
        public static void MergeStyles(this WordprocessingDocument source, WordprocessingDocument target, XDocument sourceStyles, XDocument targetStyles, IEnumerable<XElement> newContent)
        {
            //var newIds = new Dictionary<string, string>();

            if (sourceStyles.Root == null)
                return;

            foreach (XElement style in sourceStyles.Root.Elements(W.style))
            {
                var fromId = (string)style.Attribute(W.styleId);
                var fromName = (string)style.Elements(W.name).Attributes(W.val).FirstOrDefault();

                var toStyle = targetStyles
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(st => (string)st.Elements(W.name).Attributes(W.val).FirstOrDefault() == fromName);

                if (toStyle == null)
                {
                    //var linkElement = style.Element(W.link);
                    //string linkedId;
                    //if (linkElement != null && newIds.TryGetValue(linkElement.Attribute(W.val).Value, out linkedId))
                    //{
                    //    var linkedStyle = toStyles.Root.Elements(W.style)
                    //        .First(o => o.Attribute(W.styleId).Value == linkedId);
                    //    if (linkedStyle.Element(W.link) != null)
                    //        newIds.Add(fromId, linkedStyle.Element(W.link).Attribute(W.val).Value);
                    //    continue;
                    //}

                    //string name = (string)style.Elements(W.name).Attributes(W.val).FirstOrDefault();
                    //var namedStyle = toStyles
                    //    .Root
                    //    .Elements(W.style)
                    //    .Where(st => st.Element(W.name) != null)
                    //    .FirstOrDefault(o => (string)o.Element(W.name).Attribute(W.val) == name);
                    //if (namedStyle != null)
                    //{
                    //    if (! newIds.ContainsKey(fromId))
                    //        newIds.Add(fromId, namedStyle.Attribute(W.styleId).Value);
                    //    continue;
                    //}


                    int number = 1;
                    int abstractNumber = 0;
                    XDocument oldNumbering = null;
                    XDocument newNumbering = null;
                    foreach (XElement numReference in style.Descendants(W.numPr))
                    {
                        XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                        if (idElement != null)
                        {
                            if (oldNumbering == null)
                            {
                                if (source.MainDocumentPart.NumberingDefinitionsPart != null)
                                    oldNumbering = source.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                else
                                {
                                    oldNumbering = new XDocument();
                                    oldNumbering.Declaration = Commons.Common.CreateDeclaration();
                                    oldNumbering.Add(new XElement(W.numbering, Constants.NamespaceAttributes));
                                }
                            }
                            if (newNumbering == null)
                            {
                                if (target.MainDocumentPart.NumberingDefinitionsPart != null)
                                {
                                    newNumbering = target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    newNumbering.Declaration.SetDeclaration();
                                    var numIds = newNumbering
                                        .Root
                                        .Elements(W.num)
                                        .Select(f => (int)f.Attribute(W.numId));
                                    if (numIds.Any())
                                        number = numIds.Max() + 1;
                                    numIds = newNumbering
                                        .Root
                                        .Elements(W.abstractNum)
                                        .Select(f => (int)f.Attribute(W.abstractNumId));
                                    if (numIds.Any())
                                        abstractNumber = numIds.Max() + 1;
                                }
                                else
                                {
                                    target.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                                    newNumbering = target.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    newNumbering.Declaration.SetDeclaration();
                                    newNumbering.Add(new XElement(W.numbering, Constants.NamespaceAttributes));
                                }
                            }
                            string numId = idElement.Attribute(W.val).Value;
                            if (numId != "0")
                            {
                                XElement element = oldNumbering
                                    .Descendants()
                                    .Elements(W.num)
                                    .Where(p => ((string)p.Attribute(W.numId)) == numId)
                                    .FirstOrDefault();

                                // Copy abstract numbering element, if necessary (use matching NSID)
                                string abstractNumId = string.Empty;
                                if (element != null)
                                {
                                    abstractNumId = element
                                       .Elements(W.abstractNumId)
                                       .First()
                                       .Attribute(W.val)
                                       .Value;

                                    XElement abstractElement = oldNumbering
                                        .Descendants()
                                        .Elements(W.abstractNum)
                                        .Where(p => ((string)p.Attribute(W.abstractNumId)) == abstractNumId)
                                        .FirstOrDefault();
                                    string abstractNSID = string.Empty;
                                    if (abstractElement != null)
                                    {
                                        XElement nsidElement = abstractElement
                                            .Element(W.nsid);
                                        abstractNSID = null;
                                        if (nsidElement != null)
                                            abstractNSID = (string)nsidElement
                                                .Attribute(W.val);

                                        XElement newAbstractElement = newNumbering
                                            .Descendants()
                                            .Elements(W.abstractNum)
                                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                            .Where(p =>
                                            {
                                                var thisNsidElement = p.Element(W.nsid);
                                                if (thisNsidElement == null)
                                                    return false;
                                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                                            })
                                            .FirstOrDefault();
                                        if (newAbstractElement == null)
                                        {
                                            newAbstractElement = new XElement(abstractElement);
                                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                                            abstractNumber++;
                                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                                            else
                                                newNumbering.Root.Add(newAbstractElement);

                                            foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                                            {
                                                string bulletId = (string)pictId.Attribute(W.val);
                                                XElement numPicBullet = oldNumbering
                                                    .Descendants(W.numPicBullet)
                                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                                int maxNumPicBulletId = new int[] { -1 }.Concat(
                                                    newNumbering.Descendants(W.numPicBullet)
                                                    .Attributes(W.numPicBulletId)
                                                    .Select(a => (int)a))
                                                    .Max() + 1;
                                                XElement newNumPicBullet = new XElement(numPicBullet);
                                                newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                                pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                                newNumbering.Root.AddFirst(newNumPicBullet);
                                            }
                                        }
                                        string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                                        // Copy numbering element, if necessary (use matching element with no overrides)
                                        XElement newElement = null;
                                        if (!element.Elements(W.lvlOverride).Any())
                                            newElement = newNumbering
                                                .Descendants()
                                                .Elements(W.num)
                                                .Where(p => !p.Elements(W.lvlOverride).Any() &&
                                                    ((string)p.Elements(W.abstractNumId).First().Attribute(W.val)) == newAbstractId)
                                                .FirstOrDefault();
                                        if (newElement == null)
                                        {
                                            newElement = new XElement(element);
                                            newElement
                                                .Elements(W.abstractNumId)
                                                .First()
                                                .Attribute(W.val).Value = newAbstractId;
                                            newElement.Attribute(W.numId).Value = number.ToString();
                                            number++;
                                            newNumbering.Root.Add(newElement);
                                        }
                                        idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                                    }
                                }
                            }
                        }
                    }

                    var newStyle = new XElement(style);
                    // get rid of anything not in the w: namespace
                    newStyle.Descendants().Where(d => d.Name.NamespaceName != W.w).Remove();
                    newStyle.Descendants().Attributes().Where(d => d.Name.NamespaceName != W.w).Remove();
                    targetStyles.Root.Add(newStyle);
                }
                else
                {
                    var toId = (string)toStyle.Attribute(W.styleId);
                    if (fromId != toId)
                    {
                        //if (!newIds.ContainsKey(fromId))
                        //    newIds.Add(fromId, toId);
                    }
                }
            }

#if MergeStylesWithSameNames
            if (newIds.Count > 0)
            {
                foreach (var style in toStyles
                    .Root
                    .Elements(W.style))
                {
                    ConvertToNewId(style.Element(W.basedOn), newIds);
                    ConvertToNewId(style.Element(W.next), newIds);
                }

                foreach (var item in newContent.DescendantsAndSelf()
                    .Where(d => d.Name == W.pStyle ||
                                d.Name == W.rStyle ||
                                d.Name == W.tblStyle))
                {
                    ConvertToNewId(item, newIds);
                }

                if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    var newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    ConvertNumberingPartToNewIds(newNumbering, newIds);
                }

                // Convert WmlSource document, since numberings will be copied over after styles.
                if (sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    var sourceNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    ConvertNumberingPartToNewIds(sourceNumbering, newIds);
                }
            }
#endif
        }
        public static void RemoveHeadersAndFootersFromSections(this WordprocessingDocument doc)
        {
            var mdXDoc = doc.MainDocumentPart.GetXDocument();
            var sections = mdXDoc.Descendants(W.sectPr).ToList();
            foreach (var sect in sections)
            {
                sect.Elements(W.headerReference).Remove();
                sect.Elements(W.footerReference).Remove();
            }
            doc.MainDocumentPart.PutXDocument();
        }
        public static void RemoveSections(this WordprocessingDocument doc)
        {
            var mdXDoc = doc.MainDocumentPart.GetXDocument();
            var sections = mdXDoc.Descendants(W.sectPr).ToList();
            foreach (var sect in sections)
            {
                sect.RemoveAll();
                sect.Remove();
            }
            doc.MainDocumentPart.PutXDocument();
        }
        public static void NormalisePageLayout(this WordprocessingDocument doc, WmlPackage package)
        {
            var body = doc.MainDocumentPart.GetBody();
            if (package.Section == null)
            {
                package.Section = body.Descendants(W.sectPr).FirstOrDefault();
            }
            else
            {
                var sects = body.Descendants(W.sectPr).ToList();
                var source = package.Section;
                foreach (var sect in sects)
                {
                    sect.AddBeforeSelf(source);
                    sect.Remove();
                    //sect.Element(W.type).Value = source.Element(W.type).Value;
                    //sect.Element(W.pgSz).Value = source.Element(W.pgSz).Value;
                    //sect.Element(W.pgMar).Value = source.Element(W.pgMar).Value;
                    //sect.Element(W.cols).Value = source.Element(W.cols).Value;
                    //sect.Element(W.titlePg).Value = source.Element(W.titlePg).Value;
                }
                doc.MainDocumentPart.PutXDocument();
            }
        }

        #endregion

        #region Queries
        public static bool AnySections(this MainDocumentPart mainPart) => mainPart.GetXDocument().Root.Descendants(W.sectPr).Any();
        public static XElement GetBody(this MainDocumentPart mainPart)
        {
            return mainPart.GetXDocument().Root.Element(W.body);
        }
        public static IEnumerable<XElement> GetBodyElements(this MainDocumentPart mainPart)
        {
            return mainPart.GetBody().Elements();
        }
        public static IEnumerable<XElement> GetContents(this MainDocumentPart mainPart, int start = 0, int count = int.MaxValue)
        {
            return mainPart.GetBodyElements().Skip(start).Take(count).ToList();
        }
        public static IEnumerable<FooterPart> GetFooterParts(this MainDocumentPart mainPart) => mainPart.FooterParts;
        public static IEnumerable<HeaderPart> GetHeaderParts(this MainDocumentPart mainPart) => mainPart.HeaderParts;
        public static IEnumerable<Atbid> GetDivs(this MainDocumentPart mainPart)
        {
            return mainPart.GetXDocument().Root.Element(W.body).Elements()
                .Select((p, i) => new Atbid { BlockLevelContent = p, Index = i, })
                .Rollup(new Atbid { BlockLevelContent = null, Index = -1, Div = 0, }, (b, p) =>
                {
                    XElement elementBefore = b.BlockLevelContent.SiblingsBeforeSelfReverseDocumentOrder().FirstOrDefault();

                    return (elementBefore != null && elementBefore.Descendants(W.sectPr).Any())
                        ? new Atbid { BlockLevelContent = b.BlockLevelContent, Index = b.Index, Div = p.Div + 1, }
                        : new Atbid { BlockLevelContent = b.BlockLevelContent, Index = b.Index, Div = p.Div };
                });
        }
        public static Dictionary<string, string> GetStyleNameMap(this MainDocumentPart mainPart)
        {
            return mainPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.style)
                .ToDictionary(z => (string)z.Elements(W.name).Attributes(W.val).FirstOrDefault(), z => (string)z.Attribute(W.styleId));
        }
        public static XElement GetLastElement(this MainDocumentPart mainPart) => mainPart.GetXDocument().Root.Element(W.body).Elements().LastOrDefault();
        public static XDocument GetFontTablePart(this MainDocumentPart mainPart) => mainPart.FontTablePart.GetXDocument();
        public static XDocument GetEndnotesPart(this MainDocumentPart mainPart) => mainPart.EndnotesPart.GetXDocument();
        public static XDocument GetStylePart(this MainDocumentPart mainPart) => mainPart.StyleDefinitionsPart.GetXDocument();
        public static XDocument GetStylesWithEffectsPart(this MainDocumentPart mainPart) => mainPart.StylesWithEffectsPart.GetXDocument();
        #endregion

        public static void TestForUnsupportedDocument(this WordprocessingDocument doc, int sourceNumber)
        {
            //What does not work:
            //- sub docs
            //- bidi text appears to work but has not been tested
            //- languages other than en-us appear to work but have not been tested
            //- documents with activex controls
            //- mail merge WmlSource documents (look for dataSource in settings)
            //- documents with ink
            //- documents with frame sets and frames

            if (doc.MainDocumentPart.GetXDocument().Root == null)
                throw new DocumentBuilderException(string.Format("Source {0} is an invalid document - MainDocumentPart contains no content.", sourceNumber));

            if ((string)doc.MainDocumentPart.GetXDocument().Root.Name.NamespaceName == "http://purl.oclc.org/ooxml/wordprocessingml/main")
                throw new DocumentBuilderException(string.Format("Source {0} is saved in strict mode, not supported", sourceNumber));

            // note: if ever want to support section changes, need to address the code that rationalizes headers and footers, propagating to sections that inherit headers/footers from prev section
            foreach (var d in doc.MainDocumentPart.GetXDocument().Descendants())
            {
                if (d.Name == W.sectPrChange)
                    throw new DocumentBuilderException(string.Format("Source {0} contains section changes (w:sectPrChange), not supported", sourceNumber));

                // note: if ever want to support Open-Xml-PowerTools attributes, need to make sure that all attributes are propagated in all cases
                //if (d.Name.Namespace == PtOpenXml.ptOpenXml ||
                //    d.Name.Namespace == PtOpenXml.pt)
                //    throw new DocumentBuilderException(string.Format("Source {0} contains Open-Xml-PowerTools markup, not supported", sourceNumber));
                //if (d.Attributes().Any(a => a.Name.Namespace == PtOpenXml.ptOpenXml || a.Name.Namespace == PtOpenXml.pt))
                //    throw new DocumentBuilderException(string.Format("Source {0} contains Open-Xml-PowerTools markup, not supported", sourceNumber));
            }

            doc.MainDocumentPart.TestPartForUnsupportedContent(sourceNumber);
            foreach (var hdr in doc.MainDocumentPart.HeaderParts)
                hdr.TestPartForUnsupportedContent(sourceNumber);
            foreach (var ftr in doc.MainDocumentPart.FooterParts)
                ftr.TestPartForUnsupportedContent(sourceNumber);
            if (doc.MainDocumentPart.FootnotesPart != null)
                doc.MainDocumentPart.FootnotesPart.TestPartForUnsupportedContent(sourceNumber);
            if (doc.MainDocumentPart.EndnotesPart != null)
                doc.MainDocumentPart.EndnotesPart.TestPartForUnsupportedContent(sourceNumber);

            if (doc.MainDocumentPart.DocumentSettingsPart != null &&
                doc.MainDocumentPart.DocumentSettingsPart.GetXDocument().Descendants().Any(d => d.Name == W.src ||
                d.Name == W.recipientData || d.Name == W.mailMerge))
                throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains Mail Merge content",
                    sourceNumber));
            if (doc.MainDocumentPart.WebSettingsPart != null &&
                doc.MainDocumentPart.WebSettingsPart.GetXDocument().Descendants().Any(d => d.Name == W.frameset))
                throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains a frameset", sourceNumber));
            var numberingElements = doc.MainDocumentPart
                .GetXDocument()
                .Descendants(W.numPr)
                .Where(n =>
                {
                    bool zeroId = (int?)n.Attribute(W.id) == 0;
                    bool hasChildInsId = n.Elements(W.ins).Any();
                    if (zeroId || hasChildInsId)
                        return false;
                    return true;
                })
                .ToList();
            if (numberingElements.Any() &&
                doc.MainDocumentPart.NumberingDefinitionsPart == null)
                throw new DocumentBuilderException(String.Format(
                    "Source {0} is invalid document - contains numbering markup but no numbering part", sourceNumber));
        }
        public static void TestPartForUnsupportedContent(this OpenXmlPart part, int sourceNumber)
        {
            XNamespace[] obsoleteNamespaces = new[]
                {
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2007/5/30/wordml"),
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2008/9/16/wordprocessingDrawing"),
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2009/2/wordml"),
                };
            XDocument xDoc = part.GetXDocument();
            XElement invalidElement = xDoc.Descendants()
                .FirstOrDefault(d =>
                {
                    bool b = d.Name == W.subDoc ||
                        d.Name == W.control ||
                        d.Name == W.altChunk ||
                        d.Name.LocalName == "contentPart" ||
                        obsoleteNamespaces.Contains(d.Name.Namespace);
                    bool b2 = b ||
                        d.Attributes().Any(a => obsoleteNamespaces.Contains(a.Name.Namespace));
                    return b2;
                });
            if (invalidElement != null)
            {
                if (invalidElement.Name == W.subDoc)
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains sub document",
                        sourceNumber));
                if (invalidElement.Name == W.control)
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains ActiveX controls",
                        sourceNumber));
                if (invalidElement.Name == W.altChunk)
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains altChunk",
                        sourceNumber));
                if (invalidElement.Name.LocalName == "contentPart")
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains contentPart content",
                        sourceNumber));
                if (obsoleteNamespaces.Contains(invalidElement.Name.Namespace) ||
                    invalidElement.Attributes().Any(a => obsoleteNamespaces.Contains(a.Name.Namespace)))
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains obsolete namespace",
                        sourceNumber));
            }
        }

        #endregion

        #region Collections - At bottom because they don't collapse with collapse all.
        public static CachedHeaderFooter[] CachedHeadersAndFooters = new CachedHeaderFooter[]
        {
            new CachedHeaderFooter() { Ref = W.headerReference, Type = "first" },
            new CachedHeaderFooter() { Ref = W.headerReference, Type = "even" },
            new CachedHeaderFooter() { Ref = W.headerReference, Type = "default" },
            new CachedHeaderFooter() { Ref = W.footerReference, Type = "first" },
            new CachedHeaderFooter() { Ref = W.footerReference, Type = "even" },
            new CachedHeaderFooter() { Ref = W.footerReference, Type = "default" },
        };
        public static Dictionary<XName, int> ParagraphBordersOrder = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.bottom, 30 },
            { W.right, 40 },
            { W.between, 50 },
            { W.bar, 60 },
        };
        public static Dictionary<XName, int> ParagraphPropertiesOrder = new Dictionary<XName, int>
        {
            { W.pStyle, 10 },
            { W.keepNext, 20 },
            { W.keepLines, 30 },
            { W.pageBreakBefore, 40 },
            { W.framePr, 50 },
            { W.widowControl, 60 },
            { W.numPr, 70 },
            { W.suppressLineNumbers, 80 },
            { W.pBdr, 90 },
            { W.shd, 100 },
            { W.tabs, 120 },
            { W.suppressAutoHyphens, 130 },
            { W.kinsoku, 140 },
            { W.wordWrap, 150 },
            { W.overflowPunct, 160 },
            { W.topLinePunct, 170 },
            { W.autoSpaceDE, 180 },
            { W.autoSpaceDN, 190 },
            { W.bidi, 200 },
            { W.adjustRightInd, 210 },
            { W.snapToGrid, 220 },
            { W.spacing, 230 },
            { W.ind, 240 },
            { W.contextualSpacing, 250 },
            { W.mirrorIndents, 260 },
            { W.suppressOverlap, 270 },
            { W.jc, 280 },
            { W.textDirection, 290 },
            { W.textAlignment, 300 },
            { W.textboxTightWrap, 310 },
            { W.outlineLvl, 320 },
            { W.divId, 330 },
            { W.cnfStyle, 340 },
            { W.rPr, 350 },
            { W.sectPr, 360 },
            { W.pPrChange, 370 },
        };
        public static Dictionary<XName, int> RunPropertiesOrder = new Dictionary<XName, int>
        {
            { W.moveFrom, 5 },
            { W.moveTo, 7 },
            { W.ins, 10 },
            { W.del, 20 },
            { W.rStyle, 30 },
            { W.rFonts, 40 },
            { W.b, 50 },
            { W.bCs, 60 },
            { W.i, 70 },
            { W.iCs, 80 },
            { W.caps, 90 },
            { W.smallCaps, 100 },
            { W.strike, 110 },
            { W.dstrike, 120 },
            { W.outline, 130 },
            { W.shadow, 140 },
            { W.emboss, 150 },
            { W.imprint, 160 },
            { W.noProof, 170 },
            { W.snapToGrid, 180 },
            { W.vanish, 190 },
            { W.webHidden, 200 },
            { W.color, 210 },
            { W.spacing, 220 },
            { W._w, 230 },
            { W.kern, 240 },
            { W.position, 250 },
            { W.sz, 260 },
            { W14.wShadow, 270 },
            { W14.wTextOutline, 280 },
            { W14.wTextFill, 290 },
            { W14.wScene3d, 300 },
            { W14.wProps3d, 310 },
            { W.szCs, 320 },
            { W.highlight, 330 },
            { W.u, 340 },
            { W.effect, 350 },
            { W.bdr, 360 },
            { W.shd, 370 },
            { W.fitText, 380 },
            { W.vertAlign, 390 },
            { W.rtl, 400 },
            { W.cs, 410 },
            { W.em, 420 },
            { W.lang, 430 },
            { W.eastAsianLayout, 440 },
            { W.specVanish, 450 },
            { W.oMath, 460 },
        };
        public static Dictionary<XName, int> SettingsOrder = new Dictionary<XName, int>
        {
            { W.writeProtection, 10},
            { W.view, 20},
            { W.zoom, 30},
            { W.removePersonalInformation, 40},
            { W.removeDateAndTime, 50},
            { W.doNotDisplayPageBoundaries, 60},
            { W.displayBackgroundShape, 70},
            { W.printPostScriptOverText, 80},
            { W.printFractionalCharacterWidth, 90},
            { W.printFormsData, 100},
            { W.embedTrueTypeFonts, 110},
            { W.embedSystemFonts, 120},
            { W.saveSubsetFonts, 130},
            { W.saveFormsData, 140},
            { W.mirrorMargins, 150},
            { W.alignBordersAndEdges, 160},
            { W.bordersDoNotSurroundHeader, 170},
            { W.bordersDoNotSurroundFooter, 180},
            { W.gutterAtTop, 190},
            { W.hideSpellingErrors, 200},
            { W.hideGrammaticalErrors, 210},
            { W.activeWritingStyle, 220},
            { W.proofState, 230},
            { W.formsDesign, 240},
            { W.attachedTemplate, 250},
            { W.linkStyles, 260},
            { W.stylePaneFormatFilter, 270},
            { W.stylePaneSortMethod, 280},
            { W.documentType, 290},
            { W.mailMerge, 300},
            { W.revisionView, 310},
            { W.trackRevisions, 320},
            { W.doNotTrackMoves, 330},
            { W.doNotTrackFormatting, 340},
            { W.documentProtection, 350},
            { W.autoFormatOverride, 360},
            { W.styleLockTheme, 370},
            { W.styleLockQFSet, 380},
            { W.defaultTabStop, 390},
            { W.autoHyphenation, 400},
            { W.consecutiveHyphenLimit, 410},
            { W.hyphenationZone, 420},
            { W.doNotHyphenateCaps, 430},
            { W.showEnvelope, 440},
            { W.summaryLength, 450},
            { W.clickAndTypeStyle, 460},
            { W.defaultTableStyle, 470},
            { W.evenAndOddHeaders, 480},
            { W.bookFoldRevPrinting, 490},
            { W.bookFoldPrinting, 500},
            { W.bookFoldPrintingSheets, 510},
            { W.drawingGridHorizontalSpacing, 520},
            { W.drawingGridVerticalSpacing, 530},
            { W.displayHorizontalDrawingGridEvery, 540},
            { W.displayVerticalDrawingGridEvery, 550},
            { W.doNotUseMarginsForDrawingGridOrigin, 560},
            { W.drawingGridHorizontalOrigin, 570},
            { W.drawingGridVerticalOrigin, 580},
            { W.doNotShadeFormData, 590},
            { W.noPunctuationKerning, 600},
            { W.characterSpacingControl, 610},
            { W.printTwoOnOne, 620},
            { W.strictFirstAndLastChars, 630},
            { W.noLineBreaksAfter, 640},
            { W.noLineBreaksBefore, 650},
            { W.savePreviewPicture, 660},
            { W.doNotValidateAgainstSchema, 670},
            { W.saveInvalidXml, 680},
            { W.ignoreMixedContent, 690},
            { W.alwaysShowPlaceholderText, 700},
            { W.doNotDemarcateInvalidXml, 710},
            { W.saveXmlDataOnly, 720},
            { W.useXSLTWhenSaving, 730},
            { W.saveThroughXslt, 740},
            { W.showXMLTags, 750},
            { W.alwaysMergeEmptyNamespace, 760},
            { W.updateFields, 770},
            { W.footnotePr, 780},
            { W.endnotePr, 790},
            { W.compat, 800},
            { W.docVars, 810},
            { W.rsids, 820},
            { M.mathPr, 830},
            { W.attachedSchema, 840},
            { W.themeFontLang, 850},
            { W.clrSchemeMapping, 860},
            { W.doNotIncludeSubdocsInStats, 870},
            { W.doNotAutoCompressPictures, 880},
            { W.forceUpgrade, 890}, 
            //{W.captions, 900}, 
            { W.readModeInkLockDown, 910},
            { W.smartTagType, 920}, 
            //{W.sl:schemaLibrary, 930}, 
            { W.doNotEmbedSmartTags, 940},
            { W.decimalSymbol, 950},
            { W.listSeparator, 960},
        };
        public static Dictionary<XName, int> TableBordersOrder = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.start, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
        };
        public static Dictionary<XName, int> TableCellBordersOrder = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.start, 20 },
            { W.left, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
            { W.tl2br, 90 },
            { W.tr2bl, 100 },
        };
        public static Dictionary<XName, int> TableCellPropertiesOrder = new Dictionary<XName, int>
        {
            { W.cnfStyle, 10 },
            { W.tcW, 20 },
            { W.gridSpan, 30 },
            { W.hMerge, 40 },
            { W.vMerge, 50 },
            { W.tcBorders, 60 },
            { W.shd, 70 },
            { W.noWrap, 80 },
            { W.tcMar, 90 },
            { W.textDirection, 100 },
            { W.tcFitText, 110 },
            { W.vAlign, 120 },
            { W.hideMark, 130 },
            { W.headers, 140 },
        };
        public static Dictionary<XName, int> TablePropertiesOrder = new Dictionary<XName, int>
        {
            { W.tblStyle, 10 },
            { W.tblpPr, 20 },
            { W.tblOverlap, 30 },
            { W.bidiVisual, 40 },
            { W.tblStyleRowBandSize, 50 },
            { W.tblStyleColBandSize, 60 },
            { W.tblW, 70 },
            { W.jc, 80 },
            { W.tblCellSpacing, 90 },
            { W.tblInd, 100 },
            { W.tblBorders, 110 },
            { W.shd, 120 },
            { W.tblLayout, 130 },
            { W.tblCellMar, 140 },
            { W.tblLook, 150 },
            { W.tblCaption, 160 },
            { W.tblDescription, 170 },
        };
        public static Dictionary<XName, XName[]> RelationshipMarkup => new Dictionary<XName, XName[]>()
        {
            //{ button,           new [] { image }},
            { A.blip,             new [] { R.embed, R.link }},
            { A.hlinkClick,       new [] { R.id }},
            { A.relIds,           new [] { R.cs, R.dm, R.lo, R.qs }},
            //{ a14:imgLayer,     new [] { R.embed }},
            //{ ax:ocx,           new [] { R.id }},
            { C.chart,            new [] { R.id }},
            { C.externalData,     new [] { R.id }},
            { C.userShapes,       new [] { R.id }},
            { DGM.relIds,         new [] { R.cs, R.dm, R.lo, R.qs }},
            { O.OLEObject,        new [] { R.id }},
            { VML.fill,           new [] { R.id }},
            { VML.imagedata,      new [] { R.href, R.id, R.pict }},
            { VML.stroke,         new [] { R.id }},
            { W.altChunk,         new [] { R.id }},
            { W.attachedTemplate, new [] { R.id }},
            { W.control,          new [] { R.id }},
            { W.dataSource,       new [] { R.id }},
            { W.embedBold,        new [] { R.id }},
            { W.embedBoldItalic,  new [] { R.id }},
            { W.embedItalic,      new [] { R.id }},
            { W.embedRegular,     new [] { R.id }},
            { W.footerReference,  new [] { R.id }},
            { W.headerReference,  new [] { R.id }},
            { W.headerSource,     new [] { R.id }},
            { W.hyperlink,        new [] { R.id }},
            { W.printerSettings,  new [] { R.id }},
            { W.recipientData,    new [] { R.id }},  // Mail merge, not required
            { W.saveThroughXslt,  new [] { R.id }},
            { W.sourceFileName,   new [] { R.id }},  // Framesets, not required
            { W.src,              new [] { R.id }},  // Mail merge, not required
            { W.subDoc,           new [] { R.id }},  // Sub documents, not required
            //{ w14:contentPart,  new [] { R.id }},
            { WNE.toolbarData,    new [] { R.id }},
        };
        public static string[] Extensions = new[] { ".docx", ".docm", ".dotx", ".dotm", };
        #endregion
    }
}