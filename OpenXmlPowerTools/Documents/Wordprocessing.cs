using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Commons;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    public static class Wordprocessing
    {
        public static Dictionary<XName, XName[]> RelationshipMarkup => Constants.WordprocessingRelationshipMarkup;

        private static HashSet<string> UnknownFonts = new HashSet<string>();
        private static HashSet<string> KnownFamilies = null;

        public static string[] Extensions = new[] {
            ".docx",
            ".docm",
            ".dotx",
            ".dotm",
        };
        public static Dictionary<XName, int> OrderPBdr = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.bottom, 30 },
            { W.right, 40 },
            { W.between, 50 },
            { W.bar, 60 },
        };
        public static Dictionary<XName, int> OrderSettings = new Dictionary<XName, int>
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
        public static Dictionary<XName, int> OrderTableBorders = new Dictionary<XName, int>
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
        public static Dictionary<XName, int> OrderTableCellBorders = new Dictionary<XName, int>
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
        public static Dictionary<XName, int> OrderTableCellPart = new Dictionary<XName, int>
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
        public static Dictionary<XName, int> OrderTablePart = new Dictionary<XName, int>
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
        public static bool IsWordprocessing(string ext) => Extensions.Contains(ext.ToLower());

        public static int CalcWidthOfRunInTwips(XElement r)
        {
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
            if (runText.Length <= 2)
                multiplier = 100;
            else if (runText.Length <= 4)
                multiplier = 50;
            else if (runText.Length <= 8)
                multiplier = 25;
            else if (runText.Length <= 16)
                multiplier = 12;
            else if (runText.Length <= 32)
                multiplier = 6;
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

        public static bool GetBoolProp(XElement runProps, XName xName)
        {
            var p = runProps.Element(xName);
            if (p == null)
                return false;
            var v = p.Attribute(W.val);
            if (v == null)
                return true;
            var s = v.Value.ToLower();
            if (s == "0" || s == "false")
                return false;
            if (s == "1" || s == "true")
                return true;
            return false;
        }

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

        public static int? AttributeToTwips(XAttribute attribute)
        {
            if (attribute == null)
            {
                return null;
            }

            string twipsOrPoints = (string)attribute;

            // if the pos value is in points, not twips
            if (twipsOrPoints.EndsWith("pt"))
            {
                decimal decimalValue = decimal.Parse(twipsOrPoints.Substring(0, twipsOrPoints.Length - 2));
                return (int)(decimalValue * 20);
            }
            if (twipsOrPoints.Contains('.'))
            {
                decimal decimalValue = decimal.Parse(twipsOrPoints);
                return (int)decimalValue;
            }
            return int.Parse(twipsOrPoints);
        }

        public static XElement CoalesceAdjacentRunsWithIdenticalFormatting(XElement runContainer)
        {
            const string dontConsolidate = "DontConsolidate";

            IEnumerable<IGrouping<string, XElement>> groupedAdjacentRunsWithIdenticalFormatting =
                runContainer
                    .Elements()
                    .GroupAdjacent(ce =>
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
        public static object WmlOrderElementsPerStandard(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.pPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Constants.OrderPPr.ContainsKey(e.Name))
                                return Constants.OrderPPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.rPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Constants.OrderRPr.ContainsKey(e.Name))
                                return Constants.OrderRPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (OrderTablePart.ContainsKey(e.Name))
                                return OrderTablePart[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (OrderTableCellPart.ContainsKey(e.Name))
                                return OrderTableCellPart[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (OrderTableBorders.ContainsKey(e.Name))
                                return OrderTableBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (OrderTableBorders.ContainsKey(e.Name))
                                return OrderTableBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.pBdr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (OrderPBdr.ContainsKey(e.Name))
                                return OrderPBdr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.p)
                {
                    var newP = new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                        element.Elements().Where(e => e.Name != W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));
                    return newP;
                }

                if (element.Name == W.r)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                        element.Elements().Where(e => e.Name != W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));

                if (element.Name == W.settings)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (OrderSettings.ContainsKey(e.Name))
                                return OrderSettings[e.Name];
                            return 999;
                        }));

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => WmlOrderElementsPerStandard(n)));
            }
            return node;
        }

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


        #region Document
        public static void AddReferenceToExistingHeaderOrFooter(this MainDocumentPart mainDocPart, XElement sect, string rId, XName reference, string toType)
        {
            if (reference == W.headerReference)
            {
                var referenceToAdd = new XElement(W.headerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
            }
            else
            {
                var referenceToAdd = new XElement(W.footerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
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

        public static OpenXmlPowerToolsDocument CreateSplitDocument(this WordprocessingDocument source, List<XElement> contents, string newFileName)
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

        /// <summary>
        /// Create a new main document part
        /// </summary>
        /// <returns></returns>
        public static XDocument CreateMainDocumentPart()
        {
            return new XDocument(CreateDocument());
        }

        /// <summary>
        /// Create document part
        /// </summary>
        /// <returns></returns>
        public static XElement CreateDocument()
        {
            return new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XElement(W.body));
        }
        #endregion
        
        #region Source
        public static Source AddAndRectify(this Source src, Dictionary<string, string> styleNameMap, HashSet<string> styleIds)
        {
            bool modified = false;
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(src.WmlDocument.DocumentByteArray, 0, src.WmlDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    Dictionary<string, string> correctionList = new Dictionary<string, string>();
                    var thisStyleNameMap = GetStyleNameMap(wDoc);
                    foreach (var pair in thisStyleNameMap)
                    {
                        var styleName = pair.Key;
                        var styleId = pair.Value;
                        // if the styleNameMap does not contain an entry for this name
                        if (!styleNameMap.ContainsKey(styleName))
                        {
                            // if the id is already used
                            if (styleIds.Contains(styleId))
                            {
                                // this style uses a styleId that is used for another style.
                                // randomly generate new styleId
                                while (true)
                                {
                                    var newStyleId = styleName.GenStyleIdFromStyleName();
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
                        // but if the styleNameMap does contain an entry for this name
                        else
                        {
                            // if the id is the same as the existing ID, then nothing to do
                            if (styleNameMap[styleName] == styleId)
                                continue;
                            correctionList.Add(styleId, styleNameMap[styleName]);
                        }
                    }
                    if (correctionList.Any())
                    {
                        modified = true;
                        AdjustStyleIdsForDocument(wDoc, correctionList);
                    }
                }
                if (modified)
                {
                    var newWmlDocument = new WmlDocument(src.WmlDocument.FileName, ms.ToArray());
                    var newSrc = new Source(newWmlDocument, src.Start, src.Count, src.KeepSections);
                    newSrc.DiscardHeadersAndFootersInKeptSections = src.DiscardHeadersAndFootersInKeptSections;
                    newSrc.InsertId = src.InsertId;
                    return newSrc;
                }
            }
            return src;
        }
        public static List<Source> NormalizeStyleNamesAndIds(this List<Source> sources)
        {
            // there are two scenarios that need to be handled
            // - if I find a style name that maps to a style ID different from one already mapped
            // - if a style name maps to a style ID that is already used for a different style
            // - then need to correct things
            //   - make a complete list of all things that need to be changed, for every correction
            //   - do the corrections all at one
            //   - mark the document as changed, and change it in the sources.
            Dictionary<string, string> styleNameMap = new Dictionary<string, string>();
            HashSet<string> styleIds = new HashSet<string>();
            List<Source> newSources = new List<Source>();

            foreach (var src in sources)
            {
                var newSrc = AddAndRectify(src, styleNameMap, styleIds);
                newSources.Add(newSrc);
            }
            return newSources;
        }
        public static WmlDocument CoalesceGlossaryDocumentParts(this IEnumerable<Source> sources)
        {
            List<Source> allGlossaryDocuments = sources
                .Select(s => s.WmlDocument.ExtractGlossaryDocument())
                .Where(s => s != null)
                .Select(s => new Source(s))
                .ToList();

            if (!allGlossaryDocuments.Any())
                return null;

            WmlDocument coalescedRaw = DocumentBuilder.BuildDocument(allGlossaryDocuments);

            // now need to do some fix up
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(coalescedRaw.DocumentByteArray, 0, coalescedRaw.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var mainXDoc = wDoc.MainDocumentPart.GetXDocument();

                    var newBody = new XElement(W.body,
                        new XElement(W.docParts,
                            mainXDoc.Root.Element(W.body).Elements(W.docParts).Elements(W.docPart)));

                    mainXDoc.Root.Element(W.body).ReplaceWith(newBody);

                    wDoc.MainDocumentPart.PutXDocument();
                }

                WmlDocument coalescedGlossaryDocument = new WmlDocument("Coalesced.docx", ms.ToArray());

                return coalescedGlossaryDocument;
            }
        }
        #endregion

        #region WmlDocument
        public static IEnumerable<WmlDocument> SplitOnSections(this WmlDocument doc)
        {
            List<TempSource> tempSourceList;
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
            {
                XDocument mainDocument = document.MainDocumentPart.GetXDocument();
                var divs = mainDocument
                    .Root
                    .Element(W.body)
                    .Elements()
                    .Select((p, i) => new Atbi
                    {
                        BlockLevelContent = p,
                        Index = i,
                    })
                    .Rollup(new Atbid
                    {
                        BlockLevelContent = (XElement)null,
                        Index = -1,
                        Div = 0,
                    },
                        (b, p) =>
                        {
                            XElement elementBefore = b.BlockLevelContent
                                .SiblingsBeforeSelfReverseDocumentOrder()
                                .FirstOrDefault();
                            if (elementBefore != null && elementBefore.Descendants(W.sectPr).Any())
                                return new Atbid
                                {
                                    BlockLevelContent = b.BlockLevelContent,
                                    Index = b.Index,
                                    Div = p.Div + 1,
                                };
                            return new Atbid
                            {
                                BlockLevelContent = b.BlockLevelContent,
                                Index = b.Index,
                                Div = p.Div,
                            };
                        });
                var groups = divs
                    .GroupAdjacent(b => b.Div);
                tempSourceList = groups
                    .Select(g => new TempSource
                    {
                        Start = g.First().Index,
                        Count = g.Count(),
                    })
                    .ToList();
                foreach (var ts in tempSourceList)
                {
                    List<Source> sources = new List<Source>()
                    {
                        new Source(doc, ts.Start, ts.Count, true)
                    };
                    WmlDocument newDoc = DocumentBuilder.BuildDocument(sources);
                    newDoc = AdjustSectionBreak(newDoc);
                    yield return newDoc;
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
                    XElement lastElement = mainXDoc.Root
                        .Element(W.body)
                        .Elements()
                        .LastOrDefault();
                    if (lastElement != null)
                    {
                        if (lastElement.Name != W.sectPr &&
                            lastElement.Descendants(W.sectPr).Any())
                        {
                            mainXDoc.Root.Element(W.body).Add(lastElement.Descendants(W.sectPr).First());
                            lastElement.Descendants(W.sectPr).Remove();
                            if (!lastElement.Elements()
                                .Where(e => e.Name != W.pPr)
                                .Any())
                                lastElement.Remove();
                            document.MainDocumentPart.PutXDocument();
                        }
                    }
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }
        public static WmlDocument ExtractGlossaryDocument(this WmlDocument wmlGlossaryDocument)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlGlossaryDocument.DocumentByteArray, 0, wmlGlossaryDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    if (wDoc.MainDocumentPart.GlossaryDocumentPart == null)
                        return null;

                    var fromXd = wDoc.MainDocumentPart.GlossaryDocumentPart.GetXDocument();
                    if (fromXd.Root == null)
                        return null;

                    using (MemoryStream outMs = new MemoryStream())
                    {
                        using (WordprocessingDocument outWDoc = WordprocessingDocument.Create(outMs, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                        {
                            List<ImageData> images = new List<ImageData>();

                            MainDocumentPart mdp = outWDoc.AddMainDocumentPart();
                            var mdpXd = mdp.GetXDocument();
                            XElement root = new XElement(W.document);
                            if (mdpXd.Root == null)
                                mdpXd.Add(root);
                            else
                                mdpXd.Root.ReplaceWith(root);
                            root.Add(new XElement(W.body,
                                fromXd.Root.Elements(W.docParts)));
                            mdp.PutXDocument();

                            var newContent = fromXd.Root.Elements(W.docParts);
                            wDoc.CopyGlossaryDocumentPartsFromGD(outWDoc, newContent, images);
                            wDoc.MainDocumentPart.GlossaryDocumentPart.CopyRelatedPartsForContentParts(mdp, RelationshipMarkup, newContent, images);
                        }
                        return new WmlDocument("Glossary.docx", outMs.ToArray());
                    }
                }
            }
        }
        public static void WriteGlossaryDocumentPart(this WmlDocument wmlGlossaryDocument, WordprocessingDocument output, List<ImageData> images)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlGlossaryDocument.DocumentByteArray, 0, wmlGlossaryDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var fromXDoc = wDoc.MainDocumentPart.GetXDocument();
                    var newXDoc = CreateGlossary(fromXDoc);
                    var outputGlossaryDocumentPart = output.MainDocumentPart.AddNewPart<GlossaryDocumentPart>();
                    outputGlossaryDocumentPart.PutXDocument(newXDoc);

                    CopyGlossaryDocumentPartsToGD(wDoc, output, fromXDoc.Root.Descendants(W.docPart), images);
                    wDoc.MainDocumentPart.CopyRelatedPartsForContentParts(outputGlossaryDocumentPart, RelationshipMarkup, new[] { fromXDoc.Root }, images);
                }
            }

            XDocument CreateGlossary(XDocument xDoc)
            {
                return new XDocument(
                        Commons.Common.CreateDeclaration(),
                        new XElement(W.glossaryDocument,
                            Commons.Constants.NamespaceAttributes,
                            new XElement(
                                W.docParts,
                                xDoc.Descendants(W.docPart))));
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
        public static void AddSectionAndDependencies(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, XElement sectionMarkup, List<ImageData> images)
        {
            var headerReferences = sectionMarkup.Elements(W.headerReference);
            foreach (var headerReference in headerReferences)
            {
                string oldRid = headerReference.Attribute(R.id).Value;
                HeaderPart oldHeaderPart = null;
                try
                {
                    oldHeaderPart = (HeaderPart)sourceDocument.MainDocumentPart.GetPartById(oldRid);
                }
                catch (ArgumentOutOfRangeException)
                {
                    var message = string.Format("ArgumentOutOfRangeException, attempting to get header rId={0}", oldRid);
                    throw new OpenXmlPowerToolsException(message);
                }
                XDocument oldHeaderXDoc = oldHeaderPart.GetXDocument();
                if (oldHeaderXDoc != null && oldHeaderXDoc.Root != null)
                    CopyNumbering(sourceDocument, newDocument, new[] { oldHeaderXDoc.Root }, images);
                HeaderPart newHeaderPart = newDocument.MainDocumentPart.AddNewPart<HeaderPart>();
                XDocument newHeaderXDoc = newHeaderPart.GetXDocument();
                newHeaderXDoc.Declaration.SetDeclaration();
                newHeaderXDoc.Add(oldHeaderXDoc.Root);
                headerReference.Attribute(R.id).Value = newDocument.MainDocumentPart.GetIdOfPart(newHeaderPart);
                oldHeaderPart.AddRelationships(newHeaderPart, RelationshipMarkup, new[] { newHeaderXDoc.Root });
                oldHeaderPart.CopyRelatedPartsForContentParts(newHeaderPart, RelationshipMarkup, new[] { newHeaderXDoc.Root }, images);
            }

            var footerReferences = sectionMarkup.Elements(W.footerReference);
            foreach (var footerReference in footerReferences)
            {
                string oldRid = footerReference.Attribute(R.id).Value;
                var oldFooterPart2 = sourceDocument.MainDocumentPart.GetPartById(oldRid);
                if (!(oldFooterPart2 is FooterPart))
                    throw new DocumentBuilderException("Invalid document - invalid footer part.");

                FooterPart oldFooterPart = (FooterPart)oldFooterPart2;
                XDocument oldFooterXDoc = oldFooterPart.GetXDocument();
                if (oldFooterXDoc != null && oldFooterXDoc.Root != null)
                    CopyNumbering(sourceDocument, newDocument, new[] { oldFooterXDoc.Root }, images);
                FooterPart newFooterPart = newDocument.MainDocumentPart.AddNewPart<FooterPart>();
                XDocument newFooterXDoc = newFooterPart.GetXDocument();
                newFooterXDoc.Declaration.SetDeclaration();
                newFooterXDoc.Add(oldFooterXDoc.Root);
                footerReference.Attribute(R.id).Value = newDocument.MainDocumentPart.GetIdOfPart(newFooterPart);
                oldFooterPart.AddRelationships(newFooterPart, RelationshipMarkup, new[] { newFooterXDoc.Root });
                oldFooterPart.CopyRelatedPartsForContentParts(newFooterPart, RelationshipMarkup, new[] { newFooterXDoc.Root }, images);
            }
        }
        public static void AdjustUniqueIds(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent)
        {
            // adjust bookmark unique ids
            int maxId = 0;
            if (newDocument.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart).Any())
                maxId = newDocument.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart)
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
        public static void AdjustDocPrIds(this WordprocessingDocument newDocument)
        {
            int docPrId = 0;
            foreach (var item in newDocument.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
                item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            foreach (var header in newDocument.MainDocumentPart.HeaderParts)
                foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            foreach (var footer in newDocument.MainDocumentPart.FooterParts)
                foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            if (newDocument.MainDocumentPart.FootnotesPart != null)
                foreach (var item in newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            if (newDocument.MainDocumentPart.EndnotesPart != null)
                foreach (var item in newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
        }
        public static void AdjustStyleIdsForDocument(this WordprocessingDocument wDoc, Dictionary<string, string> correctionList)
        {
            // update styles part
            wDoc.MainDocumentPart.StyleDefinitionsPart.UpdateStyleIdsForStylePart(correctionList);
            if (wDoc.MainDocumentPart.StylesWithEffectsPart != null)
                wDoc.MainDocumentPart.StylesWithEffectsPart.UpdateStyleIdsForStylePart(correctionList);

            // update content parts
            wDoc.MainDocumentPart.UpdateStyleIdsForContentPart(correctionList);
            foreach (var part in wDoc.MainDocumentPart.HeaderParts)
                part.UpdateStyleIdsForContentPart(correctionList);
            foreach (var part in wDoc.MainDocumentPart.FooterParts)
                part.UpdateStyleIdsForContentPart(correctionList);
            if (wDoc.MainDocumentPart.FootnotesPart != null)
                wDoc.MainDocumentPart.FootnotesPart.UpdateStyleIdsForContentPart(correctionList);
            if (wDoc.MainDocumentPart.EndnotesPart != null)
                wDoc.MainDocumentPart.EndnotesPart.UpdateStyleIdsForContentPart(correctionList);
            if (wDoc.MainDocumentPart.WordprocessingCommentsPart != null)
                wDoc.MainDocumentPart.WordprocessingCommentsPart.UpdateStyleIdsForContentPart(correctionList);
            if (wDoc.MainDocumentPart.WordprocessingCommentsExPart != null)
                wDoc.MainDocumentPart.WordprocessingCommentsExPart.UpdateStyleIdsForContentPart(correctionList);

            // update numbering part
            wDoc.MainDocumentPart.NumberingDefinitionsPart.UpdateStyleIdsForContentPart(correctionList);
        }
        public static void AppendDocument(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, List<XElement> newContent, bool keepSection, string insertId, List<ImageData> images)
        {
            // Rules for sections
            // - if KeepSections for all documents in the source collection are false, then it takes the section
            //   from the first document.
            // - if you specify true for any document, and if the last section is part of the specified content,
            //   then that section is copied.  If any paragraph in the content has a section, then that section
            //   is copied.
            // - if you specify true for any document, and there are no sections for any paragraphs, then no
            //   sections are copied.

            sourceDocument.MainDocumentPart.GetXDocument().FixRanges(newContent);
            sourceDocument.MainDocumentPart.AddRelationships(newDocument.MainDocumentPart, RelationshipMarkup, newContent);
            sourceDocument.MainDocumentPart.CopyRelatedPartsForContentParts(newDocument.MainDocumentPart, RelationshipMarkup, newContent, images);

            // Append contents
            XDocument newMainXDoc = newDocument.MainDocumentPart.GetXDocument();
            newMainXDoc.Declaration.SetDeclaration();
            if (keepSection == false)
            {
                List<XElement> adjustedContents = newContent.Where(e => e.Name != W.sectPr).ToList();
                adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
                newContent = adjustedContents;
            }
            var listOfSectionProps = newContent.DescendantsAndSelf(W.sectPr).ToList();
            foreach (var sectPr in listOfSectionProps)
                AddSectionAndDependencies(sourceDocument, newDocument, sectPr, images);
            CopyStylesAndFonts(sourceDocument, newDocument, newContent);
            CopyNumbering(sourceDocument, newDocument, newContent, images);
            CopyComments(sourceDocument, newDocument, newContent, images);
            CopyFootnotes(sourceDocument, newDocument, newContent, images);
            CopyEndnotes(sourceDocument, newDocument, newContent, images);
            AdjustUniqueIds(sourceDocument, newDocument, newContent);
            newContent.RemoveGfxdata();
            CopyCustomXmlPartsForDataBoundContentControls(sourceDocument, newDocument, newContent);
            CopyWebExtensions(sourceDocument, newDocument);
            if (insertId != null)
            {
                XElement insertElementToReplace = newMainXDoc
                    .Descendants(PtOpenXml.Insert)
                    .FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == insertId);
                if (insertElementToReplace != null)
                    insertElementToReplace.AddAnnotation(new ReplaceSemaphore());
                newMainXDoc.Element(W.document).ReplaceWith((XElement)newMainXDoc.Root.InsertTransform(newContent));
            }
            else
                newMainXDoc.Root.Element(W.body).Add(newContent);

            if (newMainXDoc.Descendants().Any(d =>
            {
                if (d.Name.Namespace == PtOpenXml.pt || d.Name.Namespace == PtOpenXml.ptOpenXml)
                    return true;
                if (d.Attributes().Any(att => att.Name.Namespace == PtOpenXml.pt || att.Name.Namespace == PtOpenXml.ptOpenXml))
                    return true;
                return false;
            }))
            {
                var root = newMainXDoc.Root;
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
        public static void AppendDocument(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, OpenXmlPart part, List<XElement> newContent, bool keepSection, string insertId, List<ImageData> images)
        {
            // Append contents
            XDocument partXDoc = part.GetXDocument();
            partXDoc.Declaration.SetDeclaration();

            part.GetXDocument().FixRanges(newContent);
            sourceDocument.MainDocumentPart.AddRelationships(part, RelationshipMarkup, newContent);
            sourceDocument.MainDocumentPart.CopyRelatedPartsForContentParts(part, RelationshipMarkup, newContent, images);

            // never keep sections for content to be inserted into a header/footer
            List<XElement> adjustedContents = newContent.Where(e => e.Name != W.sectPr).ToList();
            adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
            newContent = adjustedContents;

            CopyNumbering(sourceDocument, newDocument, newContent, images);
            CopyComments(sourceDocument, newDocument, newContent, images);
            AdjustUniqueIds(sourceDocument, newDocument, newContent);
            newContent.RemoveGfxdata();

            if (insertId == null)
                throw new OpenXmlPowerToolsException("Internal error");

            XElement insertElementToReplace = partXDoc
                .Descendants(PtOpenXml.Insert)
                .FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == insertId);
            if (insertElementToReplace != null)
                insertElementToReplace.AddAnnotation(new ReplaceSemaphore());
            partXDoc.Elements().First().ReplaceWith((XElement)partXDoc.Root.InsertTransform(newContent));
        }
        public static void CopyComments(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> commentIdMap = new Dictionary<int, int>();
            int number = 0;
            XDocument oldComments = null;
            XDocument newComments = null;
            foreach (XElement comment in newContent.DescendantsAndSelf(W.commentReference))
            {
                if (oldComments == null)
                    oldComments = sourceDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                if (newComments == null)
                {
                    if (newDocument.MainDocumentPart.WordprocessingCommentsPart != null)
                    {
                        newComments = newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.SetDeclaration();
                        var ids = newComments.Root.Elements(W.comment).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                        newComments = newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
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
            foreach (var item in newContent.DescendantsAndSelf()
                .Where(d => d.Name == W.commentReference ||
                            d.Name == W.commentRangeStart ||
                            d.Name == W.commentRangeEnd)
                .ToList())
            {
                if (commentIdMap.ContainsKey((int)item.Attribute(W.id)))
                    item.Attribute(W.id).Value = commentIdMap[(int)item.Attribute(W.id)].ToString();
            }
            if (sourceDocument.MainDocumentPart.WordprocessingCommentsPart != null &&
                newDocument.MainDocumentPart.WordprocessingCommentsPart != null)
            {
                sourceDocument.MainDocumentPart.WordprocessingCommentsPart.AddRelationships(newDocument.MainDocumentPart.WordprocessingCommentsPart, 
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root });
                sourceDocument.MainDocumentPart.WordprocessingCommentsPart.CopyRelatedPartsForContentParts(newDocument.MainDocumentPart.WordprocessingCommentsPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root },images);
            }
        }
        public static void CopyCustomXmlPartsForDataBoundContentControls(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent)
        {
            List<string> itemList = new List<string>();
            foreach (string itemId in newContent
                .Descendants(W.dataBinding)
                .Select(e => (string)e.Attribute(W.storeItemID)))
                if (!itemList.Contains(itemId))
                    itemList.Add(itemId);
            foreach (CustomXmlPart customXmlPart in sourceDocument.MainDocumentPart.CustomXmlParts)
            {
                OpenXmlPart propertyPart = customXmlPart
                    .Parts
                    .Select(p => p.OpenXmlPart)
                    .Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                    .FirstOrDefault();
                if (propertyPart != null)
                {
                    XDocument propertyPartDoc = propertyPart.GetXDocument();
                    if (itemList.Contains(propertyPartDoc.Root.Attribute(DS.itemID).Value))
                    {
                        CustomXmlPart newPart = newDocument.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
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
        public static void CopyEndnotes(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            int number = 0;
            XDocument oldEndnotes = null;
            XDocument newEndnotes = null;
            foreach (XElement endnote in newContent.DescendantsAndSelf(W.endnoteReference))
            {
                if (oldEndnotes == null)
                    oldEndnotes = sourceDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                if (newEndnotes == null)
                {
                    if (newDocument.MainDocumentPart.EndnotesPart != null)
                    {
                        newEndnotes = newDocument
                            .MainDocumentPart
                            .EndnotesPart
                            .GetXDocument();
                        var ids = newEndnotes
                            .Root
                            .Elements(W.endnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<EndnotesPart>();
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.SetDeclaration();
                        newEndnotes.Add(new XElement(W.endnotes, Constants.NamespaceAttributes));
                    }
                }
                string id = (string)endnote.Attribute(W.id);
                XElement element = oldEndnotes
                    .Descendants()
                    .Elements(W.endnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .First();
                XElement newElement = new XElement(element);
                newElement.Attribute(W.id).Value = number.ToString();
                newEndnotes.Root.Add(newElement);
                endnote.Attribute(W.id).Value = number.ToString();
                number++;
            }
            if (sourceDocument.MainDocumentPart.EndnotesPart != null &&
                newDocument.MainDocumentPart.EndnotesPart != null)
            {
                sourceDocument.MainDocumentPart.EndnotesPart.AddRelationships(newDocument.MainDocumentPart.EndnotesPart, 
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Root });
                sourceDocument.MainDocumentPart.EndnotesPart.CopyRelatedPartsForContentParts(newDocument.MainDocumentPart.EndnotesPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Root }, images);
            }
        }
        public static void CopyEndnotesPart(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, XDocument settingsXDoc, List<ImageData> images)
        {
            int number = 0;
            XDocument oldEndnotes = null;
            XDocument newEndnotes = null;
            XElement endnotePr = settingsXDoc.Root.Element(W.endnotePr);
            if (endnotePr == null)
                return;
            if (sourceDocument.MainDocumentPart.EndnotesPart == null)
                return;
            foreach (XElement endnote in endnotePr.Elements(W.endnote))
            {
                if (oldEndnotes == null)
                    oldEndnotes = sourceDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                if (newEndnotes == null)
                {
                    if (newDocument.MainDocumentPart.EndnotesPart != null)
                    {
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.SetDeclaration();
                        var ids = newEndnotes.Root
                            .Elements(W.endnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<EndnotesPart>();
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
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
        public static void CopyFootnotes(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            int number = 0;
            XDocument oldFootnotes = null;
            XDocument newFootnotes = null;
            foreach (XElement footnote in newContent.DescendantsAndSelf(W.footnoteReference))
            {
                if (oldFootnotes == null)
                    oldFootnotes = sourceDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                if (newFootnotes == null)
                {
                    if (newDocument.MainDocumentPart.FootnotesPart != null)
                    {
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        var ids = newFootnotes
                            .Root
                            .Elements(W.footnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<FootnotesPart>();
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.SetDeclaration();
                        newFootnotes.Add(new XElement(W.footnotes, Constants.NamespaceAttributes));
                    }
                }
                string id = (string)footnote.Attribute(W.id);
                XElement element = oldFootnotes
                    .Descendants()
                    .Elements(W.footnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    XElement newElement = new XElement(element);
                    newElement.Attribute(W.id).Value = number.ToString();
                    newFootnotes.Root.Add(newElement);
                    footnote.Attribute(W.id).Value = number.ToString();
                    number++;
                }
            }
            if (sourceDocument.MainDocumentPart.FootnotesPart != null &&
                newDocument.MainDocumentPart.FootnotesPart != null)
            {
                sourceDocument.MainDocumentPart.FootnotesPart.AddRelationships(newDocument.MainDocumentPart.FootnotesPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Root });
               sourceDocument.MainDocumentPart.FootnotesPart.CopyRelatedPartsForContentParts(newDocument.MainDocumentPart.FootnotesPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Root }, images);
            }
        }
        public static void CopyFootnotesPart(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, XDocument settingsXDoc, List<ImageData> images)
        {
            int number = 0;
            XDocument oldFootnotes = null;
            XDocument newFootnotes = null;
            XElement footnotePr = settingsXDoc.Root.Element(W.footnotePr);
            if (footnotePr == null)
                return;
            if (sourceDocument.MainDocumentPart.FootnotesPart == null)
                return;
            foreach (XElement footnote in footnotePr.Elements(W.footnote))
            {
                if (oldFootnotes == null)
                    oldFootnotes = sourceDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                if (newFootnotes == null)
                {
                    if (newDocument.MainDocumentPart.FootnotesPart != null)
                    {
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.SetDeclaration();
                        var ids = newFootnotes.Root.Elements(W.footnote).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<FootnotesPart>();
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
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
        public static void CopyGlossaryDocumentPartsFromGD(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StyleDefinitionsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newStyles.Declaration.SetDeclaration();
                    newStyles.Add(oldStyles.Root);
                    newDocument.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                }
                else
                {
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    newDocument.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                }
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart != null)
            {
                XDocument oldFontTable = sourceDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                if (newDocument.MainDocumentPart.FontTablePart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    newFontTable.Declaration.SetDeclaration();
                    newFontTable.Add(oldFontTable.Root);
                    newDocument.MainDocumentPart.FontTablePart.PutXDocument();
                }
                else
                {
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    oldFontTable.MergeFontTables(newFontTable);
                    newDocument.MainDocumentPart.FontTablePart.PutXDocument();
                }
            }

            DocumentSettingsPart oldSettingsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                DocumentSettingsPart newSettingsPart = newDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                oldSettingsPart.AddRelationships(newSettingsPart, RelationshipMarkup, new[] { settingsXDoc.Root });
                //CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                //CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                XDocument newXDoc = newDocument.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(settingsXDoc.Root);
                oldSettingsPart.CopyRelatedPartsForContentParts(newSettingsPart, RelationshipMarkup, new[] { newXDoc.Root }, images);
                newSettingsPart.PutXDocument(newXDoc);
            }

            WebSettingsPart oldWebSettingsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                WebSettingsPart newWebSettingsPart = newDocument.MainDocumentPart.AddNewPart<WebSettingsPart>();
                XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                oldWebSettingsPart.AddRelationships(newWebSettingsPart, RelationshipMarkup, new[] { settingsXDoc.Root });
                XDocument newXDoc = newDocument.MainDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(settingsXDoc.Root);
                newWebSettingsPart.PutXDocument(newXDoc);
            }

            NumberingDefinitionsPart oldNumberingDefinitionsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart;
            if (oldNumberingDefinitionsPart != null)
            {
                CopyNumberingForGlossaryDocumentPartFromGD(oldNumberingDefinitionsPart, newDocument, newContent, images);
            }
        }
        public static void CopyGlossaryDocumentPartsToGD(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.StyleDefinitionsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<StyleDefinitionsPart>();
                XDocument newStyles = newDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                newStyles.Declaration.SetDeclaration();
                newStyles.Add(oldStyles.Root);
                newDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.PutXDocument();
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.FontTablePart != null)
            {
                XDocument oldFontTable = sourceDocument.MainDocumentPart.FontTablePart.GetXDocument();
                newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<FontTablePart>();
                XDocument newFontTable = newDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                newFontTable.Declaration.SetDeclaration();
                newFontTable.Add(oldFontTable.Root);
                newDocument.MainDocumentPart.FontTablePart.PutXDocument();
            }

            DocumentSettingsPart oldSettingsPart = sourceDocument.MainDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                DocumentSettingsPart newSettingsPart = newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<DocumentSettingsPart>();
                XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                oldSettingsPart.AddRelationships(newSettingsPart, RelationshipMarkup, new[] { settingsXDoc.Root });
                //CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                //CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                XDocument newXDoc = newDocument.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(settingsXDoc.Root);
                oldSettingsPart.CopyRelatedPartsForContentParts(newSettingsPart, RelationshipMarkup, new[] { newXDoc.Root }, images);
                newSettingsPart.PutXDocument(newXDoc);
            }

            WebSettingsPart oldWebSettingsPart = sourceDocument.MainDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                WebSettingsPart newWebSettingsPart = newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<WebSettingsPart>();
                XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                oldWebSettingsPart.AddRelationships(newWebSettingsPart, RelationshipMarkup, new[] { settingsXDoc.Root });
                XDocument newXDoc = newDocument.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(settingsXDoc.Root);
                newWebSettingsPart.PutXDocument(newXDoc);
            }

            NumberingDefinitionsPart oldNumberingDefinitionsPart = sourceDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (oldNumberingDefinitionsPart != null)
            {
                CopyNumberingForGlossaryDocumentPartToGD(oldNumberingDefinitionsPart, newDocument, newContent, images);
            }
        }
        public static void CopyNumbering(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> numIdMap = new Dictionary<int, int>();
            int number = 1;
            int abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (XElement numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    if (oldNumbering == null)
                        oldNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
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
                            newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
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

            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null &&
                sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
            {
                sourceDocument.MainDocumentPart.NumberingDefinitionsPart.AddRelationships(newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                sourceDocument.MainDocumentPart.NumberingDefinitionsPart.CopyRelatedPartsForContentParts(newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
        }
        public static void CopyOrCacheHeaderOrFooter(this WordprocessingDocument doc, CachedHeaderFooter[] cachedHeaderFooter, XElement sect, XName referenceXName, string type)
        {
            var referenceElement = sect.FindReference(referenceXName, type);
            if (referenceElement == null)
            {
                var cachedPartRid = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type).CachedPartRid;
                doc.MainDocumentPart.AddReferenceToExistingHeaderOrFooter(sect, cachedPartRid, referenceXName, type);
            }
            else
            {
                var cachedPart = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type);
                cachedPart.CachedPartRid = (string)referenceElement.Attribute(R.id);
            }
        }
        public static void CopySpecifiedCustomXmlParts(this WordprocessingDocument sourceDocument, WordprocessingDocument output)
        {
            //if (settings.CustomXmlGuidList == null || !settings.CustomXmlGuidList.Any())
            //    return;

            foreach (CustomXmlPart customXmlPart in sourceDocument.MainDocumentPart.CustomXmlParts)
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
                            CustomXmlPart newPart = output.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
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
        public static void CopyStartingParts(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, List<ImageData> images)
        {
            // A Core File Properties part does not have implicit or explicit relationships to other parts.
            CoreFilePropertiesPart corePart = sourceDocument.CoreFilePropertiesPart;
            if (corePart != null && corePart.GetXDocument().Root != null)
            {
                newDocument.AddCoreFilePropertiesPart();
                XDocument newXDoc = newDocument.CoreFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                XDocument sourceXDoc = corePart.GetXDocument();
                newXDoc.Add(sourceXDoc.Root);
            }

            // An application attributes part does not have implicit or explicit relationships to other parts.
            ExtendedFilePropertiesPart extPart = sourceDocument.ExtendedFilePropertiesPart;
            if (extPart != null)
            {
                OpenXmlPart newPart = newDocument.AddExtendedFilePropertiesPart();
                XDocument newXDoc = newDocument.ExtendedFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(extPart.GetXDocument().Root);
            }

            // An custom file properties part does not have implicit or explicit relationships to other parts.
            CustomFilePropertiesPart customPart = sourceDocument.CustomFilePropertiesPart;
            if (customPart != null)
            {
                newDocument.AddCustomFilePropertiesPart();
                XDocument newXDoc = newDocument.CustomFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(customPart.GetXDocument().Root);
            }

            DocumentSettingsPart oldSettingsPart = sourceDocument.MainDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                DocumentSettingsPart newSettingsPart = newDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                oldSettingsPart.AddRelationships(newSettingsPart, RelationshipMarkup, new[] { settingsXDoc.Root });
                CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                XDocument newXDoc = newDocument.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(settingsXDoc.Root);
                oldSettingsPart.CopyRelatedPartsForContentParts(newSettingsPart, RelationshipMarkup, new[] { newXDoc.Root }, images);
            }

            WebSettingsPart oldWebSettingsPart = sourceDocument.MainDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                WebSettingsPart newWebSettingsPart = newDocument.MainDocumentPart.AddNewPart<WebSettingsPart>();
                XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                oldWebSettingsPart.AddRelationships(newWebSettingsPart, RelationshipMarkup, new[] { settingsXDoc.Root });
                XDocument newXDoc = newDocument.MainDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(settingsXDoc.Root);
            }

            ThemePart themePart = sourceDocument.MainDocumentPart.ThemePart;
            if (themePart != null)
            {
                ThemePart newThemePart = newDocument.MainDocumentPart.AddNewPart<ThemePart>();
                XDocument newXDoc = newDocument.MainDocumentPart.ThemePart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(themePart.GetXDocument().Root);
                themePart.CopyRelatedPartsForContentParts(newThemePart, RelationshipMarkup, new[] { newThemePart.GetXDocument().Root }, images);
            }

            // If needed to handle GlossaryDocumentPart in the future, then
            // this code should handle the following parts:
            //   MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart
            //   MainDocumentPart.GlossaryDocumentPart.StylesWithEffectsPart

            // A Style Definitions part shall not have implicit or explicit relationships to any other part.
            StyleDefinitionsPart stylesPart = sourceDocument.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart != null)
            {
                newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                XDocument newXDoc = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                newXDoc.Add(new XElement(W.styles,
                    new XAttribute(XNamespace.Xmlns + "w", W.w)

                    //,
                    //stylesPart.GetXDocument().Descendants(W.docDefaults)

                    //,
                    //new XElement(W.latentStyles, stylesPart.GetXDocument().Descendants(W.latentStyles).Attributes())

                    ));
                stylesPart.GetXDocument().MergeDocDefaultStyles(newXDoc);
                MergeStyles(sourceDocument, newDocument, stylesPart.GetXDocument(), newXDoc, Enumerable.Empty<XElement>());
                stylesPart.GetXDocument().MergeLatentStyles(newXDoc);
            }

            // A Font Table part shall not have any implicit or explicit relationships to any other part.
            FontTablePart fontTablePart = sourceDocument.MainDocumentPart.FontTablePart;
            if (fontTablePart != null)
            {
                newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                XDocument newXDoc = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                newXDoc.Declaration.SetDeclaration();
                sourceDocument.MainDocumentPart.FontTablePart.CopyFontTable(newDocument.MainDocumentPart.FontTablePart);
                newXDoc.Add(fontTablePart.GetXDocument().Root);
            }
        }
        public static void CopyStylesAndFonts(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.StyleDefinitionsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StyleDefinitionsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newStyles.Declaration.SetDeclaration();
                    newStyles.Add(oldStyles.Root);
                }
                else
                {
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    oldStyles.MergeLatentStyles(newStyles);
                }
            }

            // Copy all styles with effects to the new document
            if (sourceDocument.MainDocumentPart.StylesWithEffectsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StylesWithEffectsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StylesWithEffectsPart>();
                    XDocument newStyles = newDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                    newStyles.Declaration.SetDeclaration();
                    newStyles.Add(oldStyles.Root);
                }
                else
                {
                    XDocument newStyles = newDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    oldStyles.MergeLatentStyles(newStyles);
                }
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.FontTablePart != null)
            {
                XDocument oldFontTable = sourceDocument.MainDocumentPart.FontTablePart.GetXDocument();
                if (newDocument.MainDocumentPart.FontTablePart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    newFontTable.Declaration.SetDeclaration();
                    newFontTable.Add(oldFontTable.Root);
                }
                else
                {
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    oldFontTable.MergeFontTables(newFontTable);
                }
            }
        }
        public static void CopyWebExtensions(this WordprocessingDocument sourceDocument, WordprocessingDocument newDocument)
        {
            if (sourceDocument.WebExTaskpanesPart != null && newDocument.WebExTaskpanesPart == null)
            {
                newDocument.AddWebExTaskpanesPart();
                newDocument.WebExTaskpanesPart.GetXDocument().Add(sourceDocument.WebExTaskpanesPart.GetXDocument().Root);

                foreach (var sourceWebExtensionPart in sourceDocument.WebExTaskpanesPart.WebExtensionParts)
                {
                    var newWebExtensionpart = newDocument.WebExTaskpanesPart.AddNewPart<WebExtensionPart>(
                        sourceDocument.WebExTaskpanesPart.GetIdOfPart(sourceWebExtensionPart));
                    newWebExtensionpart.GetXDocument().Add(sourceWebExtensionPart.GetXDocument().Root);
                }
            }
        }
        // New method to support new functionality

        // swap to wordproccessingdocument extension
        public static void CopyNumberingForGlossaryDocumentPartFromGD(this NumberingDefinitionsPart sourceNumberingPart, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> numIdMap = new Dictionary<int, int>();
            int number = 1;
            int abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (XElement numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    if (oldNumbering == null)
                        oldNumbering = sourceNumberingPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
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
                            newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
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

            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null &&
                sourceNumberingPart != null)
            {
                sourceNumberingPart.AddRelationships(newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                sourceNumberingPart.CopyRelatedPartsForContentParts(newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                newDocument.MainDocumentPart.NumberingDefinitionsPart.PutXDocument();
        }
        // swap to wordproccessingdocument extension
        public static void CopyNumberingForGlossaryDocumentPartToGD(this NumberingDefinitionsPart sourceNumberingPart, WordprocessingDocument newDocument, IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> numIdMap = new Dictionary<int, int>();
            int number = 1;
            int abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (XElement numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    if (oldNumbering == null)
                        oldNumbering = sourceNumberingPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
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
                            newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
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

            if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null &&
                sourceNumberingPart != null)
            {
                sourceNumberingPart.AddRelationships(newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                sourceNumberingPart.CopyRelatedPartsForContentParts(newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                    RelationshipMarkup, new[] { newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
            if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.PutXDocument();
        }
        public static void FixUpSectionProperties(this WordprocessingDocument newDocument)
        {
            XDocument mainDocumentXDoc = newDocument.MainDocumentPart.GetXDocument();
            mainDocumentXDoc.Declaration.SetDeclaration();
            XElement body = mainDocumentXDoc.Root.Element(W.body);
            var sectionPropertiesToMove = body
                .Elements()
                .Take(body.Elements().Count() - 1)
                .Where(e => e.Name == W.sectPr)
                .ToList();
            foreach (var s in sectionPropertiesToMove)
            {
                var p = s.SiblingsBeforeSelfReverseDocumentOrder().First();
                if (p.Element(W.pPr) == null)
                    p.AddFirst(new XElement(W.pPr));
                p.Element(W.pPr).Add(s);
            }
            foreach (var s in sectionPropertiesToMove)
                s.Remove();
        }
        public static Dictionary<string, string> GetStyleNameMap(this WordprocessingDocument wDoc)
        {
            var sxDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var thisDocumentDictionary = sxDoc
                .Root
                .Elements(W.style)
                .ToDictionary(
                    z => (string)z.Elements(W.name).Attributes(W.val).FirstOrDefault(),
                    z => (string)z.Attribute(W.styleId));
            return thisDocumentDictionary;
        }
        public static void MergeStyles(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, XDocument fromStyles, XDocument toStyles, IEnumerable<XElement> newContent)
        {
            //var newIds = new Dictionary<string, string>();

            if (fromStyles.Root == null)
                return;

            foreach (XElement style in fromStyles.Root.Elements(W.style))
            {
                var fromId = (string)style.Attribute(W.styleId);
                var fromName = (string)style.Elements(W.name).Attributes(W.val).FirstOrDefault();

                var toStyle = toStyles
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
                                if (sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                                    oldNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                else
                                {
                                    oldNumbering = new XDocument();
                                    oldNumbering.Declaration = Commons.Common.CreateDeclaration();
                                    oldNumbering.Add(new XElement(W.numbering, Constants.NamespaceAttributes));
                                }
                            }
                            if (newNumbering == null)
                            {
                                if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                                {
                                    newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
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
                                    newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                                    newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
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
                    toStyles.Root.Add(newStyle);
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

                // Convert source document, since numberings will be copied over after styles.
                if (sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    var sourceNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    ConvertNumberingPartToNewIds(sourceNumbering, newIds);
                }
            }
#endif
        }
        public static void ProcessSectionsForLinkToPreviousHeadersAndFooters(this WordprocessingDocument doc)
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
                            doc.MainDocumentPart.AddReferenceToExistingHeaderOrFooter(sect, (string)headerDefault.Attribute(R.id), W.headerReference, "even");
                        else
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, W.headerReference, "even");
                    }

                    if (headerFirst == null)
                    {
                        if (headerDefault != null)
                            doc.MainDocumentPart.AddReferenceToExistingHeaderOrFooter(sect, (string)headerDefault.Attribute(R.id), W.headerReference, "first");
                        else
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, W.headerReference, "first");
                    }

                    if (footerEven == null)
                    {
                        if (footerDefault != null)
                            doc.MainDocumentPart.AddReferenceToExistingHeaderOrFooter(sect, (string)footerDefault.Attribute(R.id), W.footerReference, "even");
                        else
                            doc.MainDocumentPart.InitEmptyHeaderOrFooter(sect, W.footerReference, "even");
                    }

                    if (footerFirst == null)
                    {
                        if (footerDefault != null)
                            doc.MainDocumentPart.AddReferenceToExistingHeaderOrFooter(sect, (string)footerDefault.Attribute(R.id), W.footerReference, "first");
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
        public static void TestForUnsupportedDocument(this WordprocessingDocument doc, int sourceNumber)
        {
            //What does not work:
            //- sub docs
            //- bidi text appears to work but has not been tested
            //- languages other than en-us appear to work but have not been tested
            //- documents with activex controls
            //- mail merge source documents (look for dataSource in settings)
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

        #endregion
    }
}
