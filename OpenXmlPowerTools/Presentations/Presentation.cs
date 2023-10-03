using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Spreadsheets;
using System.Text.RegularExpressions;
using System.Text;

namespace OpenXmlPowerTools.Presentations
{
    public static class Presentation
    {
        public static Dictionary<XName, XName[]> RelationshipMarkup => Constants.PresentationRelationshipMarkup;
        public static Dictionary<XName, int> PresentationOrder => Constants.PresentationOrder;
        public static string[] Extensions = new[] {
            ".pptx",
            ".potx",
            ".ppsx",
            ".pptm",
            ".potm",
            ".ppsm",
            ".ppam",
        };
        public static bool IsPresentation(string ext) => Extensions.Contains(ext.ToLower());
        public static OpenXmlMemoryStreamDocument CreatePresentationDocument()
        {
            MemoryStream stream = new MemoryStream();
            using (PresentationDocument doc = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
            {
                doc.AddPresentationPart();
                XNamespace ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
                XNamespace relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                XNamespace drawingns = "http://schemas.openxmlformats.org/drawingml/2006/main";
                doc.PresentationPart.PutXDocument(new XDocument(
                    new XElement(ns + "presentation",
                        new XAttribute(XNamespace.Xmlns + "a", drawingns),
                        new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                        new XAttribute(XNamespace.Xmlns + "p", ns),
                        new XElement(ns + "sldMasterIdLst"),
                        new XElement(ns + "sldIdLst"),
                        new XElement(ns + "notesSz", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000")))));
                doc.Close();
                return new OpenXmlMemoryStreamDocument(stream);
            }
        }

        #region PresentationDocument
        // Copy handout master, notes master, presentation properties and view properties, if they exist
        public static void CopyPresentationParts(this PresentationDocument sourceDocument, PresentationDocument newDocument, List<ImageData> images, List<MediaData> mediaList)
        {
            // TODO need to handle the following
            //{ P.custShowLst, 80 },
            //    { P.photoAlbum, 90 },
            //    { P.custDataLst, 100 },
            //    { P.kinsoku, 120 },
            //    { P.modifyVerifier, 150 },

            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();

            // Copy slide and note slide sizes
            XDocument oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();

            foreach (var att in oldPresentationDoc.Root.Attributes())
            {
                if (!att.IsNamespaceDeclaration && newPresentation.Root.Attribute(att.Name) == null)
                    newPresentation.Root.Add(oldPresentationDoc.Root.Attribute(att.Name));
            }

            XElement oldElement = oldPresentationDoc.Root.Elements(P.sldSz).FirstOrDefault();
            if (oldElement != null)
                newPresentation.Root.Add(oldElement);

            // Copy Font Parts
            if (oldPresentationDoc.Root.Element(P.embeddedFontLst) != null)
            {
                XElement newFontLst = new XElement(P.embeddedFontLst);
                foreach (var font in oldPresentationDoc.Root.Element(P.embeddedFontLst).Elements(P.embeddedFont))
                {
                    XElement newRegular = null, newBold = null, newItalic = null, newBoldItalic = null;
                    if (font.Element(P.regular) != null)
                        newRegular = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.regular);
                    if (font.Element(P.bold) != null)
                        newBold = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.bold);
                    if (font.Element(P.italic) != null)
                        newItalic = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.italic);
                    if (font.Element(P.boldItalic) != null)
                        newBoldItalic = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.boldItalic);
                    XElement newEmbeddedFont = new XElement(P.embeddedFont,
                        font.Elements(P.font),
                        newRegular,
                        newBold,
                        newItalic,
                        newBoldItalic);
                    newFontLst.Add(newEmbeddedFont);
                }
                newPresentation.Root.Add(newFontLst);
            }

            newPresentation.Root.Add(oldPresentationDoc.Root.Element(P.defaultTextStyle));
            newPresentation.Root.Add(oldPresentationDoc.Root.Elements(P.extLst));

            //<p:embeddedFont xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            //                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            //  <p:font typeface="Perpetua" panose="02020502060401020303" pitchFamily="18" charset="0" />
            //  <p:regular r:id="rId5" />
            //  <p:bold r:id="rId6" />
            //  <p:italic r:id="rId7" />
            //  <p:boldItalic r:id="rId8" />
            //</p:embeddedFont>

            // Copy Handout Master
            if (sourceDocument.PresentationPart.HandoutMasterPart != null)
            {
                HandoutMasterPart oldMaster = sourceDocument.PresentationPart.HandoutMasterPart;
                HandoutMasterPart newMaster = newDocument.PresentationPart.AddNewPart<HandoutMasterPart>();

                // Copy theme for master
                ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
                newThemePart.PutXDocument(oldMaster.ThemePart.GetXDocument());
                CopyRelatedPartsForContentParts(newDocument, oldMaster.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images, mediaList);

                // Copy master
                newMaster.PutXDocument(oldMaster.GetXDocument());
                oldMaster.AddRelationships(newMaster, RelationshipMarkup, new[] { newMaster.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldMaster, newMaster, new[] { newMaster.GetXDocument().Root }, images, mediaList);

                newPresentation.Root.Add(
                    new XElement(P.handoutMasterIdLst, new XElement(P.handoutMasterId,
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster)))));
            }

            // Copy Notes Master
            CopyNotesMaster(sourceDocument, newDocument, images, mediaList);

            // Copy Presentation Properties
            if (sourceDocument.PresentationPart.PresentationPropertiesPart != null)
            {
                PresentationPropertiesPart newPart = newDocument.PresentationPart.AddNewPart<PresentationPropertiesPart>();
                XDocument xd1 = sourceDocument.PresentationPart.PresentationPropertiesPart.GetXDocument();
                xd1.Descendants(P.custShow).Remove();
                newPart.PutXDocument(xd1);
            }

            // Copy View Properties
            if (sourceDocument.PresentationPart.ViewPropertiesPart != null)
            {
                ViewPropertiesPart newPart = newDocument.PresentationPart.AddNewPart<ViewPropertiesPart>();
                XDocument xd = sourceDocument.PresentationPart.ViewPropertiesPart.GetXDocument();
                xd.Descendants(P.outlineViewPr).Elements(P.sldLst).Remove();
                newPart.PutXDocument(xd);
            }

            foreach (var legacyDocTextInfo in sourceDocument.PresentationPart.Parts.Where(p => p.OpenXmlPart.RelationshipType == "http://schemas.microsoft.com/office/2006/relationships/legacyDocTextInfo"))
            {
                LegacyDiagramTextInfoPart newPart = newDocument.PresentationPart.AddNewPart<LegacyDiagramTextInfoPart>();
                using (var stream = legacyDocTextInfo.OpenXmlPart.GetStream())
                {
                    newPart.FeedData(stream);
                }
            }

            var listOfRootChildren = newPresentation.Root.Elements().ToList();
            foreach (var rc in listOfRootChildren)
                rc.Remove();
            newPresentation.Root.Add(
                listOfRootChildren.OrderBy(e =>
                {
                    if (PresentationOrder.ContainsKey(e.Name))
                        return PresentationOrder[e.Name];
                    return 999;
                }));
        }
        public static XElement CreatedEmbeddedFontPart(this PresentationDocument sourceDocument, PresentationDocument newDocument, XElement font, XName fontXName)
        {
            XElement newRegular;
            FontPart oldFontPart = (FontPart)sourceDocument.PresentationPart.GetPartById((string)font.Element(fontXName).Attributes(R.id).FirstOrDefault());
            FontPartType fpt;
            if (oldFontPart.ContentType == "application/x-fontdata")
                fpt = FontPartType.FontData;
            else if (oldFontPart.ContentType == "application/x-font-ttf")
                fpt = FontPartType.FontTtf;
            else
                fpt = FontPartType.FontOdttf;
            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            var newFontPart = newDocument.PresentationPart.AddFontPart(fpt, newId);
            using (var stream = oldFontPart.GetStream())
            {
                newFontPart.FeedData(stream);
            }
            newRegular = new XElement(fontXName,
                new XAttribute(R.id, newId));
            return newRegular;
        }
        public static SlideMasterPart AppendSlides(this PresentationDocument sourceDocument, PresentationDocument newDocument, int start, int count, bool keepMaster, List<ImageData> images, SlideMasterPart currentMasterPart, List<MediaData> mediaList)
        {
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();
            if (newPresentation.Root.Element(P.sldIdLst) == null)
                newPresentation.Root.Add(new XElement(P.sldIdLst));
            uint newID = 256;
            var ids = newPresentation.Root.Descendants(P.sldId).Select(f => (uint)f.Attribute(NoNamespace.id));
            if (ids.Any())
                newID = ids.Max() + 1;
            var slideList = sourceDocument.PresentationPart.GetXDocument().Root.Descendants(P.sldId);
            if (slideList.Count() == 0 && (currentMasterPart == null || keepMaster))
            {
                var slideMasterPart = sourceDocument.PresentationPart.SlideMasterParts.FirstOrDefault();
                if (slideMasterPart != null)
                    currentMasterPart = CopyMasterSlide(sourceDocument, slideMasterPart, newDocument, newPresentation, images, mediaList);
                return currentMasterPart;
            }
            while (count > 0 && start < slideList.Count())
            {
                SlidePart slide = (SlidePart)sourceDocument.PresentationPart.GetPartById(slideList.ElementAt(start).Attribute(R.id).Value);
                if (currentMasterPart == null || keepMaster)
                    currentMasterPart = CopyMasterSlide(sourceDocument, slide.SlideLayoutPart.SlideMasterPart, newDocument, newPresentation, images, mediaList);
                SlidePart newSlide = newDocument.PresentationPart.AddNewPart<SlidePart>();
                newSlide.PutXDocument(slide.GetXDocument());
                slide.AddRelationships(newSlide, RelationshipMarkup, new[] { newSlide.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, slide, newSlide, new[] { newSlide.GetXDocument().Root }, images, mediaList);
                CopyTableStyles(sourceDocument, newDocument, slide, newSlide);
                if (slide.NotesSlidePart != null)
                {
                    if (newDocument.PresentationPart.NotesMasterPart == null)
                        CopyNotesMaster(sourceDocument, newDocument, images, mediaList);
                    NotesSlidePart newPart = newSlide.AddNewPart<NotesSlidePart>();
                    newPart.PutXDocument(slide.NotesSlidePart.GetXDocument());
                    newPart.AddPart(newSlide);
                    newPart.AddPart(newDocument.PresentationPart.NotesMasterPart);
                    slide.NotesSlidePart.AddRelationships(newPart, RelationshipMarkup, new[] { newPart.GetXDocument().Root });
                    CopyRelatedPartsForContentParts(newDocument, slide.NotesSlidePart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);
                }

                string layoutName = slide.SlideLayoutPart.GetXDocument().Root.Element(P.cSld).Attribute(NoNamespace.name).Value;
                foreach (SlideLayoutPart layoutPart in currentMasterPart.SlideLayoutParts)
                    if (layoutPart.GetXDocument().Root.Element(P.cSld).Attribute(NoNamespace.name).Value == layoutName)
                    {
                        newSlide.AddPart(layoutPart);
                        break;
                    }
                if (newSlide.SlideLayoutPart == null)
                    newSlide.AddPart(currentMasterPart.SlideLayoutParts.First());  // Cannot find matching layout part

                if (slide.SlideCommentsPart != null)
                    CopyComments(sourceDocument, newDocument, slide, newSlide);

                newPresentation.Root.Element(P.sldIdLst).Add(new XElement(P.sldId,
                    new XAttribute(NoNamespace.id, newID.ToString()),
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newSlide))));
                newID++;
                start++;
                count--;
            }
            return currentMasterPart;
        }
        public static SlideMasterPart CopyMasterSlide(this PresentationDocument sourceDocument, SlideMasterPart sourceMasterPart, PresentationDocument newDocument, XDocument newPresentation, List<ImageData> images, List<MediaData> mediaList)
        {
            // Search for existing master slide with same theme name
            XDocument oldTheme = sourceMasterPart.ThemePart.GetXDocument();
            String themeName = oldTheme.Root.Attribute(NoNamespace.name).Value;
            foreach (SlideMasterPart master in newDocument.PresentationPart.GetPartsOfType<SlideMasterPart>())
            {
                XDocument themeDoc = master.ThemePart.GetXDocument();
                if (themeDoc.Root.Attribute(NoNamespace.name).Value == themeName)
                    return master;
            }

            SlideMasterPart newMaster = newDocument.PresentationPart.AddNewPart<SlideMasterPart>();
            XDocument sourceMaster = sourceMasterPart.GetXDocument();

            // Add to presentation slide master list, need newID for layout IDs also
            uint newID = 2147483648;
            var ids = newPresentation.Root.Descendants(P.sldMasterId).Select(f => (uint)f.Attribute(NoNamespace.id));
            if (ids.Any())
            {
                newID = ids.Max();
                XElement maxMaster = newPresentation.Root.Descendants(P.sldMasterId).Where(f => (uint)f.Attribute(NoNamespace.id) == newID).FirstOrDefault();
                SlideMasterPart maxMasterPart = (SlideMasterPart)newDocument.PresentationPart.GetPartById(maxMaster.Attribute(R.id).Value);
                newID += (uint)maxMasterPart.GetXDocument().Root.Descendants(P.sldLayoutId).Count() + 1;
            }
            newPresentation.Root.Element(P.sldMasterIdLst).Add(new XElement(P.sldMasterId,
                new XAttribute(NoNamespace.id, newID.ToString()),
                new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster))));
            newID++;

            ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
            if (newDocument.PresentationPart.ThemePart == null)
                newThemePart = newDocument.PresentationPart.AddPart(newThemePart);
            newThemePart.PutXDocument(oldTheme);
            CopyRelatedPartsForContentParts(newDocument, sourceMasterPart.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images, mediaList);
            foreach (SlideLayoutPart layoutPart in sourceMasterPart.SlideLayoutParts)
            {
                SlideLayoutPart newLayout = newMaster.AddNewPart<SlideLayoutPart>();
                newLayout.PutXDocument(layoutPart.GetXDocument());
                layoutPart.AddRelationships(newLayout, RelationshipMarkup, new[] { newLayout.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, layoutPart, newLayout, new[] { newLayout.GetXDocument().Root }, images, mediaList);
                newLayout.AddPart(newMaster);
                string resID = sourceMasterPart.GetIdOfPart(layoutPart);
                XElement entry = sourceMaster.Root.Descendants(P.sldLayoutId).Where(f => f.Attribute(R.id).Value == resID).FirstOrDefault();
                entry.Attribute(R.id).SetValue(newMaster.GetIdOfPart(newLayout));
                entry.SetAttributeValue(NoNamespace.id, newID.ToString());
                newID++;
            }
            newMaster.PutXDocument(sourceMaster);
            sourceMasterPart.AddRelationships(newMaster, RelationshipMarkup, new[] { newMaster.GetXDocument().Root });
            CopyRelatedPartsForContentParts(newDocument, sourceMasterPart, newMaster, new[] { newMaster.GetXDocument().Root }, images, mediaList);

            return newMaster;
        }
        // Copies notes master and notesSz element from presentation
        public static void CopyNotesMaster(this PresentationDocument sourceDocument, PresentationDocument newDocument, List<ImageData> images, List<MediaData> mediaList)
        {
            // Copy notesSz element from presentation
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();
            XDocument oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();
            XElement oldElement = oldPresentationDoc.Root.Element(P.notesSz);
            newPresentation.Root.Element(P.notesSz).ReplaceWith(oldElement);

            // Copy Notes Master
            if (sourceDocument.PresentationPart.NotesMasterPart != null)
            {
                NotesMasterPart oldMaster = sourceDocument.PresentationPart.NotesMasterPart;
                NotesMasterPart newMaster = newDocument.PresentationPart.AddNewPart<NotesMasterPart>();

                // Copy theme for master
                if (oldMaster.ThemePart != null)
                {
                    ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
                    newThemePart.PutXDocument(oldMaster.ThemePart.GetXDocument());
                    CopyRelatedPartsForContentParts(newDocument, oldMaster.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images, mediaList);
                }

                // Copy master
                newMaster.PutXDocument(oldMaster.GetXDocument());
                oldMaster.AddRelationships(newMaster, RelationshipMarkup, new[] { newMaster.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldMaster, newMaster, new[] { newMaster.GetXDocument().Root }, images, mediaList);

                newPresentation.Root.Add(
                    new XElement(P.notesMasterIdLst, new XElement(P.notesMasterId,
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster)))));
            }
        }
        public static void CopyComments(this PresentationDocument oldDocument, PresentationDocument newDocument, SlidePart oldSlide, SlidePart newSlide)
        {
            newSlide.AddNewPart<SlideCommentsPart>();
            newSlide.SlideCommentsPart.PutXDocument(oldSlide.SlideCommentsPart.GetXDocument());
            XDocument newSlideComments = newSlide.SlideCommentsPart.GetXDocument();
            XDocument oldAuthors = oldDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
            foreach (XElement comment in newSlideComments.Root.Elements(P.cm))
            {
                XElement newAuthor = FindCommentsAuthor(newDocument, comment, oldAuthors);
                // Update last index value for new comment
                comment.Attribute(NoNamespace.authorId).SetValue(newAuthor.Attribute(NoNamespace.id).Value);
                uint lastIndex = Convert.ToUInt32(newAuthor.Attribute(NoNamespace.lastIdx).Value);
                comment.Attribute(NoNamespace.idx).SetValue(lastIndex.ToString());
                newAuthor.Attribute(NoNamespace.lastIdx).SetValue(Convert.ToString(lastIndex + 1));
            }
        }
        public static XElement FindCommentsAuthor(this PresentationDocument newDocument, XElement comment, XDocument oldAuthors)
        {
            XElement oldAuthor = oldAuthors.Root.Elements(P.cmAuthor).Where(
                f => f.Attribute(NoNamespace.id).Value == comment.Attribute(NoNamespace.authorId).Value).FirstOrDefault();
            XElement newAuthor = null;
            if (newDocument.PresentationPart.CommentAuthorsPart == null)
            {
                newDocument.PresentationPart.AddNewPart<CommentAuthorsPart>();
                newDocument.PresentationPart.CommentAuthorsPart.PutXDocument(new XDocument(new XElement(P.cmAuthorLst,
                    new XAttribute(XNamespace.Xmlns + "a", A.a),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XAttribute(XNamespace.Xmlns + "p", P.p))));
            }
            XDocument authors = newDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
            newAuthor = authors.Root.Elements(P.cmAuthor).Where(
                f => f.Attribute(NoNamespace.initials).Value == oldAuthor.Attribute(NoNamespace.initials).Value).FirstOrDefault();
            if (newAuthor == null)
            {
                uint newID = 0;
                var ids = authors.Root.Descendants(P.cmAuthor).Select(f => (uint)f.Attribute(NoNamespace.id));
                if (ids.Any())
                    newID = ids.Max() + 1;

                newAuthor = new XElement(P.cmAuthor, new XAttribute(NoNamespace.id, newID.ToString()),
                    new XAttribute(NoNamespace.name, oldAuthor.Attribute(NoNamespace.name).Value),
                    new XAttribute(NoNamespace.initials, oldAuthor.Attribute(NoNamespace.initials).Value),
                    new XAttribute(NoNamespace.lastIdx, "1"), new XAttribute(NoNamespace.clrIdx, newID.ToString()));
                authors.Root.Add(newAuthor);
            }

            return newAuthor;
        }
        public static void CopyTableStyles(this PresentationDocument oldDocument, PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart)
        {
            foreach (XElement table in newContentPart.GetXDocument().Descendants(A.tableStyleId))
            {
                string styleId = table.Value;
                if (string.IsNullOrEmpty(styleId))
                    continue;

                // Find old style
                if (oldDocument.PresentationPart.TableStylesPart == null)
                    continue;
                XDocument oldTableStyles = oldDocument.PresentationPart.TableStylesPart.GetXDocument();
                XElement oldStyle = oldTableStyles.Root.Elements(A.tblStyle).Where(f => f.Attribute(NoNamespace.styleId).Value == styleId).FirstOrDefault();
                if (oldStyle == null)
                    continue;

                // Create new TableStylesPart, if needed
                XDocument tableStyles = null;
                if (newDocument.PresentationPart.TableStylesPart == null)
                {
                    TableStylesPart newStylesPart = newDocument.PresentationPart.AddNewPart<TableStylesPart>();
                    tableStyles = new XDocument(new XElement(A.tblStyleLst,
                        new XAttribute(XNamespace.Xmlns + "a", A.a),
                        new XAttribute(NoNamespace.def, styleId)));
                    newStylesPart.PutXDocument(tableStyles);
                }
                else
                    tableStyles = newDocument.PresentationPart.TableStylesPart.GetXDocument();

                // Search new TableStylesPart to see if it contains the ID
                if (tableStyles.Root.Elements(A.tblStyle).Where(f => f.Attribute(NoNamespace.styleId).Value == styleId).FirstOrDefault() != null)
                    continue;

                // Copy style to new part
                tableStyles.Root.Add(oldStyle);
            }

        }
        public static void CopyRelatedPartsForContentParts(this PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart, IEnumerable<XElement> newContent, List<ImageData> images, List<MediaData> mediaList)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == VML.imagedata || d.Name == VML.fill || d.Name == VML.stroke || d.Name == A.blip || d.Name == SVG.svgBlip)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                oldContentPart.CopyRelatedImage(newContentPart, imageReference, R.embed, images);
                oldContentPart.CopyRelatedImage(newContentPart, imageReference, R.pict, images);
                oldContentPart.CopyRelatedImage(newContentPart, imageReference, R.id, images);
                oldContentPart.CopyRelatedImage(newContentPart, imageReference, O.relid, images);
            }

            relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == A.videoFile || d.Name == A.quickTimeFile)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                oldContentPart.CopyRelatedMedia(newContentPart, imageReference, R.link, mediaList, "video");
            }

            relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == P14.media || d.Name == PAV.srcMedia)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                oldContentPart.CopyRelatedMedia(newContentPart, imageReference, R.embed, mediaList, "media");
                oldContentPart.CopyRelatedMediaExternalRelationship(newContentPart, imageReference, R.link, "media");
            }

            foreach (XElement extendedReference in newContent.DescendantsAndSelf(A14.imgLayer))
            {
                oldContentPart.CopyExtendedPart(newContentPart, extendedReference, R.embed);
            }

            foreach (XElement contentPartReference in newContent.DescendantsAndSelf(P.contentPart))
            {
                oldContentPart.CopyInkPart(newContentPart, contentPartReference, R.id);
            }

            foreach (XElement contentPartReference in newContent.DescendantsAndSelf(P.control))
            {
                oldContentPart.CopyActiveXPart(newContentPart, contentPartReference, R.id);
            }

            foreach (XElement contentPartReference in newContent.DescendantsAndSelf(Plegacy.textdata))
            {
                oldContentPart.CopyLegacyDiagramText(newContentPart, contentPartReference, "id");
            }

            foreach (XElement diagramReference in newContent.DescendantsAndSelf().Where(d => d.Name == DGM.relIds || d.Name == A.relIds))
            {
                // dm attribute
                string relId = diagramReference.Attribute(R.dm).Value;
                var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair != null)
                    continue;

                ExternalRelationship tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr != null)
                    continue;

                OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                OpenXmlPart newPart = newContentPart.AddNewPart<DiagramDataPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.dm).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, RelationshipMarkup, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);

                // lo attribute
                relId = diagramReference.Attribute(R.lo).Value;
                var tempPartIdPair2 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair2 != null)
                    continue;

                ExternalRelationship tempEr2 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr2 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramLayoutDefinitionPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.lo).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, RelationshipMarkup, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);

                // qs attribute
                relId = diagramReference.Attribute(R.qs).Value;
                var tempPartIdPair3 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair3 != null)
                    continue;

                ExternalRelationship tempEr3 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr3 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramStylePart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.qs).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, RelationshipMarkup, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);

                // cs attribute
                relId = diagramReference.Attribute(R.cs).Value;
                var tempPartIdPair4 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair4 != null)
                    continue;

                ExternalRelationship tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr4 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramColorsPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.cs).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, RelationshipMarkup, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);
            }

            foreach (XElement oleReference in newContent.DescendantsAndSelf().Where(d => d.Name == P.oleObj || d.Name == P.externalData))
            {
                string relId = oleReference.Attribute(R.id).Value;

                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                var tempPartIdPair5 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair5 != null)
                    continue;

                ExternalRelationship tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr5 != null)
                    continue;

                var oldPartIdPair = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair != null)
                {
                    OpenXmlPart oldPart = oldPartIdPair.OpenXmlPart;
                    OpenXmlPart newPart = null;
                    if (oldPart is EmbeddedObjectPart)
                    {
                        if (newContentPart is DialogsheetPart)
                            newPart = ((DialogsheetPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        else if (newContentPart is HandoutMasterPart)
                            newPart = ((HandoutMasterPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        else if (newContentPart is NotesMasterPart)
                            newPart = ((NotesMasterPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        else if (newContentPart is NotesSlidePart)
                            newPart = ((NotesSlidePart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        else if (newContentPart is SlideLayoutPart)
                            newPart = ((SlideLayoutPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        else if (newContentPart is SlideMasterPart)
                            newPart = ((SlideMasterPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        else if (newContentPart is SlidePart)
                            newPart = ((SlidePart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                    }
                    else if (oldPart is EmbeddedPackagePart)
                    {
                        if (newContentPart is ChartPart)
                            newPart = ((ChartPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        else if (newContentPart is HandoutMasterPart)
                            newPart = ((HandoutMasterPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        else if (newContentPart is NotesMasterPart)
                            newPart = ((NotesMasterPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        else if (newContentPart is NotesSlidePart)
                            newPart = ((NotesSlidePart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        else if (newContentPart is SlideLayoutPart)
                            newPart = ((SlideLayoutPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        else if (newContentPart is SlideMasterPart)
                            newPart = ((SlideMasterPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        else if (newContentPart is SlidePart)
                            newPart = ((SlidePart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                    }
                    using (Stream oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                    using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                    {
                        int byteCount;
                        byte[] buffer = new byte[65536];
                        while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                            newObject.Write(buffer, 0, byteCount);
                    }
                    oleReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                }
                else
                {
                    ExternalRelationship er = oldContentPart.GetExternalRelationship(relId);
                    ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    oleReference.Attribute(R.id).Value = newEr.Id;
                }
            }

            foreach (XElement chartReference in newContent.DescendantsAndSelf(C.chart))
            {
                string relId = (string)chartReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair6 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair6 != null)
                    continue;

                ExternalRelationship tempEr6 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr6 != null)
                    continue;

                var oldPartIdPair2 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair2 != null)
                {
                    ChartPart oldPart = oldPartIdPair2.OpenXmlPart as ChartPart;
                    if (oldPart != null)
                    {
                        XDocument oldChart = oldPart.GetXDocument();
                        ChartPart newPart = newContentPart.AddNewPart<ChartPart>();
                        XDocument newChart = newPart.GetXDocument();
                        newChart.Add(oldChart.Root);
                        chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                        oldPart.CopyChartObjects(newPart);
                        CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newChart.Root }, images, mediaList);
                    }
                }
            }

            foreach (XElement userShape in newContent.DescendantsAndSelf(C.userShapes))
            {
                string relId = (string)userShape.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair7 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair7 != null)
                    continue;

                ExternalRelationship tempEr7 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr7 != null)
                    continue;

                var oldPartIdPair3 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair3 != null)
                {
                    ChartDrawingPart oldPart = oldPartIdPair3.OpenXmlPart as ChartDrawingPart;
                    if (oldPart != null)
                    {
                        XDocument oldXDoc = oldPart.GetXDocument();
                        ChartDrawingPart newPart = newContentPart.AddNewPart<ChartDrawingPart>();
                        XDocument newXDoc = newPart.GetXDocument();
                        newXDoc.Add(oldXDoc.Root);
                        userShape.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                        oldPart.AddRelationships(newPart, RelationshipMarkup, newContent);
                        CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newXDoc.Root }, images, mediaList);
                    }
                }
            }

            foreach (XElement tags in newContent.DescendantsAndSelf(P.tags))
            {
                string relId = (string)tags.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair8 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair8 != null)
                    continue;

                ExternalRelationship tempEr8 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr8 != null)
                    continue;

                var oldPartIdPair4 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair4 != null)
                {
                    UserDefinedTagsPart oldPart = oldPartIdPair4.OpenXmlPart as UserDefinedTagsPart;
                    if (oldPart != null)
                    {
                        XDocument oldXDoc = oldPart.GetXDocument();
                        UserDefinedTagsPart newPart = newContentPart.AddNewPart<UserDefinedTagsPart>();
                        XDocument newXDoc = newPart.GetXDocument();
                        newXDoc.Add(oldXDoc.Root);
                        tags.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                    }
                }
            }

            foreach (XElement custData in newContent.DescendantsAndSelf(P.custData))
            {
                string relId = (string)custData.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair9 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair9 != null)
                    continue;

                var oldPartIdPair9 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair9 != null)
                {
                    CustomXmlPart newPart = newDocument.PresentationPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                    using (var stream = oldPartIdPair9.OpenXmlPart.GetStream())
                    {
                        newPart.FeedData(stream);
                    }
                    foreach (var itemProps in oldPartIdPair9.OpenXmlPart.Parts.Where(p => p.OpenXmlPart.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"))
                    {
                        var newId2 = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
                        CustomXmlPropertiesPart cxpp = newPart.AddNewPart<CustomXmlPropertiesPart>("application/vnd.openxmlformats-officedocument.customXmlProperties+xml", newId2);
                        using (var stream = itemProps.OpenXmlPart.GetStream())
                        {
                            cxpp.FeedData(stream);
                        }
                    }
                    var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
                    newContentPart.CreateRelationshipToPart(newPart, newId);
                    custData.Attribute(R.id).Value = newId;
                }
            }

            foreach (XElement soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == A.audioFile))
                CopyRelatedSound(newDocument, oldContentPart, newContentPart, soundReference, R.link);

            if ((oldContentPart is ChartsheetPart && newContentPart is ChartsheetPart) ||
                (oldContentPart is DialogsheetPart && newContentPart is DialogsheetPart) ||
                (oldContentPart is HandoutMasterPart && newContentPart is HandoutMasterPart) ||
                (oldContentPart is InternationalMacroSheetPart && newContentPart is InternationalMacroSheetPart) ||
                (oldContentPart is MacroSheetPart && newContentPart is MacroSheetPart) ||
                (oldContentPart is NotesMasterPart && newContentPart is NotesMasterPart) ||
                (oldContentPart is NotesSlidePart && newContentPart is NotesSlidePart) ||
                (oldContentPart is SlideLayoutPart && newContentPart is SlideLayoutPart) ||
                (oldContentPart is SlideMasterPart && newContentPart is SlideMasterPart) ||
                (oldContentPart is SlidePart && newContentPart is SlidePart) ||
                (oldContentPart is WorksheetPart && newContentPart is WorksheetPart))
            {
                foreach (XElement soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == P.snd || d.Name == P.sndTgt || d.Name == A.wavAudioFile || d.Name == A.snd || d.Name == PAV.srcMedia))
                    CopyRelatedSound(newDocument, oldContentPart, newContentPart, soundReference, R.embed);

                IEnumerable<VmlDrawingPart> vmlDrawingParts = null;
                if (oldContentPart is ChartsheetPart)
                    vmlDrawingParts = ((ChartsheetPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is DialogsheetPart)
                    vmlDrawingParts = ((DialogsheetPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is HandoutMasterPart)
                    vmlDrawingParts = ((HandoutMasterPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is InternationalMacroSheetPart)
                    vmlDrawingParts = ((InternationalMacroSheetPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is MacroSheetPart)
                    vmlDrawingParts = ((MacroSheetPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is NotesMasterPart)
                    vmlDrawingParts = ((NotesMasterPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is NotesSlidePart)
                    vmlDrawingParts = ((NotesSlidePart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is SlideLayoutPart)
                    vmlDrawingParts = ((SlideLayoutPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is SlideMasterPart)
                    vmlDrawingParts = ((SlideMasterPart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is SlidePart)
                    vmlDrawingParts = ((SlidePart)oldContentPart).VmlDrawingParts;
                if (oldContentPart is WorksheetPart)
                    vmlDrawingParts = ((WorksheetPart)oldContentPart).VmlDrawingParts;

                if (vmlDrawingParts != null)
                {
                    // Transitional: Copy VML Drawing parts, implicit relationship
                    foreach (VmlDrawingPart vmlPart in vmlDrawingParts)
                    {
                        VmlDrawingPart newVmlPart = null;
                        if (newContentPart is ChartsheetPart)
                            newVmlPart = ((ChartsheetPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is DialogsheetPart)
                            newVmlPart = ((DialogsheetPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is HandoutMasterPart)
                            newVmlPart = ((HandoutMasterPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is InternationalMacroSheetPart)
                            newVmlPart = ((InternationalMacroSheetPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is MacroSheetPart)
                            newVmlPart = ((MacroSheetPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is NotesMasterPart)
                            newVmlPart = ((NotesMasterPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is NotesSlidePart)
                            newVmlPart = ((NotesSlidePart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is SlideLayoutPart)
                            newVmlPart = ((SlideLayoutPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is SlideMasterPart)
                            newVmlPart = ((SlideMasterPart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is SlidePart)
                            newVmlPart = ((SlidePart)newContentPart).AddNewPart<VmlDrawingPart>();
                        if (newContentPart is WorksheetPart)
                            newVmlPart = ((WorksheetPart)newContentPart).AddNewPart<VmlDrawingPart>();

                        XDocument xd = vmlPart.GetXDocument();
                        foreach (var item in xd.Descendants(O.ink))
                        {
                            if (item.Attribute("i") != null)
                            {
                                var i = item.Attribute("i").Value;
                                i = i.Replace(" ", "\r\n");
                                item.Attribute("i").Value = i;
                            }
                        }
                        newVmlPart.PutXDocument(xd);
                        vmlPart.AddRelationships(newVmlPart, RelationshipMarkup, new[] { newVmlPart.GetXDocument().Root });
                        CopyRelatedPartsForContentParts(newDocument, vmlPart, newVmlPart, new[] { newVmlPart.GetXDocument().Root }, images, mediaList);
                    }
                }
            }
        }
        public static void CopyStartingParts(this PresentationDocument sourceDocument, PresentationDocument newDocument)
        {
            // A Core File Properties part does not have implicit or explicit relationships to other parts.
            CoreFilePropertiesPart corePart = sourceDocument.CoreFilePropertiesPart;
            if (corePart != null && corePart.GetXDocument().Root != null)
            {
                newDocument.AddCoreFilePropertiesPart();
                XDocument newXDoc = newDocument.CoreFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = "yes";
                newXDoc.Declaration.Encoding = "UTF-8";
                XDocument sourceXDoc = corePart.GetXDocument();
                newXDoc.Add(sourceXDoc.Root);
            }

            // An application attributes part does not have implicit or explicit relationships to other parts.
            ExtendedFilePropertiesPart extPart = sourceDocument.ExtendedFilePropertiesPart;
            if (extPart != null)
            {
                OpenXmlPart newPart = newDocument.AddExtendedFilePropertiesPart();
                XDocument newXDoc = newDocument.ExtendedFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = "yes";
                newXDoc.Declaration.Encoding = "UTF-8";
                newXDoc.Add(extPart.GetXDocument().Root);
            }

            // An custom file properties part does not have implicit or explicit relationships to other parts.
            CustomFilePropertiesPart customPart = sourceDocument.CustomFilePropertiesPart;
            if (customPart != null)
            {
                newDocument.AddCustomFilePropertiesPart();
                XDocument newXDoc = newDocument.CustomFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = "yes";
                newXDoc.Declaration.Encoding = "UTF-8";
                newXDoc.Add(customPart.GetXDocument().Root);
            }
        }
        public static void CopyRelatedSound(this PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement soundReference, XName attributeName)
        {
            string relId = (string)soundReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            ExternalRelationship alreadyExistingExternalRelationship = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (alreadyExistingExternalRelationship != null)
                return;

            ReferenceRelationship alreadyExistingReferenceRelationship = newContentPart.DataPartReferenceRelationships.FirstOrDefault(er => er.Id == relId);
            if (alreadyExistingReferenceRelationship != null)
                return;

            if (oldContentPart.GetReferenceRelationship(relId) is AudioReferenceRelationship)
            {
                AudioReferenceRelationship temp = (AudioReferenceRelationship)oldContentPart.GetReferenceRelationship(relId);
                MediaDataPart newSound = newDocument.CreateMediaDataPart(temp.DataPart.ContentType);
                using (var stream = temp.DataPart.GetStream())
                {
                    newSound.FeedData(stream);
                }
                AudioReferenceRelationship newRel = null;

                if (newContentPart is SlidePart)
                    newRel = ((SlidePart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is SlideLayoutPart)
                    newRel = ((SlideLayoutPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is SlideMasterPart)
                    newRel = ((SlideMasterPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is HandoutMasterPart)
                    newRel = ((HandoutMasterPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is NotesMasterPart)
                    newRel = ((NotesMasterPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is NotesSlidePart)
                    newRel = ((NotesSlidePart)newContentPart).AddAudioReferenceRelationship(newSound);
                soundReference.Attribute(attributeName).Value = newRel.Id;
            }
            if (oldContentPart.GetReferenceRelationship(relId) is MediaReferenceRelationship)
            {
                MediaReferenceRelationship temp = (MediaReferenceRelationship)oldContentPart.GetReferenceRelationship(relId);
                MediaDataPart newSound = newDocument.CreateMediaDataPart(temp.DataPart.ContentType);
                using (var stream = temp.DataPart.GetStream())
                {
                    newSound.FeedData(stream);
                }
                MediaReferenceRelationship newRel = null;

                if (newContentPart is SlidePart)
                    newRel = ((SlidePart)newContentPart).AddMediaReferenceRelationship(newSound);
                else if (newContentPart is SlideLayoutPart)
                    newRel = ((SlideLayoutPart)newContentPart).AddMediaReferenceRelationship(newSound);
                else if (newContentPart is SlideMasterPart)
                    newRel = ((SlideMasterPart)newContentPart).AddMediaReferenceRelationship(newSound);
                soundReference.Attribute(attributeName).Value = newRel.Id;
            }
        }
        public static void FixUpPresentationDocument(PresentationDocument pDoc)
        {
            foreach (var part in pDoc.GetAllParts())
            {
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.theme+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml" ||
                    part.ContentType == "application/vnd.ms-office.drawingml.diagramDrawing+xml")
                {
                    XDocument xd = part.GetXDocument();
                    xd.Descendants().Attributes("smtClean").Remove();
                    xd.Descendants().Attributes("smtId").Remove();
                    part.PutXDocument();
                }
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.vmlDrawing")
                {
                    string fixedContent = null;

                    using (var stream = part.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        using (var sr = new StreamReader(stream))
                        {
                            //string input = @"    <![if gte mso 9]><v:imagedata o:relid=""rId15""";
                            var input = sr.ReadToEnd();
                            string pattern = @"<!\[(?<test>.*)\]>";
                            //string replacement = "<![CDATA[${test}]]>";
                            //fixedContent = Regex.Replace(input, pattern, replacement, RegexOptions.Multiline);
                            fixedContent = Regex.Replace(input, pattern, m =>
                            {
                                var g = m.Groups[1].Value;
                                if (g.StartsWith("CDATA["))
                                    return "<![" + g + "]>";
                                else
                                    return "<![CDATA[" + g + "]]>";
                            },
                            RegexOptions.Multiline);

                            //var input = @"xxxxx o:relid=""rId1"" o:relid=""rId1"" xxxxx";
                            pattern = @"o:relid=[""'](?<id1>.*)[""'] o:relid=[""'](?<id2>.*)[""']";
                            fixedContent = Regex.Replace(fixedContent, pattern, m =>
                            {
                                var g = m.Groups[1].Value;
                                return @"o:relid=""" + g + @"""";
                            },
                            RegexOptions.Multiline);

                            fixedContent = fixedContent.Replace("</xml>ml>", "</xml>");

                            stream.SetLength(fixedContent.Length);
                        }
                    }
                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(fixedContent)))
                        part.FeedData(ms);
                }
            }
        }
        #endregion

        #region SlideSource
        public static void BuildPresentation(this List<SlideSource> sources, string fileName)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = CreatePresentationDocument())
            {
                using (PresentationDocument output = streamDoc.GetPresentationDocument())
                {
                    BuildPresentation(sources, output);
                    output.Close();
                }
                streamDoc.GetModifiedDocument().SaveAs(fileName);
            }
        }
        public static PmlDocument BuildPresentation(this List<SlideSource> sources)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = CreatePresentationDocument())
            {
                using (PresentationDocument output = streamDoc.GetPresentationDocument())
                {
                    BuildPresentation(sources, output);
                    output.Close();
                }
                return streamDoc.GetModifiedPmlDocument();
            }
        }
        public static void BuildPresentation(this List<SlideSource> sources, PresentationDocument output)
        {
            List<ImageData> images = new List<ImageData>();
            List<MediaData> mediaList = new List<MediaData>();
            XDocument mainPart = output.PresentationPart.GetXDocument();
            mainPart.Declaration.Standalone = "yes";
            mainPart.Declaration.Encoding = "UTF-8";
            output.PresentationPart.PutXDocument();

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].PmlDocument))
            using (PresentationDocument doc = streamDoc.GetPresentationDocument())
            {
                CopyStartingParts(doc, output);
            }

            int sourceNum = 0;
            SlideMasterPart currentMasterPart = null;
            foreach (SlideSource source in sources)
            {
                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.PmlDocument))
                using (PresentationDocument doc = streamDoc.GetPresentationDocument())
                {
                    try
                    {
                        if (sourceNum == 0)
                            CopyPresentationParts(doc, output, images, mediaList);
                        currentMasterPart = AppendSlides(doc, output, source.Start, source.Count, source.KeepMaster, images, currentMasterPart, mediaList);
                    }
                    catch (PresentationBuilderInternalException dbie)
                    {
                        if (dbie.Message.Contains("{0}"))
                            throw new PresentationBuilderException(string.Format(dbie.Message, sourceNum));
                        else
                            throw dbie;
                    }
                }
                sourceNum++;
            }
            foreach (var part in output.GetAllParts())
            {
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.theme+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml" ||
                    part.ContentType == "application/vnd.ms-office.drawingml.diagramDrawing+xml")
                {
                    XDocument xd = part.GetXDocument();
                    xd.Descendants().Attributes("smtClean").Remove();
                    part.PutXDocument();
                }
                else if (part.Annotation<XDocument>() != null)
                    part.PutXDocument();
            }
        }
        #endregion


    }
}