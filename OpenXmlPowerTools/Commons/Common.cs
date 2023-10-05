using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools.Documents;
using OpenXmlPowerTools.Presentations;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Commons
{
    public static class Common
    {

        #region Declaration
        /// <summary>
        /// Get new XDeclaration
        /// </summary>
        public static XDeclaration CreateDeclaration() => new XDeclaration(Constants.OnePointZero, Constants.Utf8, Constants.Yes);
        public static void SetDeclaration(this XDeclaration dec)
        {
            dec.Standalone = Constants.Yes;
            dec.Encoding = Constants.Utf8;
        }
        #endregion

        #region OpenXml Parts
        public static void AddRelationships(this OpenXmlPart oldPart, OpenXmlPart newPart, Dictionary<XName, XName[]> relationshipMarkup, IEnumerable<XElement> newContent)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => relationshipMarkup.ContainsKey(d.Name) &&
                    d.Attributes().Any(a => relationshipMarkup[d.Name].Contains(a.Name))).ToList();
            foreach (var e in relevantElements)
            {
                if (e.Name == W.hyperlink || e.Name == A.hlinkHover || e.Name == A.hlinkMouseOver)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                    {
                        // handle the following:
                        //<a:hlinkClick r:id=""
                        //              action="ppaction://customshow?id=0" />

                        //var action = (string)e.Attribute("action");
                        //if (action != null)
                        //{
                        //    if (action.Contains("customshow"))
                        //        e.Attribute("action").Remove();
                        //}
                        continue;
                    }
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    //string newRid = "R" + g.ToString().Replace("-", "");

                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null)
                    {
                        //throw new DocumentBuilderInternalException("Internal Error 0002");
                        //TODO Issue with reference to another part: var temp = oldPart.GetPartById(relId);
                        //RemoveContent(newContent, e.Name, relId);
                        continue;
                    }
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    newContent.UpdateContent(relationshipMarkup, e.Name, relId, newRid);
                }
                if (e.Name == W.attachedTemplate || e.Name == W.saveThroughXslt)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new DocumentBuilderInternalException("Source {0} is invalid document - hyperlink contains invalid references");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    newContent.UpdateContent(relationshipMarkup, e.Name, relId, newRid);
                }
                if (e.Name == A.hlinkClick || e.Name == A.hlinkHover || e.Name == A.hlinkMouseOver)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null)
                        continue;
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    newContent.UpdateContent(relationshipMarkup, e.Name, relId, newRid);
                }
                if (e.Name == VML.imagedata)
                {
                    string relId = (string)e.Attribute(R.href);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    //string newRid = "R" + g.ToString().Replace("-", "");

                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new DocumentBuilderInternalException("Internal Error 0006");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    newContent.UpdateContent(relationshipMarkup, e.Name, relId, newRid);
                }
                if (e.Name == A.blip || e.Name == A14.imgLayer || e.Name == A.audioFile || e.Name == A.videoFile || e.Name == A.quickTimeFile)
                {
                    // <a:blip r:embed="rId6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
                    string relId = (string)e.Attribute(R.link);
                    //if (relId == null)
                    //    relId = (string)e.Attribute(R.embed);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    //string newRid = "R" + g.ToString().Replace("-", "");

                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        continue;
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    newContent.UpdateContent(relationshipMarkup, e.Name, relId, newRid);
                }
            }
        }
        public static void CopyActiveXPart(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement activeXPartReference, XName attributeName)
        {
            string relId = (string)activeXPartReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair != null)
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            EmbeddedControlPersistencePart newPart = newContentPart.AddNewPart<EmbeddedControlPersistencePart>("application/vnd.ms-office.activeX+xml", newId);

            using (var stream = oldPart.GetStream())
            {
                newPart.FeedData(stream);
            }
            activeXPartReference.Attribute(attributeName).Value = newId;

            if (newPart.ContentType == "application/vnd.ms-office.activeX+xml")
            {
                XDocument axc = newPart.GetXDocument();
                if (axc.Root.Attribute(R.id) != null)
                {
                    var oldPersistencePart = oldPart.GetPartById((string)axc.Root.Attribute(R.id));

                    var newId2 = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
                    EmbeddedControlPersistenceBinaryDataPart newPersistencePart = newPart.AddNewPart<EmbeddedControlPersistenceBinaryDataPart>("application/vnd.ms-office.activeX", newId2);

                    using (var stream = oldPersistencePart.GetStream())
                    {
                        newPersistencePart.FeedData(stream);
                    }
                    axc.Root.Attribute(R.id).Value = newId2;
                    newPart.PutXDocument();
                }
            }
        }
        public static void CopyChartObjects(this ChartPart oldChart, ChartPart newChart)
        {
            foreach (XElement dataReference in newChart.GetXDocument().Descendants(C.externalData))
            {
                string relId = dataReference.Attribute(R.id).Value;

                var oldPartIdPair = oldChart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair != null)
                {
                    EmbeddedPackagePart oldPart = oldPartIdPair.OpenXmlPart as EmbeddedPackagePart;
                    if (oldPart != null)
                    {
                        EmbeddedPackagePart newPart = newChart.AddEmbeddedPackagePart(oldPart.ContentType);
                        using (Stream oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                        using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                        {
                            int byteCount;
                            byte[] buffer = new byte[65536];
                            while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                                newObject.Write(buffer, 0, byteCount);
                        }
                        dataReference.Attribute(R.id).Value = newChart.GetIdOfPart(newPart);
                        continue;
                    }
                    ExtendedPart extendedPart = oldPartIdPair.OpenXmlPart as ExtendedPart;
                    if (extendedPart != null)
                    {
                        ExtendedPart newPart = newChart.AddExtendedPart(extendedPart.RelationshipType, extendedPart.ContentType, ".dat");
                        using (Stream oldObject = extendedPart.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                            {
                                int byteCount;
                                byte[] buffer = new byte[65536];
                                while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                                    newObject.Write(buffer, 0, byteCount);
                            }
                        }
                        dataReference.Attribute(R.id).Value = newChart.GetIdOfPart(newPart);
                        continue;
                    }
                    EmbeddedObjectPart oldEmbeddedObjectPart = oldPartIdPair.OpenXmlPart as EmbeddedObjectPart;
                    if (oldEmbeddedObjectPart != null)
                    {
                        EmbeddedPackagePart newPart = newChart.AddEmbeddedPackagePart(oldEmbeddedObjectPart.ContentType);
                        using (Stream oldObject = oldEmbeddedObjectPart.GetStream(FileMode.Open, FileAccess.Read))
                        using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                        {
                            int byteCount;
                            byte[] buffer = new byte[65536];
                            while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                                newObject.Write(buffer, 0, byteCount);
                        }

                        var rId = newChart.GetIdOfPart(newPart);
                        dataReference.Attribute(R.id).Value = rId;

                        // following is a hack to fix the package because the Open XML SDK does not let us create
                        // a relationship from a chart with the oleObject relationship type.

                        var pkg = newChart.OpenXmlPackage.Package;
                        var fromPart = pkg.GetParts().FirstOrDefault(p => p.Uri == newChart.Uri);
                        var rel = fromPart.GetRelationships().FirstOrDefault(p => p.Id == rId);
                        var targetUri = rel.TargetUri;

                        fromPart.DeleteRelationship(rId);
                        fromPart.CreateRelationship(targetUri, System.IO.Packaging.TargetMode.Internal,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", rId);

                        continue;
                    }
                }
                else
                {
                    ExternalRelationship oldRelationship = oldChart.GetExternalRelationship(relId);
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldRel = oldChart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new PresentationBuilderInternalException("Internal Error 0007");
                    newChart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    dataReference.Attribute(R.id).Value = newRid;
                }
            }
        }
        public static void CopyExtendedPart(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement extendedReference, XName attributeName)
        {
            string relId = (string)extendedReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;
            try
            {
                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair != null)
                    return;

                var tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr != null)
                    return;

                ExtendedPart oldPart = (ExtendedPart)oldContentPart.GetPartById(relId);
                FileInfo fileInfo = new FileInfo(oldPart.Uri.OriginalString);
                ExtendedPart newPart = null;

#if !NET35
                if (newContentPart is ChartColorStylePart)
                    newPart = ((ChartColorStylePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else
#endif
                if (newContentPart is ChartDrawingPart)
                    newPart = ((ChartDrawingPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ChartPart)
                    newPart = ((ChartPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ChartsheetPart)
                    newPart = ((ChartsheetPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
#if !NET35
                else if (newContentPart is ChartStylePart)
                    newPart = ((ChartStylePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
#endif
                else if (newContentPart is CommentAuthorsPart)
                    newPart = ((CommentAuthorsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ConnectionsPart)
                    newPart = ((ConnectionsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ControlPropertiesPart)
                    newPart = ((ControlPropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CoreFilePropertiesPart)
                    newPart = ((CoreFilePropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomDataPart)
                    newPart = ((CustomDataPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomDataPropertiesPart)
                    newPart = ((CustomDataPropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomFilePropertiesPart)
                    newPart = ((CustomFilePropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomizationPart)
                    newPart = ((CustomizationPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomPropertyPart)
                    newPart = ((CustomPropertyPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomUIPart)
                    newPart = ((CustomUIPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomXmlMappingsPart)
                    newPart = ((CustomXmlMappingsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomXmlPart)
                    newPart = ((CustomXmlPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is CustomXmlPropertiesPart)
                    newPart = ((CustomXmlPropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is DiagramColorsPart)
                    newPart = ((DiagramColorsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is DiagramDataPart)
                    newPart = ((DiagramDataPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is DiagramLayoutDefinitionPart)
                    newPart = ((DiagramLayoutDefinitionPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is DiagramPersistLayoutPart)
                    newPart = ((DiagramPersistLayoutPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is DiagramStylePart)
                    newPart = ((DiagramStylePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is DigitalSignatureOriginPart)
                    newPart = ((DigitalSignatureOriginPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is DrawingsPart)
                    newPart = ((DrawingsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is EmbeddedControlPersistenceBinaryDataPart)
                    newPart = ((EmbeddedControlPersistenceBinaryDataPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is EmbeddedControlPersistencePart)
                    newPart = ((EmbeddedControlPersistencePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is EmbeddedObjectPart)
                    newPart = ((EmbeddedObjectPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is EmbeddedPackagePart)
                    newPart = ((EmbeddedPackagePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ExtendedFilePropertiesPart)
                    newPart = ((ExtendedFilePropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ExtendedPart)
                    newPart = ((ExtendedPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is FontPart)
                    newPart = ((FontPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is FontTablePart)
                    newPart = ((FontTablePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is HandoutMasterPart)
                    newPart = ((HandoutMasterPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is InternationalMacroSheetPart)
                    newPart = ((InternationalMacroSheetPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is LegacyDiagramTextInfoPart)
                    newPart = ((LegacyDiagramTextInfoPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is LegacyDiagramTextPart)
                    newPart = ((LegacyDiagramTextPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is MacroSheetPart)
                    newPart = ((MacroSheetPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is NotesMasterPart)
                    newPart = ((NotesMasterPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is NotesSlidePart)
                    newPart = ((NotesSlidePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is PresentationPart)
                    newPart = ((PresentationPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is PresentationPropertiesPart)
                    newPart = ((PresentationPropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is QuickAccessToolbarCustomizationsPart)
                    newPart = ((QuickAccessToolbarCustomizationsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is RibbonAndBackstageCustomizationsPart)
                    newPart = ((RibbonAndBackstageCustomizationsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is RibbonExtensibilityPart)
                    newPart = ((RibbonExtensibilityPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is SingleCellTablePart)
                    newPart = ((SingleCellTablePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is SlideCommentsPart)
                    newPart = ((SlideCommentsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is SlideLayoutPart)
                    newPart = ((SlideLayoutPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is SlideMasterPart)
                    newPart = ((SlideMasterPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is SlidePart)
                    newPart = ((SlidePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is SlideSyncDataPart)
                    newPart = ((SlideSyncDataPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is StyleDefinitionsPart)
                    newPart = ((StyleDefinitionsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is StylesPart)
                    newPart = ((StylesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is StylesWithEffectsPart)
                    newPart = ((StylesWithEffectsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is TableDefinitionPart)
                    newPart = ((TableDefinitionPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is TableStylesPart)
                    newPart = ((TableStylesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ThemeOverridePart)
                    newPart = ((ThemeOverridePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ThemePart)
                    newPart = ((ThemePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ThumbnailPart)
                    newPart = ((ThumbnailPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
#if !NET35
                else if (newContentPart is TimeLineCachePart)
                    newPart = ((TimeLineCachePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is TimeLinePart)
                    newPart = ((TimeLinePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
#endif
                else if (newContentPart is UserDefinedTagsPart)
                    newPart = ((UserDefinedTagsPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is VbaDataPart)
                    newPart = ((VbaDataPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is VbaProjectPart)
                    newPart = ((VbaProjectPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is ViewPropertiesPart)
                    newPart = ((ViewPropertiesPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is VmlDrawingPart)
                    newPart = ((VmlDrawingPart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);
                else if (newContentPart is XmlSignaturePart)
                    newPart = ((XmlSignaturePart)newContentPart).AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension);

                relId = newContentPart.GetIdOfPart(newPart);
                using (Stream sourceStream = oldPart.GetStream())
                {
                    newPart.FeedData(sourceStream);
                }
                extendedReference.Attribute(attributeName).Value = relId;
            }
            catch (ArgumentOutOfRangeException)
            {
                try
                {
                    ExternalRelationship er = oldContentPart.GetExternalRelationship(relId);
                    ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    extendedReference.Attribute(R.id).Value = newEr.Id;
                }
                catch (KeyNotFoundException)
                {
                    var fromPart = newContentPart.OpenXmlPackage.Package.GetParts().FirstOrDefault(p => p.Uri == newContentPart.Uri);
                    fromPart.CreateRelationship(new Uri("NULL", UriKind.RelativeOrAbsolute), System.IO.Packaging.TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", relId);
                }
            }
        }
        public static void CopyFontTable(this FontTablePart oldFontTablePart, FontTablePart newFontTablePart)
        {
            var relevantElements = oldFontTablePart.GetXDocument().Descendants().Where(d => d.Name == W.embedRegular ||
                d.Name == W.embedBold || d.Name == W.embedItalic || d.Name == W.embedBoldItalic).ToList();
            foreach (XElement fontReference in relevantElements)
            {
                string relId = (string)fontReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var ipp1 = newFontTablePart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp1 != null)
                {
                    OpenXmlPart tempPart = ipp1.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr1 = newFontTablePart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr1 != null)
                    continue;

                var oldPart2 = oldFontTablePart.GetPartById(relId);
                if (oldPart2 == null || (!(oldPart2 is FontPart)))
                    throw new DocumentBuilderException("Invalid document - FontTablePart contains invalid relationship");

                FontPart oldPart = (FontPart)oldPart2;
                FontPart newPart = newFontTablePart.AddFontPart(oldPart.ContentType);
                var ResourceID = newFontTablePart.GetIdOfPart(newPart);
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
        public static void CopyInkPart(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement contentPartReference, XName attributeName)
        {
            string relId = (string)contentPartReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair != null)
                return;

            var tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (tempEr != null)
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            CustomXmlPart newPart = newContentPart.AddNewPart<CustomXmlPart>("application/inkml+xml", newId);

            using (var stream = oldPart.GetStream())
            {
                newPart.FeedData(stream);
            }
            contentPartReference.Attribute(attributeName).Value = newId;
        }
        public static void CopyLegacyDiagramText(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement textdataReference, XName attributeName)
        {
            string relId = (string)textdataReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair != null)
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            LegacyDiagramTextPart newPart = newContentPart.AddNewPart<LegacyDiagramTextPart>(newId);

            using (var stream = oldPart.GetStream())
            {
                newPart.FeedData(stream);
            }
            textdataReference.Attribute(attributeName).Value = newId;
        }
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null) return partXDocument;

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument();
                    partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                        partXDocument = XDocument.Load(partXmlReader);
                }
            }

            part.AddAnnotation(partXDocument);
            return partXDocument;
        }
        public static XDocument GetXDocument(this OpenXmlPart part, out XmlNamespaceManager namespaceManager)
        {
            if (part == null) throw new ArgumentNullException("part");

            namespaceManager = part.Annotation<XmlNamespaceManager>();
            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
            {
                if (namespaceManager != null) return partXDocument;

                namespaceManager = GetManagerFromXDocument(partXDocument);
                part.AddAnnotation(namespaceManager);

                return partXDocument;
            }

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument();
                    partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");

                    part.AddAnnotation(partXDocument);

                    return partXDocument;
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                    {
                        partXDocument = XDocument.Load(partXmlReader);
                        namespaceManager = new XmlNamespaceManager(partXmlReader.NameTable);

                        part.AddAnnotation(partXDocument);
                        part.AddAnnotation(namespaceManager);

                        return partXDocument;
                    }
                }
            }
        }
        public static bool GetValidReferenceId(this OpenXmlPart part, XElement reference, XName attributeName, out string refId)
        {
            var id = (string)reference.Attribute(attributeName);
            refId = id;
            if (!string.IsNullOrEmpty(id))
            {
                ExternalRelationship extRel = part.ExternalRelationships.FirstOrDefault(er => er.Id == id);

                if (extRel != null)
                {
                    return true;
                }
            }
            return false;
        }
        public static void SetExternalReferenceId(this OpenXmlPart oldPart, OpenXmlPart newPart, XElement element, string refId)
        {
            ExternalRelationship extRel = oldPart.ExternalRelationships.FirstOrDefault(r => r.Id == refId);
            if (extRel != null)
            {
                ExternalRelationship newExtRel = newPart.AddExternalRelationship(extRel.RelationshipType, extRel.Uri);
                element.Attribute(R.id).Value = newExtRel.Id;
            }
            else
            {
                var fromPart = newPart.OpenXmlPackage.Package.GetParts().FirstOrDefault(p => p.Uri == newPart.Uri);
                fromPart.CreateRelationship(new Uri("NULL", UriKind.RelativeOrAbsolute), System.IO.Packaging.TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", refId);
            }

        }
        public static void CopyRelatedImage(this OpenXmlPart oldPart, OpenXmlPart newPart, XElement imageReference, XName attributeName, IList<ImageData> images)
        {
            if(!GetValidReferenceId(oldPart, imageReference, attributeName, out string refId)) return;

            var part = oldPart.Parts.FirstOrDefault(ipp => ipp.RelationshipId == refId);
            if (part?.OpenXmlPart is ImagePart oldImagePart == true)
            {
                ImageData imageData = ManageImageCopy(oldImagePart, newPart, images);
                if (imageData.ImagePart == null) CopyImageData(imageData, newPart);
                else
                {
                    var refRel = newPart.Parts.FirstOrDefault(pip =>
                    {
                        var rel = imageData.ContentPartRelTypeIdList.FirstOrDefault(cpr =>
                        {
                            var found = cpr.ContentPart == newPart;
                            return found;
                        });
                        return rel != null;
                    });
                    if (refRel != null)
                    {
                        imageReference.Attribute(attributeName).Value = imageData.ContentPartRelTypeIdList.First(cpr =>
                        {
                            var found = cpr.ContentPart == newPart;
                            return found;
                        }).RelationshipId;
                        return;
                    }
                    var g = new Guid();
                    var newId = $"R{g:N}".Substring(0, 16);
                    newPart.CreateRelationshipToPart(imageData.ImagePart, newId);
                    imageReference.Attribute(R.id).Value = newId;

                }
                //var refRel = newPart.DataPartReferenceRelationships.FirstOrDefault(rr =>
                //{
                //    var rel = temp.ContentPartRelTypeIdList.FirstOrDefault(cpr =>
                //    {
                //        var found = cpr.ContentPart == newPart && cpr.RelationshipId == rr.Id;
                //        return found;
                //    });
                //    if (rel != null)
                //        return true;
                //    return false;
                //});
                //if (refRel != null)
                //{
                //    imageReference.Attribute(attributeName).Value = temp.ContentPartRelTypeIdList.First((Func<ContentPartRelTypeIdTuple, bool>)(cpr =>
                //    {
                //        var found = cpr.ContentPart == newPart && cpr.RelationshipId == refRel.Id;
                //        return found;
                //    })).RelationshipId;
                //    return;
                //}

                //var cpr2 = temp.ContentPartRelTypeIdList.FirstOrDefault(c => c.ContentPart == newPart);
                //if (cpr2 != null)
                //{
                //    imageReference.Attribute(attributeName).Value = cpr2.RelationshipId;
                //}
                //else
                //{
                //    ImagePart imagePart = (ImagePart)temp.ImagePart;
                //    var existingImagePart = newPart.AddPart<ImagePart>(imagePart);
                //    var newId = newPart.GetIdOfPart(existingImagePart);
                //    temp.AddContentPartRelTypeResourceIdTupple(newPart, imagePart.RelationshipType, newId);
                //    imageReference.Attribute(attributeName).Value = newId;
                //}
            }
            else SetExternalReferenceId(oldPart, newPart, imageReference, refId);

            void CopyImageData(ImageData target, OpenXmlPart source)
            {
                ImagePart newImagePart = CastAndAddImagePart(source);
                target.ImagePart = newPart;
                var partId = newPart.GetIdOfPart(newImagePart);
                target.AddContentPartRelTypeResourceIdTupple(newImagePart, newPart.RelationshipType, partId);
                imageReference.Attribute(attributeName).Value = partId;
                target.WriteImage(newImagePart);
            }
            ImagePart CastAndAddImagePart(OpenXmlPart source)
            {
                ImagePart imagePart = null;
                if (source is MainDocumentPart)
                    imagePart = ((MainDocumentPart)source).AddImagePart(oldPart.ContentType);
                if (source is HeaderPart)
                    imagePart = ((HeaderPart)source).AddImagePart(oldPart.ContentType);
                if (source is FooterPart)
                    imagePart = ((FooterPart)source).AddImagePart(oldPart.ContentType);
                if (source is EndnotesPart)
                    imagePart = ((EndnotesPart)source).AddImagePart(oldPart.ContentType);
                if (source is FootnotesPart)
                    imagePart = ((FootnotesPart)source).AddImagePart(oldPart.ContentType);
                if (source is ThemePart)
                    imagePart = ((ThemePart)source).AddImagePart(oldPart.ContentType);
                if (source is WordprocessingCommentsPart)
                    imagePart = ((WordprocessingCommentsPart)source).AddImagePart(oldPart.ContentType);
                if (source is DocumentSettingsPart)
                    imagePart = ((DocumentSettingsPart)source).AddImagePart(oldPart.ContentType);
                if (source is ChartPart)
                    imagePart = ((ChartPart)source).AddImagePart(oldPart.ContentType);
                if (source is NumberingDefinitionsPart)
                    imagePart = ((NumberingDefinitionsPart)source).AddImagePart(oldPart.ContentType);
                if (source is DiagramDataPart)
                    imagePart = ((DiagramDataPart)source).AddImagePart(oldPart.ContentType);
                if (source is ChartDrawingPart)
                    imagePart = ((ChartDrawingPart)source).AddImagePart(oldPart.ContentType);
                return imagePart;
            }
        }
        public static void CopyRelatedMedia(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName, IList<MediaData> mediaList, string mediaRelationshipType)
        {
            string relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            // First look to see if this relId has already been added to the new document.
            var existingDataPartRefRel2 = newContentPart.DataPartReferenceRelationships.FirstOrDefault(dpr => dpr.Id == relId);
            if (existingDataPartRefRel2 != null)
                return;

            var oldRel = oldContentPart.DataPartReferenceRelationships.FirstOrDefault(dpr => dpr.Id == relId);
            if (oldRel == null)
                return;

            DataPart oldPart = oldRel.DataPart;
            MediaData temp = oldPart.ManageMediaCopy(mediaList);
            if (temp.DataPart == null)
            {
                var ct = oldPart.ContentType;
                var ext = Path.GetExtension(oldPart.Uri.OriginalString);
                MediaDataPart newPart = newContentPart.OpenXmlPackage.CreateMediaDataPart(ct, ext);
                using (var stream = oldPart.GetStream())
                {
                    newPart.FeedData(stream);
                }
                string id = null;
                string relationshipType = null;

                if (mediaRelationshipType == "media")
                {
                    MediaReferenceRelationship mrr = null;

                    if (newContentPart is SlidePart)
                        mrr = ((SlidePart)newContentPart).AddMediaReferenceRelationship(newPart);
                    else if (newContentPart is SlideLayoutPart)
                        mrr = ((SlideLayoutPart)newContentPart).AddMediaReferenceRelationship(newPart);
                    else if (newContentPart is SlideMasterPart)
                        mrr = ((SlideMasterPart)newContentPart).AddMediaReferenceRelationship(newPart);

                    id = mrr.Id;
                    relationshipType = "http://schemas.microsoft.com/office/2007/relationships/media";
                }
                else if (mediaRelationshipType == "video")
                {
                    VideoReferenceRelationship vrr = null;

                    if (newContentPart is SlidePart)
                        vrr = ((SlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is HandoutMasterPart)
                        vrr = ((HandoutMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is NotesMasterPart)
                        vrr = ((NotesMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is NotesSlidePart)
                        vrr = ((NotesSlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is SlideLayoutPart)
                        vrr = ((SlideLayoutPart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is SlideMasterPart)
                        vrr = ((SlideMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);

                    id = vrr.Id;
                    relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
                }
                temp.DataPart = newPart;
                temp.AddContentPartRelTypeResourceIdTupple(newContentPart, relationshipType, id);
                imageReference.Attribute(attributeName).Value = id;
            }
            else
            {
                string desiredRelType = null;
                if (mediaRelationshipType == "media")
                    desiredRelType = "http://schemas.microsoft.com/office/2007/relationships/media";
                if (mediaRelationshipType == "video")
                    desiredRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
                var existingRel = temp.ContentPartRelTypeIdList.FirstOrDefault(cp => cp.ContentPart == newContentPart && cp.RelationshipType == desiredRelType);
                if (existingRel != null)
                {
                    imageReference.Attribute(attributeName).Value = existingRel.RelationshipId;
                }
                else
                {
                    MediaDataPart newPart = (MediaDataPart)temp.DataPart;
                    string id = null;
                    string relationshipType = null;
                    if (mediaRelationshipType == "media")
                    {
                        MediaReferenceRelationship mrr = null;

                        if (newContentPart is SlidePart)
                            mrr = ((SlidePart)newContentPart).AddMediaReferenceRelationship(newPart);
                        if (newContentPart is SlideLayoutPart)
                            mrr = ((SlideLayoutPart)newContentPart).AddMediaReferenceRelationship(newPart);
                        if (newContentPart is SlideMasterPart)
                            mrr = ((SlideMasterPart)newContentPart).AddMediaReferenceRelationship(newPart);

                        id = mrr.Id;
                        relationshipType = mrr.RelationshipType;
                    }
                    else if (mediaRelationshipType == "video")
                    {
                        VideoReferenceRelationship vrr = null;

                        if (newContentPart is SlidePart)
                            vrr = ((SlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is HandoutMasterPart)
                            vrr = ((HandoutMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is NotesMasterPart)
                            vrr = ((NotesMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is NotesSlidePart)
                            vrr = ((NotesSlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is SlideLayoutPart)
                            vrr = ((SlideLayoutPart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is SlideMasterPart)
                            vrr = ((SlideMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);

                        id = vrr.Id;
                        relationshipType = vrr.RelationshipType;
                    }
                    temp.AddContentPartRelTypeResourceIdTupple(newContentPart, relationshipType, id);
                    imageReference.Attribute(attributeName).Value = id;
                }
            }
        }
        public static void CopyRelatedMediaExternalRelationship(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName, string mediaRelationshipType)
        {
            string relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var existingExternalReference = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (existingExternalReference != null)
                return;

            var oldRel = oldContentPart.ExternalRelationships.FirstOrDefault(dpr => dpr.Id == relId);
            if (oldRel == null)
                return;

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            newContentPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newId);

            imageReference.Attribute(attributeName).Value = newId;
        }
        public static void CopyRelatedPartsForContentParts(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, Dictionary<XName, XName[]> relationshipMarkup, IEnumerable<XElement> newContent, IList<ImageData> images)
        {
            // this should be recursive
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == VML.imagedata || d.Name == VML.fill || d.Name == VML.stroke || d.Name == A.blip)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                oldContentPart.CopyRelatedImage(newContentPart, imageReference, R.embed, images);
                oldContentPart.CopyRelatedImage(newContentPart, imageReference, R.pict, images);
                oldContentPart.CopyRelatedImage(newContentPart, imageReference, R.id, images);
            }

            foreach (XElement diagramReference in newContent.DescendantsAndSelf().Where(d => d.Name == DGM.relIds || d.Name == A.relIds))
            {
                // dm attribute
                string relId = diagramReference.Attribute(R.dm).Value;
                var ipp = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp != null)
                {
                    OpenXmlPart tempPart = ipp.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er2 => er2.Id == relId);
                if (tempEr != null)
                    continue;

                OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                OpenXmlPart newPart = newContentPart.AddNewPart<DiagramDataPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.dm).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root });
                oldPart.CopyRelatedPartsForContentParts(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root }, images);

                // lo attribute
                relId = diagramReference.Attribute(R.lo).Value;
                var ipp2 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp2 != null)
                {
                    OpenXmlPart tempPart = ipp2.OpenXmlPart;
                    continue;
                }


                ExternalRelationship tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(er3 => er3.Id == relId);
                if (tempEr4 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramLayoutDefinitionPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.lo).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root });
                oldPart.CopyRelatedPartsForContentParts(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root }, images);

                // qs attribute
                relId = diagramReference.Attribute(R.qs).Value;
                var ipp5 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp5 != null)
                {
                    OpenXmlPart tempPart = ipp5.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr5 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramStylePart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.qs).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root });
                oldPart.CopyRelatedPartsForContentParts(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root }, images);

                // cs attribute
                relId = diagramReference.Attribute(R.cs).Value;
                var ipp6 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp6 != null)
                {
                    OpenXmlPart tempPart = ipp6.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr6 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr6 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramColorsPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.cs).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.AddRelationships(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root });
                oldPart.CopyRelatedPartsForContentParts(newPart, relationshipMarkup, new[] { newPart.GetXDocument().Root }, images);
            }

            foreach (XElement oleReference in newContent.DescendantsAndSelf(O.OLEObject))
            {
                string relId = (string)oleReference.Attribute(R.id);

                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                var ipp1 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp1 != null)
                {
                    OpenXmlPart tempPart = ipp1.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr1 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr1 != null)
                    continue;

                var ipp4 = oldContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp4 != null)
                {
                    OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                    OpenXmlPart newPart = null;
                    if (oldPart is EmbeddedObjectPart)
                    {
                        if (newContentPart is HeaderPart)
                            newPart = ((HeaderPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is FooterPart)
                            newPart = ((FooterPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is MainDocumentPart)
                            newPart = ((MainDocumentPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is FootnotesPart)
                            newPart = ((FootnotesPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is EndnotesPart)
                            newPart = ((EndnotesPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is WordprocessingCommentsPart)
                            newPart = ((WordprocessingCommentsPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                    }
                    else if (oldPart is EmbeddedPackagePart)
                    {
                        if (newContentPart is HeaderPart)
                            newPart = ((HeaderPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is FooterPart)
                            newPart = ((FooterPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is MainDocumentPart)
                            newPart = ((MainDocumentPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is FootnotesPart)
                            newPart = ((FootnotesPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is EndnotesPart)
                            newPart = ((EndnotesPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is WordprocessingCommentsPart)
                            newPart = ((WordprocessingCommentsPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is ChartPart)
                            newPart = ((ChartPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
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
                    if (relId != null)
                    {
                        ExternalRelationship er = oldContentPart.GetExternalRelationship(relId);
                        ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                        oleReference.Attribute(R.id).Value = newEr.Id;
                    }
                }
            }

            foreach (XElement chartReference in newContent.DescendantsAndSelf(C.chart))
            {
                string relId = (string)chartReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;
                var ipp2 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp2 != null)
                {
                    OpenXmlPart tempPart = ipp2.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr2 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr2 != null)
                    continue;

                var ipp3 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp3 == null)
                    continue;
                ChartPart oldPart = (ChartPart)ipp3.OpenXmlPart;
                XDocument oldChart = oldPart.GetXDocument();
                ChartPart newPart = newContentPart.AddNewPart<ChartPart>();
                XDocument newChart = newPart.GetXDocument();
                newChart.Add(oldChart.Root);
                chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                oldPart.CopyChartObjects(newPart);
                oldPart.CopyRelatedPartsForContentParts(newPart, relationshipMarkup, new[] { newChart.Root }, images);
            }

            foreach (XElement userShape in newContent.DescendantsAndSelf(C.userShapes))
            {
                string relId = (string)userShape.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var ipp4 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp4 != null)
                {
                    OpenXmlPart tempPart = ipp4.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr4 != null)
                    continue;

                var ipp5 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp5 != null)
                {
                    ChartDrawingPart oldPart = (ChartDrawingPart)ipp5.OpenXmlPart;
                    XDocument oldXDoc = oldPart.GetXDocument();
                    ChartDrawingPart newPart = newContentPart.AddNewPart<ChartDrawingPart>();
                    XDocument newXDoc = newPart.GetXDocument();
                    newXDoc.Add(oldXDoc.Root);
                    userShape.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                    oldPart.AddRelationships(newPart, relationshipMarkup, newContent);
                    oldPart.CopyRelatedPartsForContentParts(newPart, relationshipMarkup, new[] { newXDoc.Root }, images);
                }
            }
        }
        public static string GenStyleIdFromStyleName(this string styleName)
        {
            var newStyleId = styleName
                .Replace("_", "")
                .Replace("#", "")
                .Replace(".", "") + ((new Random()).Next(990) + 9).ToString();
            return newStyleId;
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
        public static ImageData ManageImageCopy(this ImagePart oldImage, OpenXmlPart newContentPart, IList<ImageData> images)
        {
            ImageData oldImageData = new ImageData(oldImage);
            foreach (ImageData item in images)
            {
                if (newContentPart != item.ImagePart) continue;
                if (item.Compare(oldImageData)) return item;
            }
            images.Add(oldImageData);
            return oldImageData;
        }
        public static MediaData ManageMediaCopy(this DataPart oldMedia, IList<MediaData> mediaList)
        {
            MediaData oldMediaData = new MediaData(oldMedia);
            foreach (MediaData item in mediaList)
            {
                if (item.Compare(oldMediaData))
                    return item;
            }
            mediaList.Add(oldMediaData);
            return oldMediaData;
        }
        public static void PutXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
#if true
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                    partXDocument.Save(partXmlWriter);
#else
                byte[] array = Encoding.UTF8.GetBytes(partXDocument.ToString(SaveOptions.DisableFormatting));
                using (MemoryStream ms = new MemoryStream(array))
                    part.FeedData(ms);
#endif
            }
        }
        public static void PutXDocumentWithFormatting(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                {
                    XmlWriterSettings settings = new XmlWriterSettings();
                    settings.Indent = true;
                    settings.OmitXmlDeclaration = true;
                    settings.NewLineOnAttributes = true;
                    using (XmlWriter partXmlWriter = XmlWriter.Create(partStream, settings))
                        partXDocument.Save(partXmlWriter);
                }
            }
        }
        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            if (part == null) throw new ArgumentNullException("part");
            if (document == null) throw new ArgumentNullException("document");

            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                document.Save(partXmlWriter);

            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }
        public static void UpdateStyleIdsForNumberingPart(this NumberingDefinitionsPart part, Dictionary<string, string> correctionList)
        {
            var numXDoc = part.GetXDocument();
            var numAttributeChangeList = correctionList.Select(cor => new
            {
                NewId = cor.Value,
                PStyleAttributesToChange = numXDoc.Descendants(W.pStyle).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                NumStyleLinkAttributesToChange = numXDoc.Descendants(W.numStyleLink).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
                StyleLinkAttributesToChange = numXDoc.Descendants(W.styleLink).Attributes(W.val).Where(a => a.Value == cor.Key).ToList(),
            }).ToList();

            foreach (var item in numAttributeChangeList)
            {
                foreach (var att in item.PStyleAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.NumStyleLinkAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.StyleLinkAttributesToChange)
                    att.Value = item.NewId;
            }
            part.PutXDocument();
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

        #region XDocument
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
        public static void ConvertNumberingPartToNewIds(this XDocument newNumbering, Dictionary<string, string> newIds)
        {
            foreach (var abstractNum in newNumbering
                .Root
                .Elements(W.abstractNum))
            {
                ConvertToNewId(abstractNum.Element(W.styleLink), newIds);
                ConvertToNewId(abstractNum.Element(W.numStyleLink), newIds);
            }

            foreach (var item in newNumbering
                .Descendants()
                .Where(d => d.Name == W.pStyle ||
                            d.Name == W.rStyle ||
                            d.Name == W.tblStyle))
            {
                ConvertToNewId(item, newIds);
            }
        }
        #endregion

        #region XElement
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
        public static void ConvertToNewId(this XElement element, Dictionary<string, string> newIds)
        {
            if (element == null)
                return;

            var valueAttribute = element.Attribute(W.val);
            string newId;
            if (newIds.TryGetValue(valueAttribute.Value, out newId))
            {
                valueAttribute.Value = newId;
            }
        }
        public static XElement FindReference(this XElement sect, XName reference, string type)
        {
            return sect.Elements(reference).FirstOrDefault(z =>
            {
                return (string)z.Attribute(W.type) == type;
            });
        }
        public static bool? GetBoolProp(XElement rPr, XName propertyName)
        {
            XElement propAtt = rPr.Element(propertyName);
            if (propAtt == null)
                return null;

            XAttribute val = propAtt.Attribute(W.val);
            if (val == null)
                return true;

            string s = ((string)val).ToLower();
            if (s == "1")
                return true;
            if (s == "0")
                return false;
            if (s == "true")
                return true;
            if (s == "false")
                return false;
            if (s == "on")
                return true;
            if (s == "off")
                return false;

            return (bool)propAtt.Attribute(W.val);
        }
        public static IEnumerable<XElement> LogicalChildrenContent(this XElement element)
        {
            if (element.Name == W.document)
                return element.Descendants(W.body).Take(1);

            if (element.Name == W.body ||
                element.Name == W.tc ||
                element.Name == W.txbxContent)
                return element
                    .DescendantsTrimmed(e =>
                        e.Name == W.tbl ||
                        e.Name == W.p)
                    .Where(e =>
                        e.Name == W.p ||
                        e.Name == W.tbl);

            if (element.Name == W.tbl)
                return element
                    .DescendantsTrimmed(W.tr)
                    .Where(e => e.Name == W.tr);

            if (element.Name == W.tr)
                return element
                    .DescendantsTrimmed(W.tc)
                    .Where(e => e.Name == W.tc);

            if (element.Name == W.p)
                return element
                    .DescendantsTrimmed(e => e.Name == W.r ||
                        e.Name == W.pict ||
                        e.Name == W.drawing)
                    .Where(e => e.Name == W.r ||
                        e.Name == W.pict ||
                        e.Name == W.drawing);

            if (element.Name == W.r)
                return element
                    .DescendantsTrimmed(e => W.SubRunLevelContent.Contains(e.Name))
                    .Where(e => W.SubRunLevelContent.Contains(e.Name));

            if (element.Name == MC.AlternateContent)
                return element
                    .DescendantsTrimmed(e =>
                        e.Name == W.pict ||
                        e.Name == W.drawing ||
                        e.Name == MC.Fallback)
                    .Where(e =>
                        e.Name == W.pict ||
                        e.Name == W.drawing);

            if (element.Name == W.pict || element.Name == W.drawing)
                return element
                    .DescendantsTrimmed(W.txbxContent)
                    .Where(e => e.Name == W.txbxContent);

            return XElement.EmptySequence;
        }
        public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source)
        {
            foreach (XElement e1 in source)
                foreach (XElement e2 in e1.LogicalChildrenContent())
                    yield return e2;
        }
        public static IEnumerable<XElement> LogicalChildrenContent(this XElement element, XName name)
        {
            return element.LogicalChildrenContent().Where(e => e.Name == name);
        }
        public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source, XName name)
        {
            foreach (XElement e1 in source)
                foreach (XElement e2 in e1.LogicalChildrenContent(name))
                    yield return e2;
        }
        public static void RemoveGfxdata(this IEnumerable<XElement> newContent)
        {
            newContent.DescendantsAndSelf().Attributes(O.gfxdata).Remove();
        }
        public static void UpdateContent(this IEnumerable<XElement> newContent, Dictionary<XName, XName[]> relationshipMarkup, XName elementToModify, string oldRid, string newRid)
        {
            foreach (var attributeName in relationshipMarkup[elementToModify])
            {
                var elementsToUpdate = newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid);
                foreach (var element in elementsToUpdate)
                    element.Attribute(attributeName).Value = newRid;
            }
        }
        public static void RemoveContent(this IEnumerable<XElement> newContent, Dictionary<XName, XName[]> relationshipMarkup, XName elementToModify,  string oldRid)
        {
            foreach (var attributeName in relationshipMarkup[elementToModify])
            {
                newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid).Remove();
            }
        }
        #endregion

        #region XNode
        public static object InsertTransform(this XNode node, IEnumerable<XElement> newContent)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Annotation<ReplaceSemaphore>() != null)
                    return newContent;
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => InsertTransform(n, newContent)));
            }
            return node;
        }
        #endregion


        public static XAttribute GetXmlSpaceAttribute(string value)
        {
            return (value.Length > 0) && ((value[0] == ' ') || (value[value.Length - 1] == ' '))
                ? new XAttribute(XNamespace.Xml + "space", "preserve")
                : null;
        }

        public static XAttribute GetXmlSpaceAttribute(char value)
        {
            return value == ' ' ? new XAttribute(XNamespace.Xml + "space", "preserve") : null;
        }

        private static XmlNamespaceManager GetManagerFromXDocument(XDocument xDocument)
        {
            XmlReader reader = xDocument.CreateReader();
            XDocument newXDoc = XDocument.Load(reader);

            XElement rootElement = xDocument.Elements().FirstOrDefault();
            rootElement.ReplaceWith(newXDoc.Root);

            XmlNameTable nameTable = reader.NameTable;
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(nameTable);
            return namespaceManager;
        }


        public static IEnumerable<OpenXmlPart> ContentParts(this WordprocessingDocument doc)
        {
            yield return doc.MainDocumentPart;

            foreach (var hdr in doc.MainDocumentPart.HeaderParts)
                yield return hdr;

            foreach (var ftr in doc.MainDocumentPart.FooterParts)
                yield return ftr;

            if (doc.MainDocumentPart.FootnotesPart != null)
                yield return doc.MainDocumentPart.FootnotesPart;

            if (doc.MainDocumentPart.EndnotesPart != null)
                yield return doc.MainDocumentPart.EndnotesPart;
        }

        /// <summary>
        /// Creates a complete list of all parts contained in the <see cref="OpenXmlPartContainer"/>.
        /// </summary>
        /// <param name="container">
        /// A <see cref="WordprocessingDocument"/>, <see cref="SpreadsheetDocument"/>, or
        /// <see cref="PresentationDocument"/>.
        /// </param>
        /// <returns>list of <see cref="OpenXmlPart"/>s contained in the <see cref="OpenXmlPartContainer"/>.</returns>
        public static List<OpenXmlPart> GetAllParts(this OpenXmlPartContainer container)
        {
            // Use a HashSet so that parts are processed only once.
            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();

            foreach (IdPartPair p in container.Parts)
                AddPart(partList, p.OpenXmlPart);

            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
        }

        private static void AddPart(HashSet<OpenXmlPart> partList, OpenXmlPart part)
        {
            if (partList.Contains(part)) return;

            partList.Add(part);
            foreach (IdPartPair p in part.Parts)
                AddPart(partList, p.OpenXmlPart);
        }

    }
}
