using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools.Documents;
using OpenXmlPowerTools.Presentations;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Commons
{
    public static class Common
    {
        public static object GetPropValue(this object src, string propName)
        {
            return src.GetType().GetProperty(propName)?.GetValue(src, null) ?? null;
        }
        public static T GetPart<T>(this object src) where T : OpenXmlPart, IFixedContentTypePart
        {
            return (T)src.GetPropValue(typeof(T).Name);
        }

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
        public static void AddRelationships(this OpenXmlPart oldPart, OpenXmlPart newPart, IDictionary<XName, XName[]> relationshipMarkup, IEnumerable<XElement> newContent)
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
        public static void CopyRelatedImage(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName, IList<ImageData> images)
        {
            string relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            // First look to see if this relId has already been added to the new document.
            // This is necessary for those parts that get processed with both old and new ids, such as the comments
            // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
            // in that case.
            var tempPartIdPair5 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair5 != null)
                return;

            ExternalRelationship tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (tempEr5 != null)
                return;

            var ipp2 = oldContentPart.Parts.FirstOrDefault(ipp => ipp.RelationshipId == relId);
            if (ipp2 != null)
            {
                var oldPart2 = ipp2.OpenXmlPart;
                if (!(oldPart2 is ImagePart))
                    throw new DocumentBuilderException("Invalid document - target part is not ImagePart");

                ImagePart oldPart = (ImagePart)ipp2.OpenXmlPart;
                ImageData temp = ManageImageCopy(oldPart, newContentPart, images);
                if (temp.ImagePart == null)
                {
                    ImagePart newPart = null;
                    if (newContentPart is MainDocumentPart)
                        newPart = ((MainDocumentPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is HeaderPart)
                        newPart = ((HeaderPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is FooterPart)
                        newPart = ((FooterPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is EndnotesPart)
                        newPart = ((EndnotesPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is FootnotesPart)
                        newPart = ((FootnotesPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ThemePart)
                        newPart = ((ThemePart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is WordprocessingCommentsPart)
                        newPart = ((WordprocessingCommentsPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is DocumentSettingsPart)
                        newPart = ((DocumentSettingsPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ChartPart)
                        newPart = ((ChartPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is NumberingDefinitionsPart)
                        newPart = ((NumberingDefinitionsPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is DiagramDataPart)
                        newPart = ((DiagramDataPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ChartDrawingPart)
                        newPart = ((ChartDrawingPart)newContentPart).AddImagePart(oldPart.ContentType);
                    temp.ImagePart = newPart;
                    var id = newContentPart.GetIdOfPart(newPart);
                    temp.AddContentPartRelTypeResourceIdTupple(newContentPart, newPart.RelationshipType, id);
                    imageReference.Attribute(attributeName).Value = id;
                    temp.WriteImage(newPart);
                }
                else
                {
                    var refRel = newContentPart.Parts.FirstOrDefault(pip =>
                    {
                        var rel = temp.ContentPartRelTypeIdList.FirstOrDefault(cpr =>
                        {
                            var found = cpr.ContentPart == newContentPart;
                            return found;
                        });
                        return rel != null;
                    });
                    if (refRel != null)
                    {
                        imageReference.Attribute(attributeName).Value = temp.ContentPartRelTypeIdList.First(cpr =>
                        {
                            var found = cpr.ContentPart == newContentPart;
                            return found;
                        }).RelationshipId;
                        return;
                    }
                    var g = new Guid();
                    var newId = $"R{g:N}".Substring(0, 16);
                    newContentPart.CreateRelationshipToPart(temp.ImagePart, newId);
                    imageReference.Attribute(R.id).Value = newId;
                }
            }
            else
            {
                ExternalRelationship er = oldContentPart.ExternalRelationships.FirstOrDefault(er1 => er1.Id == relId);
                if (er != null)
                {
                    ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    imageReference.Attribute(R.id).Value = newEr.Id;
                }
                throw new DocumentBuilderInternalException("Source {0} is unsupported document - contains reference to NULL image");
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
        public static void CopyRelatedPartsForContentParts(this OpenXmlPart oldContentPart, OpenXmlPart newContentPart, IDictionary<XName, XName[]> relationshipMarkup, IEnumerable<XElement> newContent, IList<ImageData> images)
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
        #endregion

        #region XAttribute
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
        #endregion

        #region XDocument
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
        public static void UpdateContent(this IEnumerable<XElement> newContent, IDictionary<XName, XName[]> relationshipMarkup, XName elementToModify, string oldRid, string newRid)
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
        public static void RemoveContent(this IEnumerable<XElement> newContent, IDictionary<XName, XName[]> relationshipMarkup, XName elementToModify,  string oldRid)
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
