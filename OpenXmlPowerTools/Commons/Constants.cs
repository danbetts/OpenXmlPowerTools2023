﻿using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Commons
{
    public static class Constants
    {
        public static XNamespace SpreadsheetNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public static XNamespace RelationshipsNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static string Yes = "yes";
        public static string Utf8 = "UTF-8";
        public static string OnePointZero = "1.0";
        public static XAttribute[] NamespaceAttributes =
        {
            new XAttribute(XNamespace.Xmlns + "wpc", WPC.wpc),
            new XAttribute(XNamespace.Xmlns + "mc", MC.mc),
            new XAttribute(XNamespace.Xmlns + "o", O.o),
            new XAttribute(XNamespace.Xmlns + "r", R.r),
            new XAttribute(XNamespace.Xmlns + "m", M.m),
            new XAttribute(XNamespace.Xmlns + "v", VML.vml),
            new XAttribute(XNamespace.Xmlns + "wp14", WP14.wp14),
            new XAttribute(XNamespace.Xmlns + "wp", WP.wp),
            new XAttribute(XNamespace.Xmlns + "w10", W10.w10),
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XAttribute(XNamespace.Xmlns + "w14", W14.w14),
            new XAttribute(XNamespace.Xmlns + "wpg", WPG.wpg),
            new XAttribute(XNamespace.Xmlns + "wpi", WPI.wpi),
            new XAttribute(XNamespace.Xmlns + "wne", WNE.wne),
            new XAttribute(XNamespace.Xmlns + "wps", WPS.wps),
            new XAttribute(MC.Ignorable, "w14 wp14"),
        };
        public static Dictionary<XName, XName[]> WordprocessingRelationshipMarkup => new Dictionary<XName, XName[]>()
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
        public static Dictionary<XName, XName[]> PresentationRelationshipMarkup = new Dictionary<XName, XName[]>()
        {
            { A.audioFile,        new [] { R.link }},
            { A.videoFile,        new [] { R.link }},
            { A.quickTimeFile,    new [] { R.link }},
            { A.wavAudioFile,     new [] { R.embed }},
            { A.blip,             new [] { R.embed, R.link }},
            { A.hlinkClick,       new [] { R.id }},
            { A.hlinkMouseOver,   new [] { R.id }},
            { A.hlinkHover,       new [] { R.id }},
            { A.relIds,           new [] { R.cs, R.dm, R.lo, R.qs }},
            { C.chart,            new [] { R.id }},
            { C.externalData,     new [] { R.id }},
            { C.userShapes,       new [] { R.id }},
            { DGM.relIds,         new [] { R.cs, R.dm, R.lo, R.qs }},
            { A14.imgLayer,       new [] { R.embed }},
            { P14.media,          new [] { R.embed, R.link }},
            { P.oleObj,           new [] { R.id }},
            { P.externalData,     new [] { R.id }},
            { P.control,          new [] { R.id }},
            { P.snd,              new [] { R.embed }},
            { P.sndTgt,           new [] { R.embed }},
            { PAV.srcMedia,       new [] { R.embed, R.link }},
            { P.contentPart,      new [] { R.id }},
            { VML.fill,           new [] { R.id }},
            { VML.imagedata,      new [] { R.href, R.id, R.pict, O.relid }},
            { VML.stroke,         new [] { R.id }},
            { WNE.toolbarData,    new [] { R.id }},
            { Plegacy.textdata,   new [] { XName.Get("id") }},
        };
        public static Dictionary<XName, int> PresentationOrder = new Dictionary<XName, int>
        {
            { P.sldMasterIdLst, 10 },
            { P.notesMasterIdLst, 20 },
            { P.handoutMasterIdLst, 30 },
            { P.sldIdLst, 40 },
            { P.sldSz, 50 },
            { P.notesSz, 60 },
            { P.embeddedFontLst, 70 },
            { P.custShowLst, 80 },
            { P.photoAlbum, 90 },
            { P.custDataLst, 100 },
            { P.kinsoku, 120 },
            { P.defaultTextStyle, 130 },
            { P.modifyVerifier, 150 },
            { P.extLst, 160 },
        };



        public static Dictionary<XName, int> OrderPPr = new Dictionary<XName, int>
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

        public static Dictionary<XName, int> OrderRPr = new Dictionary<XName, int>
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





        public static readonly List<XName> AdditionalRunContainerNames = new List<XName>
        {
            W.w + "bdo",
            W.customXml,
            W.dir,
            W.fldSimple,
            W.hyperlink,
            W.moveFrom,
            W.moveTo,
            W.sdtContent
        };
    }
}