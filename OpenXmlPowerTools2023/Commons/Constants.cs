using System.Collections.Generic;
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