using OpenXmlPowerTools.Commons;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// Docx metrics class
    /// </summary>
    public class DocxMetrics
    {
        public string FileName;
        public int ActiveX;
        public int AltChunk;
        public int AsciiCharCount;
        public int AsciiRunCount;
        public int AverageParagraphLength;
        public int ComplexField;
        public int ContentControlCount;
        public XmlDocument ContentControls;
        public int CSCharCount;
        public int CSRunCount;
        public bool DocumentProtection;
        public int EastAsiaCharCount;
        public int EastAsiaRunCount;
        public int ElementCount;
        public bool EmbeddedXlsx;
        public int HAnsiCharCount;
        public int HAnsiRunCount;
        public int Hyperlink;
        public bool InvalidSaveThroughXslt;
        public string Languages;
        public int LegacyFrame;
        public int MultiFontRun;
        public string NumberingFormatList;
        public int ReferenceToNullImage;
        public bool RevisionTracking;
        public int RunCount;
        public int SimpleField;
        public XmlDocument StyleHierarchy;
        public int SubDocument;
        public int Table;
        public int TextBox;
        public bool TrackRevisionsEnabled;
        public bool Valid;
        public int ZeroLengthText;

        public DocxMetrics(string fileName)
        {
            SetMetrics(fileName);
        }

        /// <summary>
        /// Set up metrics
        /// </summary>
        /// <param name="fileName"></param>
        public void SetMetrics(string fileName)
        {
            WmlDocument wmlDoc = new WmlDocument(fileName);
            MetricsGetterSettings settings = new MetricsGetterSettings();
            settings.IncludeTextInContentControls = false;
            settings.IncludeXlsxTableCellData = false;
            var metricsXml = MetricsGetter.GetDocxMetrics(wmlDoc, settings);
            FileName = wmlDoc.FileName;
            StyleHierarchy = GetXmlDocumentForMetrics(metricsXml, H.StyleHierarchy);
            ContentControls = GetXmlDocumentForMetrics(metricsXml, H.Parts);
            TextBox = GetIntForMetrics(metricsXml, H.TextBox);
            ContentControlCount = GetIntForMetrics(metricsXml, H.ContentControl);
            ComplexField = GetIntForMetrics(metricsXml, H.ComplexField);
            SimpleField = GetIntForMetrics(metricsXml, H.SimpleField);
            AltChunk = GetIntForMetrics(metricsXml, H.AltChunk);
            Table = GetIntForMetrics(metricsXml, H.Table);
            Hyperlink = GetIntForMetrics(metricsXml, H.Hyperlink);
            LegacyFrame = GetIntForMetrics(metricsXml, H.LegacyFrame);
            ActiveX = GetIntForMetrics(metricsXml, H.ActiveX);
            SubDocument = GetIntForMetrics(metricsXml, H.SubDocument);
            ReferenceToNullImage = GetIntForMetrics(metricsXml, H.ReferenceToNullImage);
            ElementCount = GetIntForMetrics(metricsXml, H.ElementCount);
            AverageParagraphLength = GetIntForMetrics(metricsXml, H.AverageParagraphLength);
            RunCount = GetIntForMetrics(metricsXml, H.RunCount);
            ZeroLengthText = GetIntForMetrics(metricsXml, H.ZeroLengthText);
            MultiFontRun = GetIntForMetrics(metricsXml, H.MultiFontRun);
            AsciiCharCount = GetIntForMetrics(metricsXml, H.AsciiCharCount);
            CSCharCount = GetIntForMetrics(metricsXml, H.CSCharCount);
            EastAsiaCharCount = GetIntForMetrics(metricsXml, H.EastAsiaCharCount);
            HAnsiCharCount = GetIntForMetrics(metricsXml, H.HAnsiCharCount);
            AsciiRunCount = GetIntForMetrics(metricsXml, H.AsciiRunCount);
            CSRunCount = GetIntForMetrics(metricsXml, H.CSRunCount);
            EastAsiaRunCount = GetIntForMetrics(metricsXml, H.EastAsiaRunCount);
            HAnsiRunCount = GetIntForMetrics(metricsXml, H.HAnsiRunCount);
            RevisionTracking = GetBoolForMetrics(metricsXml, H.RevisionTracking);
            EmbeddedXlsx = GetBoolForMetrics(metricsXml, H.EmbeddedXlsx);
            InvalidSaveThroughXslt = GetBoolForMetrics(metricsXml, H.InvalidSaveThroughXslt);
            TrackRevisionsEnabled = GetBoolForMetrics(metricsXml, H.TrackRevisionsEnabled);
            DocumentProtection = GetBoolForMetrics(metricsXml, H.DocumentProtection);
            Valid = GetBoolForMetrics(metricsXml, H.Valid);
            Languages = GetStringForMetrics(metricsXml, H.Languages);
            NumberingFormatList = GetStringForMetrics(metricsXml, H.NumberingFormatList);

            string GetStringForMetrics(XElement xElement, XName xName)
            {
                var ele = xElement.Element(xName);
                if (ele == null)
                    return "";
                return (string)ele.Attribute(H.Val);
            }

            bool GetBoolForMetrics(XElement xElement, XName xName)
            {
                var ele = xElement.Element(xName);
                if (ele == null)
                    return false;
                return (bool)ele.Attribute(H.Val);
            }

            int GetIntForMetrics(XElement xElement, XName xName)
            {
                var ele = xElement.Element(xName);
                if (ele == null)
                    return 0;
                return (int)ele.Attribute(H.Val);
            }

            XmlDocument GetXmlDocumentForMetrics(XElement xElement, XName xName)
            {
                var ele = xElement.Element(xName);
                if (ele == null)
                    return null;
                return (new XDocument(metricsXml.Element(xName))).GetXmlDocument();
            }

        }
    }
}