
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Converters;
using OpenXmlPowerTools.Documents;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Converters
{
    [TestClass]
    public class HtmlToWmlConverterTests : ConverterTestsBase
    {
        protected override string FeatureFolder { get; } = @".\HtmlWmlConverter";
        private static bool s_ProduceAnnotatedHtml = true;
        private static string userCss = @"";
        private static string defaultCss =
    @"html, address,
blockquote,
body, dd, div,
dl, dt, fieldset, form,
frame, frameset,
h1, h2, h3, h4,
h5, h6, noframes,
ol, p, ul, center,
dir, hr, menu, pre { display: block; unicode-bidi: embed }
li { display: list-item }
head { display: none }
table { display: table }
tr { display: table-row }
thead { display: table-header-group }
tbody { display: table-row-group }
tfoot { display: table-footer-group }
col { display: table-column }
colgroup { display: table-column-group }
td, th { display: table-cell }
caption { display: table-caption }
th { font-weight: bolder; text-align: center }
caption { text-align: center }
body { margin: auto; }
h1 { font-size: 2em; margin: auto; }
h2 { font-size: 1.5em; margin: auto; }
h3 { font-size: 1.17em; margin: auto; }
h4, p,
blockquote, ul,
fieldset, form,
ol, dl, dir,
menu { margin: auto }
a { color: blue; }
h5 { font-size: .83em; margin: auto }
h6 { font-size: .75em; margin: auto }
h1, h2, h3, h4,
h5, h6, b,
strong { font-weight: bolder }
blockquote { margin-left: 40px; margin-right: 40px }
i, cite, em,
var, address { font-style: italic }
pre, tt, code,
kbd, samp { font-family: monospace }
pre { white-space: pre }
button, textarea,
input, select { display: inline-block }
big { font-size: 1.17em }
small, sub, sup { font-size: .83em }
sub { vertical-align: sub }
sup { vertical-align: super }
table { border-spacing: 2px; }
thead, tbody,
tfoot { vertical-align: middle }
td, th, tr { vertical-align: inherit }
s, strike, del { text-decoration: line-through }
hr { border: 1px inset }
ol, ul, dir,
menu, dd { margin-left: 40px }
ol { list-style-type: decimal }
ol ul, ul ol,
ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
u, ins { text-decoration: underline }
br:before { content: ""\A""; white-space: pre-line }
center { text-align: center }
:link, :visited { text-decoration: underline }
:focus { outline: thin dotted invert }
/* Begin bidirectionality settings (do not change) */
BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }

";

        [TestMethod]
        [DataRow("T0010.html")]
        [DataRow("T0011.html")]
        [DataRow("T0012.html")]
        [DataRow("T0013.html")]
        [DataRow("T0014.html")]
        [DataRow("T0015.html")]
        [DataRow("T0020.html")]
        [DataRow("T0030.html")]
        [DataRow("T0040.html")]
        [DataRow("T0050.html")]
        [DataRow("T0060.html")]
        [DataRow("T0070.html")]
        [DataRow("T0080.html")]
        [DataRow("T0090.html")]
        [DataRow("T0100.html")]
        [DataRow("T0110.html")]
        [DataRow("T0111.html")]
        [DataRow("T0112.html")]
        [DataRow("T0120.html")]
        [DataRow("T0130.html")]
        [DataRow("T0140.html")]
        [DataRow("T0150.html")]
        [DataRow("T0160.html")]
        [DataRow("T0170.html")]
        [DataRow("T0180.html")]
        [DataRow("T0190.html")]
        [DataRow("T0200.html")]
        [DataRow("T0210.html")]
        [DataRow("T0220.html")]
        [DataRow("T0230.html")]
        [DataRow("T0240.html")]
        [DataRow("T0250.html")]
        [DataRow("T0251.html")]
        [DataRow("T0260.html")]
        [DataRow("T0270.html")]
        [DataRow("T0280.html")]
        [DataRow("T0290.html")]
        [DataRow("T0300.html")]
        [DataRow("T0310.html")]
        [DataRow("T0320.html")]
        [DataRow("T0330.html")]
        [DataRow("T0340.html")]
        [DataRow("T0350.html")]
        [DataRow("T0360.html")]
        [DataRow("T0370.html")]
        [DataRow("T0380.html")]
        [DataRow("T0390.html")]
        [DataRow("T0400.html")]
        [DataRow("T0410.html")]
        [DataRow("T0420.html")]
        [DataRow("T0430.html")]
        [DataRow("T0431.html")]
        [DataRow("T0432.html")]
        [DataRow("T0440.html")]
        [DataRow("T0450.html")]
        [DataRow("T0460.html")]
        [DataRow("T0470.html")]
        [DataRow("T0480.html")]
        [DataRow("T0490.html")]
        [DataRow("T0500.html")]
        [DataRow("T0510.html")]
        [DataRow("T0520.html")]
        [DataRow("T0530.html")]
        [DataRow("T0540.html")]
        [DataRow("T0550.html")]
        [DataRow("T0560.html")]
        [DataRow("T0570.html")]
        [DataRow("T0580.html")]
        [DataRow("T0590.html")]
        [DataRow("T0600.html")]
        [DataRow("T0610.html")]
        [DataRow("T0620.html")]
        [DataRow("T0622.html")]
        [DataRow("T0630.html")]
        [DataRow("T0640.html")]
        [DataRow("T0650.html")]
        [DataRow("T0651.html")]
        [DataRow("T0660.html")]
        [DataRow("T0670.html")]
        [DataRow("T0680.html")]
        [DataRow("T0690.html")]
        [DataRow("T0691.html")]
        [DataRow("T0692.html")]
        [DataRow("T0700.html")]
        [DataRow("T0710.html")]
        [DataRow("T0720.html")]
        [DataRow("T0730.html")]
        [DataRow("T0740.html")]
        [DataRow("T0742.html")]
        [DataRow("T0745.html")]
        [DataRow("T0750.html")]
        [DataRow("T0760.html")]
        [DataRow("T0770.html")]
        [DataRow("T0780.html")]
        [DataRow("T0790.html")]
        [DataRow("T0791.html")]
        [DataRow("T0792.html")]
        [DataRow("T0793.html")]
        [DataRow("T0794.html")]
        [DataRow("T0795.html")]
        [DataRow("T0802.html")]
        [DataRow("T0804.html")]
        [DataRow("T0805.html")]
        [DataRow("T0810.html")]
        [DataRow("T0812.html")]
        [DataRow("T0814.html")]
        [DataRow("T0820.html")]
        [DataRow("T0821.html")]
        [DataRow("T0830.html")]
        [DataRow("T0840.html")]
        [DataRow("T0850.html")]
        [DataRow("T0851.html")]
        [DataRow("T0860.html")]
        [DataRow("T0870.html")]
        [DataRow("T0880.html")]
        [DataRow("T0890.html")]
        [DataRow("T0900.html")]
        [DataRow("T0910.html")]
        [DataRow("T0920.html")]
        [DataRow("T0921.html")]
        [DataRow("T0922.html")]
        [DataRow("T0923.html")]
        [DataRow("T0924.html")]
        [DataRow("T0925.html")]
        [DataRow("T0926.html")]
        [DataRow("T0927.html")]
        [DataRow("T0928.html")]
        [DataRow("T0929.html")]
        [DataRow("T0930.html")]
        [DataRow("T0931.html")]
        [DataRow("T0932.html")]
        [DataRow("T0933.html")]
        [DataRow("T0934.html")]
        [DataRow("T0935.html")]
        [DataRow("T0936.html")]
        [DataRow("T0940.html")]
        [DataRow("T0945.html")]
        [DataRow("T0948.html")]
        [DataRow("T0950.html")]
        [DataRow("T0955.html")]
        [DataRow("T0960.html")]
        [DataRow("T0968.html")]
        [DataRow("T0970.html")]
        [DataRow("T0980.html")]
        [DataRow("T0990.html")]
        [DataRow("T1000.html")]
        [DataRow("T1010.html")]
        [DataRow("T1020.html")]
        [DataRow("T1030.html")]
        [DataRow("T1040.html")]
        [DataRow("T1050.html")]
        [DataRow("T1060.html")]
        [DataRow("T1070.html")]
        [DataRow("T1080.html")]
        [DataRow("T1100.html")]
        [DataRow("T1110.html")]
        [DataRow("T1111.html")]
        [DataRow("T1112.html")]
        [DataRow("T1120.html")]
        [DataRow("T1130.html")]
        [DataRow("T1131.html")]
        [DataRow("T1132.html")]
        [DataRow("T1140.html")]
        [DataRow("T1150.html")]
        [DataRow("T1160.html")]
        [DataRow("T1170.html")]
        [DataRow("T1180.html")]
        [DataRow("T1190.html")]
        [DataRow("T1200.html")]
        [DataRow("T1201.html")]
        [DataRow("T1210.html")]
        [DataRow("T1220.html")]
        [DataRow("T1230.html")]
        [DataRow("T1240.html")]
        [DataRow("T1241.html")]
        [DataRow("T1242.html")]
        [DataRow("T1250.html")]
        [DataRow("T1251.html")]
        [DataRow("T1260.html")]
        [DataRow("T1270.html")]
        [DataRow("T1280.html")]
        [DataRow("T1290.html")]
        [DataRow("T1300.html")]
        [DataRow("T1310.html")]
        [DataRow("T1320.html")]
        [DataRow("T1330.html")]
        [DataRow("T1340.html")]
        [DataRow("T1350.html")]
        [DataRow("T1360.html")]
        [DataRow("T1370.html")]
        [DataRow("T1380.html")]
        [DataRow("T1390.html")]
        [DataRow("T1400.html")]
        [DataRow("T1410.html")]
        [DataRow("T1420.html")]
        [DataRow("T1430.html")]
        [DataRow("T1440.html")]
        [DataRow("T1450.html")]
        [DataRow("T1460.html")]
        [DataRow("T1470.html")]
        [DataRow("T1480.html")]
        [DataRow("T1490.html")]
        [DataRow("T1500.html")]
        [DataRow("T1510.html")]
        [DataRow("T1520.html")]
        [DataRow("T1530.html")]
        [DataRow("T1540.html")]
        [DataRow("T1550.html")]
        [DataRow("T1560.html")]
        [DataRow("T1570.html")]
        [DataRow("T1580.html")]
        [DataRow("T1590.html")]
        [DataRow("T1591.html")]
        [DataRow("T1610.html")]
        [DataRow("T1620.html")]
        [DataRow("T1630.html")]
        [DataRow("T1640.html")]
        [DataRow("T1650.html")]
        [DataRow("T1660.html")]
        [DataRow("T1670.html")]
        [DataRow("T1680.html")]
        [DataRow("T1690.html")]
        [DataRow("T1700.html")]
        [DataRow("T1710.html")]
        [DataRow("T1800.html")]
        [DataRow("T1810.html")]
        [DataRow("T1820.html")]
        [DataRow("T1830.html")]
        [DataRow("T1840.html")]
        [DataRow("T1850.html")]
        [DataRow("T1860.html")]
        [DataRow("T1870.html")]
        //[DataRow("E0010.html")]
        //[DataRow("E0020.html")]
        public void Html_Converts_To_Wml(string name)
        {
            string htmlFile = GetFile(name);

            var cssFile = Path.ChangeExtension(htmlFile, ".css");
            var docxFile = Path.ChangeExtension(htmlFile, ".docx");
            var annotatedFile = Path.ChangeExtension(htmlFile, ".txt");

            CleanupTest(cssFile);
            CleanupTest(docxFile);
            CleanupTest(annotatedFile);

            XElement html = HtmlToWmlReadAsXElement.ReadAsXElement(htmlFile);
            string usedAuthorCss = HtmlToWmlConverter.CleanUpCss((string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(cssFile, usedAuthorCss);

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();
            settings.BaseUriForImages = Path.Combine(TestUtil.TempDir.FullName);
            WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, usedAuthorCss, userCss, html, settings, null, s_ProduceAnnotatedHtml ? annotatedFile : null);
        }

        private static void SaveValidateAndFormatMainDocPart(string docxFile, WmlDocument doc)
        {
            WmlDocument formattedDoc;

            doc.SaveAs(docxFile);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
                using (WordprocessingDocument document = WordprocessingDocument.Open(ms, true))
                {
                    XDocument xDoc = document.MainDocumentPart.GetXDocument();
                    document.MainDocumentPart.PutXDocumentWithFormatting();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(document);
                    var errorsString = errors
                        .Select(e => e.Description + Environment.NewLine)
                        .StringConcatenate();

                    // Assert that there were no errors in the generated document.
                    errorsString.Should().BeEquivalentTo("");
                }
                formattedDoc = new WmlDocument(docxFile, ms.ToArray());
            }
            formattedDoc.SaveAs(docxFile);
        }
    }
}