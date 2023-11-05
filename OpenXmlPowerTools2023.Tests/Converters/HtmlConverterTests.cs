using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Converters;
using OpenXmlPowerTools.Documents;
using OpenXmlPowerTools2023.Tests.Converters;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class HtmlConverterTests : ConverterTestsBase
    {
        protected override string FeatureFolder { get; } = @".\HtmlConverter";
        public static bool s_CopySourceFiles = true;
        public static bool s_CopyFormattingAssembledDocx = true;
        public static bool s_ConvertUsingWord = true;
        [TestMethod]
        [DataRow("HC001-5DayTourPlanTemplate.docx")]
        [DataRow("HC002-Hebrew-01.docx")]
        [DataRow("HC003-Hebrew-02.docx")]
        [DataRow("HC004-ResumeTemplate.docx")]
        [DataRow("HC005-TaskPlanTemplate.docx")]
        [DataRow("HC006-Test-01.docx")]
        [DataRow("HC007-Test-02.docx")]
        [DataRow("HC008-Test-03.docx")]
        [DataRow("HC009-Test-04.docx")]
        [DataRow("HC010-Test-05.docx")]
        [DataRow("HC011-Test-06.docx")]
        [DataRow("HC012-Test-07.docx")]
        [DataRow("HC013-Test-08.docx")]
        [DataRow("HC014-RTL-Table-01.docx")]
        [DataRow("HC015-Vertical-Spacing-atLeast.docx")]
        [DataRow("HC016-Horizontal-Spacing-firstLine.docx")]
        [DataRow("HC017-Vertical-Alignment-Cell-01.docx")]
        [DataRow("HC018-Vertical-Alignment-Para-01.docx")]
        [DataRow("HC019-Hidden-Run.docx")]
        [DataRow("HC020-Small-Caps.docx")]
        [DataRow("HC021-Symbols.docx")]
        [DataRow("HC022-Table-Of-Contents.docx")]
        [DataRow("HC023-Hyperlink.docx")]
        [DataRow("HC024-Tabs-01.docx")]
        [DataRow("HC025-Tabs-02.docx")]
        [DataRow("HC026-Tabs-03.docx")]
        [DataRow("HC027-Tabs-04.docx")]
        [DataRow("HC028-No-Break-Hyphen.docx")]
        [DataRow("HC029-Table-Merged-Cells.docx")]
        [DataRow("HC030-Content-Controls.docx")]
        [DataRow("HC031-Complicated-Document.docx")]
        [DataRow("HC032-Named-Color.docx")]
        [DataRow("HC033-Run-With-Border.docx")]
        [DataRow("HC034-Run-With-Position.docx")]
        [DataRow("HC035-Strike-Through.docx")]
        [DataRow("HC036-Super-Script.docx")]
        [DataRow("HC037-Sub-Script.docx")]
        [DataRow("HC038-Conflicting-Border-Weight.docx")]
        [DataRow("HC039-Bold.docx")]
        [DataRow("HC040-Hyperlink-Fieldcode-01.docx")]
        [DataRow("HC041-Hyperlink-Fieldcode-02.docx")]
        [DataRow("HC042-Image-Png.docx")]
        [DataRow("HC043-Chart.docx")]
        [DataRow("HC044-Embedded-Workbook.docx")]
        [DataRow("HC045-Italic.docx")]
        [DataRow("HC046-BoldAndItalic.docx")]
        [DataRow("HC047-No-Section.docx")]
        [DataRow("HC048-Excerpt.docx")]
        [DataRow("HC049-Borders.docx")]
        [DataRow("HC050-Shaded-Text-01.docx")]
        [DataRow("HC051-Shaded-Text-02.docx")]
        [DataRow("HC060-Image-with-Hyperlink.docx")]
        [DataRow("HC061-Hyperlink-in-Field.docx")]       
        public void HC001(string name)
        {
            var source = GetFile(name);

            var oxPtConvertedDestHtml = source.Replace(".docx", "-3-OxPt.html");
            ConvertToHtml(source, oxPtConvertedDestHtml);
        }

        [TestMethod]
        [DataRow("HC006-Test-01.docx")]
        public void HC002_NoCssClasses(string name)
        {
            var source = GetFile(name);
            var oxPtConvertedDestHtml =  source.Replace(".docx", "-5-OxPt-No-CSS-Classes.html");
            ConvertToHtmlNoCssClasses(source, oxPtConvertedDestHtml);
        }

        private static void CopyFormattingAssembledDocx(FileInfo source, FileInfo dest)
        {
            var ba = File.ReadAllBytes(source.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(ba, 0, ba.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true))
                {

                    RevisionAccepter.AcceptRevisions(wordDoc);
                    SimplifyMarkupSettings simplifyMarkupSettings = new SimplifyMarkupSettings
                    {
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveEndAndFootNotes = true,
                        RemoveFieldCodes = false,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveRsidInfo = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        RemoveGoBackBookmark = true,
                        ReplaceTabsWithSpaces = false,
                    };
                    MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);

                    FormattingAssemblerSettings formattingAssemblerSettings = new FormattingAssemblerSettings
                    {
                        RemoveStyleNamesFromParagraphAndRunProperties = false,
                        ClearStyles = false,
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        CreateHtmlConverterAnnotationAttributes = true,
                        OrderElementsPerStandard = false,
                        ListItemRetrieverSettings =
                            new ListItemRetrieverSettings()
                            {
                                ListItemTextImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations,
                            },
                    };

                    FormattingAssembler.AssembleFormatting(wordDoc, formattingAssemblerSettings);
                }
                var newBa = ms.ToArray();
                File.WriteAllBytes(dest.FullName, newBa);
            }
        }

        private static void ConvertToHtml(string sourceDocx, string destFileName)
        {
            byte[] byteArray = File.ReadAllBytes(sourceDocx);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var imageDirectoryName = destFileName.Substring(0, destFileName.Length - 5) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                        pageTitle = sourceDocx;

                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert png to jpeg.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageFileName = imageDirectoryName + "/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(XHtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlCommon.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);
                    File.WriteAllText(destFileName, htmlString, Encoding.UTF8);
                }
            }
        }

        private static void ConvertToHtmlNoCssClasses(string source, string target)
        {
            byte[] byteArray = File.ReadAllBytes(source);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var imageDirectoryName = target.Substring(0, target.Length - 5) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                        pageTitle = source;

                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = false,
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert png to jpeg.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageFileName = imageDirectoryName + "/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(XHtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlCommon.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);
                    File.WriteAllText(target, htmlString, Encoding.UTF8);
                }
            }
        }
    }
}