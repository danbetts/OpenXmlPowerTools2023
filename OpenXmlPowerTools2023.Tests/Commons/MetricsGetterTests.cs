using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Documents;
using OpenXmlPowerTools.Presentations;
using OpenXmlPowerTools.Spreadsheets;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Commons
{
    [TestClass]
    public class MetricsGetterTests : CommonTestsBase
    {
        protected override string ModuleFolder { get; } = "";
        protected override string FeatureFolder { get; } = "";

        [TestMethod]
        [DataRow(@"Presentations\Presentation.pptx")]
        [DataRow(@"Spreadsheets\Spreadsheet.xlsx")]
        [DataRow(@"Documents\DocumentAssembler\DA001-TemplateDocument.docx")]
        [DataRow(@"Documents\DocumentAssembler\DA002-TemplateDocument.docx")]
        [DataRow(@"Documents\DocumentAssembler\DA003-Select-XPathFindsNoData.docx")]
        [DataRow(@"Documents\DocumentAssembler\DA004-Select-XPathFindsNoDataOptional.docx")]
        [DataRow(@"Documents\DocumentAssembler\DA005-SelectRowData-NoData.docx")]
        [DataRow(@"Documents\DocumentAssembler\DA006-SelectTestValue-NoData.docx")]
        public void MG001(string name)
        {
            MetricsGetterSettings settings = new MetricsGetterSettings()
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            var filepath = GetFile(name);
            var extension = Path.GetExtension(filepath).ToLower();
            XElement metrics = null;
            if (Wordprocessing.IsWordprocessing(extension))
            {
                WmlDocument wmlDocument = new WmlDocument(filepath);
                metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);
            }
            else if (Spreadsheet.IsSpreadsheet(extension))
            {
                SmlDocument smlDocument = new SmlDocument(filepath);
                metrics = MetricsGetter.GetXlsxMetrics(smlDocument, settings);
            }
            else if (Presentation.IsPresentation(extension))
            {
                PmlDocument pmlDocument = new PmlDocument(filepath);
                metrics = MetricsGetter.GetPptxMetrics(pmlDocument, settings);
            }

            metrics.Should().NotBeNull();
        }
    }
}