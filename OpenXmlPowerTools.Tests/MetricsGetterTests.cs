﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Xml.Linq;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Documents;
using OpenXmlPowerTools.Presentations;
using OpenXmlPowerTools.Spreadsheets;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class MgTests
    {
        [Theory]
        [InlineData("Presentation.pptx")]
        [InlineData("Spreadsheet.xlsx")]
        [InlineData("DA001-TemplateDocument.docx")]
        [InlineData("DA002-TemplateDocument.docx")]
        [InlineData("DA003-Select-XPathFindsNoData.docx")]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx")]
        [InlineData("DA005-SelectRowData-NoData.docx")]
        [InlineData("DA006-SelectTestValue-NoData.docx")]
        public void MG001(string name)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo fi = new FileInfo(Path.Combine(sourceDir.FullName, name));

            MetricsGetterSettings settings = new MetricsGetterSettings()
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            var extension = fi.Extension.ToLower();
            XElement metrics = null;
            if (Wordprocessing.IsWordprocessing(extension))
            {
                WmlDocument wmlDocument = new WmlDocument(fi.FullName);
                metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);
            }
            else if (Spreadsheet.IsSpreadsheet(extension))
            {
                SmlDocument smlDocument = new SmlDocument(fi.FullName);
                metrics = MetricsGetter.GetXlsxMetrics(smlDocument, settings);
            }
            else if (Presentation.IsPresentation(extension))
            {
                PmlDocument pmlDocument = new PmlDocument(fi.FullName);
                metrics = MetricsGetter.GetPptxMetrics(pmlDocument, settings);
            }

            Assert.NotNull(metrics);
        }
    }
}

#endif
