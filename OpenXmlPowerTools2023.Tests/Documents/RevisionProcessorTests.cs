using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Documents;
using System;
using System.IO;
using System.Linq;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class RevisionProcessorTests : DocumentTestsBase
    {
        protected override string FeatureFolder { get; } = @".\RevisionProcessor";

        public static bool m_CopySourceFilesToTempDir = true;
        public static bool m_OpenTempDirInExplorer = false;

        [TestMethod]
        [DataRow("RP002-Deleted-Text.docx")]
        [DataRow("RP003-Inserted-Text.docx")]
        [DataRow("RP004-Deleted-Text-in-CC.docx")]
        [DataRow("RP005-Deleted-Paragraph-Mark.docx")]
        [DataRow("RP006-Inserted-Paragraph-Mark.docx")]
        [DataRow("RP007-Multiple-Deleted-Para-Mark.docx")]
        [DataRow("RP008-Multiple-Inserted-Para-Mark.docx")]
        [DataRow("RP009-Deleted-Table-Row.docx")]
        [DataRow("RP010-Inserted-Table-Row.docx")]
        [DataRow("RP011-Multiple-Deleted-Rows.docx")]
        [DataRow("RP012-Multiple-Inserted-Rows.docx")]
        [DataRow("RP013-Deleted-Math-Control-Char.docx")]
        [DataRow("RP014-Inserted-Math-Control-Char.docx")]
        [DataRow("RP015-MoveFrom-MoveTo.docx")]
        [DataRow("RP016-Deleted-CC.docx")]
        [DataRow("RP017-Inserted-CC.docx")]
        [DataRow("RP018-MoveFrom-MoveTo-CC.docx")]
        [DataRow("RP019-Deleted-Field-Code.docx")]
        [DataRow("RP020-Inserted-Field-Code.docx")]
        [DataRow("RP021-Inserted-Numbering-Properties.docx")]
        [DataRow("RP022-NumberingChange.docx")]
        [DataRow("RP023-NumberingChange.docx")]
        [DataRow("RP024-ParagraphMark-rPr-Change.docx")]
        [DataRow("RP025-Paragraph-Props-Change.docx")]
        [DataRow("RP026-NumberingChange.docx")]
        [DataRow("RP027-Change-Section.docx")]
        [DataRow("RP028-Table-Grid-Change.docx")]
        [DataRow("RP029-Table-Row-Props-Change.docx")]
        [DataRow("RP030-Table-Row-Props-Change.docx")]
        [DataRow("RP031-Table-Prop-Change.docx")]
        [DataRow("RP032-Table-Prop-Change.docx")]
        [DataRow("RP033-Table-Prop-Ex-Change.docx")]
        [DataRow("RP034-Deleted-Cells.docx")]
        [DataRow("RP035-Inserted-Cells.docx")]
        [DataRow("RP036-Vert-Merged-Cells.docx")]
        [DataRow("RP037-Changed-Style-Para-Props.docx")]
        [DataRow("RP038-Inserted-Paras-at-End.docx")]
        [DataRow("RP039-Inserted-Paras-at-End.docx")]
        [DataRow("RP040-Deleted-Paras-at-End.docx")]
        [DataRow("RP041-Cell-With-Empty-Paras-at-End.docx")]
        [DataRow("RP042-Deleted-Para-Mark-at-End.docx")]
        [DataRow("RP043-MERGEFORMAT-Field-Code.docx")]
        [DataRow("RP044-MERGEFORMAT-Field-Code.docx")]
        [DataRow("RP045-One-and-Half-Deleted-Lines-at-End.docx")]
        [DataRow("RP046-Consecutive-Deleted-Ranges.docx")]
        [DataRow("RP047-Inserted-and-Deleted-Paragraph-Mark.docx")]
        [DataRow("RP048-Deleted-Inserted-Para-Mark.docx")]
        [DataRow("RP049-Deleted-Para-Before-Table.docx")]
        [DataRow("RP050-Deleted-Footnote.docx")]
        [DataRow("RP052-Deleted-Para-Mark.docx")]
        public void RP001(string name)
        {
            var source = GetFile(name);

            WmlDocument sourceWml = new WmlDocument(source);
            WmlDocument afterRejectingWml = RevisionProcessor.RejectRevisions(sourceWml);
            WmlDocument afterAcceptingWml = RevisionProcessor.AcceptRevisions(sourceWml);

            var accepted = source.Replace(".docx", "-Accepted.docx");
            var rejected = source.Replace(".docx", "-Rejected.docx");
            afterAcceptingWml.SaveAs(accepted);
            afterRejectingWml.SaveAs(rejected);

            WmlComparerSettings settings = new WmlComparerSettings();

            Validate(accepted, afterAcceptingWml);
            Validate(rejected, afterRejectingWml);

            void Validate(string sourcePath, WmlDocument compareDoc)
            {
                var sourceDoc = new WmlDocument(sourcePath);
                WmlDocument resultDoc = WmlComparer.Compare(sourceDoc, compareDoc, settings);
                var revisions = WmlComparer.GetRevisions(resultDoc, settings);
                revisions.Should().BeEmpty(because: "Regression Error: Rejected baseline document did not match processed document");
            }
        }
    }
}