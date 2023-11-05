using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Converters;
using OpenXmlPowerTools.Documents;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class WmlComparerTests : DocumentTestsBase
    {
        protected override string FeatureFolder { get; } = @".\WmlComparer";
        public static bool m_OpenWord = false;
        public static bool m_OpenTempDirInExplorer = false;
        public WmlComparerSettings settings = new WmlComparerSettings();


        [TestMethod]
        [DataRow(@"RC\RC001-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC001-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
              <RcInfo>
                <DocName>RC/RC001-After2.docx</DocName>
                <Color>LightPink</Color>
                <Revisor>From Fred</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow(@"RC\RC002-Image.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC002-Image-After1.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow(@"RC\RC002-Image-After1.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC002-Image.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow("WC/WC027-Twenty-Paras-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>WC/WC027-Twenty-Paras-After-1.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow("WC/WC027-Twenty-Paras-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>WC/WC027-Twenty-Paras-After-3.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow(@"RC\RC003-Multi-Paras.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC003-Multi-Paras-After.docx</DocName>
                <Color>LightBlue</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow(@"RC\RC004-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC004-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
              <RcInfo>
                <DocName>RC/RC004-After2.docx</DocName>
                <Color>LightPink</Color>
                <Revisor>From Fred</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow(@"RC\RC005-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC005-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow(@"RC\RC006-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC006-After1.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        [DataRow(@"RC\RC007-Endnotes-Before.docx",
            @"<Root>
              <RcInfo>
                <DocName>RC/RC007-Endnotes-After.docx</DocName>
                <Color>LightYellow</Color>
                <Revisor>From Bob</Revisor>
              </RcInfo>
            </Root>")]
        public void WC001_Consolidate(string filename, string revisedDocumentsXml)
        {
            var source = GetFile(filename);
            var revisedDocumentsXElement = XElement.Parse(revisedDocumentsXml);
            var revisedDocumentsArray = revisedDocumentsXElement
                .Elements()
                .Select(z =>
                {
                    var revisedDocx = GetFile(z.Element("DocName").Value);
                    var wml1 = new WmlDocument(source);
                    var wml2 = Wordprocessing.BreakLinkToTemplate(wml1);
                    wml2.SaveAs(revisedDocx);
                    return new WmlRevisedDocumentInfo()
                    {
                        RevisedDocument = new WmlDocument(revisedDocx),
                        Color = ColorParser.FromName(z.Element("Color")?.Value),
                        Revisor = z.Element("Revisor")?.Value,
                    };
                })
                .ToList();

            var consolidatedDocxName = source.Replace(".docx", "-Consolidated.docx");

            WmlDocument source1Wml = new WmlDocument(source);
            WmlDocument consolidatedWml = WmlComparer.Consolidate(source1Wml, revisedDocumentsArray, settings);
            var wml3 = Wordprocessing.BreakLinkToTemplate(consolidatedWml);
            wml3.SaveAs(consolidatedDocxName);
            Validate(consolidatedWml);
        }

        [TestMethod]
        [DataRow("CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx")]
        [DataRow("WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx")]
        [DataRow("WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx")]
        [DataRow("WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx")]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx")]
        [DataRow("WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx")]
        [DataRow("WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx")]
        [DataRow("WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx")]
        [DataRow("WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx")]
        [DataRow("WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx")]
        [DataRow("WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx")]
        [DataRow("WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx")]
        [DataRow("WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx")]
        [DataRow("WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx")]
        [DataRow("WC/WC011-Before.docx", "WC/WC011-After.docx")]
        [DataRow("WC/WC012-Math-Before.docx", "WC/WC012-Math-After.docx")]
        [DataRow("WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx")]
        [DataRow("WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx")]
        [DataRow("WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx")]
        [DataRow("WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx")]
        [DataRow("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx")]
        [DataRow("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx")]
        [DataRow("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx")]
        [DataRow("WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx")]
        [DataRow("WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx")]
        [DataRow("WC/WC017-Image.docx", "WC/WC017-Image-After.docx")]
        [DataRow("WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx")]
        [DataRow("WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx")]
        [DataRow("WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx")]
        [DataRow("WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx")]
        [DataRow("WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx")]
        [DataRow("WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx")]
        [DataRow("WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx")]
        [DataRow("WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx")]
        [DataRow("WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx")]
        [DataRow("WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx")]
        [DataRow("WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx")]
        [DataRow("WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx")]
        [DataRow("WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx")]
        [DataRow("WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx")]
        [DataRow("WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx")]
        [DataRow("WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx")]
        [DataRow("WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx")]
        [DataRow("WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx")]
        [DataRow("WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx")]
        [DataRow("WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx")]
        [DataRow("WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx")]
        [DataRow("WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx")]
        [DataRow("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx")]
        [DataRow("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx")]
        [DataRow("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx")]
        [DataRow("WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx")]
        [DataRow("WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx")]
        [DataRow("WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx")]
        [DataRow("WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx")]
        [DataRow("WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx")]
        [DataRow("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx")]
        [DataRow("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx")]
        [DataRow("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx")]
        [DataRow("WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx")]
        [DataRow("WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx")]
        [DataRow("WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx")]
        [DataRow("WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx")]
        [DataRow("WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx")]
        [DataRow("WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx")]
        [DataRow("RC/RC001-Before.docx", "RC/RC001-After1.docx")]
        [DataRow("RC/RC002-Image.docx", "RC/RC002-Image-After1.docx")]
        public void WC002_Consolidate_Bulk_Test(string name1, string name2)
        {
            var source1 = GetFile(name1);
            var source2 = GetFile(name2);

            WmlDocument source1Wml = new WmlDocument(source1);
            WmlDocument source2Wml = new WmlDocument(source2);
            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);

            var output1 = source1.Replace(".docx", "-Revised.docx");
            var output2 = source1.Replace(".docx", "-Consolidated.docx");

            Wordprocessing.BreakLinkToTemplate(comparedWml).SaveAs(output1);

            List<WmlRevisedDocumentInfo> revisedDocInfo = new List<WmlRevisedDocumentInfo>()
            {
                new WmlRevisedDocumentInfo()
                {
                    RevisedDocument = source2Wml,
                    Color = Color.LightBlue,
                    Revisor = "Revised by Eric White",
                }
            };
            WmlDocument consolidatedWml = WmlComparer.Consolidate(source1Wml, revisedDocInfo, settings);
            Wordprocessing.BreakLinkToTemplate(consolidatedWml).SaveAs(output2);
            Validate(consolidatedWml);
        }

        [TestMethod]
        [DataRow("CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx", 1)]
        [DataRow("WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx", 4)]
        [DataRow("WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx", 1)]
        [DataRow("WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx", 1)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx", 2)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx", 2)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx", 1)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx", 1)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx", 1)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx", 1)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx", 1)]
        [DataRow("WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx", 1)]
        [DataRow("WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx", 1)]
        [DataRow("WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx", 1)]
        [DataRow("WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx", 1)]
        [DataRow("WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx", 2)]
        [DataRow("WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx", 2)]
        [DataRow("WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx", 1)]
        [DataRow("WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx", 2)]
        [DataRow("WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx", 1)]
        [DataRow("WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx", 3)]
        [DataRow("WC/WC011-Before.docx", "WC/WC011-After.docx", 2)]
        [DataRow("WC/WC012-Math-Before.docx", "WC/WC012-Math-After.docx", 2)]
        [DataRow("WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx", 2)]
        [DataRow("WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx", 2)]
        [DataRow("WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx", 2)]
        [DataRow("WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx", 2)]
        [DataRow("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx", 2)]
        [DataRow("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx", 3)]
        [DataRow("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx", 1)]
        [DataRow("WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx", 3)]
        [DataRow("WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx", 1)]
        [DataRow("WC/WC017-Image.docx", "WC/WC017-Image-After.docx", 3)]
        [DataRow("WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx", 2)]
        [DataRow("WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx", 3)]
        [DataRow("WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx", 3)]
        [DataRow("WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx", 5)]
        [DataRow("WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx", 3)]
        [DataRow("WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx", 5)]
        [DataRow("WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx", 9)]
        [DataRow("WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx", 6)]
        [DataRow("WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx", 10)]
        [DataRow("WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx", 7)]
        [DataRow("WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx", 1)]
        [DataRow("WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx", 7)]
        [DataRow("WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx", 4)]
        [DataRow("WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx", 2)]
        //[DataRow("WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx", 2)]
        //[DataRow("WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx", 2)]
        [DataRow("WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx", 4)]
        [DataRow("WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx", 2)]
        [DataRow("WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx", 4)]
        [DataRow("WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx", 3)]
        [DataRow("WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx", 2)]
        [DataRow("WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx", 4)]
        [DataRow("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx", 1)]
        [DataRow("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx", 4)]
        [DataRow("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx", 3)]
        [DataRow("WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx", 3)]
        [DataRow("WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx", 2)]
        [DataRow("WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx", 2)]
        [DataRow("WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx", 5)]
        [DataRow("WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx", 5)]
        [DataRow("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx", 1)]
        [DataRow("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx", 4)]
        [DataRow("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx", 7)]
        [DataRow("WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx", 7)]
        [DataRow("WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx", 2)]
        [DataRow("WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx", 2)]
        [DataRow("WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx", 6)]
        [DataRow("WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx", 6)]
        [DataRow("WC/WC037-Textbox-Before.docx", "WC/WC037-Textbox-After1.docx", 2)]
        [DataRow("WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx", 2)]
        //[DataRow("RC/RC001-Before.docx", "RC/RC001-After1.docx", 2)]
        //[DataRow("RC/RC002-Image.docx", "RC/RC002-Image-After1.docx", 1)]
        [DataRow("WC/WC039-Break-In-Row.docx", "WC/WC039-Break-In-Row-After1.docx", 1)]
        [DataRow("WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx", 2)]
        [DataRow("WC/WC042-Table-5.docx", "WC/WC042-Table-5-Mod.docx", 2)]
        [DataRow("WC/WC043-Nested-Table.docx", "WC/WC043-Nested-Table-Mod.docx", 2)]
        [DataRow("WC/WC044-Text-Box.docx", "WC/WC044-Text-Box-Mod.docx", 2)]
        [DataRow("WC/WC045-Text-Box.docx", "WC/WC045-Text-Box-Mod.docx", 2)]
        [DataRow("WC/WC046-Two-Text-Box.docx", "WC/WC046-Two-Text-Box-Mod.docx", 2)]
        [DataRow("WC/WC047-Two-Text-Box.docx", "WC/WC047-Two-Text-Box-Mod.docx", 2)]
        [DataRow("WC/WC048-Text-Box-in-Cell.docx", "WC/WC048-Text-Box-in-Cell-Mod.docx", 6)]
        [DataRow("WC/WC049-Text-Box-in-Cell.docx", "WC/WC049-Text-Box-in-Cell-Mod.docx", 5)]
        [DataRow("WC/WC050-Table-in-Text-Box.docx", "WC/WC050-Table-in-Text-Box-Mod.docx", 8)]
        [DataRow("WC/WC051-Table-in-Text-Box.docx", "WC/WC051-Table-in-Text-Box-Mod.docx", 9)]
        [DataRow("WC/WC052-SmartArt-Same.docx", "WC/WC052-SmartArt-Same-Mod.docx", 2)]
        [DataRow("WC/WC053-Text-in-Cell.docx", "WC/WC053-Text-in-Cell-Mod.docx", 2)]
        [DataRow("WC/WC054-Text-in-Cell.docx", "WC/WC054-Text-in-Cell-Mod.docx", 0)]
        [DataRow("WC/WC055-French.docx", "WC/WC055-French-Mod.docx", 0)]
        [DataRow("WC/WC056-French.docx", "WC/WC056-French-Mod.docx", 0)]
        [DataRow("WC/WC057-Table-Merged-Cell.docx", "WC/WC057-Table-Merged-Cell-Mod.docx", 4)]
        [DataRow("WC/WC058-Table-Merged-Cell.docx", "WC/WC058-Table-Merged-Cell-Mod.docx", 6)]
        [DataRow("WC/WC059-Footnote.docx", "WC/WC059-Footnote-Mod.docx", 5)]
        [DataRow("WC/WC060-Endnote.docx", "WC/WC060-Endnote-Mod.docx", 3)]
        [DataRow("WC/WC061-Style-Added.docx", "WC/WC061-Style-Added-Mod.docx", 1)]
        [DataRow("WC/WC062-New-Char-Style-Added.docx", "WC/WC062-New-Char-Style-Added-Mod.docx", 2)]
        [DataRow("WC/WC063-Footnote.docx", "WC/WC063-Footnote-Mod.docx", 1)]
        [DataRow("WC/WC063-Footnote-Mod.docx", "WC/WC063-Footnote.docx", 1)]
        [DataRow("WC/WC064-Footnote.docx", "WC/WC064-Footnote-Mod.docx", 0)]
        [DataRow("WC/WC065-Textbox.docx", "WC/WC065-Textbox-Mod.docx", 2)]
        [DataRow("WC/WC066-Textbox-Before-Ins.docx", "WC/WC066-Textbox-Before-Ins-Mod.docx", 1)]
        [DataRow("WC/WC066-Textbox-Before-Ins-Mod.docx", "WC/WC066-Textbox-Before-Ins.docx", 1)]
        [DataRow("WC/WC067-Textbox-Image.docx", "WC/WC067-Textbox-Image-Mod.docx", 2)]
        public void WC003_Compare(string name1, string name2, int revisionCount)
        {
            var source1 = GetFile(name1);
            var source2 = GetFile(name2);

            WmlDocument source1Wml = new WmlDocument(source1);
            WmlDocument source2Wml = new WmlDocument(source2);
            //settings.DebugTempFileDi = thisTestTempDir;
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            var output1 = source1.Replace(".docx", "-Revised.docx");
            comparedWml.SaveAs(output1);

            // validate generated document
            Validate(comparedWml);

            WmlComparerSettings settings2 = new WmlComparerSettings();

            WmlDocument revisionWml = new WmlDocument(output1);
            var revisions = WmlComparer.GetRevisions(revisionWml, settings);
            revisions.Count().Should().Be(revisionCount);

            var afterRejectingWml = RevisionProcessor.RejectRevisions(revisionWml);

            var WRITE_TEMP_FILES = true;

            if (WRITE_TEMP_FILES)
            {
                var afterRejectingFi = source1.Replace(".docx", "AfterRejecting.docx");
                afterRejectingWml.SaveAs(afterRejectingFi);
            }

            WmlDocument afterRejectingComparedWml = WmlComparer.Compare(source1Wml, afterRejectingWml, settings);
            var sanityCheck1 = WmlComparer.GetRevisions(afterRejectingComparedWml, settings);

            if (WRITE_TEMP_FILES)
            {
                var afterRejectingComparedFi = source1.Replace(".docx", "AfterRejectingCompared.docx");
                afterRejectingComparedWml.SaveAs(afterRejectingComparedFi);
            }

            var afterAcceptingWml = RevisionProcessor.AcceptRevisions(revisionWml);

            if (WRITE_TEMP_FILES)
            {
                var afterAcceptingFi = source1.Replace(".docx", "AfterAccepting.docx");
                afterAcceptingWml.SaveAs(afterAcceptingFi);
            }

            WmlDocument afterAcceptingComparedWml = WmlComparer.Compare(source2Wml, afterAcceptingWml, settings);
            var sanityCheck2 = WmlComparer.GetRevisions(afterAcceptingComparedWml, settings);

            if (WRITE_TEMP_FILES)
            {
                var afterAcceptingComparedFi = source2.Replace(".docx", "AfterAcceptingCompared.docx");
                afterAcceptingComparedWml.SaveAs(afterAcceptingComparedFi);
            }

            if (sanityCheck1.Count() != 0)
                Assert.Fail("Sanity Check #1 failed");
            if (sanityCheck2.Count() != 0)
                Assert.Fail("Sanity Check #2 failed");
        }

        [TestMethod]
        [DataRow("WC/WC001-Digits.docx")]
        [DataRow("WC/WC001-Digits-Deleted-Paragraph.docx")]
        [DataRow("WC/WC001-Digits-Mod.docx")]
        [DataRow("WC/WC002-DeleteAtBeginning.docx")]
        [DataRow("WC/WC002-DeleteAtEnd.docx")]
        [DataRow("WC/WC002-DeleteInMiddle.docx")]
        [DataRow("WC/WC002-DiffAtBeginning.docx")]
        [DataRow("WC/WC002-DiffInMiddle.docx")]
        [DataRow("WC/WC002-InsertAtBeginning.docx")]
        [DataRow("WC/WC002-InsertAtEnd.docx")]
        [DataRow("WC/WC002-InsertInMiddle.docx")]
        [DataRow("WC/WC002-Unmodified.docx")]
      //[DataRow("WC/WC004-Large.docx")]
      //[DataRow("WC/WC004-Large-Mod.docx")]
        [DataRow("WC/WC006-Table.docx")]
        [DataRow("WC/WC006-Table-Delete-Contests-of-Row.docx")]
        [DataRow("WC/WC006-Table-Delete-Row.docx")]
        [DataRow("WC/WC007-Deleted-at-Beginning-of-Para.docx")]
        [DataRow("WC/WC007-Longest-At-End.docx")]
        [DataRow("WC/WC007-Moved-into-Table.docx")]
        [DataRow("WC/WC007-Unmodified.docx")]
        [DataRow("WC/WC009-Table-Cell-1-1-Mod.docx")]
        [DataRow("WC/WC009-Table-Unmodified.docx")]
        [DataRow("WC/WC010-Para-Before-Table-Mod.docx")]
        [DataRow("WC/WC010-Para-Before-Table-Unmodified.docx")]
        [DataRow("WC/WC011-After.docx")]
        [DataRow("WC/WC011-Before.docx")]
        [DataRow("WC/WC012-Math-After.docx")]
        [DataRow("WC/WC012-Math-Before.docx")]
        [DataRow("WC/WC013-Image-After.docx")]
        [DataRow("WC/WC013-Image-After2.docx")]
        [DataRow("WC/WC013-Image-Before.docx")]
        [DataRow("WC/WC013-Image-Before2.docx")]
        [DataRow("WC/WC014-SmartArt-After.docx")]
        [DataRow("WC/WC014-SmartArt-Before.docx")]
        [DataRow("WC/WC014-SmartArt-With-Image-After.docx")]
        [DataRow("WC/WC014-SmartArt-With-Image-Before.docx")]
        [DataRow("WC/WC014-SmartArt-With-Image-Deleted-After.docx")]
        [DataRow("WC/WC014-SmartArt-With-Image-Deleted-After2.docx")]
        [DataRow("WC/WC015-Three-Paragraphs.docx")]
        [DataRow("WC/WC015-Three-Paragraphs-After.docx")]
        [DataRow("WC/WC016-Para-Image-Para.docx")]
        [DataRow("WC/WC016-Para-Image-Para-w-Deleted-Image.docx")]
        [DataRow("WC/WC017-Image.docx")]
        [DataRow("WC/WC017-Image-After.docx")]
        [DataRow("WC/WC018-Field-Simple-After-1.docx")]
        [DataRow("WC/WC018-Field-Simple-After-2.docx")]
        [DataRow("WC/WC018-Field-Simple-Before.docx")]
        [DataRow("WC/WC019-Hyperlink-After-1.docx")]
        [DataRow("WC/WC019-Hyperlink-After-2.docx")]
        [DataRow("WC/WC019-Hyperlink-Before.docx")]
        [DataRow("WC/WC020-FootNote-After-1.docx")]
        [DataRow("WC/WC020-FootNote-After-2.docx")]
        [DataRow("WC/WC020-FootNote-Before.docx")]
        [DataRow("WC/WC021-Math-After-1.docx")]
        [DataRow("WC/WC021-Math-Before-1.docx")]
        [DataRow("WC/WC022-Image-Math-Para-After.docx")]
        [DataRow("WC/WC022-Image-Math-Para-Before.docx")]
        public void WC004_Compare_To_Self(string name)
        {
            var source = GetFile(name);

            WmlDocument source1Wml = new WmlDocument(source);
            WmlDocument source2Wml = new WmlDocument(source);

            var compare1 = source.Replace(".docx", "-Compare1.docx");
            var compare2 = source.Replace(".docx", "-Compare2.docx");

            WmlDocument compare1Wml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            compare1Wml.SaveAs(compare1);
            Validate(compare1Wml);

            WmlDocument compare2Wml = WmlComparer.Compare(compare1Wml, source1Wml, settings);
            compare2Wml.SaveAs(compare2);
            Validate(compare2Wml);
        }

        [TestMethod]
        [DataRow("WCI-1000", "WC/WC040-Case-Before.docx", "WC/WC040-Case-After.docx", 2)]

        public void WC005_Compare_CaseInsensitive(string testId, string name1, string name2, int revisionCount)
        {
            var source1 = GetFile(name1);
            var source2 = GetFile(name2);

            WmlDocument source1Wml = new WmlDocument(source1);
            WmlDocument source2Wml = new WmlDocument(source2);
            settings.CaseInsensitive = true;
            settings.CultureInfo = System.Globalization.CultureInfo.CurrentCulture;
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            
            var compared = source1.Replace(".docx", "Compared.docx");
            comparedWml.SaveAs(compared);

            Validate(comparedWml);

            WmlDocument revisionWml = new WmlDocument(compared);
            var revisions = WmlComparer.GetRevisions(revisionWml, settings);
            revisions.Count().Should().Be(revisionCount);
        }

        [TestMethod]
        [DataRow("CZ/CZ001-Plain.docx", "CZ/CZ001-Plain-Mod.docx", 1)]
        [DataRow("CZ/CZ002-Multi-Paragraphs.docx", "CZ/CZ002-Multi-Paragraphs-Mod.docx", 1)]
        [DataRow("CZ/CZ003-Multi-Paragraphs.docx", "CZ/CZ003-Multi-Paragraphs-Mod.docx", 1)]
        [DataRow("CZ/CZ004-Multi-Paragraphs-in-Cell.docx", "CZ/CZ004-Multi-Paragraphs-in-Cell-Mod.docx", 1)]
        public void CZ001_CompareTrackedInPrev(string name1, string name2, int revisionCount)
        {
            // TODO: Do we need to keep the revision count parameter?
            revisionCount.Should().Be(1);

            var source1 = GetFile(name1);
            var source2 = GetFile(name2);

            WmlDocument source1Wml = new WmlDocument(source1);
            WmlDocument source2Wml = new WmlDocument(source2);
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);

            var copmared1 = source1.Replace(".docx", "Compared.docx");
            comparedWml.SaveAs(copmared1);
            Validate(comparedWml);
        }

        private static void Validate(WmlDocument doc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                    if (errors.Count() > 0)
                    {

                        var ind = "  ";
                        var sb = new StringBuilder();
                        foreach (var err in errors)
                        {
#if true
                            sb.Append("Error" + Environment.NewLine);
                            sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                            sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                            sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                            sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
#else
                        sb.Append("            \"" + err.Description + "\"," + Environment.NewLine);
#endif
                        }
                        var sbs = sb.ToString();
                        sbs.Should().Be("");
                    }
                }
            }
        }
    }
}