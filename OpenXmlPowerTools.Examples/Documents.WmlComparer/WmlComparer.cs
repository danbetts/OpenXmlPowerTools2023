// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using OpenXmlPowerTools.Documents;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace OpenXmlPowerTools.Examples
{
    class WmlComparer
    {
        static void Main(string[] args)
        {
            Example01();
        }

        private static void Example01()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument result = Documents.WmlComparer.Compare(
                new WmlDocument("../../Source1.docx"),
                new WmlDocument("../../Source2.docx"),
                settings);
            result.SaveAs(Path.Combine(tempDi.FullName, "Compared.docx"));

            var revisions = Documents.WmlComparer.GetRevisions(result, settings);
            foreach (var rev in revisions)
            {
                Console.WriteLine("Author: " + rev.Author);
                Console.WriteLine("Revision type: " + rev.RevisionType);
                Console.WriteLine("Revision text: " + rev.Text);
                Console.WriteLine();
            }
        }

        private static void Example02()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WmlDocument originalWml = new WmlDocument("../../Original.docx");
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList = new List<WmlRevisedDocumentInfo>()
            {
                new WmlRevisedDocumentInfo()
                {
                    RevisedDocument = new WmlDocument("../../RevisedByBob.docx"),
                    Revisor = "Bob",
                    Color = Color.LightBlue,
                },
                new WmlRevisedDocumentInfo()
                {
                    RevisedDocument = new WmlDocument("../../RevisedByMary.docx"),
                    Revisor = "Mary",
                    Color = Color.LightYellow,
                },
            };
            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument consolidatedWml = Documents.WmlComparer.Consolidate(
                originalWml,
                revisedDocumentInfoList,
                settings);
            consolidatedWml.SaveAs(Path.Combine(tempDi.FullName, "Consolidated.docx"));
        }
    }
}
