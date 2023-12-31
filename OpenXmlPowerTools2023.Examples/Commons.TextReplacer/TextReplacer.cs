﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System;
using System.IO;

namespace OpenXmlPowerTools.Examples
{
    class TestPmlTextReplacer
    {
        static void Main(string[] args)
        {
            Example01();
            Example02();
        }
        private static void Example01()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            File.Copy("../../Test01.pptx", Path.Combine(tempDi.FullName, "Test01out.pptx"));
            using (PresentationDocument pDoc =
                PresentationDocument.Open(Path.Combine(tempDi.FullName, "Test01out.pptx"), true))
            {
                TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
            }
            File.Copy("../../Test02.pptx", Path.Combine(tempDi.FullName, "Test02out.pptx"));
            using (PresentationDocument pDoc =
                PresentationDocument.Open(Path.Combine(tempDi.FullName, "Test02out.pptx"), true))
            {
                TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
            }
            File.Copy("../../Test03.pptx", Path.Combine(tempDi.FullName, "Test03out.pptx"));
            using (PresentationDocument pDoc =
                PresentationDocument.Open(Path.Combine(tempDi.FullName, "Test03out.pptx"), true))
            {
                TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", false);
            }
        }

        private static void Example02()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            DirectoryInfo di2 = new DirectoryInfo("../../../");
            foreach (var file in di2.GetFiles("*.docx"))
                file.CopyTo(Path.Combine(tempDi.FullName, file.Name));

            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test01.docx"), true))
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test02.docx"), true))
                    TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception) { }
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test03.docx"), true))
                    TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception) { }
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test04.docx"), true))
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test05.docx"), true))
                TextReplacer.SearchAndReplace(doc, "is on", "is above", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test06.docx"), true))
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test07.docx"), true))
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test08.docx"), true))
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test09.docx"), true))
                TextReplacer.SearchAndReplace(doc, "===== Replace this text =====", "***zzz***", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Path.Combine(tempDi.FullName, "Test09.docx"), true))
                TextReplacer.SearchAndReplace(doc, "***zzz***", "", true);
        }
    }
}