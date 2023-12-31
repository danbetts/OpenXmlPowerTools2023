﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using OpenXmlPowerTools.Commons;
using System;
using System.IO;

namespace OpenXmlPowerTools.Examples
{
    class MetricsGetter01
    {
        static void Main(string[] args)
        {
            MetricsGetterSettings settings = null;
            FileInfo fi = null;

            fi = new FileInfo("../../ContentControls.docx");
            settings = new MetricsGetterSettings();
            settings.IncludeTextInContentControls = false;
            Console.WriteLine("============== No text from content controls ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../ContentControls.docx");
            settings = new MetricsGetterSettings();
            settings.IncludeTextInContentControls = true;
            Console.WriteLine("============== With text from content controls ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../TrackedRevisions.docx");
            settings = new MetricsGetterSettings();
            settings.IncludeTextInContentControls = true;
            Console.WriteLine("============== Tracked Revisions ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../Styles.docx");
            settings = new MetricsGetterSettings();
            settings.IncludeTextInContentControls = false;
            Console.WriteLine("============== Style Hierarchy ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();

            fi = new FileInfo("../../Tables.xlsx");
            settings = new MetricsGetterSettings();
            settings.IncludeTextInContentControls = false;
            settings.IncludeXlsxTableCellData = true;
            Console.WriteLine("============== Spreadsheet Tables ==============");
            Console.WriteLine(fi.FullName);
            Console.WriteLine(MetricsGetter.GetMetrics(fi.FullName, settings));
            Console.WriteLine();
        }
    }
}
