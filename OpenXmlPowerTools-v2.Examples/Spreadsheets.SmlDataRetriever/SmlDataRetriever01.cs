// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using OpenXmlPowerTools.Spreadsheets;
using System;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Examples
{
    class SmlDataRetriever01
    {
        static void Main(string[] args)
        {
            FileInfo fi = null;
            fi = new FileInfo("../../SampleSpreadsheet.xlsx");

            // Retrieve range from Sheet1
            XElement data = SmlDataRetriever.RetrieveRange(fi.FullName, "Sheet1", "A1:C3");
            Console.WriteLine(data);

            // Retrieve entire sheet
            data = SmlDataRetriever.RetrieveSheet(fi.FullName, "Sheet1");
            Console.WriteLine(data);

            // Retrieve table
            data = SmlDataRetriever.RetrieveTable(fi.FullName, "VehicleTable");
            Console.WriteLine(data);
        }
    }
}
