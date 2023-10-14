using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace OpenXmlPowerTools
{
    public class WordAutomationUtilities
    {
        public static void ProcessFilesUsingWordAutomation(List<string> fileNames)
        {
            Application app = new Application();
            app.Visible = false;
            foreach (string fileName in fileNames)
            {
                FileInfo fi = new FileInfo(fileName);
                try
                {
                    Document doc = app.Documents.Open(fi.FullName);
                    doc.Save();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Console.WriteLine("Caught unexpected COM exception.");
                    app.Quit();
                    Environment.Exit(0);
                }
            }
            app.Quit();
        }
    }
}
