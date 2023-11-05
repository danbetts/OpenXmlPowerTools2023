using System;
using System.IO;
using System.Diagnostics;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    public class WordRunner
    {
        public static void RunWord(FileInfo executablePath, FileInfo docxPath)
        {
            if (executablePath.Exists)
            {
                using (Process proc = new Process())
                {
                    proc.StartInfo.FileName = executablePath.FullName;
                    proc.StartInfo.Arguments = docxPath.FullName;
                    proc.StartInfo.WorkingDirectory = docxPath.DirectoryName;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.Start();
                }
            }
            else
            {
                throw new ArgumentException("Invalid executable path.", "executablePath");
            }
        }
    }
}