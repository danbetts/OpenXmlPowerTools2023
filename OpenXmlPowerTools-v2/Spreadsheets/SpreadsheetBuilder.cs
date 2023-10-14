using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using OpenXmlPowerTools.Presentations;

namespace OpenXmlPowerTools.Spreadsheets
{
    public class SpreadsheetBuilder
    {

        private string _fileName { get; set; } = string.Empty;
        public List<SlideSource> Sources { get; set; } = new List<SlideSource>();
        private int _start { get; set; } = 0;
        private int _count { get; set; } = 0;
        private HashSet<string> _customXmlGuidList { get; set; } = null;
        private bool _normalizeStyleIds { get; set; } = false;
        private static Dictionary<XName, XName[]> _relationshipMarkup { get; set; } = null;

        /// <summary>
        /// Set file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public SpreadsheetBuilder FileName(string fileName)
        {
            _fileName = fileName;
            return this;
        }

        /// <summary>
        /// Set sources to a collection of sources
        /// </summary>
        /// <param name="sources"></param>
        /// <returns></returns>
        public SpreadsheetBuilder SetSources(IEnumerable<SlideSource> sources)
        {
            Sources = sources.ToList();
            return this;
        }

        /// <summary>
        /// Add a new source
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public SpreadsheetBuilder AddSource(SlideSource source)
        {
            Sources.Add(source);
            return this;
        }

        /// <summary>
        /// Add a range of new sources to existing sources
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public SpreadsheetBuilder AppendSources(IEnumerable<SlideSource> source)
        {
            Sources.AddRange(source);
            return this;
        }

        /// <summary>
        /// Build and saves presentation document
        /// </summary>
        public static void Save()
        {

        }

        /// <summary>
        /// Build PresentationDocument
        /// </summary>
        /// <returns></returns>
        //public static PresentationDocument Build()
        //{
        //    return new PresentationDocument();
        //}

        public static void Write(string fileName, SpreadsheetWorkbook workbook)
        {
            try
            {
                if (fileName == null) throw new ArgumentNullException("fileName");
                if (workbook == null) throw new ArgumentNullException("workbook");

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                    fi.Delete();

                // create the blank workbook
                char[] base64CharArray = Spreadsheet.GetEmptySpreadsheet()
                    .Where(c => c != '\r' && c != '\n').ToArray();
                byte[] byteArray =
                    System.Convert.FromBase64CharArray(base64CharArray,
                    0, base64CharArray.Length);
                File.WriteAllBytes(fi.FullName, byteArray);

                // open the workbook, and create the TableProperties sheet, populate it
                using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(fi.FullName, true))
                {
                    WorkbookPart workbookPart = sDoc.WorkbookPart;
                    XDocument wXDoc = workbookPart.GetXDocument();
                    XElement sheetElement = wXDoc
                        .Root
                        .Elements(S.sheets)
                        .Elements(S.sheet)
                        .Where(s => (string)s.Attribute(SSNoNamespace.name) == "Sheet1")
                        .FirstOrDefault();
                    if (sheetElement == null)
                        throw new SpreadsheetBuilderInternalException();
                    string id = (string)sheetElement.Attribute(R.id);
                    sheetElement.Remove();
                    workbookPart.PutXDocument();

                    WorksheetPart sPart = (WorksheetPart)workbookPart.GetPartById(id);
                    workbookPart.DeletePart(sPart);

                    XDocument appXDoc = sDoc
                        .ExtendedFilePropertiesPart
                        .GetXDocument();
                    XElement vector = appXDoc
                        .Root
                        .Elements(EP.TitlesOfParts)
                        .Elements(VT.vector)
                        .FirstOrDefault();
                    if (vector != null)
                    {
                        vector.SetAttributeValue(SSNoNamespace.size, 0);
                        XElement lpstr = vector.Element(VT.lpstr);
                        lpstr.Remove();
                    }
                    XElement vector2 = appXDoc
                        .Root
                        .Elements(EP.HeadingPairs)
                        .Elements(VT.vector)
                        .FirstOrDefault();
                    XElement variant = vector2
                        .Descendants(VT.i4)
                        .FirstOrDefault();
                    if (variant != null)
                        variant.Value = "1";
                    sDoc.ExtendedFilePropertiesPart.PutXDocument();

                    if (workbook.Worksheets != null)
                        foreach (var worksheet in workbook.Worksheets)
                            sDoc.AddWorksheet(worksheet);

                    workbookPart.WorkbookStylesPart.PutXDocument();
                }
            }
            catch
            {
                throw;
            }
        }

        public static OpenXmlMemoryStreamDocument CreateSpreadsheetDocument()
        {
            MemoryStream stream = new MemoryStream();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                doc.AddWorkbookPart();
                doc.WorkbookPart.PutXDocument(Spreadsheet.CreateWorkbook());
                doc.Close();
                return new OpenXmlMemoryStreamDocument(stream);
            }
        }
    }
}
