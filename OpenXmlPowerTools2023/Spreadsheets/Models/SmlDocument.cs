using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools.Commons;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Spreadsheets
{
    public class SmlDocument : PowerToolsDocument
    {
        public SmlDocument(PowerToolsDocument original) : base(original)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument)) throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(PowerToolsDocument original, bool convertToTransitional) : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument)) throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName) : base(fileName)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument)) throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, bool convertToTransitional) : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument)) throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, byte[] byteArray) : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(SpreadsheetDocument)) throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, byte[] byteArray, bool convertToTransitional) : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(SpreadsheetDocument)) throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, MemoryStream memStream) : base(fileName, memStream)
        {
        }

        public SmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional) : base(fileName, memStream, convertToTransitional)
        {
        }

        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public XElement ConvertToHtml(SmlToHtmlConverterSettings htmlConverterSettings, string tableName)
        {
            return SmlToHtmlConverter.ConvertTableToHtml(this, htmlConverterSettings, tableName);
        }

        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public XElement ConvertTableToHtml(string tableName)
        {
            SmlToHtmlConverterSettings settings = new SmlToHtmlConverterSettings();
            return SmlToHtmlConverter.ConvertTableToHtml(this, settings, tableName);
        }
    }
}