using System;

namespace OpenXmlPowerTools.Spreadsheets
{
    /// <summary>
    /// Spreadsheet internal exception
    /// </summary>
    public class SpreadsheetBuilderInternalException : Exception
    {
        public SpreadsheetBuilderInternalException() : base("Internal error - unexpected content in _EmptyXlsx.") { }
    }

    /// <summary>
    /// Invalid sheet name exception
    /// </summary>
    public class InvalidSheetNameException : Exception
    {
        public InvalidSheetNameException(string name) : base(string.Format("The supplied name ({0}) is not a valid XLSX worksheet name.", name)) { }
    }

    /// <summary>
    /// Column reference out of range exception
    /// </summary>
    public class ColumnReferenceOutOfRange : Exception
    {
        public ColumnReferenceOutOfRange(string columnReference)
            : base(string.Format("Column reference ({0}) is out of range.", columnReference))
        {
        }
    }

    /// <summary>
    /// Worksheet already exists exception
    /// </summary>
    public class WorksheetAlreadyExistsException : Exception
    {
        public WorksheetAlreadyExistsException(string sheetName)
            : base(string.Format("The worksheet ({0}) already exists.", sheetName))
        {
        }
    }
}
