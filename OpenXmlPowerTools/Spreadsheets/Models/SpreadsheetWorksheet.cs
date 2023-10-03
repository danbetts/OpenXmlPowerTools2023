using System.Collections.Generic;

namespace OpenXmlPowerTools.Spreadsheets
{
    public class SpreadsheetWorksheet
    {
        public string Name;
        public string TableName;
        public IEnumerable<SpreadsheetCell> ColumnHeadings;
        public IEnumerable<SpreadsheetRow> Rows;
    }
}
