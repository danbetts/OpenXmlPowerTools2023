using System;

namespace OpenXmlPowerTools.Commons
{
    /// <summary>
    /// Powertools document exception
    /// </summary>
    public class PowerToolsDocumentException : Exception
    {
        public PowerToolsDocumentException(string message) : base(message) { }
    }

    /// <summary>
    /// PowerTools Invalid Data Exception
    /// </summary>
    public class PowerToolsInvalidDataException : Exception
    {
        public PowerToolsInvalidDataException(string message) : base(message) { }
    }

    /// <summary>
    /// Invalid OpenXml Document Exception
    /// </summary>
    public class InvalidOpenXmlDocumentException : Exception
    {
        public InvalidOpenXmlDocumentException(string message) : base(message) { }
    }

    /// <summary>
    /// OpenXml PowerTools Exception
    /// </summary>
    public class OpenXmlPowerToolsException : Exception
    {
        public OpenXmlPowerToolsException(string message) : base(message) { }
    }
}