using System;

namespace OpenXmlPowerTools.Documents
{
    /// <summary>
    /// Document builder exception
    /// </summary>
    public class DocumentBuilderException : Exception
    {
        public DocumentBuilderException(string message) : base(message) { }
    }

    /// <summary>
    /// Document builder internal exception
    /// </summary>
    public class DocumentBuilderInternalException : Exception
    {
        public DocumentBuilderInternalException(string message) : base(message) { }
    }
}