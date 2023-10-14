using System;

namespace OpenXmlPowerTools.Presentations
{
    /// <summary>
    /// Presentation builder exception
    /// </summary>
    public class PresentationBuilderException : Exception
    {
        public PresentationBuilderException(string message) : base(message) { }
    }

    /// <summary>
    /// Presentation builder internal exception
    /// </summary>
    public class PresentationBuilderInternalException : Exception
    {
        public PresentationBuilderInternalException(string message) : base(message) { }
    }
}
