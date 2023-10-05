using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Commons
{
    /// <summary>
    /// Cached header and footer
    /// </summary>
    public class CachedHeaderFooter
    {
        public XName Ref;
        public string Type;
        public string CachedPartRid;
    };

    public struct TempSource
    {
        public int Start;
        public int Count;
    };

    public class Atbid
    {
        public XElement BlockLevelContent;
        public int Index;
        public int Div;
    }
    public class ReplaceSemaphore { }
    public class FromPreviousSourceSemaphore { };
}