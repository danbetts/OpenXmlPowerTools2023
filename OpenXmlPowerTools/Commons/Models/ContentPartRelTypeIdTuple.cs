using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools.Commons
{
    public class ContentPartRelTypeIdTuple
    {
        public OpenXmlPart ContentPart { get; set; }
        public string RelationshipType { get; set; }
        public string RelationshipId { get; set; }
    }
}