using OpenXmlPowerTools.Documents;

namespace OpenXmlPowerTools.Commons
{
    public interface ISource
    {
        PowerToolsDocument WmlDocument { get; set; }
        int Start { get; set; }
        int Count { get; set; }
        string InsertId { get; set; }
        bool KeepHeadersAndFooters { get; set; }
        bool KeepSections { get; set; }
    }
}