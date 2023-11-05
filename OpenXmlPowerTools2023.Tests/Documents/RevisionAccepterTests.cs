using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Documents;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    public class RevisionAccepterTests : DocumentTestsBase
    {
        protected override string FeatureFolder { get; } = @".\RevisionAccepter";

        [TestMethod]
        [DataRow("RA001-Tracked-Revisions-01.docx")]
        [DataRow("RA001-Tracked-Revisions-02.docx")]

        public void RA001(string name)
        {
            WmlDocument afterAccepting = RevisionAccepter.AcceptRevisions(new WmlDocument(GetFile(name)));
            CleanupTest(OutputFile);
            afterAccepting.SaveAs(OutputFile);
        }
    }
}