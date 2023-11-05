using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using System.IO;

namespace OpenXmlPowerTools2023.Tests.Commons
{
    [TestClass]
    public class PtUtilTests : CommonTestsBase
    {
        [TestMethod]
        [DataRow("PU001-Test001.mht")]
        public void PU001(string name)
        {
            var filePath = GetFile(name);
            var src = File.ReadAllText(filePath);
            var p = MhtParser.Parse(src);

            p.ContentType.Should().NotBeNull();
            p.MimeVersion.Should().NotBeNull();
            p.Parts.Length.Should().NotBe(0);
            p.Parts.Should().NotContain(part => part.ContentType == null || part.ContentLocation == null);
        }
    }
}