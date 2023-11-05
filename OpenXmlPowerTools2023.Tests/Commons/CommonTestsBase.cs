using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace OpenXmlPowerTools2023.Tests.Commons
{
    [TestClass]
    [DeploymentItem(TestResourcePath, OutputPath)]
    public abstract class CommonTestsBase : TestsBase
    {
        protected override string Extension { get; } = ".*";
        protected override string ModuleFolder { get; } = @".\Commons";
        protected override string OutputFile { get; set; } = "Output.pptx";
        protected override string FeatureFolder { get; } = "";
        protected static void CreateEmptyWordprocessingDocument(Stream stream) => TestUtil.CreateEmptyWordprocessingDocument(stream);
    }
}