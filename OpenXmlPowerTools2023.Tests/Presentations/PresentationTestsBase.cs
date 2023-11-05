using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlPowerTools2023.Tests.Presentations
{
    [TestClass]
    [DeploymentItem(TestResourcePath, OutputPath)]
    public abstract class PresentationTestsBase : TestsBase
    {
        protected override string Extension { get; } = ".pptx";
        protected override string ModuleFolder { get; } = @".\Presentations";
        protected override string OutputFile { get; set; } = "Output.docx";
        protected override string FeatureFolder { get; } = "";

    }
}