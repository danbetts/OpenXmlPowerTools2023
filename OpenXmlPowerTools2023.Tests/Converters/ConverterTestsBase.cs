using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace OpenXmlPowerTools2023.Tests.Converters
{
    [TestClass]
    [DeploymentItem(TestResourcePath, OutputPath)]
    public abstract class ConverterTestsBase : TestsBase
    {
        protected override string Extension { get; } = ".*";
        protected override string ModuleFolder { get; } = @".\Converters";
        protected override string OutputFile { get; set; } = "Output.pptx";
    }
}