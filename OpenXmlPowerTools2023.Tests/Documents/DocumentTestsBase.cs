using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Documents;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpenXmlPowerTools2023.Tests.Documents
{
    [TestClass]
    [DeploymentItem(TestResourcePath, OutputPath)]
    public abstract class DocumentTestsBase : TestsBase
    {
        protected override string Extension { get; } = ".docx";
        protected override string ModuleFolder { get; } = @".\Documents";
        protected override string OutputFile { get; set; } = @".\Output.docx";
        protected DocumentBuilder Builder { get; set; } = new DocumentBuilder();

        //[ClassInitialize]
        //public static void TestInit(TestContext context)
        //{
        //    Builder = new DocumentBuilder();
        //}

        protected static void CreateEmptyWordprocessingDocument(Stream stream) => TestUtil.CreateEmptyWordprocessingDocument(stream);
        protected string Validate(string filepath, out string output)
        {
            output = BuildAndSave();
            base.Validate(output);
            return output;
        }

        protected string ValidateAgainstExpected(string filepath, out string output)
        {
            output = BuildAndSave();
            base.ValidateAgainstExpected(output);
            return output;
        }

        protected string BuildAndSave()
        {
            var output = ToOutputPath(OutputFile);
            CleanupTest(output);
            Builder.SaveAs(output);
            return output;
        }
    }
}