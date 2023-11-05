using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests
{
    [TestClass]
    [DeploymentItem(TestResourcePath, OutputPath)]
    public abstract class TestsBase
    {
        private TestContext testContext;
        public TestContext TestContext
        {
            get => testContext;
            set => testContext = value;
        }
        protected OpenXmlValidator Validator = new OpenXmlValidator();


        protected static IList<string> ExpectedErrors = new List<string>()
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
            "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:sectPr'.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:p'.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:updateFields'.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:name' has invalid value 'useWord2013TrackBottomHyphenation'. The Enumeration constraint failed.",
            "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:ins'. List of possible elements expected: <http://schemas.openxmlformats.org/officeDocument/2006/math:rPr>.",
            "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:del'. List of possible elements expected: <http://schemas.openxmlformats.org/officeDocument/2006/math:rPr>.",
            "The 'http://schemas.microsoft.com/office/word/2012/wordml:restartNumberingAfterBreak' attribute is not declared.",
            "Attribute 'id' should have unique value. Its current value '",
            "The 'urn:schemas-microsoft-com:mac:vml:blur' attribute is not declared.",
            "Attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:id' should have unique value. Its current value '",
            "The element has unexpected child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The element has invalid child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The 'urn:schemas-microsoft-com:mac:vml:complextextbox' attribute is not declared.",
            "http://schemas.microsoft.com/office/word/2010/wordml:",
            "http://schemas.microsoft.com/office/word/2008/9/12/wordml:",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        };

        protected const string OutputPath = @".\TestResources";
        protected const string TestResourcePath = @"TestResources\";
        protected abstract string ModuleFolder { get; }
        protected abstract string FeatureFolder { get; }
        protected abstract string Extension { get; }
        protected abstract string OutputFile { get; set; }

        protected static void CleanupTest(string path)
        {
            if (File.Exists(path)) File.Delete(path);
        }

        protected string ToOutputPath(string path)
        {
            return Path.Combine(TestContext.TestResultsDirectory, OutputFile);
        }

        protected IDictionary<string, string> GetFiles(string subfolder, string extension = null)
        {
            if (extension == null) extension = Extension;
            return Directory.GetFiles(Path.Combine(TestResourcePath, subfolder), $"*{extension}")
                        .ToDictionary(k => Path.GetFileName(k), v => v);
        }

        protected string GetFile(string fileName, string moduleFolder = null, string featureFolder = null)
        {
            if (moduleFolder == null) moduleFolder = ModuleFolder;
            if (featureFolder == null) featureFolder = FeatureFolder;
            return Path.Combine(Path.Combine(Path.Combine(OutputPath, moduleFolder), featureFolder), fileName);
        }

        protected static string InnerText(XContainer e)
        {
            return e.Descendants(W.r)
                .Where(r => r.Parent.Name != W.del)
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
        }

        protected static string InnerDelText(XContainer e)
        {
            return e.Descendants(W.delText)
                .Select(delText => delText.Value)
                .StringConcatenate();
        }

        protected static string ToTempPath(string path, string ext = null)
        {
            if (ext != null) path = Path.ChangeExtension(Path.GetFileName(path), ext);

            return Path.Combine(Path.GetTempPath(), path);
        }

        protected void Validate(string filepath)
        {
            IEnumerable<ValidationErrorInfo> validationErrors = new List<ValidationErrorInfo>();
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filepath, true))
            {
                validationErrors = Validator.Validate(wDoc).ToList();
            }

            // Assert
            validationErrors.Count().Should().Be(0, because: validationErrors.ToString());
        }

        protected void ValidateAgainstExpected(string filepath)
        {
            IEnumerable<ValidationErrorInfo> validationErrors = new List<ValidationErrorInfo>();
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filepath, true))
            {
                validationErrors = Validator.Validate(wDoc).Where(ve =>
                {
                    var found = ExpectedErrors.Any(xe => ve.Description.Contains(xe));
                    return !found;
                });
            }

            // Assert
            validationErrors.Count().Should().Be(0);
        }
    }
}