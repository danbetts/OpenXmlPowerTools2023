using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Converters;

namespace OpenXmlPowerTools2023.Tests.Converters
{
    [TestClass]
    public class ListItemTextGetterRuTests : ConverterTestsBase
    {
        protected override string FeatureFolder { get; } = "";

        [TestMethod]
        [DataRow(1, "1-ый")]
        [DataRow(2, "2-ой")]
        [DataRow(3, "3-ий")]
        [DataRow(4, "4-ый")]
        [DataRow(5, "5-ый")]
        [DataRow(6, "6-ой")]
        [DataRow(7, "7-ой")]
        [DataRow(8, "8-ой")]
        [DataRow(9, "9-ый")]
        [DataRow(10, "10-ый")]
        [DataRow(11, "11-ый")]
        [DataRow(12, "12-ый")]
        [DataRow(13, "13-ый")]
        [DataRow(14, "14-ый")]
        [DataRow(16, "16-ый")]
        [DataRow(17, "17-ый")]
        [DataRow(18, "18-ый")]
        [DataRow(19, "19-ый")]
        [DataRow(20, "20-ый")]
        [DataRow(23, "23-ий")]
        [DataRow(25, "25-ый")]
        [DataRow(50, "50-ый")]
        [DataRow(56, "56-ой")]
        [DataRow(67, "67-ой")]
        [DataRow(78, "78-ой")]
        [DataRow(100, "100-ый")]
        [DataRow(123, "123-ий")]
        [DataRow(125, "125-ый")]
        [DataRow(1050, "1050-ый")]
        public void GetListItemText_Ordinal(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_ru_RU.GetListItemText("", integer, "ordinal");
            actualText.Should().Be(expectedText);
        }
    }
}