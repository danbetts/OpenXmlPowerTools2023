using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools.Commons;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools2023.Tests.Commons
{
    [TestClass]
    public class UnicodeMapperTests : CommonTestsBase
    {
        private const string PreserveSpacingXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">The following space is retained: </w:t>
      </w:r>
      <w:r>
        <w:t>but this one is not: </w:t>
      </w:r>
      <w:r>
        <w:t xml:space=""preserve"">. Similarly these two lines should have only a space between them: </w:t>
      </w:r>
      <w:r>
        <w:t>
          Line 1!
Line 2!
        </w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

        [TestMethod]
        public void CanStringifyRunAndTextElements()
        {
            const string textValue = "Hello World!";
            var textElement = new XElement(W.t, textValue);
            var runElement = new XElement(W.r, textElement);
            var formattedRunElement = new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), textElement);

            UnicodeMapper.RunToString(textElement).Should().Be(textValue);
            UnicodeMapper.RunToString(runElement).Should().Be(textValue);
            UnicodeMapper.RunToString(formattedRunElement).Should().Be(textValue);
        }

        [TestMethod]
        public void CanStringifySpecialElements()
        {
            UnicodeMapper.RunToString(new XElement(W.cr)).First().Should().Be(UnicodeMapper.CarriageReturn);
            UnicodeMapper.RunToString(new XElement(W.br)).First().Should().Be(UnicodeMapper.CarriageReturn);
            UnicodeMapper.RunToString(new XElement(W.br, new XAttribute(W.type, "page"))).First().Should().Be(UnicodeMapper.FormFeed);
            UnicodeMapper.RunToString(new XElement(W.noBreakHyphen)).First().Should().Be(UnicodeMapper.NonBreakingHyphen);
            UnicodeMapper.RunToString(new XElement(W.softHyphen)).First().Should().Be(UnicodeMapper.SoftHyphen);
            UnicodeMapper.RunToString(new XElement(W.tab)).First().Should().Be(UnicodeMapper.HorizontalTabulation);
        }

        [TestMethod]
        public void CanCreateRunChildElementsFromSpecialCharacters()
        {
            UnicodeMapper.CharToRunChild(UnicodeMapper.CarriageReturn).Name.Should().Be(W.br);
            UnicodeMapper.CharToRunChild(UnicodeMapper.NonBreakingHyphen).Name.Should().Be(W.noBreakHyphen);
            UnicodeMapper.CharToRunChild(UnicodeMapper.SoftHyphen).Name.Should().Be(W.softHyphen);
            UnicodeMapper.CharToRunChild(UnicodeMapper.HorizontalTabulation).Name.Should().Be(W.tab);

            XElement element = UnicodeMapper.CharToRunChild(UnicodeMapper.FormFeed);

            element.Name.Should().Be(W.br);
            element.Attribute(W.type).Value.Should().Be("page");
            UnicodeMapper.CharToRunChild('\r').Name.Should().Be(W.br);
        }

        [TestMethod]
        public void CanCreateCoalescedRuns()
        {
            const string textString = "This is only text.";
            const string mixedString = "First\tSecond\tThird";

            List<XElement> textRuns = UnicodeMapper.StringToCoalescedRunList(textString, null);
            List<XElement> mixedRuns = UnicodeMapper.StringToCoalescedRunList(mixedString, null);

            textRuns.Count.Should().Be(1);
            mixedRuns.Count.Should().Be(5);
            mixedRuns.Elements(W.t).Skip(0).First().Value.Should().Be("First");
            mixedRuns.Elements(W.t).Skip(1).First().Value.Should().Be("Second");
            mixedRuns.Elements(W.t).Skip(2).First().Value.Should().Be("Third");
        }

        [TestMethod]
        public void CanMapSymbols()
        {
            var sym1 = new XElement(W.sym,
                new XAttribute(W.font, "Wingdings"),
                new XAttribute(W._char, "F028"));
            char charFromSym1 = UnicodeMapper.SymToChar(sym1);
            XElement symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);

            var sym2 = new XElement(W.sym,
                new XAttribute(W._char, "F028"),
                new XAttribute(W.font, "Wingdings"));
            char charFromSym2 = UnicodeMapper.SymToChar(sym2);

            var sym3 = new XElement(W.sym,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(W.font, "Wingdings"),
                new XAttribute(W._char, "F028"));
            char charFromSym3 = UnicodeMapper.SymToChar(sym3);

            var sym4 = new XElement(W.sym,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(W.font, "Webdings"),
                new XAttribute(W._char, "F028"));
            char charFromSym4 = UnicodeMapper.SymToChar(sym4);
            XElement symFromChar4 = UnicodeMapper.CharToRunChild(charFromSym4);

            charFromSym2.Should().Be(charFromSym1);
            charFromSym3.Should().Be(charFromSym1);
            charFromSym4.Should().NotBe(charFromSym1);

            symFromChar1.Attribute(W._char).Value.Should().Be("F028");
            symFromChar1.Attribute(W.font).Value.Should().Be("Wingdings");
            symFromChar4.Attribute(W._char).Value.Should().Be("F028");
            symFromChar4.Attribute(W.font).Value.Should().Be("Webdings");
        }

        [TestMethod]
        public void CanStringifySymbols()
        {
            char charFromSym1 = UnicodeMapper.SymToChar("Wingdings", '\uF028');
            char charFromSym2 = UnicodeMapper.SymToChar("Wingdings", 0xF028);
            char charFromSym3 = UnicodeMapper.SymToChar("Wingdings", "F028");

            XElement symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);
            XElement symFromChar2 = UnicodeMapper.CharToRunChild(charFromSym2);
            XElement symFromChar3 = UnicodeMapper.CharToRunChild(charFromSym3);
            
            charFromSym2.Should().Be(charFromSym1);
            charFromSym3.Should().Be(charFromSym1);

            symFromChar2.ToString(SaveOptions.None).Should().Be(symFromChar1.ToString(SaveOptions.None));
            symFromChar3.ToString(SaveOptions.None).Should().Be(symFromChar1.ToString(SaveOptions.None));
        }

        [TestMethod]
        public void HonorsXmlSpace()
        {
            XDocument partDocument = XDocument.Parse(PreserveSpacingXmlString);
            XElement p = partDocument.Descendants(W.p).Last();
            string innerText = p.Descendants(W.r)
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
            innerText.Should().Be(@"The following space is retained: but this one is not:. Similarly these two lines should have only a space between them: Line 1! Line 2!");
        }
    }
}