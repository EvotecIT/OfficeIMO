using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordSectionBreakType.NextPage, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.Continuous, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.NextColumn, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.OddPage, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.EvenPage, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.NextPage, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.Continuous, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.NextColumn, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.OddPage, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.EvenPage, WordFileFormat.Doc)]
    public void SectionStartContract_RoundTripsEachSectionsOwnType(WordSectionBreakType breakType, WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("FirstSection");
        WordSection second = document.AddSection(breakType);
        second.AddParagraph("SecondSection");
        WordSection third = document.AddSection();
        third.AddParagraph("ThirdSection");
        Assert.Equal(breakType, second.BreakType);
        Assert.Equal(WordSectionBreakType.NextPage, third.BreakType);
        Assert.Empty(new OpenXmlValidator().Validate(document._wordprocessingDocument));

        byte[] bytes = document.ToBytes(format);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal(new[] { WordSectionBreakType.NextPage, breakType, WordSectionBreakType.NextPage },
            loaded.Sections.Select(section => section.BreakType));
        for (int index = 0; index < 3; index++) {
            Assert.Contains(new[] { "FirstSection", "SecondSection", "ThirdSection" }[index],
                loaded.Sections[index].Paragraphs.Select(paragraph => paragraph.Text));
        }
        if (format == WordFileFormat.Doc) {
            string storedText = Encoding.ASCII.GetString(ReadCompoundStream(bytes, "WordDocument"));
            Assert.Contains("FirstSection\r\fSecondSection\r\fThirdSection\r", storedText);
        }
    }

    [Fact]
    public void SectionStartContract_SetterPreservesOtherSettingsAndRejectsInvalidType() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        Assert.Equal(WordSectionBreakType.NextPage, section.BreakType);
        section.Margins.Left = 1234;
        section.BreakType = WordSectionBreakType.Continuous;
        Assert.Throws<ArgumentOutOfRangeException>(() => section.BreakType = (WordSectionBreakType)999);
        Assert.Equal(WordSectionBreakType.Continuous, section.BreakType);
        Assert.Equal(1234U, section.Margins.Left);
        section._sectionProperties.RemoveAllChildren<SectionType>();
        Assert.Equal(WordSectionBreakType.NextPage, section.BreakType);
    }

    [Fact]
    public void SectionStartContract_TypedSectionPreservesPrecedingHeaderAndStart() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].BreakType = WordSectionBreakType.Continuous;
        document.AddHeadersAndFooters();
        RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("OriginalHeader");
        document.AddParagraph("FirstSection");
        document.AddSection(WordSectionBreakType.OddPage).AddParagraph("SecondSection");
        Assert.Equal(WordSectionBreakType.Continuous, document.Sections[0].BreakType);
        Assert.Contains("OriginalHeader", RequireSectionHeader(document, 0, HeaderFooterValues.Default).Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Equal(WordSectionBreakType.OddPage, document.Sections[1].BreakType);
    }
}
