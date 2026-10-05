using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Theory]
    [InlineData(WordSectionBreakType.Continuous, RtfSectionBreakKind.Continuous)]
    [InlineData(WordSectionBreakType.NextColumn, RtfSectionBreakKind.Column)]
    [InlineData(WordSectionBreakType.NextPage, RtfSectionBreakKind.NextPage)]
    [InlineData(WordSectionBreakType.OddPage, RtfSectionBreakKind.OddPage)]
    [InlineData(WordSectionBreakType.EvenPage, RtfSectionBreakKind.EvenPage)]
    public void Word_Rtf_Bridge_Preserves_Each_Sections_Own_Start_Before_And_After_Docx_Save(
        WordSectionBreakType start, RtfSectionBreakKind expected) {
        using WordDocument word = WordDocument.Create();
        word.Sections[0].BreakType = WordSectionBreakType.Continuous;
        word.Sections[0].AddParagraph("First");
        word.AddSection(start).AddParagraph("Second");
        WordSectionBreakType last = start == WordSectionBreakType.EvenPage
            ? WordSectionBreakType.OddPage : WordSectionBreakType.EvenPage;
        word.AddSection(last).AddParagraph("Third");
        RtfSectionBreakKind lastExpected = last == WordSectionBreakType.OddPage
            ? RtfSectionBreakKind.OddPage : RtfSectionBreakKind.EvenPage;

        AssertStarts(word);
        using var stream = new MemoryStream();
        word.Save(stream);
        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        AssertStarts(reopened);

        void AssertStarts(WordDocument source) {
            RtfDocument direct = source.ToRtfDocument();
            Assert.Equal(RtfSectionBreakKind.Continuous, direct.Sections[0].BreakKind);
            Assert.Equal(expected, direct.Sections[1].BreakKind);
            Assert.Equal(lastExpected, direct.Sections[2].BreakKind);
            RtfDocument parsed = RtfDocument.Parse(source.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
            Assert.Equal(RtfSectionBreakKind.Continuous, parsed.Sections[0].BreakKind);
            Assert.Equal(expected, parsed.Sections[1].BreakKind);
            Assert.Equal(lastExpected, parsed.Sections[2].BreakKind);
            using WordDocument roundTrip = parsed.ToWordDocument();
            Assert.Equal(WordSectionBreakType.Continuous, roundTrip.Sections[0].BreakType);
            Assert.Equal(start, roundTrip.Sections[1].BreakType);
            Assert.Equal(last, roundTrip.Sections[2].BreakType);
            Assert.Equal(new[] { "First", "Second", "Third" },
                roundTrip.Sections.Select(section => string.Concat(section.Paragraphs.Select(paragraph => paragraph.Text))));
            var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator()
                .Validate(roundTrip._wordprocessingDocument).Select(error => error.Description).ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
        }
    }
}
