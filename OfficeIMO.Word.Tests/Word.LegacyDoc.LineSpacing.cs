using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordLineSpacingRule.Auto, 276)]
    [InlineData(WordLineSpacingRule.Exact, 320)]
    [InlineData(WordLineSpacingRule.AtLeast, 360)]
    public void LegacyDoc_LineSpacingRuleSurvivesNativeSave(WordLineSpacingRule rule, int value) {
        using WordDocument source = WordDocument.Create();
        var first = source.AddParagraph("First");
        first.LineSpacing = value;
        first.LineSpacingRule = rule;
        first.LineSpacingAfter = 160;
        var second = source.AddParagraph("Second");
        second.LineSpacing = value;
        second.LineSpacingRule = rule;
        second.LineSpacingAfter = 160;
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        byte[] document = ReadCompoundStream(bytes, "WordDocument");
        byte[] table = ReadCompoundStream(bytes, "1Table");
        int bins = BitConverter.ToInt32(document, 0x102);
        int page = BitConverter.ToInt32(table, bins + 8) * 512;
        // The external DOC paragraph-boundary contract must retain both equally
        // formatted paragraphs, rather than relying on the importer splitting CRs.
        int textStart = BitConverter.ToInt32(document, 0x18);
        Assert.Equal(textStart, BitConverter.ToInt32(document, page));
        Assert.Equal(textStart + "First\r".Length, BitConverter.ToInt32(document, page + 4));
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        var paragraphs = reopened.Paragraphs.Where(p => p.Text is "First" or "Second").ToArray();
        Assert.Equal(2, paragraphs.Length);
        Assert.All(paragraphs, p => {
            Assert.Equal(value, p.LineSpacing);
            Assert.Equal(rule, p.LineSpacingRule);
            Assert.Equal(160, p.LineSpacingAfter);
        });
    }

    [Theory]
    [InlineData(5)]
    [InlineData(40)]
    public void LegacyDoc_PapxPreservesPlainParagraphBoundariesAcrossPages(int count) {
        using WordDocument source = WordDocument.Create();
        var expectedOffsets = new List<int>();
        int characters = 0;
        for (int i = 0; i < count; i++) {
            string text = "Paragraph " + i;
            expectedOffsets.Add(characters);
            source.AddParagraph(text);
            characters += text.Length + 1;
        }
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        byte[] document = ReadCompoundStream(bytes, "WordDocument");
        byte[] table = ReadCompoundStream(bytes, "1Table");
        int binOffset = BitConverter.ToInt32(document, 0x102);
        int binCount = (BitConverter.ToInt32(document, 0x106) - 4) / 8;
        int textStart = BitConverter.ToInt32(document, 0x18);
        var actualOffsets = new List<int>();
        for (int bin = 0; bin < binCount; bin++) {
            int page = BitConverter.ToInt32(table, binOffset + (binCount + 1) * 4 + bin * 4) * 512;
            int paragraphs = document[page + 511];
            Assert.InRange(paragraphs, 1, 29);
            for (int i = 0; i < paragraphs; i++) {
                int offset = BitConverter.ToInt32(document, page + i * 4) - textStart;
                if (offset < characters) actualOffsets.Add(offset);
            }
        }
        Assert.Equal(expectedOffsets, actualOffsets);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_PieceTableUsesTerminalPaddingOnlyForAdditionalStories(bool withHeader) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Main body");
        if (withHeader) {
            source.AddHeadersAndFooters();
            source.Header.Default!.AddParagraph("Header");
        }
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        byte[] document = ReadCompoundStream(bytes, "WordDocument");
        byte[] table = ReadCompoundStream(bytes, "1Table");
        int textCharacters = BitConverter.ToInt32(document, 0x4C);
        int headerCharacters = BitConverter.ToInt32(document, 0x54);
        Assert.Equal("Main body\r".Length, textCharacters);
        Assert.Equal(withHeader, headerCharacters > 0);
        // Native one-piece PlcPcd: Pcdt begins at byte zero in 1Table;
        // the second CP is its exclusive text limit.
        Assert.Equal(textCharacters + headerCharacters + (withHeader ? 1 : 0), BitConverter.ToInt32(table, 9));
    }

    [Fact]
    public void LegacyDoc_IndependentNormalStylePreservesMultipleLineSpacing() {
        using WordDocument word = WordDocument.Load(GetFixtureDoc(Path.Combine("LegacyDocCorpus", "ComCroppedInlinePicture.doc")));
        void AssertSpacing(WordDocument document) {
            var style = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>(),
                item => item.StyleId?.Value == "Normal");
            var spacing = style.StyleParagraphProperties!.GetFirstChild<SpacingBetweenLines>()!;
            Assert.Equal(LineSpacingRuleValues.Auto, spacing.LineRule!.Value);
            Assert.Equal("278", spacing.Line!.Value);
            Assert.Equal("160", spacing.After!.Value);
        }
        AssertSpacing(word);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes(WordFileFormat.Doc)));
        AssertSpacing(reopened);
    }
}
