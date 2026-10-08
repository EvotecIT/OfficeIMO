using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Compatibility_PageBreakMarkReadDoesNotCreateSettings() {
        using WordDocument document = WordDocument.Create();
        var main = document._wordprocessingDocument.MainDocumentPart!;
        main.DeletePart(main.DocumentSettingsPart!);
        Assert.False(document.CompatibilitySettings.SplitPageBreakAndParagraphMark);
        Assert.Null(main.DocumentSettingsPart);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, false)]
    [InlineData(WordFileFormat.Docx, true)]
    [InlineData(WordFileFormat.Doc, true)]
    public void Compatibility_PageBreakMarkExplicitValuesSurviveRoundTrips(WordFileFormat format, bool enabled) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Preserved content");
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        document.CompatibilitySettings.DoNotBalanceTextColumns = true;
        document.Settings.MirrorMargins = true;
        document.CompatibilitySettings.SplitPageBreakAndParagraphMark = enabled;
        byte[] bytes = document.ToBytes(format);
        if (format == WordFileFormat.Doc) {
            byte[] word = ReadCompoundStream(bytes, "WordDocument");
            Assert.Equal(0x00C1, BitConverter.ToUInt16(word, 2));
            Assert.Equal(500, BitConverter.ToInt32(word, 0x196));
        }
        for (int cycle = 0; cycle < 2; cycle++) {
            using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
            Assert.Equal(enabled, loaded.CompatibilitySettings.SplitPageBreakAndParagraphMark);
            Assert.True(loaded.CompatibilitySettings.DoNotBalanceTextColumns);
            Assert.True(loaded.Settings.MirrorMargins);
            Assert.Equal("Preserved content", Assert.Single(loaded.Paragraphs).Text);
            bytes = loaded.ToBytes(format);
        }
    }

    [Fact]
    public void Compatibility_PageBreakMarkSetterRetainsExplicitFalseAndSchemaOrder() {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        Compatibility compatibility = document._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.GetFirstChild<Compatibility>()!;
        compatibility.AddChild(new UseWord2002TableStyleRules(), true);
        compatibility.AddChild(new SplitPageBreakAndParagraphMark(), true);
        Assert.True(document.CompatibilitySettings.SplitPageBreakAndParagraphMark);
        foreach (bool enabled in new[] { false, true, false }) {
            document.CompatibilitySettings.SplitPageBreakAndParagraphMark = enabled;
            Assert.Equal(enabled, document.CompatibilitySettings.SplitPageBreakAndParagraphMark);
            Assert.Equal(enabled, Assert.Single(compatibility.Elements<SplitPageBreakAndParagraphMark>()).Val!.Value);
            Assert.NotNull(compatibility.GetFirstChild<UseWord2002TableStyleRules>());
            Assert.Equal(WordCompatibilityMode.Word2010, document.CompatibilitySettings.CompatibilityMode);
            Assert.Empty(new OpenXmlValidator().Validate(compatibility));
        }
    }

    [Fact]
    public void LegacyDoc_PageBreakMarkDefaultProjectsTheLegacyLayoutIntoDocx() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Plain Word97 document");
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.Equal(0x00C1, BitConverter.ToUInt16(ReadCompoundStream(bytes, "WordDocument"), 2));
        using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
        Assert.True(loaded.CompatibilitySettings.SplitPageBreakAndParagraphMark);
        Assert.Equal(WordCompatibilityMode.Word2003, loaded.CompatibilitySettings.CompatibilityMode);
        using WordDocument docx = WordDocument.Load(new MemoryStream(loaded.ToBytes(WordFileFormat.Docx)));
        Assert.True(docx.CompatibilitySettings.SplitPageBreakAndParagraphMark);
        Assert.Equal(WordCompatibilityMode.Word2003, docx.CompatibilitySettings.CompatibilityMode);
    }

    [Theory]
    [InlineData(WordCompatibilityMode.Word2003)]
    [InlineData(WordCompatibilityMode.Word2013)]
    public void LegacyDoc_PageBreakMarkSavingRejectsUnsupportedDisabledSetting(WordCompatibilityMode mode) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Preserved content");
        source.CompatibilitySettings.CompatibilityMode = mode;
        source.CompatibilitySettings.SplitPageBreakAndParagraphMark = false;
        var error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Contains("Save as DOCX", error.Message);
        using WordDocument docx = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Docx)));
        Assert.False(docx.CompatibilitySettings.SplitPageBreakAndParagraphMark);
        Assert.Equal(mode, docx.CompatibilitySettings.CompatibilityMode);
        Assert.Equal("Preserved content", Assert.Single(docx.Paragraphs).Text);
    }
}
