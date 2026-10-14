using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Compatibility_NumberingTabSettingRetainsExplicitValuesAndSchemaOrder() {
        using WordDocument document = WordDocument.Create();
        Assert.False(document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop);
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        document.CompatibilitySettings.SplitPageBreakAndParagraphMark = true;
        Compatibility compatibility = document._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.GetFirstChild<Compatibility>()!;
        compatibility.AddChild(new DoNotUseIndentAsNumberingTabStop(), true);
        Assert.True(document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop);
        foreach (bool enabled in new[] { false, true, false }) {
            document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop = enabled;
            using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
            Assert.Equal(enabled, restored.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop);
            Assert.Equal(enabled, Assert.Single(compatibility.Elements<DoNotUseIndentAsNumberingTabStop>()).Val!.Value);
            Assert.True(restored.CompatibilitySettings.SplitPageBreakAndParagraphMark);
            Assert.Equal(WordCompatibilityMode.Word2010, restored.CompatibilitySettings.CompatibilityMode);
            Assert.Empty(new OpenXmlValidator().Validate(compatibility));
        }
    }

    [Fact]
    public void LegacyDoc_NumberingTabEffectiveLayoutSurvivesDocxProjectionAndNativeSaving() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Word", "LegacyLists", "WordListNumberTabs.doc");
        using WordDocument source = WordDocument.Load(path);
        Assert.True(source.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop);
        using WordDocument docx = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Docx)));
        using WordDocument native = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.True(docx.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop);
        Assert.True(native.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop);
        Assert.Equal(WordCompatibilityMode.Word2003, docx.CompatibilitySettings.CompatibilityMode);
        AssertListNumberTab(native, 3000);
    }

    [Theory]
    [InlineData(WordCompatibilityMode.Word2003)]
    [InlineData(WordCompatibilityMode.Word2013)]
    public void LegacyDoc_NumberingTabSavingRejectsAnExplicitlyDisabledSetting(WordCompatibilityMode mode) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Retained content");
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop = false;
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToBytes(WordFileFormat.Doc));
        Assert.Contains("DoNotUseIndentAsNumberingTabStop=false", error.Message);
        using WordDocument docx = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        Assert.False(docx.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop);
        Assert.Equal("Retained content", Assert.Single(docx.Paragraphs).Text);
    }
}
