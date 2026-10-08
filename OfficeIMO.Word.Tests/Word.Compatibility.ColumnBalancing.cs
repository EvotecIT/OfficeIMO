using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Compatibility_ColumnBalancingReadsDefaultWithoutCreatingSettings() {
        using WordDocument document = WordDocument.Create();
        var main = document._wordprocessingDocument.MainDocumentPart!;
        main.DeletePart(main.DocumentSettingsPart!);
        Assert.False(document.CompatibilitySettings.DoNotBalanceTextColumns);
        document.CompatibilitySettings.DoNotBalanceTextColumns = false;
        Assert.Null(main.DocumentSettingsPart);
        document.CompatibilitySettings.DoNotBalanceTextColumns = true;
        Assert.True(document.CompatibilitySettings.DoNotBalanceTextColumns);
        Assert.NotNull(main.DocumentSettingsPart);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, null)]
    [InlineData(WordFileFormat.Docx, true)]
    [InlineData(WordFileFormat.Docx, false)]
    [InlineData(WordFileFormat.Doc, null)]
    [InlineData(WordFileFormat.Doc, true)]
    [InlineData(WordFileFormat.Doc, false)]
    public void Compatibility_ColumnBalancingImportsImplicitAndExplicitValues(WordFileFormat format, bool? value) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Column policy");
        Settings settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        Compatibility compatibility = settings.GetFirstChild<Compatibility>() ?? new Compatibility();
        if (compatibility.Parent == null) settings.AddChild(compatibility, true);
        var flag = new NoColumnBalance();
        if (value.HasValue) flag.Val = value.Value;
        compatibility.AddChild(flag, true);
        bool expected = value ?? true;
        Assert.Equal(expected, document.CompatibilitySettings.DoNotBalanceTextColumns);
        byte[] bytes = document.ToBytes(format);
        if (format == WordFileFormat.Doc && expected) {
            byte[] word = ReadCompoundStream(bytes, "WordDocument");
            byte[] table = ReadCompoundStream(bytes, "1Table");
            int offset = BitConverter.ToInt32(word, 0x192);
            Assert.True(BitConverter.ToInt32(word, 0x196) >= 88);
            Assert.Equal(0x20, BitConverter.ToUInt16(table, offset + 8) & 0x20);
            Assert.Equal(BitConverter.ToUInt16(table, offset + 8), BitConverter.ToUInt16(table, offset + 84));
        }
        using var stream = new MemoryStream(bytes);
        using WordDocument loaded = WordDocument.Load(stream);
        Assert.Equal(expected, loaded.CompatibilitySettings.DoNotBalanceTextColumns);
        Assert.Equal("Column policy", Assert.Single(loaded.Paragraphs).Text);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void Compatibility_ColumnBalancingCanBeClearedWithoutLosingDocumentOptions(WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Preserved options");
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        document.Settings.MirrorMargins = true;
        document.Settings.GutterAtTop = true;
        document.Settings.TrackRevisions = true;
        document.Sections[0].DifferentOddAndEvenPages = true;
        document.Sections[0].Margins.Gutter = 400;
        document.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal,
            WordEndnotePosition.DocumentEnd, WordNoteNumberRestart.Continuous, startNumber: 1);
        document.CompatibilitySettings.DoNotBalanceTextColumns = true;
        using var stream = new MemoryStream(document.ToBytes(format));
        using WordDocument loaded = WordDocument.Load(stream);
        Assert.True(loaded.CompatibilitySettings.DoNotBalanceTextColumns);
        loaded.CompatibilitySettings.DoNotBalanceTextColumns = false;
        using var clearedStream = new MemoryStream(loaded.ToBytes(format));
        using WordDocument cleared = WordDocument.Load(clearedStream);
        Assert.False(cleared.CompatibilitySettings.DoNotBalanceTextColumns);
        Assert.True(cleared.Settings.MirrorMargins);
        Assert.True(cleared.Settings.GutterAtTop);
        Assert.True(cleared.Settings.TrackRevisions);
        Assert.True(cleared.Sections[0].DifferentOddAndEvenPages);
        Assert.Equal(400U, cleared.Sections[0].Margins.Gutter);
        Assert.Equal(EndnotePositionValues.DocumentEnd,
            cleared.Sections[0].EndnoteProperties.EndnotePosition!.Val!.Value);
        Assert.Equal("Preserved options", Assert.Single(cleared.Paragraphs).Text);
        if (format == WordFileFormat.Docx)
            Assert.Equal(WordCompatibilityMode.Word2010, cleared.CompatibilitySettings.CompatibilityMode);
    }

    [Fact]
    public void Compatibility_ColumnBalancingSetterPreservesOtherCompatibilitySettingsAndSchemaOrder() {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        Settings settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        Compatibility compatibility = settings.GetFirstChild<Compatibility>()!;
        compatibility.AddChild(new UseWord2002TableStyleRules(), true);
        foreach (bool enabled in new[] { true, false, true }) {
            document.CompatibilitySettings.DoNotBalanceTextColumns = enabled;
            Assert.Equal(enabled, document.CompatibilitySettings.DoNotBalanceTextColumns);
            Assert.Equal(WordCompatibilityMode.Word2010, document.CompatibilitySettings.CompatibilityMode);
            Assert.NotNull(compatibility.GetFirstChild<UseWord2002TableStyleRules>());
            Assert.Empty(new OpenXmlValidator().Validate(settings));
        }
        Assert.Single(compatibility.Elements<NoColumnBalance>());
    }
}
