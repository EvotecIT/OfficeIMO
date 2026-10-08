using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_PageLayoutSettingsPreserveMarginsAndOtherDocumentFlags(bool mirrorMargins, bool gutterAtTop) {
        using WordDocument document = WordDocument.Create();
        document.Settings.MirrorMargins = mirrorMargins;
        document.Settings.GutterAtTop = gutterAtTop;
        document.Settings.TrackRevisions = true;
        document.Sections[0].DifferentOddAndEvenPages = true;
        document.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal,
            WordEndnotePosition.DocumentEnd, WordNoteNumberRestart.Continuous, startNumber: 1);
        document.Sections[0].Margins.Gutter = 400;
        document.AddParagraph("Margin settings");
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);

        // Check the MS-DOC wire contract independently of the importer.
        byte[] wordStream = ReadCompoundStream(bytes, "WordDocument");
        byte[] tableStream = ReadCompoundStream(bytes, "1Table");
        int dopOffset = BitConverter.ToInt32(wordStream, 0x192);
        int dopLength = BitConverter.ToInt32(wordStream, 0x196);
        Assert.InRange(dopOffset, 0, tableStream.Length - dopLength);
        Assert.True(dopLength >= (gutterAtTop ? 84 : 56));
        Assert.Equal(mirrorMargins, (tableStream[dopOffset + 6] & 0x20) != 0);
        if (dopLength >= 84) Assert.Equal(gutterAtTop, (tableStream[dopOffset + 83] & 0x80) != 0);

        using var source = new MemoryStream(bytes);
        using WordDocument loaded = WordDocument.Load(source);
        Assert.Equal(mirrorMargins, loaded.Settings.MirrorMargins);
        Assert.Equal(gutterAtTop, loaded.Settings.GutterAtTop);
        Assert.True(loaded.Settings.TrackRevisions);
        Assert.True(loaded.Sections[0].DifferentOddAndEvenPages);
        Assert.Equal(400U, loaded.Sections[0].Margins.Gutter);
        Assert.Equal(EndnotePositionValues.DocumentEnd,
            loaded.Sections[0].EndnoteProperties.EndnotePosition!.Val!.Value);
        Assert.Equal("Margin settings", Assert.Single(loaded.Paragraphs).Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageLayout_ImplicitOnOffValuesRemainEnabledThroughNativeDoc(bool topGutter) {
        using WordDocument document = WordDocument.Create();
        if (topGutter) {
            var settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            settings.AddChild(new GutterAtTop(), true);
            Assert.True(document.Settings.GutterAtTop);
        } else {
            document.Sections[0]._sectionProperties.AddChild(new GutterOnRight(), true);
            Assert.True(document.Sections[0].RtlGutter);
        }
        document.AddParagraph("Implicit gutter");
        using var source = new MemoryStream(document.ToBytes(WordFileFormat.Doc));
        using WordDocument loaded = WordDocument.Load(source);
        Assert.Equal(topGutter, loaded.Settings.GutterAtTop);
        Assert.Equal(!topGutter, loaded.Sections[0].RtlGutter);
        loaded.Settings.GutterAtTop = false;
        loaded.Sections[0].RtlGutter = false;
        Assert.False(loaded.Settings.GutterAtTop);
        Assert.False(loaded.Sections[0].RtlGutter);
    }

    [Fact]
    public void PageLayout_RightGutterInsertionRetainsSectionSchemaOrder() {
        using WordDocument document = WordDocument.Create();
        var properties = document.Sections[0]._sectionProperties;
        properties.AddChild(new DocGrid { Type = DocGridValues.Lines }, true);
        var validator = new OpenXmlValidator();
        Assert.Empty(validator.Validate(properties));
        document.Sections[0].RtlGutter = true;
        Assert.Empty(validator.Validate(properties));
    }
}
