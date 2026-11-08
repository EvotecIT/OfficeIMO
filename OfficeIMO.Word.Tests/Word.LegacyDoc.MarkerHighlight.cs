using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_LoadReportsUnrepresentableListMarkerHighlightAndRetainsListFormatting() {
        using WordDocument source = CreateNativeListDefinitionControl();
        Level level = source._wordprocessingDocument!.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single().Elements<Level>().First();
        level.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new Bold(), new Italic());
        Assert.Empty(source.ValidateDocument());
        byte[] nativeBytes = source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
        Assert.True(OfficeCompoundFileReader.TryRead(nativeBytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"], table = (byte[])compound.Streams["1Table"].Clone();
        int firstLevel = BitConverter.ToInt32(word, 0x2E2) + BitConverter.ToInt32(word, 0x2E6);
        int chpx = firstLevel + 28 + table[firstLevel + 25];
        Assert.Equal(6, table[firstLevel + 24]);
        Assert.Equal(new byte[] { 0x35, 0x08, 1 }, table.Skip(chpx).Take(3));
        // Replace the bold SPRM with recognized highlight, retaining the italic SPRM.
        table[chpx] = 0x0C;
        table[chpx + 1] = 0x2A;
        table[chpx + 2] = 7;
        byte[] highlighted = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["1Table"] = table });
        LegacyDocDocument native = LegacyDocDocument.Load(new MemoryStream(highlighted));
        Assert.Equal(LegacyDocHighlightColorKind.Yellow, native.Numbering.Definitions.Single().Levels[0].CharacterFormat.Highlight);

        using WordDocument migrated = WordDocument.Load(new MemoryStream(highlighted), options: null, out OfficeConversionReport report);

        OfficeDiagnostic diagnostic = Assert.Single(report.Diagnostics,
            item => item.Code == "doc-numbering-invalid");
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Contains("highlight cannot be represented", diagnostic.Message);
        Level actual = migrated._wordprocessingDocument!.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single().Elements<Level>().First();
        Assert.NotNull(actual.NumberingSymbolRunProperties!.GetFirstChild<Italic>());
        Assert.Null(actual.NumberingSymbolRunProperties.GetFirstChild<Highlight>());
        var markers = WordDocumentTraversal.BuildListMarkers(migrated);
        Assert.Equal("12.", markers[migrated.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Equal("13.", markers[migrated.Paragraphs.First(paragraph => paragraph.Text == "Second")].Marker);
        Assert.Empty(migrated.ValidateDocument());
        Assert.Throws<NotSupportedException>(() => migrated.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_SaveRejectsListMarkerHighlightBeforeWritingDestination(bool explicitNone) {
        using WordDocument source = CreateNativeListDefinitionControl();
        Level level = source._wordprocessingDocument!.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single().Elements<Level>().First();
        level.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new Italic());
        Assert.Empty(source.ValidateDocument());
        level.NumberingSymbolRunProperties.Append(new Highlight { Val = explicitNone ? HighlightColorValues.None : HighlightColorValues.Yellow });
        byte[] original = { 17, 34, 51, 68 };
        using var destination = new MemoryStream();
        destination.Write(original, 0, original.Length);
        destination.Position = 2;

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            source.Save(destination, new WordSaveOptions { Format = WordFileFormat.Doc }));

        Assert.Contains("highlight in numbering symbol properties", exception.Message);
        Assert.Equal(original, destination.ToArray());
        Assert.Equal(2, destination.Position);
    }

    [Fact]
    public void LegacyDoc_MigrationPreservesParagraphMarkerHighlight() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Paragraph marker");
        paragraph._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphMarkRunProperties(
            new Bold(), new FontSize { Val = "28" }, new Highlight { Val = HighlightColorValues.Yellow },
            new Underline { Val = UnderlineValues.Single }));
        var sourceErrors = source.ValidateDocument();
        Assert.True(!sourceErrors.Any(), string.Join("\n", sourceErrors.Select(error => error.Description)));

        byte[] nativeBytes = source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
        LegacyDocDocument native = LegacyDocDocument.Load(new MemoryStream(nativeBytes));
        Assert.DoesNotContain(native.Diagnostics, diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error);
        LegacyDocCharacterFormat nativeFormat = Assert.IsType<LegacyDocCharacterFormat>(native.ParagraphFormats.First().ParagraphMarkFormat);
        Assert.Equal(LegacyDocHighlightColorKind.Yellow, nativeFormat.Highlight);

        using WordDocument migrated = WordDocument.Load(new MemoryStream(nativeBytes), options: null, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        AssertMarkerHighlight(migrated);
        migrated.Paragraphs[0].Text += " edited";
        string editedText = migrated.Paragraphs[0].Text;

        using WordDocument reopened = WordDocument.Load(new MemoryStream(migrated.ToBytes(new WordSaveOptions { Format = WordFileFormat.Docx })));
        Assert.Equal(editedText, reopened.Paragraphs[0].Text);
        AssertMarkerHighlight(reopened);
        Assert.Empty(reopened.ValidateDocument());

        void AssertMarkerHighlight(WordDocument document) {
            OpenXmlCompositeElement? properties = document.Paragraphs[0]._paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
            Assert.NotNull(properties);
            Assert.Equal(HighlightColorValues.Yellow, properties!.GetFirstChild<Highlight>()?.Val?.Value);
            Assert.NotNull(properties.GetFirstChild<Bold>());
            Assert.Equal("28", properties.GetFirstChild<FontSize>()?.Val?.Value);
            Assert.Equal(UnderlineValues.Single, properties.GetFirstChild<Underline>()?.Val?.Value);
        }
    }
}
