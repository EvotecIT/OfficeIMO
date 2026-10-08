using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeListDefinitions_SharedAuthoredIdentityKeepsTheCanonicalSequence(bool differentFormatting) {
        using WordDocument source = CreateNativeListDefinitionControl();
        WordList list = source.Lists.Single();
        list.AddItem("Third"); list.AddItem("Fourth");
        Numbering numbering = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance first = numbering.Elements<NumberingInstance>().Single();
        var definition = (AbstractNum)numbering.Elements<AbstractNum>().Single().CloneNode(true);
        definition.AbstractNumberId = 1;
        if (differentFormatting) {
            Level level = definition.Elements<Level>().First();
            level.StartNumberingValue!.Val = 7; level.LevelText!.Val = "%1)";
        }
        numbering.InsertBefore(definition, first);
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 2 });
        source.Paragraphs.First(paragraph => paragraph.Text == "Third")._paragraph.ParagraphProperties!
            .NumberingProperties!.NumberingId!.Val = 2;
        string before = numbering.OuterXml;
        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        string[] names = { "First", "Second", "Third", "Fourth" };
        Assert.Equal(new[] { "12.", "13.", "14.", "15." }, names.Select(name => markers[reopened.Paragraphs.First(paragraph => paragraph.Text == name)].Marker));
        Assert.Equal(before, numbering.OuterXml);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void NativeListDefinitions_AbsentOptionalRestartAttributeSupportsNativeRewrite() {
        using WordDocument source = CreateNativeListDefinitionControl();
        source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single().RemoveAttribute("restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml");
        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        using WordDocument rewritten = WordDocument.Load(new MemoryStream(reopened.ToBytes(WordFileFormat.Doc)));
        Assert.Equal("12.", WordDocumentTraversal.BuildListMarkers(rewritten)[rewritten.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Empty(rewritten.ValidateDocument());
    }

    [Fact]
    public void NativeListDefinitions_FormattingOnlyOverrideDoesNotCreateARestart() {
        using WordDocument source = CreateNativeListDefinitionControl();
        Numbering numbering = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance instance = numbering.Elements<NumberingInstance>().Single();
        Level level = (Level)numbering.Elements<AbstractNum>().Single().Elements<Level>().First().CloneNode(true);
        level.StartNumberingValue!.Val = 5;
        level.LevelText!.Val = "%1)";
        instance.Append(new LevelOverride(level) { LevelIndex = 0 });
        string before = numbering.OuterXml;
        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        NumberingInstance actual = reopened._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<NumberingInstance>().Single();
        LevelOverride actualOverride = Assert.Single(actual.Elements<LevelOverride>());
        Assert.Null(actualOverride.StartOverrideNumberingValue);
        Assert.Equal(12, actualOverride.Level!.StartNumberingValue!.Val!.Value);
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        Assert.Equal("12)", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Equal("13)", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "Second")].Marker);
        Assert.Equal(before, numbering.OuterXml);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void NativeListDefinitions_SectionBreakRestartFailsBeforeWriting() {
        using WordDocument source = CreateNativeListDefinitionControl();
        source.Lists.Single().RestartNumberingAfterBreak = true;
        source.AddSection();
        string before = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml;
        Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Equal(before, source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml);
    }

    [Fact]
    public void NativeListDefinitions_ExcessivePlaceholderCountFailsBeforeWriting() {
        using WordDocument source = CreateNativeListDefinitionControl();
        source.Lists.Single().Numbering.Levels[0].LevelText = "%1/%1";
        Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
    }

    [Fact]
    public void NativeListDefinitions_ExcessiveNativePlaceholderCountReportsLoss() {
        using WordDocument source = CreateNativeListDefinitionControl();
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"], table = (byte[])compound.Streams["1Table"].Clone();
        int level = BitConverter.ToInt32(word, 0x2E2) + BitConverter.ToInt32(word, 0x2E6);
        int text = level + 28 + table[level + 24] + table[level + 25] + 2;
        table[level + 7] = 2; // A second valid-position placeholder exceeds level zero's native count limit.
        table[text + 2] = 0; table[text + 3] = 0;
        byte[] altered = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["1Table"] = table });
        using WordDocument reopened = WordDocument.Load(new MemoryStream(altered));
        Assert.Contains(reopened.LegacyDocUnsupportedFeatures, item => item.Kind == OfficeIMO.Word.LegacyDoc.Model.LegacyDocUnsupportedFeatureKind.Numbering);
    }

    [Theory]
    [InlineData("instance")]
    [InlineData("level")]
    [InlineData("override")]
    public void NativeListDefinitions_SavingMissingListReferenceFailsBeforeProducingNativeBytes(string missing) {
        using WordDocument source = CreateNativeListDefinitionControl();
        Numbering numbering = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        if (missing == "instance") {
            numbering.RemoveAllChildren<NumberingInstance>();
        } else {
            foreach (Level level in numbering.Elements<AbstractNum>().Single().Elements<Level>().Skip(1).ToArray()) level.Remove();
            if (missing == "override") numbering.Elements<NumberingInstance>().Single()
                .Append(new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 1 });
            else source.Paragraphs.First(paragraph => paragraph.Text == "First")
                ._paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val = 1;
        }
        string before = numbering.OuterXml;
        Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Equal(before, numbering.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeListDefinitions_MissingOrMalformedTablesReportNumberingLoss(bool malformed) {
        using WordDocument source = CreateNativeListDefinitionControl();
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = (byte[])compound!.Streams["WordDocument"].Clone();
        BitConverter.GetBytes(malformed ? 29 : 0).CopyTo(word, 0x2E6);
        if (!malformed) BitConverter.GetBytes(0).CopyTo(word, 0x2EE);
        byte[] altered = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["WordDocument"] = word });
        using WordDocument reopened = WordDocument.Load(new MemoryStream(altered));
        Assert.Contains(reopened.LegacyDocUnsupportedFeatures, item => item.Kind == OfficeIMO.Word.LegacyDoc.Model.LegacyDocUnsupportedFeatureKind.Numbering);
        Assert.Throws<NotSupportedException>(() => reopened.ToBytes(WordFileFormat.Doc));
    }

    [Theory]
    [InlineData("bullet", "•", "•")]
    [InlineData("lowerRoman", "xii.", "xiii.")]
    [InlineData("none", "", "")]
    public void NativeListDefinitions_NativeFormatsRetainTheirMarkers(string formatName, string first, string second) {
        using WordDocument source = CreateNativeListDefinitionControl();
        Level level = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single().Elements<Level>().First();
        NumberFormatValues format = formatName == "bullet" ? NumberFormatValues.Bullet
            : formatName == "none" ? NumberFormatValues.None : NumberFormatValues.LowerRoman;
        level.NumberingFormat = new NumberingFormat { Val = format };
        if (format == NumberFormatValues.Bullet || format == NumberFormatValues.None)
            level.LevelText = new LevelText { Val = first };
        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Level actual = reopened._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().First().Elements<Level>().First();
        Assert.Equal(format, actual.NumberingFormat!.Val!.Value);
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        Assert.Equal(first, markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Equal(second, markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "Second")].Marker);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeListDefinitions_LevelPropertyLengthExcludesParagraphPagePadding(bool enabled) {
        using WordDocument source = CreateNativeListDefinitionControl();
        AbstractNum definition = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single();
        definition.Elements<Level>().First().PreviousParagraphProperties = new PreviousParagraphProperties(
            new ContextualSpacing { Val = enabled });
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"], table = compound.Streams["1Table"];
        int firstLevel = BitConverter.ToInt32(word, 0x2E2) + BitConverter.ToInt32(word, 0x2E6);
        Assert.Equal(3, table[firstLevel + 25]);
        Assert.Equal((byte)(enabled ? 1 : 0), table[firstLevel + 30]);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        Level level = reopened._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().First().Elements<Level>().First();
        Assert.Equal(enabled, level.PreviousParagraphProperties!.GetFirstChild<ContextualSpacing>()!.Val!.Value);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(9)]
    public void NativeListDefinitions_ExplicitInstanceStartOverridesAbstractStart(int levelCount) {
        using WordDocument source = CreateNativeListDefinitionControl();
        NumberingInstance instance = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<NumberingInstance>().Single();
        for (int level = 0; level < levelCount; level++)
            instance.Append(new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = level });
        string before = instance.OuterXml;
        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        Assert.Equal("1.", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Equal("2.", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "Second")].Marker);
        Assert.Equal(before, instance.OuterXml);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void NativeListDefinitions_SparseInstanceIdsKeepTheirParagraphReferences() {
        using WordDocument source = CreateNativeListDefinitionControl();
        NumberingInstance instance = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<NumberingInstance>().Single();
        instance.NumberID = 7;
        foreach (WordParagraph paragraph in source.Paragraphs.Where(paragraph => paragraph.Text == "First" || paragraph.Text == "Second"))
            paragraph._paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 7;
        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        Assert.Equal("12.", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Equal("13.", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "Second")].Marker);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void NativeListDefinitions_SaveAndImportPreserveAbstractStart() {
        using WordDocument source = CreateNativeListDefinitionControl();
        string numberingBefore = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml;
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        Assert.Equal("12.", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Equal("13.", markers[reopened.Paragraphs.First(paragraph => paragraph.Text == "Second")].Marker);
        Assert.Equal(numberingBefore, source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void NativeListDefinitions_SaveWritesDefinitionAndInstanceTables() {
        using WordDocument source = CreateNativeListDefinitionControl();
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"], table = compound.Streams["1Table"];
        int listsOffset = BitConverter.ToInt32(word, 0x2E2), listsLength = BitConverter.ToInt32(word, 0x2E6);
        int instancesOffset = BitConverter.ToInt32(word, 0x2EA), instancesLength = BitConverter.ToInt32(word, 0x2EE);
        Assert.True(listsLength >= 30, "Native Word paragraphs need a real PlfLst definition.");
        Assert.True(listsOffset >= 0 && listsOffset <= table.Length - listsLength);
        Assert.Equal(1, BitConverter.ToUInt16(table, listsOffset));
        Assert.True(instancesLength >= 24, "Native Word list indices need a matching PlfLfo instance.");
        Assert.True(instancesOffset >= 0 && instancesOffset <= table.Length - instancesLength);
        Assert.Equal(1u, BitConverter.ToUInt32(table, instancesOffset));
        Assert.Equal(BitConverter.ToInt32(table, listsOffset + 2), BitConverter.ToInt32(table, instancesOffset + 4));
    }

    private static WordDocument CreateNativeListDefinitionControl() {
        WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Numbered);
        list.Numbering.Levels[0].StartNumberingValue = 12;
        list.AddItem("First"); list.AddItem("Second");
        // Isolate the native codec contract from the separate fresh-DOCX-list precedence fix.
        document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<NumberingInstance>().Single().RemoveAllChildren<LevelOverride>();
        return document;
    }
}
