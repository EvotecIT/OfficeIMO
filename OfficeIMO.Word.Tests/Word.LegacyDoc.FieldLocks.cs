using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_FieldLocks_EditHistoryRetainsCacheWithoutInferringInvalidation() {
        byte[] original = CreateLockedNativeDoc();
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"].ToArray();
        int start = BitConverter.ToInt32(main, 0x11A);
        int records = (BitConverter.ToInt32(main, 0x11E) - 4) / 6;
        int properties = start + (records + 1) * 4;
        table[properties + (records - 1) * 2 + 1] |= 0x0C;
        using WordDocument document = WordDocument.Load(new MemoryStream(RebuildFieldMetadataCarrier(compound, main, table)));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        Assert.Contains(document.LegacyDocImportDiagnostics, finding => finding.Code == "doc-field-cache-edit-history");
        W.SimpleField field = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.Document.Descendants<W.SimpleField>());
        Assert.True(field.FieldLock?.Value);
        Assert.False(field.Dirty?.Value ?? false);
        Assert.Equal("FIRSTSECOND", field.InnerText);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        W.SimpleField saved = Assert.Single(reopened._wordprocessingDocument.MainDocumentPart!.Document.Descendants<W.SimpleField>());
        Assert.True(saved.FieldLock?.Value);
        Assert.False(saved.Dirty?.Value ?? false);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_FieldLocks_NativeSaveRejectsUnqualifiedInvalidationRatherThanDiscardingIt(bool complex) {
        using WordDocument source = WordDocument.Create();
        AppendLockedDateCache(source.AddParagraph()._paragraph, complex, true, "FIRST", "SECOND");
        if (complex) source._wordprocessingDocument.MainDocumentPart!.Document.Descendants<W.FieldChar>()
            .First(field => field.FieldCharType == W.FieldCharValues.Begin).Dirty = true;
        else Assert.Single(source._wordprocessingDocument.MainDocumentPart!.Document.Descendants<W.SimpleField>()).Dirty = true;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.Save(path));
            Assert.Contains("dirty", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(path));
        } finally { DeleteIfExists(path); }
    }

    [Theory]
    [InlineData(0x01, "doc-field-display-approximation")]
    [InlineData(0x02, "doc-field-state-unqualified")]
    [InlineData(0x20, "doc-field-state-unqualified")]
    public void LegacyDoc_FieldLocks_UnmodeledBehaviorReportsStrictLoss(int flag, string code) {
        byte[] original = CreateLockedNativeDoc();
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"].ToArray();
        int start = BitConverter.ToInt32(main, 0x11A);
        int records = (BitConverter.ToInt32(main, 0x11E) - 4) / 6;
        int properties = start + (records + 1) * 4;
        int globalEnd = BitConverter.ToInt32(table, start + (records - 1) * 4);
        table[properties + (records - 1) * 2 + 1] |= (byte)flag;
        using WordDocument document = WordDocument.Load(new MemoryStream(RebuildFieldMetadataCarrier(compound, main, table)));
        var finding = Assert.Single(document.LegacyDocUnsupportedFeatures, feature => feature.Code == code);
        Assert.Contains(globalEnd.ToString(), finding.Description);
        using var output = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_FieldLocks_SurviveNativeModernAndNativeResaveAcrossStories(bool complex) {
        using WordDocument source = WordDocument.Create();
        foreach (W.Paragraph paragraph in CreateFieldFormattingStories(source)) {
            AppendLockedDateCache(paragraph, complex, locked: true, "FIRST", "SECOND");
            AppendLockedDateCache(paragraph, complex, locked: false, "OPEN", "CACHE");
        }
        string first = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        string second = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            source.Save(first);
            using WordDocument imported = WordDocument.Load(first);
            AssertFieldLocks(imported);
            using var modern = new MemoryStream();
            imported.Save(modern);
            modern.Position = 0;
            using WordDocument reopenedModern = WordDocument.Load(modern);
            AssertFieldLocks(reopenedModern);
            reopenedModern.Save(second);
            using WordDocument reopenedNative = WordDocument.Load(second);
            AssertFieldLocks(reopenedNative);
        } finally { DeleteIfExists(first); DeleteIfExists(second); }

        void AssertFieldLocks(WordDocument document) {
            W.SimpleField[] fields = EnumerateFieldFormattingRoots(document._wordprocessingDocument.MainDocumentPart!)
                .SelectMany(root => root.Descendants<W.SimpleField>()).ToArray();
            Assert.Equal(12, fields.Length);
            Assert.Equal(6, fields.Count(field => field.FieldLock?.Value == true));
            Assert.All(fields, field => {
                Assert.Equal("DATE", field.Instruction!.Value!.Trim());
                bool locked = field.InnerText == "FIRSTSECOND";
                Assert.True(locked || field.InnerText == "OPENCACHE");
                Assert.Equal(locked, field.FieldLock?.Value ?? false);
                W.Run last = field.Descendants<W.Run>().Last();
                Assert.NotNull(last.RunProperties?.Bold);
                Assert.True(last.RunProperties.Bold.Val?.Value ?? true);
            });
            Assert.Empty(document.ValidateDocument());
        }
    }

    [Theory]
    [InlineData("range")]
    [InlineData("length")]
    [InlineData("cp-order")]
    [InlineData("delimiter")]
    [InlineData("has-separator")]
    public void LegacyDoc_FieldLocks_InvalidDeclaredMetadataReportsLossWithoutGuessing(string mutation) {
        byte[] original = CreateLockedNativeDoc();
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        byte[] table = compound.Streams["1Table"].ToArray();
        int start = BitConverter.ToInt32(main, 0x11A);
        int length = BitConverter.ToInt32(main, 0x11E);
        int records = (length - 4) / 6;
        int properties = start + (records + 1) * 4;
        switch (mutation) {
            case "range": WriteInt(main, 0x11A, table.Length + 1); break;
            case "length": WriteInt(main, 0x11E, length - 1); break;
            case "cp-order": WriteInt(table, start + 4, BitConverter.ToInt32(table, start)); break;
            case "delimiter": table[properties] = 0x12; break;
            case "has-separator": table[properties + (records - 1) * 2 + 1] &= 0x7F; break;
        }
        byte[] input = RebuildFieldMetadataCarrier(compound, main, table);
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Contains(document.LegacyDocUnsupportedFeatures, finding => finding.Code == "doc-field-plc-invalid");
        using var output = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public void LegacyDoc_FieldLocks_IgnoreReservedCharacterBitsSeparatorByteAndTerminalExtent() {
        byte[] original = CreateLockedNativeDoc();
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"].ToArray();
        int start = BitConverter.ToInt32(main, 0x11A);
        int records = (BitConverter.ToInt32(main, 0x11E) - 4) / 6;
        int properties = start + (records + 1) * 4;
        for (int index = 0; index < records; index++) {
            table[properties + index * 2] |= 0xE0;
            if ((table[properties + index * 2] & 0x1F) == 0x14) table[properties + index * 2 + 1] = 0xFF;
        }
        WriteInt(table, start + records * 4, int.MaxValue);
        using WordDocument document = WordDocument.Load(new MemoryStream(RebuildFieldMetadataCarrier(compound, main, table)));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        W.SimpleField field = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.Document.Descendants<W.SimpleField>());
        Assert.True(field.FieldLock?.Value);
        Assert.Equal("FIRSTSECOND", field.InnerText);
        using var output = new MemoryStream();
        document.Save(output);
        Assert.True(output.Length > 0);
    }

    private static byte[] CreateLockedNativeDoc() {
        using WordDocument source = WordDocument.Create();
        AppendLockedDateCache(source.AddParagraph()._paragraph, false, true, "FIRST", "SECOND");
        return source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
    }

    private static byte[] RebuildFieldMetadataCarrier(OfficeCompoundFile compound, byte[] main, byte[] table) =>
        OfficeCompoundFileWriter.Write(compound.Streams.Select(stream => new OfficeCompoundStream(stream.Key,
            stream.Key == "WordDocument" ? main : stream.Key == "1Table" ? table : stream.Value)).ToArray());

    private static void AppendLockedDateCache(W.Paragraph paragraph, bool complex, bool locked, string first, string second) {
        var result = new OpenXmlElement[] { new W.Run(new W.Text(first)), new W.Run(new W.RunProperties(new W.Bold()), new W.Text(second)) };
        if (!complex) { paragraph.Append(new W.SimpleField(result) { Instruction = " DATE ", FieldLock = locked }); return; }
        paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin, FieldLock = locked }),
            new W.Run(new W.FieldCode(" DATE ")), new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }));
        paragraph.Append(result);
        paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
    }
}
