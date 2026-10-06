using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_DocumentOptionsHaveTheDeclaredFormatsCompleteSize(bool nested) {
        using WordDocument source = CreateLegacyMetadataDocument(nested);
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        byte[] fib = ReadCompoundStream(bytes, "WordDocument");
        byte[] table = ReadCompoundStream(bytes, "1Table");
        int offset = BitConverter.ToInt32(fib, 0x192);
        int length = BitConverter.ToInt32(fib, 0x196);
        Assert.Equal(nested ? 544 : 500, length);
        Assert.InRange(offset, 1, table.Length - length);
        Assert.Equal(720, BitConverter.ToUInt16(table, offset + 10));
        Assert.Equal(1, BitConverter.ToUInt16(table, offset + 2) >> 2);
        Assert.Equal(1u, (BitConverter.ToUInt32(table, offset + 52) >> 2) & 0x3FFF);
        Assert.Equal(3u, (BitConverter.ToUInt32(table, offset + 52) >> 16) & 3);
        Assert.Equal(BitConverter.ToUInt16(table, offset + 8), BitConverter.ToUInt16(table, offset + 84));
        if (nested) Assert.Equal(BitConverter.ToUInt32(table, offset + 84), BitConverter.ToUInt32(table, offset + 508));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_CustomDefaultTabStopSurvivesNativeRoundTrips(bool nested) {
        using WordDocument source = CreateLegacyMetadataDocument(nested);
        source.Settings.DefaultTabStop = 960;
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        for (int cycle = 0; cycle < 2; cycle++) {
            using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
            Assert.Equal(960, restored.Settings.DefaultTabStop);
            byte[] fib = ReadCompoundStream(bytes, "WordDocument");
            byte[] table = ReadCompoundStream(bytes, "1Table");
            Assert.Equal(960, BitConverter.ToUInt16(table, BitConverter.ToInt32(fib, 0x192) + 10));
            bytes = restored.ToBytes(WordFileFormat.Doc);
        }
    }

    [Fact]
    public void LegacyDoc_RequiredAssociatedStringsDoNotReplaceSummaryProperties() {
        using WordDocument source = CreateLegacyMetadataDocument(false);
        source.BuiltinDocumentProperties.Title = "Document metadata";
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        byte[] fib = ReadCompoundStream(bytes, "WordDocument");
        byte[] table = ReadCompoundStream(bytes, "1Table");
        int offset = BitConverter.ToInt32(fib, 0x19A);
        int length = BitConverter.ToInt32(fib, 0x19E);
        Assert.Equal(42, length);
        Assert.InRange(offset, 1, table.Length - length);
        Assert.Equal(0xFFFF, BitConverter.ToUInt16(table, offset));
        Assert.Equal(18, BitConverter.ToUInt16(table, offset + 2));
        Assert.Equal(0, BitConverter.ToUInt16(table, offset + 4));
        for (int index = 0; index < 18; index++) Assert.Equal(0, BitConverter.ToUInt16(table, offset + 6 + index * 2));
        using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal("Document metadata", restored.BuiltinDocumentProperties.Title);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_DocumentOptionsCarryTheNoteSettingsUsedByWord(bool nested) {
        using WordDocument source = CreateLegacyMetadataDocument(nested);
        source.Sections[0].AddFootnoteProperties(WordNumberFormat.UpperLetter, WordFootnotePosition.BeneathText,
            WordNoteNumberRestart.EachPage, startNumber: 3);
        source.Sections[0].AddEndnoteProperties(WordNumberFormat.LowerLetter, WordEndnotePosition.DocumentEnd,
            WordNoteNumberRestart.EachSection, startNumber: 9);
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        byte[] fib = ReadCompoundStream(bytes, "WordDocument");
        byte[] table = ReadCompoundStream(bytes, "1Table");
        int offset = BitConverter.ToInt32(fib, 0x192);
        Assert.Equal(2, (BitConverter.ToUInt16(table, offset) >> 5) & 3);
        Assert.Equal((3 << 2) | 2, BitConverter.ToUInt16(table, offset + 2));
        Assert.Equal((3u << 16) | (9u << 2) | 1, BitConverter.ToUInt32(table, offset + 52));
        Assert.Equal(3, BitConverter.ToUInt16(table, offset + 492));
        Assert.Equal(4, BitConverter.ToUInt16(table, offset + 494));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_ImportsDocumentNoteSettingsWithoutSectionSprms(bool nested) {
        using WordDocument source = CreateLegacyMetadataDocument(nested);
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        byte[] fib = ReadCompoundStream(bytes, "WordDocument");
        byte[] table = ReadCompoundStream(bytes, "1Table");
        int offset = BitConverter.ToInt32(fib, 0x192);
        BitConverter.GetBytes((ushort)0x0040).CopyTo(table, offset);
        BitConverter.GetBytes((ushort)((3 << 2) | 2)).CopyTo(table, offset + 2);
        BitConverter.GetBytes((3u << 16) | (9u << 2) | 1).CopyTo(table, offset + 52);
        BitConverter.GetBytes((ushort)3).CopyTo(table, offset + 492);
        BitConverter.GetBytes((ushort)4).CopyTo(table, offset + 494);
        using var package = new MemoryStream();
        using (RootStorage root = RootStorage.Create(package, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen)) {
            using (CfbStream stream = root.CreateStream("WordDocument")) stream.Write(fib, 0, fib.Length);
            using (CfbStream stream = root.CreateStream("1Table")) stream.Write(table, 0, table.Length);
        }
        using WordDocument restored = WordDocument.Load(new MemoryStream(package.ToArray()));
        var footnotes = restored.FootnoteSettings;
        Assert.Equal(WordFootnotePosition.BeneathText, footnotes.Position);
        Assert.Equal(WordNoteNumberRestart.EachPage, footnotes.NumberingRestart);
        Assert.Equal(3, footnotes.StartNumber);
        Assert.Equal(WordNumberFormat.UpperLetter, footnotes.NumberingFormat);
        Assert.Equal(WordNoteNumberRestart.EachSection, restored.EndnoteSettings.NumberingRestart);
        Assert.Equal(9, restored.EndnoteSettings.StartNumber);
        Assert.Equal(WordNumberFormat.LowerLetter, restored.EndnoteSettings.NumberingFormat);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_RejectsDifferentSectionNoteSettingsInsteadOfSilentlyChangingThem(bool endnotes) {
        using WordDocument source = CreateLegacyMetadataDocument(false);
        WordSection second = source.AddSection();
        second.AddParagraph("SECOND SECTION");
        if (endnotes) second.AddEndnoteProperties(numberingFormat: WordNumberFormat.UpperLetter);
        else second.AddFootnoteProperties(numberingFormat: WordNumberFormat.UpperLetter);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Contains("whole document", error.Message);
    }

    private static WordDocument CreateLegacyMetadataDocument(bool nested) {
        WordDocument document = WordDocument.Create();
        document.AddParagraph("BEFORE\tAFTER");
        if (nested) {
            WordTable outer = document.AddTable(1, 1, WordTableStyle.TableNormal);
            outer.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0].Text = "NESTED";
        }
        return document;
    }
}
