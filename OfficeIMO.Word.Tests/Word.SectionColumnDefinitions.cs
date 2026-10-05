using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(".doc")]
    [InlineData(".docx")]
    public void LegacyDoc_SectionColumnDefinitions_OmittedIndividualGapRetainsItsEffectiveZero(string extension) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        try {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Unequal columns with inherited gap");
            document.Sections[0].ColumnsSpace = 720;
            document.Sections[0].ColumnDefinitions = new[] { new WordSectionColumn(3000), new WordSectionColumn(4000, 0) };
            document.Save(path);
            Assert.Null(document.Sections[0].ColumnDefinitions[0].SpaceAfterTwips);
            using WordDocument reloaded = WordDocument.Load(path);
            Assert.Equal(720, reloaded.Sections[0].ColumnsSpace);
            Assert.Equal(new[] { 3000, 4000 }, reloaded.Sections[0].ColumnDefinitions.Select(column => column.WidthTwips));
            Assert.Equal(0, reloaded.Sections[0].ColumnDefinitions[0].SpaceAfterTwips ?? 0);
            Assert.Equal(0, reloaded.Sections[0].ColumnDefinitions[1].SpaceAfterTwips);
            if (extension == ".docx") Assert.Null(reloaded.Sections[0].ColumnDefinitions[0].SpaceAfterTwips);
        } finally { DeleteIfExists(path); }
    }

    [Theory]
    [InlineData(".doc")]
    [InlineData(".docx")]
    public void SectionColumnDefinitions_SaveAndReloadPreservesWidthsAndIndividualGaps(string extension) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        try {
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Unequal section columns");
                WordSection section = document.Sections[0];
                section.ColumnCount = 2;
                section.ColumnsSpace = 720;
                section.ColumnDefinitions = new[] {
                    new WordSectionColumn(3000, 360), new WordSectionColumn(4000, 0)
                };
                document.Save(path);
            }
            using WordDocument reloaded = WordDocument.Load(path);
            Columns actual = reloaded.Sections[0]._sectionProperties.GetFirstChild<Columns>()!;
            Assert.Equal(2, reloaded.Sections[0].ColumnCount);
            Assert.False(actual.EqualWidth?.Value ?? true);
            Assert.Equal(new[] { "3000", "4000" }, actual.Elements<Column>().Select(column => column.Width?.Value));
            Assert.Equal(new[] { "360", "0" }, actual.Elements<Column>().Select(column => column.Space?.Value));
            Assert.Equal(new[] { 3000, 4000 }, reloaded.Sections[0].ColumnDefinitions.Select(column => column.WidthTwips));
            Assert.Equal(new int?[] { 360, 0 }, reloaded.Sections[0].ColumnDefinitions.Select(column => column.SpaceAfterTwips));
            Assert.Contains("Unequal section columns", reloaded.Paragraphs[0].Text);
        } finally {
            DeleteIfExists(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SectionColumnDefinitions_LoadHonorsIndexedOperandsAndFinalEqualWidthFlag(bool equalWidthOverride) {
        using WordDocument document = WordDocument.Load(new MemoryStream(
            LegacyDocTestBuilder.CreateIndexedSectionColumnDoc(equalWidthOverride)));
        Columns columns = document.Sections[0]._sectionProperties.GetFirstChild<Columns>()!;
        if (equalWidthOverride) {
            Assert.Empty(columns.Elements<Column>());
        } else {
            Assert.False(columns.EqualWidth?.Value ?? true);
            Assert.Equal(new[] { "3000", "4000" }, columns.Elements<Column>().Select(column => column.Width?.Value));
            Assert.Equal(new[] { "360", "0" }, columns.Elements<Column>().Select(column => column.Space?.Value));
        }
        Assert.Equal(2, document.Sections[0].ColumnCount);
    }

    [Fact]
    public void SectionColumnDefinitions_NativeSaveRejectsColumnCountBeyondBinaryFormatRange() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Invalid native column count");
            document.Sections[0].ColumnCount = 45;
            Assert.Throws<NotSupportedException>(() => document.Save(path));
            Assert.False(File.Exists(path));
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void SectionColumnDefinitions_SnapshotAndCountChangesRemainCoherent() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        var definitions = new[] { new WordSectionColumn(3000, 360), new WordSectionColumn(4000) };
        section.ColumnDefinitions = definitions;
        definitions[0] = new WordSectionColumn(5000);
        Assert.Equal(3000, section.ColumnDefinitions[0].WidthTwips);
        section.ColumnCount = null;
        Assert.Equal(2, section.ColumnCount);
        Assert.Throws<InvalidOperationException>(() => section.ColumnCount = 3);
        Assert.Equal(2, section.ColumnDefinitions.Count);
        Assert.Equal(2, section.ColumnCount);
        section.ColumnDefinitions = Array.Empty<WordSectionColumn>();
        Assert.Empty(section.ColumnDefinitions);
        Assert.Equal(2, section.ColumnCount);
        section.ColumnCount = 3;
        Assert.Equal(3, section.ColumnCount);
    }

    [Theory]
    [InlineData(717, 0)]
    [InlineData(32768, 0)]
    [InlineData(3000, 32768)]
    public void SectionColumnDefinitions_NativeSaveRejectsUnrepresentableDimensionsBeforeCreatingOutput(int width, int gap) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Unrepresentable native columns");
            document.Sections[0].ColumnDefinitions = new[] { new WordSectionColumn(width, gap), new WordSectionColumn(4000) };
            Assert.Throws<NotSupportedException>(() => document.Save(path));
            Assert.False(File.Exists(path));
        } finally {
            DeleteIfExists(path);
        }
    }

    [Theory]
    [InlineData("missing-width")]
    [InlineData("invalid-index")]
    [InlineData("truncated")]
    public void SectionColumnDefinitions_IncompleteIndexedNativeLayoutReportsAnImportWarning(string fault) {
        using var result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(
            LegacyDocTestBuilder.CreateIndexedSectionColumnDoc(false, fault)));
        Assert.True(result.HasDocument);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DOC-SEPX-INVALID");
        Assert.Empty(result.Document!.Sections[0].ColumnDefinitions);
        Assert.Contains("Indexed section columns", result.Document.Paragraphs[0].Text);
    }

    [Fact]
    public void SectionColumnDefinitions_LoadIgnoresCachedColumnsBeyondTheActiveCount() {
        using var result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(
            LegacyDocTestBuilder.CreateIndexedSectionColumnDoc(false, "unused-cache")));
        Assert.True(result.HasDocument);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "DOC-SEPX-INVALID");
        Assert.Equal(new[] { 3000, 4000 }, result.Document!.Sections[0].ColumnDefinitions.Select(column => column.WidthTwips));
        Assert.Equal(new int?[] { 360, 0 }, result.Document.Sections[0].ColumnDefinitions.Select(column => column.SpaceAfterTwips));
        Assert.Equal(2, result.Document.Sections[0].ColumnCount);
    }

    private static partial class LegacyDocTestBuilder {
        internal static byte[] CreateIndexedSectionColumnDoc(bool equalWidthOverride, string? fault = null) {
            const string text = "Indexed section columns\r";
            const int sepxOffset = 0x300;
            byte[] tableStream = CreateTableStream(text.Length);
            int fcPlcfSed = tableStream.Length;
            byte[] plc = CreateOneSectionDescriptorPlc(text.Length, sepxOffset);
            Array.Resize(ref tableStream, tableStream.Length + plc.Length);
            Buffer.BlockCopy(plc, 0, tableStream, fcPlcfSed, plc.Length);
            byte[] wordStream = CreateWordDocumentStream(text, fcPlcfSed: fcPlcfSed, lcbPlcfSed: plc.Length);
            var grpprl = new List<byte>(CreateSectionSepx(columnCount: 2, columnSpacing: 720).Skip(2));
            grpprl.AddRange(new byte[] { 0x05, 0x30, 0 });
            void Indexed(ushort sprm, byte index, ushort value) =>
                grpprl.AddRange(new[] { (byte)sprm, (byte)(sprm >> 8), index, (byte)value, (byte)(value >> 8) });
            Indexed(0xF203, 0, 2500);
            if (fault != "missing-width") Indexed(0xF203, 1, 4000);
            Indexed(0xF204, 0, 360);
            Indexed(0xF203, 0, 3000);
            if (equalWidthOverride) grpprl.AddRange(new byte[] { 0x05, 0x30, 1 });
            if (fault == "invalid-index") Indexed(0xF203, 44, 2000);
            if (fault == "truncated") grpprl.AddRange(new byte[] { 0x03, 0xF2, 0, 0xB8 });
            if (fault == "unused-cache") {
                Indexed(0xF203, 2, 4000);
                Indexed(0xF204, 2, 720);
            }
            byte[] sepx = new byte[grpprl.Count + 2];
            sepx[0] = (byte)grpprl.Count;
            sepx[1] = (byte)(grpprl.Count >> 8);
            grpprl.CopyTo(sepx, 2);
            WriteBytesAt(ref wordStream, sepxOffset, sepx);
            using var package = new MemoryStream();
            using (RootStorage root = RootStorage.Create(package, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen)) {
                WriteStream(root, "WordDocument", wordStream);
                WriteStream(root, "1Table", tableStream);
            }
            return package.ToArray();
        }
    }
}
