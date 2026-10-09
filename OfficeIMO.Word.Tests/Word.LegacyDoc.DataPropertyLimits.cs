using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Model;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_DataPropertyLimits_AllFormattingPagesShareConfiguredBudget(bool huge) {
        byte[] bytes = LegacyDocTestBuilder.CreateDocWithSharedDataProperties(huge);
        LegacyDocDocument complete = LegacyDocDocument.Load(bytes);
        Assert.Equal(new[] { "first", "second", "third" }, complete.Paragraphs);
        Assert.All(complete.ParagraphFormats, format => Assert.True(format.KeepLinesTogether));
        Assert.DoesNotContain(complete.Diagnostics, diagnostic => diagnostic.Code == "DOC-PAPX-INVALID");

        var options = new LegacyDocImportOptions { MaxParagraphPropertyWorkBytes = 1100, ReportUnsupportedContent = false };
        LegacyDocDocument limited = LegacyDocDocument.Load(new MemoryStream(bytes), options);
        Assert.Equal(complete.Text, limited.Text);
        Assert.True(limited.ParagraphFormats[0].KeepLinesTogether);
        Assert.True(limited.ParagraphFormats[1].KeepLinesTogether);
        Assert.Null(limited.ParagraphFormats[2].KeepLinesTogether);
        Assert.Contains(limited.Diagnostics, diagnostic => diagnostic.Code == "DOC-PAPX-INVALID"
            && diagnostic.Message.IndexOf("work limit", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void LegacyDoc_DataPropertyLimits_StyleOnlyPapxConsumesWork() {
        byte[] bytes = LegacyDocTestBuilder.CreateDocWithSharedDataProperties(false, styleOnly: true);
        var complete = LegacyDocDocument.Load(bytes);
        Assert.All(complete.ParagraphFormats, format => Assert.Equal((ushort)1, format.StyleIndex));
        var limited = LegacyDocDocument.Load(bytes, new LegacyDocImportOptions { MaxParagraphPropertyWorkBytes = 1 });
        Assert.Equal(complete.Text, limited.Text);
        Assert.Contains(limited.Diagnostics, diagnostic => diagnostic.Code == "DOC-PAPX-INVALID"
            && diagnostic.Message.IndexOf("work limit", StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_DataPropertyLimits_OverflowingReferencesPreserveTextAndReportDiagnostic(bool tableSpan) {
        byte[] bytes = LegacyDocTestBuilder.CreateDocWithSharedDataProperties(false,
            pageNumber: tableSpan ? null : int.MaxValue, overflowingTableSpan: tableSpan);
        var model = LegacyDocDocument.Load(bytes);
        Assert.Equal(new[] { "first", "second", "third" }, model.Paragraphs);
        Assert.Contains(model.Diagnostics, diagnostic => diagnostic.Code == "DOC-PAPX-INVALID");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void LegacyDoc_DataPropertyLimits_PublicLoadRejectsNonpositiveLimits(bool workLimit) {
        var options = new LegacyDocImportOptions();
        if (workLimit) options.MaxParagraphPropertyWorkBytes = 0;
        else options.MaxDataPropertyChainLength = 0;
        var error = Assert.Throws<ArgumentOutOfRangeException>(() => LegacyDocDocument.Load(Array.Empty<byte>(), options));
        Assert.Equal(workLimit ? nameof(options.MaxParagraphPropertyWorkBytes) : nameof(options.MaxDataPropertyChainLength), error.ParamName);
    }

    [Fact]
    public void LegacyDoc_DataPropertyLimits_CachedSuffixRetainsChainDepth() {
        byte[] data = DataPointerChain(3);
        var options = new LegacyDocImportOptions { MaxDataPropertyChainLength = 2 };
        var context = new LegacyDocDataPropertyContext(data, options);
        byte[] child = DataPropertyPointer(12);
        byte[] expected = data.Skip(26).Take(10).ToArray();
        Assert.Equal(expected, LegacyDocParagraphFormattingReader.ResolveDataProperties(child, 0, child.Length, data, context: context));
        byte[] parent = DataPropertyPointer(0);
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ResolveDataProperties(parent, 0, parent.Length, data, context: context));
    }

    [Fact]
    public void LegacyDoc_DataPropertyLimits_CachedExpansionsShareDocumentWorkBudget() {
        byte[] data = DataPointerChain(2);
        var context = new LegacyDocDataPropertyContext(data, new LegacyDocImportOptions { MaxParagraphPropertyWorkBytes = 45 });
        byte[] pointer = DataPropertyPointer(0);
        byte[] expected = data.Skip(14).Take(10).ToArray();
        for (int index = 0; index < 2; index++)
            Assert.Equal(expected, LegacyDocParagraphFormattingReader.ResolveDataProperties(pointer, 0, pointer.Length, data, context: context));
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ResolveDataProperties(pointer, 0, pointer.Length, data, context: context));
    }

    [Fact]
    public void LegacyDoc_DataPropertyLimits_CachedSuffixPreservesCallerPrefixAndHugeConstraints() {
        byte[] data = DataPointerChain(2);
        var context = new LegacyDocDataPropertyContext(data, new LegacyDocImportOptions());
        byte[] pointer = DataPropertyPointer(0);
        byte[] resolved = LegacyDocParagraphFormattingReader.ResolveDataProperties(pointer, 0, pointer.Length, data, context: context);
        byte[] prefixed = new byte[] { 0x05, 0x24, 1 }.Concat(pointer).ToArray();
        Assert.Equal(prefixed.Take(3).Concat(resolved), LegacyDocParagraphFormattingReader.ResolveDataProperties(prefixed, 0, prefixed.Length, data, context: context));
        byte[] huge = DataPropertyPointer(0, huge: true);
        Assert.Equal(resolved, LegacyDocParagraphFormattingReader.ResolveDataProperties(huge, 0, huge.Length, data, context: context));
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ResolveDataProperties(huge, 0, huge.Length, data, styleIndex: 1, context: context));
        // Returned property arrays must not become writable aliases of the cached record.
        resolved[0] = 0;
        Assert.Equal(0x07, LegacyDocParagraphFormattingReader.ResolveDataProperties(pointer, 0, pointer.Length, data, context: context)[0]);
    }

    private static byte[] DataPropertyPointer(int offset, bool huge = false) =>
        new[] { (byte)(huge ? 0x46 : 0x6B), (byte)(huge ? 0x66 : 0x64) }.Concat(BitConverter.GetBytes(offset)).ToArray();

    private static byte[] DataPointerChain(int records) {
        var data = new byte[records * 12];
        for (int index = 0; index < records; index++) {
            int offset = index * 12;
            data[offset] = 10;
            byte[] properties = index + 1 < records ? DataPropertyPointer(offset + 12)
                : new byte[] { 0x07, 0x94, 0x20, 3, 0x16, 0x24, 1, 0x17, 0x24, 1 };
            Buffer.BlockCopy(properties, 0, data, offset + 2, properties.Length);
        }
        return data;
    }

    private static partial class LegacyDocTestBuilder {
        internal static byte[] CreateDocWithSharedDataProperties(bool huge, bool styleOnly = false,
            int? pageNumber = null, bool overflowingTableSpan = false) {
            const string text = "first\rsecond\rthird\r";
            const int textOffset = 0x800;
            byte[] word = CreateWordDocumentStream(text, textOffset: textOffset);
            byte[] table = CreateTableStream(text.Length, textOffset);
            int plcOffset = table.Length;
            Array.Resize(ref table, table.Length + 20);
            int middle = textOffset + "first\rsecond\r".Length;
            int end = textOffset + text.Length;
            WriteInt32(table, plcOffset, textOffset);
            WriteInt32(table, plcOffset + 4, middle);
            WriteInt32(table, plcOffset + 8, end);
            WriteInt32(table, plcOffset + 12, pageNumber ?? 2);
            WriteInt32(table, plcOffset + 16, 3);
            WriteInt32(word, 0x102, overflowingTableSpan ? int.MaxValue - 7 : plcOffset);
            WriteInt32(word, 0x106, overflowingTableSpan ? 12 : 20);
            byte[] papx = styleOnly ? new byte[] { 0, 1, 1, 0 }
                : CreateParagraphPropertiesPapx(DataPropertyPointer(0, huge));
            WritePapxFkp(word, 0x400, new[] { textOffset, textOffset + "first\r".Length, middle },
                new Dictionary<int, byte[]> { [0] = papx, [1] = papx });
            WritePapxFkp(word, 0x600, new[] { middle, end }, new Dictionary<int, byte[]> { [0] = papx });
            byte[] data = DataRecord(new byte[] { 0x05, 0x24, 1, 0x06, 0x24, 1, 0x07, 0x24, 1, 0x31, 0x24, 1 });
            using var package = new MemoryStream();
            using (RootStorage root = RootStorage.Create(package, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen)) {
                WriteStream(root, "WordDocument", word);
                WriteStream(root, "1Table", table);
                WriteStream(root, "Data", data);
            }
            return package.ToArray();
        }
    }
}
