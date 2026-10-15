using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Read;
using System.Buffers;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void XlsbSharedStrings_PreserveDecoderPolicyAcrossBorrowedAndCanonicalAccess(bool largeTable, bool empty) {
        ushort[] units = empty ? Array.Empty<ushort>() : [
            0x41, 0xD800, 0x42, 0xD83D, 0xDC22, 0, 0xDC00,
            0xFEFF, 0xFFFE, 0xFFFF, 0xD800, 0xD800, 0x43, 0xDC00, 0xDC00
        ];
        string expected = empty ? string.Empty : "A\uFFFDB🐢\0\uFFFD\uFEFF\uFFFE\uFFFF\uFFFD\uFFFDC\uFFFD\uFFFD";
        long characters = 0;
        byte[] package = RewriteSst(CreateSstWorkbook(largeTable), (records, output) => {
            foreach (XlsbRecord record in records) {
                byte[] payload = record.Data;
                if (record.Type == 19) {
                    string original = DecodeSstItem(payload);
                    if (original == "Replace") {
                        payload = CreateSstItem(units);
                        characters += expected.Length;
                    } else {
                        characters += original.Length;
                    }
                }
                XlsbRecordWriter.Write(output, record.Type, payload);
            }
        });
        string materialized;
        using (ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package,
            new ExcelReadOptions { MaxSharedStringCharacters = characters })) {
            Assert.True(reader.Read());
#if NET8_0_OR_GREATER
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> held));
            Assert.Equal(Encoding.UTF8.GetBytes(expected), held.ToArray());
            Assert.True(reader.TryGetUtf8Text(1, out ReadOnlySpan<byte> other));
            Assert.Equal("Other", Encoding.UTF8.GetString(other));
#endif
            materialized = reader.GetString(0);
            Assert.Equal(expected, materialized);
            Assert.Equal(expected, Assert.IsType<string>(reader.GetValue(0)));
#if NET8_0_OR_GREATER
            Assert.Equal(Encoding.UTF8.GetBytes(expected), held.ToArray());
            Assert.Equal("Other", Encoding.UTF8.GetString(other));
#endif
            Assert.True(reader.Read());
            Assert.Equal("After", reader.GetString(0));
            Assert.Equal(expected, materialized);
        }
        Assert.Equal(expected, materialized);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => XlsbTabularWorkbook.Open(package,
            new ExcelReadOptions { MaxSharedStringCharacters = characters - 1 }));
        Assert.Contains("aggregate limit", error.Message);
    }

    [Fact]
    public void XlsbSharedStrings_CanonicalTextRemainsCorrectAcrossWorkbookOwnersAndForeignPoolEntries() {
        byte[] firstPackage = CreateSstWorkbook(false, repeatFirstValue: true);
        byte[] secondPackage = RewriteSst(firstPackage, (records, output) => {
            foreach (XlsbRecord record in records) {
                XlsbRecordWriter.Write(output, record.Type, record.Type == 19
                    ? CreateSstItem(("Second" + DecodeSstItem(record.Data)).Select(c => (ushort)c).ToArray())
                    : record.Data);
            }
        });
        // Shared pools can supply arrays returned by unrelated callers without clearing.
        string?[] foreign = ArrayPool<string?>.Shared.Rent(5);
        for (int index = 0; index < foreign.Length; index++) foreign[index] = "Foreign pool entry";
        ArrayPool<string?>.Shared.Return(foreign, clearArray: false);

        string held;
        using (XlsbTabularWorkbook firstWorkbook = XlsbTabularWorkbook.Open(firstPackage, new ExcelReadOptions()))
        using (var first = (XlsbTabularDataReader)firstWorkbook.OpenTable("Data", true, new ExcelReadOptions())) {
            Assert.True(first.Read());
            held = first.GetString(0);
            Assert.Equal("Replace", held);
            Assert.Equal("Other", first.GetString(1));
            firstWorkbook.Dispose();

            using (XlsbTabularWorkbook secondWorkbook = XlsbTabularWorkbook.Open(secondPackage, new ExcelReadOptions()))
            using (var second = (XlsbTabularDataReader)secondWorkbook.OpenTable("Data", true, new ExcelReadOptions())) {
                Assert.True(second.Read());
                Assert.Equal("SecondReplace", second.GetString(0));
                Assert.Equal("SecondOther", second.GetString(1));
                Assert.True(first.Read());
                Assert.Equal("After", first.GetString(0));
                Assert.True(first.Read());
                Assert.Equal("Replace", first.GetString(0));
                first.Close();
                secondWorkbook.Dispose();
                Assert.True(second.Read());
                Assert.Equal("SecondAfter", second.GetString(0));
                Assert.True(second.Read());
                Assert.Equal("SecondReplace", second.GetString(0));
                Assert.Equal("Replace", held);
            }
        }
        Assert.Equal("Replace", held);
    }

    [Theory]
    [InlineData("Truncated")]
    [InlineData("ItemLimit")]
    [InlineData("AggregateLimit")]
    public void XlsbSharedStrings_RejectLateUnreferencedFailuresBeforeOpenReturns(string failure) {
        byte[] original = CreateSstWorkbook(false);
        long characters = 0;
        byte[] package = RewriteSst(original, (records, output) => {
            foreach (XlsbRecord record in records) {
                byte[] payload = (byte[])record.Data.Clone();
                if (record.Type == 159) {
                    uint unique = BitConverter.ToUInt32(payload, 4);
                    BitConverter.GetBytes(BitConverter.ToUInt32(payload, 0) + 1).CopyTo(payload, 0);
                    BitConverter.GetBytes(unique + 1).CopyTo(payload, 4);
                }
                if (record.Type == 19) characters += DecodeSstItem(payload).Length;
                if (record.Type == 160) {
                    byte[] item = CreateSstItem("123456789".Select(c => (ushort)c).ToArray());
                    if (failure == "Truncated") item = item.Take(6).ToArray();
                    XlsbRecordWriter.Write(output, 19, item);
                }
                XlsbRecordWriter.Write(output, record.Type, payload);
            }
        });
        var options = new ExcelReadOptions();
        if (failure == "ItemLimit") options.MaxSharedStringItemCharacters = 8;
        if (failure == "AggregateLimit") options.MaxSharedStringCharacters = characters + 8;
        if (failure == "Truncated") {
            Assert.Throws<EndOfStreamException>(() => XlsbTabularWorkbook.Open(package, options));
        } else {
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => XlsbTabularWorkbook.Open(package, options));
            Assert.Contains(failure == "ItemLimit" ? "configured limit" : "aggregate limit", error.Message);
        }
    }

    [Fact]
    public void XlsbSharedStrings_ReleasePartWhenLateTableValidationFails() {
        using XlsbPooledPartStream part = CreateSstPart("First");
        var records = new XlsbRecordSliceReader(part.Buffer, 1024, new XlsbRecordReadBudget(100), part.DataLength);
        Assert.True(records.TryRead(out XlsbRecordSlice header));
        BitConverter.GetBytes(2U).CopyTo(part.Buffer, header.PayloadOffset);
        BitConverter.GetBytes(2U).CopyTo(part.Buffer, header.PayloadOffset + sizeof(uint));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ReadSstPart(part));
        Assert.Contains("declares 2 unique items but contains 1", error.Message);
        Assert.False(part.CanRead);
    }

    [Fact]
    public void XlsbSharedStrings_ReleasePayloadAfterLastOverlappingReaderCloses() {
        using XlsbPooledPartStream part = CreateSstPart("First", "Second");
        using XlsbSharedStringTable owner = ReadSstPart(part);
        using var firstWorksheet = CreateTabularWorksheet((0, 0U), (1, 1U));
        using var secondWorksheet = CreateTabularWorksheet((0, 1U));
        using XlsbTabularDataReader first = CreateTabularReader(firstWorksheet, owner, false, new XlsbCellReadBudget(10));
        using XlsbTabularDataReader second = CreateTabularReader(secondWorksheet, owner, false, new XlsbCellReadBudget(10));
        Assert.True(first.Read());
#if NET8_0_OR_GREATER
        Assert.True(first.TryGetUtf8Text(0, out ReadOnlySpan<byte> firstText));
#endif
        owner.Dispose();
        owner.Dispose();
        Assert.True(part.CanRead);
        Assert.True(second.Read());
#if NET8_0_OR_GREATER
        Assert.True(second.TryGetUtf8Text(0, out ReadOnlySpan<byte> secondText));
        Assert.Equal("First", Encoding.UTF8.GetString(firstText));
        Assert.Equal("Second", Encoding.UTF8.GetString(secondText));
#endif
        string materialized = second.GetString(0);
        first.Close();
        first.Close();
        Assert.True(part.CanRead);
#if NET8_0_OR_GREATER
        Assert.Equal("Second", Encoding.UTF8.GetString(secondText));
#endif
        second.Close();
        second.Close();
        Assert.False(part.CanRead);
        Assert.Equal("Second", materialized);
    }

    [Fact]
    public void XlsbSharedStrings_ReleaseFailedReaderReferenceWithoutClosingOwner() {
        using XlsbPooledPartStream part = CreateSstPart("First");
        using XlsbSharedStringTable owner = ReadSstPart(part);
        using var worksheet = new MemoryStream(new byte[] { 0x80 });
        Assert.Throws<EndOfStreamException>(() => CreateTabularReader(worksheet, owner, false, new XlsbCellReadBudget(10)));
        Assert.True(part.CanRead);
        Assert.Equal("First", owner[0]);
        owner.Dispose();
        Assert.False(part.CanRead);
    }

    [Fact]
    public void XlsbSharedStrings_WorkbookCloseLeavesExistingReaderUsable() {
        using XlsbTabularWorkbook workbook = XlsbTabularWorkbook.Open(CreateSstWorkbook(true), new ExcelReadOptions());
        using var reader = (XlsbTabularDataReader)workbook.OpenTable("Data", true, new ExcelReadOptions());
        workbook.Dispose();
        Assert.True(reader.Read());
#if NET8_0_OR_GREATER
        Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
        Assert.Equal("Replace", Encoding.UTF8.GetString(text));
#endif
        Assert.Equal("Replace", reader.GetString(0));
        Assert.True(reader.Read());
        Assert.Equal("After", reader.GetString(0));
    }

    [Fact]
    public void XlsbSharedStrings_RejectMalformedStylesAfterSstCapture() {
        byte[] bytes = CreateSstWorkbook(true);
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
            zip.GetEntry("xl/styles.bin")!.Delete();
            using Stream output = zip.CreateEntry("xl/styles.bin").Open();
            output.WriteByte(0x80);
        }
        Assert.Throws<EndOfStreamException>(() => XlsbTabularWorkbook.Open(stream.ToArray(), new ExcelReadOptions()));
    }

    private static byte[] CreateSstWorkbook(bool largeTable, bool repeatFirstValue = false) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Text");
        sheet.CellValue(1, 2, "Value");
        sheet.CellValue(2, 1, "Replace");
        sheet.CellValue(2, 2, "Other");
        sheet.CellValue(3, 1, "After");
        if (repeatFirstValue) sheet.CellValue(4, 1, "Replace");
        if (largeTable) for (int row = 4; row < 270; row++) sheet.CellValue(row, 1, "Item" + row);
        return document.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = true });
    }

    // Keep record mutation local to each test; no shared mutable output or producer encoder.
    private static byte[] RewriteSst(byte[] bytes, Action<IReadOnlyList<XlsbRecord>, Stream> rewrite) {
        using var package = new MemoryStream();
        package.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(package, ZipArchiveMode.Update, true)) {
            ZipArchiveEntry part = zip.GetEntry("xl/sharedStrings.bin")!;
            IReadOnlyList<XlsbRecord> records;
            using (Stream input = part.Open()) records = XlsbRecordReader.ReadAll(input);
            part.Delete();
            using Stream output = zip.CreateEntry("xl/sharedStrings.bin").Open();
            rewrite(records, output);
        }
        return package.ToArray();
    }

    private static string DecodeSstItem(byte[] bytes) => Encoding.Unicode.GetString(bytes, 5, checked((int)BitConverter.ToUInt32(bytes, 1) * 2));

    private static byte[] CreateSstItem(ushort[] units) {
        using var payload = new MemoryStream();
        using (var writer = new BinaryWriter(payload, Encoding.Unicode, true)) {
            writer.Write((byte)0);
            writer.Write((uint)units.Length);
            foreach (ushort unit in units) writer.Write(unit);
        }
        return payload.ToArray();
    }

    private static XlsbPooledPartStream CreateSstPart(params string[] strings) {
        using var bytes = new MemoryStream();
        byte[] header = new byte[8];
        BitConverter.GetBytes((uint)strings.Length).CopyTo(header, 0);
        BitConverter.GetBytes((uint)strings.Length).CopyTo(header, 4);
        XlsbRecordWriter.Write(bytes, 159, header);
        foreach (string value in strings) XlsbRecordWriter.Write(bytes, 19, CreateSstItem(value.Select(c => (ushort)c).ToArray()));
        XlsbRecordWriter.Write(bytes, 160);
        byte[] buffer = ArrayPool<byte>.Shared.Rent((int)bytes.Length);
        bytes.ToArray().CopyTo(buffer, 0);
        return new XlsbPooledPartStream(buffer, (int)bytes.Length);
    }

    private static XlsbSharedStringTable ReadSstPart(XlsbPooledPartStream part) => XlsbSharedStringTable.Read(
        part, "sharedStrings.bin", new XlsbImportOptions(), new XlsbRecordReadBudget(100),
        32767, 1_000_000, static () => { });
}
