using OpenMcdf;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class ExcelDataReaderXlsChartSubstreamTests {
    [Fact]
    public void NativeXlsReader_ExcludesEmbeddedChartSeriesAndKeepsOrdinarySheets() {
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
            FixturePath, new ExcelReadOptions { HasHeaderRow = false });
        string[] names = { "graphs2", "graphs1", "b", "lb", "c", "Info", "cl", "maquis", "wijn", "All" };
        int[] expectedRows = { 0, 0, 29, 26, 28, 26, 43, 23, 77, 110 };
        int[] expectedValues = { 0, 0, 333, 294, 224, 45, 502, 160, 592, 761 };
        Assert.Equal(names, reader.SheetNames);

        for (int sheet = 0; sheet < names.Length; sheet++) {
            Assert.Equal(names[sheet], reader.CurrentSheetName);
            int rows = 0;
            int values = 0;
            while (reader.Read()) {
                rows++;
                if (names[sheet] == "b" && rows == 3) {
                    Assert.Equal("text", reader.GetString(5));
                    Assert.Equal("land", reader.GetString(6));
                }
                if (names[sheet] == "Info" && rows == 1) {
                    Assert.Equal("Validation Soil water balance model La Peyne", reader.GetString(0));
                }
                if (names[sheet] == "Info" && rows == 3) {
                    Assert.Equal(new DateTime(1998, 12, 14), reader.GetDateTime(1));
                }
                for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) {
                    if (reader.IsDBNull(ordinal)) continue;
                    values++;
                }
            }
            Assert.Equal(expectedRows[sheet], rows);
            Assert.Equal(expectedValues[sheet], values);
            if (sheet < 2) Assert.Equal(0, reader.FieldCount);
            Assert.Equal(sheet < names.Length - 1, reader.NextResult());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeXlsReader_RejectsInvalidEmbeddedChartBof(bool oversizedLength) {
        string path = CopyFixture();
        try {
            MutateWorkbook(path, bytes => {
                int offset = FindFirstChartBof(bytes);
                if (oversizedLength) WriteUInt16(bytes, offset + 2, ushort.MaxValue);
                else WriteUInt16(bytes, offset + 6, 0x0010);
            });
            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void NativeXlsReader_RejectsMissingEmbeddedChartEof() {
        string path = CopyFixture();
        try {
            MutateWorkbook(path, bytes => {
                int offset = FindFirstChartBof(bytes);
                for (offset += 4 + ReadUInt16(bytes, offset + 2); offset < bytes.Length - 4;) {
                    if (ReadUInt16(bytes, offset) == 0x000A) {
                        WriteUInt16(bytes, offset, 0);
                        return;
                    }
                    offset += 4 + ReadUInt16(bytes, offset + 2);
                }
                throw new InvalidDataException("The test fixture contains no embedded chart EOF.");
            });
            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void NativeXlsReader_StillRejectsDecreasingOrdinaryCellRows() {
        string path = CopyFixture();
        try {
            MutateWorkbook(path, bytes => {
                int depth = 0;
                int previousRow = -1;
                bool worksheet = false;
                for (int offset = 0; offset < bytes.Length - 8;) {
                    int type = ReadUInt16(bytes, offset);
                    int length = ReadUInt16(bytes, offset + 2);
                    if (type == 0x0809) {
                        if (depth == 0) {
                            worksheet = ReadUInt16(bytes, offset + 6) == 0x0010;
                            previousRow = -1;
                        }
                        depth++;
                    } else if (type == 0x000A) {
                        depth--;
                    } else if (worksheet && depth == 1 && IsCell(type)) {
                        int row = ReadUInt16(bytes, offset + 4);
                        if (previousRow > 0 && row > previousRow) {
                            WriteUInt16(bytes, offset + 4, 0);
                            return;
                        }
                        previousRow = row;
                    }
                    offset += 4 + length;
                }
                throw new InvalidDataException("The fixture contains no ordinary cell row to reorder.");
            });
            Assert.Throws<InvalidDataException>(() => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
                do {
                    while (reader.Read()) { }
                } while (reader.NextResult());
            });
        } finally {
            File.Delete(path);
        }
    }

    private static string FixturePath => Path.Combine(
        AppContext.BaseDirectory, "Documents", "LegacyXlsCorpus", "openpreserve-format-corpus", "valid.xls");

    private static string CopyFixture() {
        string path = Path.Combine(Path.GetTempPath(), $"xls-chart-{Guid.NewGuid():N}.xls");
        File.Copy(FixturePath, path);
        return path;
    }

    private static void MutateWorkbook(string path, Action<byte[]> mutate) {
        using RootStorage root = RootStorage.Open(path, FileMode.Open, FileAccess.ReadWrite);
        using CfbStream workbook = root.OpenStream("Workbook");
        using var copy = new MemoryStream();
        workbook.CopyTo(copy);
        byte[] bytes = copy.ToArray();
        mutate(bytes);
        workbook.Position = 0;
        workbook.Write(bytes, 0, bytes.Length);
        root.Flush();
    }

    private static int FindFirstChartBof(byte[] bytes) {
        int depth = 0;
        bool worksheet = false;
        for (int offset = 0; offset < bytes.Length - 8;) {
            int length = ReadUInt16(bytes, offset + 2);
            int type = ReadUInt16(bytes, offset);
            if (type == 0x0809 && length >= 4) {
                int substreamType = ReadUInt16(bytes, offset + 6);
                if (worksheet && depth > 0 && substreamType == 0x0020) return offset;
                if (depth == 0) worksheet = substreamType == 0x0010;
                depth++;
            } else if (type == 0x000A) {
                depth--;
            }
            offset += 4 + length;
        }
        throw new InvalidDataException("The test fixture contains no embedded chart BOF.");
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) =>
        (ushort)(bytes[offset] | bytes[offset + 1] << 8);

    private static bool IsCell(int type) => type is
        0x0006 or 0x0201 or 0x0203 or 0x0204 or 0x0205 or 0x027E or 0x00BD or 0x00BE or 0x00FD;

    private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }
}
