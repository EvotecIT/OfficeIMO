using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Read;
using System.Data.Common;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenDataReader_XlsbRejectsMissingWorkbookEndBoundary() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MissingEndBook.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            RemoveFinalXlsbRecord(path, "xl/workbook.bin", expectedRecordType: 132);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("BrtBeginBook/BrtEndBook", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsTruncatedSharedStringHeader() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.TruncatedSstHeader.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            RemoveFirstXlsbRecordPayload(path, "xl/sharedStrings.bin", recordType: 159);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("BrtBeginSst", exception.Message, StringComparison.Ordinal);
            Assert.Contains("truncated", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    private static void RemoveFinalXlsbRecord(
        string path,
        string entryName,
        int expectedRecordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        var reader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(int.MaxValue));
        XlsbRecordSlice final = default;
        bool found = false;
        while (reader.TryRead(out XlsbRecordSlice record)) {
            final = record;
            found = true;
        }

        Assert.True(found);
        Assert.Equal(expectedRecordType, final.Type);
        Assert.Equal(bytes.Length, reader.Position);
        byte[] truncated = new byte[final.RecordOffset];
        Buffer.BlockCopy(bytes, 0, truncated, 0, truncated.Length);
        ReplaceZipEntry(path, entryName, truncated);
    }

    private static void RemoveFirstXlsbRecordPayload(
        string path,
        string entryName,
        int recordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        var reader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(int.MaxValue));
        XlsbRecordSlice target = default;
        bool found = false;
        while (reader.TryRead(out XlsbRecordSlice record)) {
            if (record.Type == recordType) {
                target = record;
                found = true;
                break;
            }
        }

        Assert.True(found);
        Assert.InRange(target.Size, 1, 127);
        Assert.Equal(target.Size, bytes[target.PayloadOffset - 1]);
        byte[] mutated = new byte[bytes.Length - target.Size];
        Buffer.BlockCopy(bytes, 0, mutated, 0, target.PayloadOffset);
        mutated[target.PayloadOffset - 1] = 0;
        Buffer.BlockCopy(
            bytes,
            target.PayloadOffset + target.Size,
            mutated,
            target.PayloadOffset,
            bytes.Length - target.PayloadOffset - target.Size);
        ReplaceZipEntry(path, entryName, mutated);
    }

    private static byte[] ReadZipEntry(string path, string entryName) {
        using var archive = ZipFile.OpenRead(path);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"The XLSB fixture has no '{entryName}' part.");
        using Stream input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static void ReplaceZipEntry(
        string path,
        string entryName,
        byte[] bytes) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry original = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"The XLSB fixture has no '{entryName}' part.");
        original.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            entryName,
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(bytes, 0, bytes.Length);
    }
}
