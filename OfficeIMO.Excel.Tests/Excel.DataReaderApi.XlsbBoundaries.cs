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

    [Fact]
    public void OpenDataReader_XlsbRejectsMissingWorksheetRelationship() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MissingWorksheetRelationship.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            DuplicateFirstXlsbBundleSheet(path, mutateRelationshipId: true);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("missing relationship", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsCaseInsensitiveDuplicateWorksheetNames() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateWorksheetName.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            DuplicateFirstXlsbBundleSheet(path, changeWorksheetNameCase: true);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("duplicate worksheet name", exception.Message, StringComparison.OrdinalIgnoreCase);
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

    private static void DuplicateFirstXlsbBundleSheet(
        string path,
        bool mutateRelationshipId = false,
        bool changeWorksheetNameCase = false) {
        byte[] bytes = ReadZipEntry(path, "xl/workbook.bin");
        var reader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(int.MaxValue));
        XlsbRecordSlice bundle = default;
        XlsbRecordSlice endBook = default;
        bool foundBundle = false;
        bool foundEndBook = false;
        while (reader.TryRead(out XlsbRecordSlice record)) {
            if (!foundBundle && record.Type == 156) {
                bundle = record;
                foundBundle = true;
            }
            if (record.Type == 132) {
                endBook = record;
                foundEndBook = true;
            }
        }

        Assert.True(foundBundle);
        Assert.True(foundEndBook);
        int bundleLength = bundle.PayloadOffset - bundle.RecordOffset + bundle.Size;
        var duplicate = new byte[bundleLength];
        Buffer.BlockCopy(
            bytes,
            bundle.RecordOffset,
            duplicate,
            0,
            bundleLength);
        int payloadOffset = bundle.PayloadOffset - bundle.RecordOffset;
        int relationshipLength = checked((int)ReadUInt32LittleEndian(
            duplicate,
            payloadOffset + 8));
        int relationshipOffset = payloadOffset + 12;
        int nameLengthOffset = checked(
            relationshipOffset + relationshipLength * 2);
        int nameLength = checked((int)ReadUInt32LittleEndian(
            duplicate,
            nameLengthOffset));
        int nameOffset = nameLengthOffset + 4;
        if (mutateRelationshipId) {
            Assert.True(relationshipLength > 0);
            duplicate[relationshipOffset] =
                duplicate[relationshipOffset] == (byte)'z'
                    ? (byte)'y'
                    : (byte)'z';
        }
        if (changeWorksheetNameCase) {
            bool changed = false;
            for (int index = 0; index < nameLength; index++) {
                int characterOffset = nameOffset + index * 2;
                char value = (char)(
                    duplicate[characterOffset]
                    | duplicate[characterOffset + 1] << 8);
                if (!char.IsLetter(value)) {
                    continue;
                }

                char replacement = char.IsUpper(value)
                    ? char.ToLowerInvariant(value)
                    : char.ToUpperInvariant(value);
                duplicate[characterOffset] = (byte)replacement;
                duplicate[characterOffset + 1] = (byte)(replacement >> 8);
                changed = true;
                break;
            }
            Assert.True(changed);
        }

        var expanded = new byte[bytes.Length + duplicate.Length];
        Buffer.BlockCopy(
            bytes,
            0,
            expanded,
            0,
            endBook.RecordOffset);
        Buffer.BlockCopy(
            duplicate,
            0,
            expanded,
            endBook.RecordOffset,
            duplicate.Length);
        Buffer.BlockCopy(
            bytes,
            endBook.RecordOffset,
            expanded,
            endBook.RecordOffset + duplicate.Length,
            bytes.Length - endBook.RecordOffset);
        ReplaceZipEntry(path, "xl/workbook.bin", expanded);
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

    private static uint ReadUInt32LittleEndian(byte[] bytes, int offset) =>
        (uint)(
            bytes[offset]
            | bytes[offset + 1] << 8
            | bytes[offset + 2] << 16
            | bytes[offset + 3] << 24);
}
