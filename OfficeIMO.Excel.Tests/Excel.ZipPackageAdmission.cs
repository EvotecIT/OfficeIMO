using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class ExcelZipPackageAdmissionTests {
    public static TheoryData<string> EntryPoints() => new() {
        "Load(path)", "Load(stream)", "OpenReader(path)", "OpenReader(stream)",
        "ExcelDocumentReader.Open(path)", "ExcelDocumentReader.Open(stream)", "ExcelDocumentReader.Open(bytes)"
    };

    [Theory]
    [MemberData(nameof(EntryPoints))]
    public void EncryptedZipEntriesAreRejectedAsUnreadableWorkbook(string entryPoint) {
        byte[] package = SetEncryptionFlag(CreateWorkbook());

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => Open(entryPoint, package));

        Assert.Contains("encrypted entries", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(EntryPoints))]
    public void Zip64PackageAboveEntryLimitIsRejectedBeforeOpening(string entryPoint) {
        byte[] package = AddEmptyEntries(CreateWorkbook(), ushort.MaxValue + 2);

        OfficePackageSecurityException exception = Assert.Throws<OfficePackageSecurityException>(() => Open(entryPoint, package));

        Assert.Equal(OfficePackageSecurityRule.PartCount, exception.Rule);
    }

    private static void Open(string entryPoint, byte[] package) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.ZipAdmission.{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(path, package);
        try {
            using var stream = new MemoryStream(package, writable: false);
            IDisposable opened = entryPoint switch {
                "Load(path)" => ExcelDocument.Load(path),
                "Load(stream)" => ExcelDocument.Load(stream),
                "OpenReader(path)" => ExcelDocument.OpenReader(path),
                "OpenReader(stream)" => ExcelDocument.OpenReader(stream),
                "ExcelDocumentReader.Open(path)" => ExcelDocumentReader.Open(path),
                "ExcelDocumentReader.Open(stream)" => ExcelDocumentReader.Open(stream),
                "ExcelDocumentReader.Open(bytes)" => ExcelDocumentReader.Open(package),
                _ => throw new ArgumentOutOfRangeException(nameof(entryPoint))
            };
            opened.Dispose();
        } finally {
            File.Delete(path);
        }
    }

    private static byte[] CreateWorkbook() {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddSheet("Data").Cell(1, 1).SetValue("Value");
        return document.ToBytes();
    }

    // Sets general-purpose bit 0 (traditional ZIP encryption) in every local and central file header.
    private static byte[] SetEncryptionFlag(byte[] package) {
        byte[] bytes = (byte[])package.Clone();
        for (int offset = 0; offset + 4 <= bytes.Length; offset++) {
            uint signature = BitConverter.ToUInt32(bytes, offset);
            if (signature == 0x04034B50U) bytes[offset + 6] |= 1;
            else if (signature == 0x02014B50U) bytes[offset + 8] |= 1;
        }
        return bytes;
    }

    // More than 65,535 entries require the ZIP64 end-of-central-directory record.
    private static byte[] AddEmptyEntries(byte[] package, int count) {
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            for (int index = 0; index < count; index++) {
                archive.CreateEntry("padding/" + index, CompressionLevel.NoCompression);
            }
        }
        return stream.ToArray();
    }
}
