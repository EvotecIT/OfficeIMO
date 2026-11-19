using System.IO.Compression;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderEncryptionDetectionTests {
    private static readonly OfficeDocumentReader Reader = OfficeIMO.Reader.Tests.ReaderTestReaders.All;

    [Fact]
    public void EncryptedOfficePackageIsReportedWithTheKindItsExtensionNames() {
        byte[] encrypted = CreatePresentation(new PowerPointSaveOptions { Password = "deck-secret" });
        byte[] plain = CreatePresentation(null);

        ReaderDetectionResult named = Reader.Detect(encrypted, "deck.pptx");
        Assert.True(named.IsEncrypted);
        Assert.Equal(ReaderInputKind.PowerPoint, named.Kind);
        Assert.Equal(ReaderInputKind.PowerPoint, named.ContentKind);
        Assert.False(named.IsMismatch);
        Assert.Contains("encryption:office-encrypted-package", named.Evidence);

        // Without an extension nothing outside the encrypted payload names the format.
        ReaderDetectionResult unnamed = Reader.Detect(encrypted, "attachment");
        Assert.True(unnamed.IsEncrypted);
        Assert.Equal(ReaderInputKind.Unknown, unnamed.Kind);

        Assert.False(Reader.Detect(plain, "deck.pptx").IsEncrypted);
    }

    [Fact]
    public async Task PdfEncryptionDictionaryIsFoundInTheTrailerWithoutOpeningTheFile() {
        byte[] plain = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.AddParagraph("Quarterly figures.")))).ToBytes();
        byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new("reader") { OwnerPassword = "owner" }).Pdf;
        // A small probe keeps the trailer out of the prefix, so the tail and startxref lookup must find it.
        var smallProbe = new ReaderDetectionOptions { MaxProbeBytes = 256 };
        Assert.True(encrypted.Length > 4 * smallProbe.MaxProbeBytes);

        ReaderDetectionResult detection = Reader.Detect(encrypted, "report.pdf", smallProbe);
        Assert.True(detection.IsEncrypted);
        Assert.Equal(ReaderInputKind.Pdf, detection.Kind);
        Assert.Contains("encryption:pdf-encrypt-dictionary", detection.Evidence);
        using var stream = new MemoryStream(encrypted);
        stream.Position = 0;
        Assert.True((await Reader.DetectAsync(stream, "report.pdf", smallProbe)).IsEncrypted);
        Assert.Equal(0, stream.Position);
        Assert.False(Reader.Detect(plain, "report.pdf", smallProbe).IsEncrypted);
    }

    [Fact]
    public void ZipArchiveWithAnEncryptedEntryIsReported() {
        byte[] plain = CreateZip("notes.txt", "second.txt");
        // Mark the second entry encrypted in its local and central headers, as a ZipCrypto archive would.
        byte[] encrypted = SetZipEntryEncrypted(plain, entryIndex: 1);

        ReaderDetectionResult detection = Reader.Detect(encrypted, "archive.zip");
        Assert.True(detection.IsEncrypted);
        Assert.Equal(ReaderInputKind.Zip, detection.Kind);
        Assert.Contains("encryption:zip-encrypted-entry", detection.Evidence);
        Assert.False(Reader.Detect(plain, "archive.zip").IsEncrypted);
    }

    [Fact]
    public void LegacyOfficeFilesThatMarkThemselvesEncryptedAreReported() {
        byte[] presentation = CreatePresentation(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt, Password = "legacy-secret" });
        byte[] plainPresentation = CreatePresentation(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
        byte[] workbook;
        byte[] plainWorkbook;
        using (ExcelDocument document = ExcelDocument.Create()) {
            document.AddSheet("Data").Cell(1, 1).SetValue("Legacy");
            workbook = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls, Password = "legacy-secret", LossPolicy = OfficeConversionLossPolicy.Allow });
            plainWorkbook = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls, LossPolicy = OfficeConversionLossPolicy.Allow });
        }

        ReaderDetectionResult ppt = Reader.Detect(presentation, "deck.ppt");
        Assert.True(ppt.IsEncrypted);
        Assert.Equal(ReaderInputKind.PowerPoint, ppt.ContentKind);
        Assert.Contains("encryption:legacy-powerpoint-current-user", ppt.Evidence);
        ReaderDetectionResult xls = Reader.Detect(workbook, "book.xls");
        Assert.True(xls.IsEncrypted);
        Assert.Equal(ReaderInputKind.Excel, xls.ContentKind);
        Assert.Contains("encryption:legacy-excel-filepass", xls.Evidence);
        Assert.False(Reader.Detect(plainPresentation, "deck.ppt").IsEncrypted);
        Assert.False(Reader.Detect(plainWorkbook, "book.xls").IsEncrypted);
    }

    private static byte[] CreatePresentation(PowerPointSaveOptions? options) {
        using PowerPointDocument document = PowerPointDocument.Create();
        document.AddSlide().AddTextBox("Encrypted detection");
        return options == null ? document.ToBytes() : document.ToBytes(options);
    }

    private static byte[] CreateZip(params string[] names) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (string name in names) {
                using var writer = new StreamWriter(archive.CreateEntry(name, CompressionLevel.NoCompression).Open());
                writer.Write("payload " + name);
            }
        }
        return output.ToArray();
    }

    private static byte[] SetZipEntryEncrypted(byte[] zip, int entryIndex) {
        byte[] copy = (byte[])zip.Clone();
        int local = -1, central = -1;
        for (int offset = 0, locals = 0, centrals = 0; offset + 4 <= copy.Length; offset++) {
            uint signature = BitConverter.ToUInt32(copy, offset);
            if (signature == 0x04034B50U && locals++ == entryIndex) local = offset;
            if (signature == 0x02014B50U && centrals++ == entryIndex) central = offset;
        }
        Assert.True(local >= 0 && central >= 0);
        copy[local + 6] |= 1;
        copy[central + 8] |= 1;
        return copy;
    }
}
