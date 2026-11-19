using System.IO.Compression;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
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
        // An owner-only file opens without a password but still carries an encryption dictionary.
        byte[] ownerOnly = PdfDocument.Load(plain).Security.Encrypt(
            PdfStandardEncryptionOptions.ForPermissions("owner", PdfStandardPermissions.Print)).Pdf;
        // A small probe keeps the trailer out of the prefix, so the tail and startxref lookup must find it.
        var smallProbe = new ReaderDetectionOptions { MaxProbeBytes = 256 };
        Assert.True(encrypted.Length > 4 * smallProbe.MaxProbeBytes);

        foreach (byte[] protectedPdf in new[] { encrypted, ownerOnly }) {
            ReaderDetectionResult detection = Reader.Detect(protectedPdf, "report.pdf", smallProbe);
            Assert.True(detection.IsEncrypted);
            Assert.Equal(ReaderInputKind.Pdf, detection.Kind);
            Assert.Contains("encryption:pdf-encrypt-dictionary", detection.Evidence);
        }
        using var stream = new MemoryStream(encrypted);
        stream.Position = 0;
        Assert.True((await Reader.DetectAsync(stream, "report.pdf", smallProbe)).IsEncrypted);
        Assert.Equal(0, stream.Position);
        Assert.False(Reader.Detect(plain, "report.pdf", smallProbe).IsEncrypted);
    }

    [Fact]
    public void OpenDocumentManifestEncryptionIsReported() {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Encrypted text document");
        byte[] encrypted = document.ToBytes(new OdfSaveOptions {
            Encryption = new OdfEncryptionOptions { Password = "odf-secret" }
        });
        byte[] plain = document.ToBytes();

        ReaderDetectionResult detection = Reader.Detect(encrypted, "letter.odt");
        Assert.True(detection.IsEncrypted);
        Assert.Equal(ReaderInputKind.OpenDocument, detection.Kind);
        Assert.Contains("encryption:opendocument-manifest", detection.Evidence);
        Assert.True(Reader.Detect(encrypted, "attachment").IsEncrypted);
        Assert.False(Reader.Detect(plain, "letter.odt").IsEncrypted);
    }

    [Theory]
    [InlineData("<< /Size 5 /Root 1 0 R >>", "BT (trailer /Encrypt 9 0 R) Tj ET", false)]
    [InlineData("<< /Size 5 /Root 1 0 R /Info << /Encrypt 1 >> >>", "BT (plain) Tj ET\n% trailer << /Encrypt 9 0 R >>", false)]
    [InlineData("<< /Size 5 /Root 1 0 R /EncryptMetadata false >>", "BT (plain) Tj ET", false)]
    [InlineData("<< /Size 5 /Root 1 0 R /Encrypt null >>", "BT (plain) Tj ET", false)]
    [InlineData("<< /Size 5 /Root 1 0 R /Encrypt 9 0 R /Info null >>", "BT (plain) Tj ET", true)]
    [InlineData("<< /Size 5 /Root 1 0 R /En#63rypt 9 0 R >>", "BT (plain) Tj ET", true)]
    [InlineData("<< /Size 5 /Root 1 0 R /ID [<00> <01>] /Encrypt 9 0 R >>", "BT (plain) Tj ET", true)]
    public void PdfEncryptionComesOnlyFromTheTrailerEncryptKey(string trailer, string content, bool encrypted) {
        // Text in content streams, comments, nested dictionaries and other keys must not count; an escaped key name must.
        byte[] pdf = BuildPdf(trailer, content);
        Assert.Equal(encrypted, Reader.Detect(pdf, "file.pdf").IsEncrypted);
        Assert.Equal(encrypted, Reader.Detect(pdf, "file.pdf", new ReaderDetectionOptions { MaxProbeBytes = 64 }).IsEncrypted);
    }

    [Theory]
    [InlineData("999999999999999999")]
    [InlineData("461168601842738791")]
    [InlineData("5")]
    public void MalformedPdfCrossReferenceIsReportedWithoutThrowing(string count) {
        // A subsection count that overruns the file, or one that does not match the entries, leaves the result unproven.
        string body = "%PDF-1.4\n1 0 obj\n<< /Type /Catalog >>\nendobj\n";
        string pdf = body + "xref\n0 " + count + "\n0000000000 65535 f \ntrailer\n<< /Size 2 /Encrypt 9 0 R >>\n" +
                     "startxref\n" + body.Length + "\n%%EOF\n";
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(pdf);

        ReaderDetectionResult detection = Reader.Detect(bytes, "broken.pdf", new ReaderDetectionOptions { MaxProbeBytes = 16 });
        Assert.Equal(ReaderInputKind.Pdf, detection.Kind);
        Assert.False(detection.IsEncrypted);
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

        // A package recognized from its entries keeps that kind and still reports the encrypted entry.
        byte[] wordPackage = SetZipEntryEncrypted(CreateZip("[Content_Types].xml", "word/document.xml"), entryIndex: 1);
        ReaderDetectionResult word = Reader.Detect(wordPackage, "letter.docx");
        Assert.Equal(ReaderInputKind.Word, word.ContentKind);
        Assert.True(word.IsEncrypted);
        Assert.Contains("encryption:zip-encrypted-entry", word.Evidence);
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

        // No writer produces encrypted .doc files, so set the FIB fEncrypted bit on a real Word 97 document.
        byte[] plainDocument = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Documents", "LegacyDocCorpus", "ComSimpleParagraphs.doc"));
        byte[] wordDocument = SetWordFibEncrypted(plainDocument);

        ReaderDetectionResult doc = Reader.Detect(wordDocument, "letter.doc");
        Assert.True(doc.IsEncrypted);
        Assert.Equal(ReaderInputKind.Word, doc.ContentKind);
        Assert.Contains("encryption:legacy-word-fib", doc.Evidence);
        Assert.False(Reader.Detect(plainDocument, "letter.doc").IsEncrypted);
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

    private static byte[] BuildPdf(string trailer, string content) {
        var builder = new System.Text.StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        void Add(string body) {
            offsets.Add(builder.Length);
            builder.Append(offsets.Count).Append(" 0 obj\n").Append(body).Append("\nendobj\n");
        }
        Add("<< /Type /Catalog /Pages 2 0 R >>");
        Add("<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
        Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>");
        Add("<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream");
        int crossReference = builder.Length;
        builder.Append("xref\n0 5\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10")).Append(" 00000 n \n");
        builder.Append("trailer\n").Append(trailer).Append("\nstartxref\n").Append(crossReference).Append("\n%%EOF\n");
        return System.Text.Encoding.ASCII.GetBytes(builder.ToString());
    }

    private static byte[] SetWordFibEncrypted(byte[] document) {
        // The WordDocument stream starts with the FIB (wIdent 0xA5EC) on a sector boundary after the 512-byte header;
        // fEncrypted is bit 8 of the 16-bit flags at FIB offset 10.
        byte[] copy = (byte[])document.Clone();
        for (int offset = 512; offset + 12 <= copy.Length; offset += 512) {
            if (copy[offset] != 0xEC || copy[offset + 1] != 0xA5) continue;
            Assert.Equal(0, copy[offset + 11] & 0x01);
            copy[offset + 11] |= 0x01;
            return copy;
        }
        throw new InvalidDataException("The Word document has no FIB on a sector boundary.");
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
