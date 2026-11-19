using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderInputFailureCodeTests {
    private const string Password = "shared-secret";

    [Fact]
    public async Task BatchAndFolderReadsContinueAndReportPasswordAndContentCodes() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-codes-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string deck = Path.Combine(root, "deck.pptx");
            string report = Path.Combine(root, "report.pdf");
            string damaged = Path.Combine(root, "damaged.docx");
            string notes = Path.Combine(root, "notes.txt");
            using (PowerPointDocument document = PowerPointDocument.Create()) {
                document.AddSlide().AddTextBox("Encrypted deck");
                File.WriteAllBytes(deck, document.ToBytes(new PowerPointSaveOptions { Password = Password }));
            }
            byte[] plainPdf = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.AddParagraph("Encrypted report")))).ToBytes();
            File.WriteAllBytes(report, PdfDocument.Load(plainPdf).Security.Encrypt(new(Password) { OwnerPassword = "owner" }).Pdf);
            File.WriteAllBytes(damaged, new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00 });
            File.WriteAllText(notes, "Plain notes");
            string[] paths = { deck, report, damaged, notes };
            OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.All;

            IReadOnlyList<ReaderDocumentReadOutcome> withoutPassword = await reader.ReadDocumentsDetailedAsync(paths);
            Assert.Equal(
                new string?[] { ReaderErrorCodes.PasswordRequired, ReaderErrorCodes.PasswordRequired, ReaderErrorCodes.InvalidContent, null },
                withoutPassword.Select(outcome => outcome.ErrorCode).ToArray());
            Assert.True(withoutPassword[3].Succeeded);

            IReadOnlyList<ReaderDocumentReadOutcome> wrongPassword = await reader.ReadDocumentsDetailedAsync(
                paths.Take(2), new ReaderOptions { OpenPassword = "not-the-password" });
            Assert.All(wrongPassword, outcome => Assert.Equal(ReaderErrorCodes.InvalidPassword, outcome.ErrorCode));

            IReadOnlyList<ReaderDocumentReadOutcome> rightPassword = await reader.ReadDocumentsDetailedAsync(
                paths.Take(2), new ReaderOptions { OpenPassword = Password });
            Assert.All(rightPassword, outcome => Assert.Null(outcome.ErrorCode));
            Assert.All(rightPassword, outcome => Assert.True(outcome.Succeeded, outcome.Error?.ToString()));

            ReaderIngestResult folder = reader.ReadFolderDetailed(root);
            Assert.Equal(ReaderErrorCodes.PasswordRequired, folder.Files.Single(file => file.Path == deck).ErrorCode);
            Assert.Equal(ReaderErrorCodes.PasswordRequired, folder.Files.Single(file => file.Path == report).ErrorCode);
            Assert.Equal(ReaderErrorCodes.InvalidContent, folder.Files.Single(file => file.Path == damaged).ErrorCode);
            Assert.Null(folder.Files.Single(file => file.Path == notes).ErrorCode);
            Assert.Equal(ReaderErrorCodes.InvalidPassword, reader.EnumerateFolderDocuments(root, options: new ReaderOptions { OpenPassword = "not-the-password" })
                .Single(document => document.Path == deck).ErrorCode);

            // Single-document reads keep the engine's exception; the reader maps it to the same code.
            Exception error = Assert.ThrowsAny<Exception>(() => reader.ReadDocument(deck));
            Assert.Equal(ReaderErrorCodes.PasswordRequired, reader.ClassifyReadError(error, deck));
            Assert.Equal(ReaderErrorCodes.PasswordRequired, reader.ClassifyReadError(Assert.ThrowsAny<Exception>(() => reader.ReadDocument(report))));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task PasswordCodesNeedThePasswordToBeTheCause() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-password-codes-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.All;

            // A legacy workbook reports a wrong password as an engine diagnostic, not a decryption exception.
            string workbook = Path.Combine(root, "book.xls");
            using (ExcelDocument document = ExcelDocument.Create()) {
                document.AddSheet("Data").Cell(1, 1).SetValue("Legacy");
                File.WriteAllBytes(workbook, document.ToBytes(new ExcelSaveOptions {
                    Format = ExcelFileFormat.Xls, Password = Password, LossPolicy = OfficeConversionLossPolicy.Allow
                }));
            }
            Assert.Equal(ReaderErrorCodes.InvalidPassword, Assert.Single(await reader.ReadDocumentsDetailedAsync(
                new[] { workbook }, new ReaderOptions { OpenPassword = "not-the-password" })).ErrorCode);

            // A damaged package that fails its integrity check after the right password is damaged content.
            string deck = Path.Combine(root, "deck.pptx");
            using (PowerPointDocument document = PowerPointDocument.Create()) {
                document.AddSlide().AddTextBox("Encrypted deck");
                File.WriteAllBytes(deck, CorruptEncryptedPackage(document.ToBytes(new PowerPointSaveOptions { Password = Password })));
            }
            ReaderDocumentReadOutcome damaged = Assert.Single(await reader.ReadDocumentsDetailedAsync(
                new[] { deck }, new ReaderOptions { OpenPassword = Password }));
            Assert.Equal(ReaderErrorCodes.InvalidContent, damaged.ErrorCode);

            // A projection the handler cannot produce is not fixed by a password.
            string encryptedDocx = Path.Combine(root, "report.docx");
            File.Copy(deck, encryptedDocx);
            OfficeDocumentReader richOnly = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.rich-only-docx",
                    Kind = ReaderInputKind.Word,
                    Extensions = new[] { ".docx" },
                    ReadDocumentPath = (_, _, _) => new OfficeDocumentReadResult()
                })
                .Build();
            Assert.Equal(ReaderErrorCodes.UnsupportedFormat, Assert.Single(await richOnly.ReadDocumentsDetailedAsync(
                new[] { encryptedDocx }, new ReaderOptions { Projection = ReaderProjection.TextAndMetadata })).ErrorCode);

            // With container inspection turned off the fallback probe cannot see the encrypted package either.
            string plainDeck = Path.Combine(root, "plain-deck.pptx");
            using (PowerPointDocument document = PowerPointDocument.Create()) {
                document.AddSlide().AddTextBox("Encrypted deck");
                File.WriteAllBytes(plainDeck, document.ToBytes(new PowerPointSaveOptions { Password = Password }));
            }
            Assert.Equal(ReaderErrorCodes.InvalidContent, Assert.Single(await reader.ReadDocumentsDetailedAsync(
                new[] { plainDeck }, new ReaderOptions { DetectionInspectContainers = false })).ErrorCode);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void CertificateProtectedPdfIsUnsupportedEncryptionNotAMissingPassword() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Pdf.Tests", "Pdf", "Fixtures", "Interoperability",
            "PublicKeyEncryption", "pyhanko-s4-rc4-128.pdf");
        OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.All;

        Exception error = Assert.ThrowsAny<Exception>(() => reader.ReadDocument(path, new ReaderOptions { OpenPassword = Password }));
        Assert.Equal(ReaderErrorCodes.UnsupportedEncryption, reader.ClassifyReadError(error, path));
    }

    [Fact]
    public async Task AThrowingClassifierKeepsTheReadFailureAndItsDefaultCode() {
        const string extension = ".brokenclassifier";
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-classifier-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "input" + extension);
            File.WriteAllText(path, "input");
            var readFailure = new InvalidDataException("Damaged input.");
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.throwing-classifier",
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadDocumentPath = (_, _, _) => throw readFailure,
                    ClassifyReadError = _ => throw new InvalidCastException("Classifier bug.")
                })
                .Build();

            ReaderDocumentReadOutcome outcome = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { path }));
            Assert.Same(readFailure, outcome.Error);
            Assert.Equal(ReaderErrorCodes.InvalidContent, outcome.ErrorCode);
            Assert.Equal(ReaderErrorCodes.InvalidContent, Assert.Single(reader.ReadFolderDetailed(root).Files).ErrorCode);
            Assert.Equal(ReaderErrorCodes.InvalidContent, reader.ClassifyReadError(readFailure));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(".xlsx", "[Content_Types].xml")]
    [InlineData(".xlsx", "_rels/.rels")]
    [InlineData(".docx", "[Content_Types].xml")]
    [InlineData(".pptx", "[Content_Types].xml")]
    public async Task DamagedOfficePackageIsInvalidContentWhateverTheEngineWrapsItIn(string extension, string damagedPart) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-damaged-xlsx-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string workbook = Path.Combine(root, "package" + extension);
            File.WriteAllBytes(workbook, CreateOfficePackage(extension));
            // Break a package part's XML. Excel wraps the parse failure in a general IOException, and the package security
            // check reports it as an IOException subtype; neither is an I/O failure.
            using (var archive = System.IO.Compression.ZipFile.Open(workbook, System.IO.Compression.ZipArchiveMode.Update)) {
                archive.GetEntry(damagedPart)!.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(damagedPart).Open());
                writer.Write("<broken");
            }

            ReaderDocumentReadOutcome outcome = Assert.Single(await OfficeIMO.Reader.Tests.ReaderTestReaders.All
                .ReadDocumentsDetailedAsync(new[] { workbook }));
            Assert.False(outcome.Succeeded);
            Assert.Equal(ReaderErrorCodes.InvalidContent, outcome.ErrorCode);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ExcelPackageWithEncryptedZipEntriesIsUnsupportedEncryption() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-zip-encryption-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            // Traditional ZIP encryption on the entries is not Office encryption; no password option can open it.
            string workbook = Path.Combine(root, "zip-encrypted.xlsx");
            File.WriteAllBytes(workbook, SetZipEncryptionFlag(CreateOfficePackage(".xlsx")));
            OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.All;

            ReaderDocumentReadOutcome outcome = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { workbook }));
            Assert.False(outcome.Succeeded);
            Assert.Equal(ReaderErrorCodes.UnsupportedEncryption, outcome.ErrorCode);
            Assert.Equal(ReaderErrorCodes.UnsupportedEncryption,
                reader.ClassifyReadError(Assert.ThrowsAny<Exception>(() => reader.ReadDocument(workbook)), workbook));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    // Sets general-purpose bit 0 (traditional ZIP encryption) in every local and central file header.
    private static byte[] SetZipEncryptionFlag(byte[] package) {
        byte[] bytes = (byte[])package.Clone();
        for (int offset = 0; offset + 4 <= bytes.Length; offset++) {
            uint signature = BitConverter.ToUInt32(bytes, offset);
            if (signature == 0x04034B50U) bytes[offset + 6] |= 1;
            else if (signature == 0x02014B50U) bytes[offset + 8] |= 1;
        }
        return bytes;
    }

    private static byte[] CreateOfficePackage(string extension) {
        switch (extension) {
            case ".xlsx": {
                using ExcelDocument document = ExcelDocument.Create();
                document.AddSheet("Data").Cell(1, 1).SetValue("Damaged");
                return document.ToBytes();
            }
            case ".pptx": {
                using PowerPointDocument document = PowerPointDocument.Create();
                document.AddSlide().AddTextBox("Damaged");
                return document.ToBytes();
            }
            default: {
                using var stream = new MemoryStream();
                using (WordDocument document = WordDocument.Create(stream)) {
                    document.AddParagraph("Damaged");
                    document.Save();
                }
                return stream.ToArray();
            }
        }
    }

    [Fact]
    public void ConventionalFileNamesUseTheRoutedExtensionsInputLimit() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.ruby",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { ".rb" },
                DefaultMaxInputBytesByExtension = new Dictionary<string, long> { [".rb"] = 16 },
                ReadDocumentStream = (_, _, _, _) => new OfficeDocumentReadResult()
            })
            .Build();
        byte[] gemfile = System.Text.Encoding.UTF8.GetBytes("source 'https://rubygems.org'\ngem 'rake'\n");

        Assert.ThrowsAny<IOException>(() => reader.ReadDocument(gemfile, "Gemfile"));
    }

    private static byte[] CorruptEncryptedPackage(byte[] encrypted) {
        // Flip one ciphertext byte inside the EncryptedPackage stream, past its 8-byte length prefix.
        byte[] name = System.Text.Encoding.Unicode.GetBytes("EncryptedPackage\0");
        for (int entry = 512; entry + 128 <= encrypted.Length; entry += 128) {
            if (!encrypted.AsSpan(entry, name.Length).SequenceEqual(name)) continue;
            int sectorSize = 1 << BitConverter.ToUInt16(encrypted, 30);
            long start = BitConverter.ToUInt32(encrypted, entry + 116);
            byte[] copy = (byte[])encrypted.Clone();
            copy[checked((int)((start + 1) * sectorSize + 64))] ^= 0x5A;
            return copy;
        }
        throw new InvalidDataException("The package has no EncryptedPackage stream.");
    }

    [Fact]
    public async Task WordPageLayoutTimeoutIsATimeoutNotTheCallersCancellation() {
        byte[] bytes;
        using (var stream = new MemoryStream()) {
            using (WordDocument document = WordDocument.Create(stream)) {
                for (int index = 0; index < 40; index++) document.AddParagraph("Paragraph " + index + " with enough text to lay out.");
                document.Save();
            }
            bytes = stream.ToArray();
        }
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordHandler(new ReaderWordOptions { IncludePageLocations = true })
            .Build();
        var limited = new ReaderOptions { LayoutTimeout = TimeSpan.FromTicks(1) };

        OfficeImageExportTimeoutException timeout = Assert.Throws<OfficeImageExportTimeoutException>(() =>
            reader.ReadDocument(bytes, "report.docx", limited));
        Assert.Equal(TimeSpan.FromTicks(1), timeout.Timeout);
        Assert.Equal(ReaderErrorCodes.Timeout, reader.ClassifyReadError(timeout));
        Assert.ThrowsAny<OperationCanceledException>(() =>
            reader.ReadDocument(bytes, "report.docx", new ReaderOptions { LayoutTimeout = TimeSpan.FromMinutes(5) }, new CancellationToken(true)));
        Assert.NotEmpty(reader.ReadDocument(bytes, "report.docx", new ReaderOptions { LayoutTimeout = TimeSpan.FromMinutes(5) }).Pages);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            reader.ReadDocument(bytes, "report.docx", new ReaderOptions { LayoutTimeout = TimeSpan.Zero }));

        string path = Path.Combine(Path.GetTempPath(), "officeimo-reader-layout-" + Guid.NewGuid().ToString("N") + ".docx");
        File.WriteAllBytes(path, bytes);
        try {
            ReaderDocumentReadOutcome outcome = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { path }, limited));
            Assert.Equal(ReaderErrorCodes.Timeout, outcome.ErrorCode);
        } finally {
            File.Delete(path);
        }
    }
}
