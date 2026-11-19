using OfficeIMO.Drawing;
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
