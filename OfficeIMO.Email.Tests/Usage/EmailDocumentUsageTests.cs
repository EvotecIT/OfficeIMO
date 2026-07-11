using OfficeIMO.Email;
using System.Text;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailDocumentUsageTests {
    [Fact]
    public void LoadAndSaveUseTheDocumentAsTheSimpleEntryPoint() {
        string directory = CreateTempDirectory();
        try {
            string emlPath = Path.Combine(directory, "message.eml");
            string msgPath = Path.Combine(directory, "message.msg");
            File.WriteAllText(emlPath, BuildEml("Simple API"), Encoding.UTF8);

            EmailDocument document = EmailDocument.Load(emlPath);
            EmailWriteResult write = document.Save(msgPath);
            EmailDocument converted = EmailDocument.Load(msgPath);

            Assert.Equal("Simple API", document.Subject);
            Assert.Equal(EmailFileFormat.OutlookMsg, converted.Format);
            Assert.Equal("Simple API", converted.Subject);
            Assert.True(write.BytesWritten > 0);
            Assert.DoesNotContain(write.Diagnostics,
                diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task AsyncSaveInfersKnownExtensionsAndExplicitFormatSupportsOtherNames() {
        string directory = CreateTempDirectory();
        try {
            string sourcePath = Path.Combine(directory, "source.eml");
            string inferredPath = Path.Combine(directory, "converted.msg");
            string explicitPath = Path.Combine(directory, "artifact.bin");
            File.WriteAllText(sourcePath, BuildEml("Async API"), Encoding.UTF8);
            EmailDocument document = await EmailDocument.LoadAsync(sourcePath);

            await document.SaveAsync(inferredPath);
            document.Save(explicitPath, EmailFileFormat.Eml);

            Assert.Equal(EmailFileFormat.OutlookMsg,
                EmailDocumentReader.DetectFormat(File.ReadAllBytes(inferredPath)));
            Assert.Equal(EmailFileFormat.Eml,
                EmailDocumentReader.DetectFormat(File.ReadAllBytes(explicitPath)));
            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(Path.Combine(directory, "ambiguous.bin")));
            Assert.Contains("explicit EmailFileFormat", exception.Message, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void LoadFailsClearlyWhileTheAdvancedReaderRetainsDiagnostics() {
        byte[] invalid = Encoding.UTF8.GetBytes("This is not an email artifact.");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => EmailDocument.Load(invalid));
        EmailReadResult result = new EmailDocumentReader().Read(invalid);

        Assert.Contains("EMAIL_FORMAT_UNKNOWN", exception.Message, StringComparison.Ordinal);
        Assert.True(result.HasErrors);
        Assert.Equal(EmailFileFormat.Unknown, result.Document.Format);
    }

    [Fact]
    public void SaveDoesNotCreatePartialOutputWhenSerializationReportsAnError() {
        string directory = CreateTempDirectory();
        try {
            string outputPath = Path.Combine(directory, "message.eml");
            string diagnosticOutputPath = Path.Combine(directory, "diagnostic-message.eml");
            var document = new EmailDocument { Subject = "Missing attachment content" };
            document.Attachments.Add(new EmailAttachment {
                FileName = "missing.bin",
                ContentType = "application/octet-stream",
                Length = 10
            });

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => document.Save(outputPath));
            EmailWriteResult diagnosticResult = new EmailDocumentWriter().Write(
                document, diagnosticOutputPath, EmailFileFormat.Eml);

            Assert.Contains("EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE", exception.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(outputPath));
            Assert.True(diagnosticResult.HasErrors);
            Assert.True(File.Exists(diagnosticOutputPath));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    private static string CreateTempDirectory() {
        string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(path);
        return path;
    }

    private static string BuildEml(string subject) => string.Join("\r\n", new[] {
        "From: Alice <alice@example.com>",
        "To: Bob <bob@example.com>",
        $"Subject: {subject}",
        "Date: Mon, 21 Jun 2021 10:00:00 +0000",
        "Message-ID: <simple-api@example.com>",
        "MIME-Version: 1.0",
        "Content-Type: text/plain; charset=utf-8",
        string.Empty,
        "Hello"
    });
}
