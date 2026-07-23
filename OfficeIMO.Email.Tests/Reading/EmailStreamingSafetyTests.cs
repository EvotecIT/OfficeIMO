using System.Threading;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailStreamingSafetyTests {
#if NET8_0_OR_GREATER
    [Fact]
    public void EmailTemporaryStorageCreatesOwnerOnlyUnixDirectories() {
        if (OperatingSystem.IsWindows()) return;
        string path = Path.Combine(Path.GetTempPath(),
            "officeimo-email-private-" + Guid.NewGuid().ToString("N"));
        try {
            EmailTemporaryStorage.CreatePrivateDirectory(path);

            UnixFileMode permissions = File.GetUnixFileMode(path);
            Assert.Equal(UnixFileMode.None, permissions &
                (UnixFileMode.GroupRead | UnixFileMode.GroupWrite | UnixFileMode.GroupExecute |
                 UnixFileMode.OtherRead | UnixFileMode.OtherWrite | UnixFileMode.OtherExecute));
        } finally {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
        }
    }
#endif

    [Fact]
    public void InvalidBase64AttachmentIsPreservedAndDiagnosed() {
        byte[] artifact = Encoding.ASCII.GetBytes(
            "From: sender@example.test\r\n" +
            "To: recipient@example.test\r\n" +
            "Subject: malformed base64\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=officeimo\r\n\r\n" +
            "--officeimo\r\nContent-Type: text/plain\r\n\r\nbody\r\n" +
            "--officeimo\r\nContent-Type: application/octet-stream; name=bad.bin\r\n" +
            "Content-Disposition: attachment; filename=bad.bin\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAA!B\r\n" +
            "--officeimo--\r\n");

        using EmailReadResult result = new EmailDocumentReader().ReadStreaming(
            new MemoryStream(artifact, writable: false), "bad.eml");
        EmailAttachment attachment = Assert.Single(result.Document.Attachments);

        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MIME_BASE64_INVALID");
        Assert.NotNull(attachment.ContentSource);
        using Stream content = attachment.OpenContentStream();
        using var reader = new StreamReader(content, Encoding.ASCII);
        Assert.Equal("AA!B", reader.ReadToEnd().Trim());
    }

    [Fact]
    public void TruncatedTnefReturnsDiagnosticsWithoutRetainingTemporaryContent() {
        var document = new EmailDocument { Subject = "truncated" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin",
            ContentType = "application/octet-stream",
            Content = Enumerable.Range(0, 1024).Select(index => unchecked((byte)index)).ToArray()
        });
        byte[] complete = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Tnef);
        byte[] truncated = complete.Take(complete.Length - 7).ToArray();

        using EmailReadResult result = new EmailDocumentReader().ReadStreaming(
            new MemoryStream(truncated, writable: false), "winmail.dat");

        Assert.True(result.HasErrors || result.Diagnostics.Count > 0);
        foreach (EmailAttachment attachment in result.Document.Attachments) {
            if (attachment.ContentSource == null) continue;
            using Stream content = attachment.OpenContentStream();
            Assert.True(content.Length <= 1024);
        }
    }

    [Fact]
    public void MalformedCompoundFileReturnsBoundedDiagnostic() {
        byte[] malformed = new byte[512];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(malformed, 0);

        using EmailReadResult result = new EmailDocumentReader(new EmailReaderOptions(
            maxInputBytes: 1024,
            maxCompoundDirectoryEntries: 16)).ReadStreaming(
                new MemoryStream(malformed, writable: false), "malformed.msg");

        Assert.True(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MSG_COMPOUND_INVALID");
        Assert.False(result.UsesFileBackedContent);
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.OutlookTemplate)]
    [InlineData(EmailFileFormat.Tnef)]
    public void DeterministicArtifactsRoundTripUnderMigrationSemantics(EmailFileFormat format) {
        EmailDocument source = CreatePortableDocument();
        var writer = new EmailDocumentWriter();

        byte[] first = writer.ToBytes(source, format);
        byte[] second = writer.ToBytes(source, format);
        using EmailReadResult reopened = new EmailDocumentReader().ReadStreaming(
            new MemoryStream(first, writable: false), FileName(format));
        using EmailReadResult reopenedAgain = new EmailDocumentReader().ReadStreaming(
            new MemoryStream(second, writable: false), FileName(format));
        EmailSemanticComparisonReport comparison = EmailSemanticComparer.Compare(
            reopened.Document, reopenedAgain.Document);

        Assert.Equal(first, second);
        Assert.True(comparison.IsMatch, string.Join(" | ", comparison.Differences.Select(
            difference => string.Concat(difference.Kind, ":", difference.Path))));
    }

    private static EmailDocument CreatePortableDocument() {
        var document = new EmailDocument {
            Subject = "deterministic synthetic message",
            MessageClass = "IPM.Note",
            Date = new DateTimeOffset(2026, 7, 17, 12, 0, 0, TimeSpan.Zero),
            From = new EmailAddress("sender@example.test", "Sender")
        };
        document.Body.Text = "deterministic body";
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("recipient@example.test", "Recipient")));
        document.Attachments.Add(new EmailAttachment {
            FileName = "synthetic.bin",
            ContentType = "application/octet-stream",
            Content = Enumerable.Range(0, 4096).Select(index => unchecked((byte)index)).ToArray(),
            Length = 4096
        });
        return document;
    }

    private static string FileName(EmailFileFormat format) {
        switch (format) {
            case EmailFileFormat.Eml: return "message.eml";
            case EmailFileFormat.OutlookMsg: return "message.msg";
            case EmailFileFormat.OutlookTemplate: return "message.oft";
            default: return "winmail.dat";
        }
    }
}
