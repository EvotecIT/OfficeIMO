using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmlxPartialMessageRecoveryTests {
    [Fact]
    public void PartialEnvelopeCannotUseSiblingRecoveryToBypassAnInvalidMimeLength() {
        using var tree = new PartialTree();
        tree.WriteMessage();
        string path = Path.Combine(tree.Root, "Messages/123.partial.emlx");
        byte[] original = File.ReadAllBytes(path);
        byte[] invalid = Encoding.ASCII.GetBytes("1000000\n").Concat(original.Skip(Array.IndexOf(original, (byte)'\n') + 1)).ToArray();
        tree.Write("Messages/123.partial.emlx", invalid);
        using EmailStoreSession session = EmailStoreSession.Open(tree.Root);
        Assert.Throws<InvalidDataException>(() => session.ReadItem(Assert.Single(session.EnumerateItems())));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedPartIdentityRestoresDistinctPayloadsAndSurvivesPortableExport(bool metadataOnly) {
        using var tree = new PartialTree();
        byte[] first = Enumerable.Range(0, 100003).Select(value => (byte)value).ToArray();
        byte[] second = new byte[] { 9, 8, 7 };
        tree.WriteMessage();
        tree.Write("Attachments/123/2.1/localized.eml", first);
        tree.Write("Attachments/123/2.2/localized.eml", second);
        IEmailContentSource? source = null;
        Stream? outstanding = null;
        using (EmailStoreSession session = EmailStoreSession.Open(tree.Root)) {
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            EmailStoreItem item = session.ReadItem(reference, new EmailStoreItemReadOptions(
                metadataOnly ? EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.AttachmentMetadata : EmailStoreItemReadParts.All,
                preferStreamingAttachmentContent: true));
            Assert.Equal(2, item.Document.Properties["Emlx:RecoveredPartCount"]);
            Assert.Equal(0, item.Document.Properties["Emlx:UnresolvedPartCount"]);
            Assert.Equal(new long[] { first.Length, second.Length }, item.Document.Attachments.Select(attachment => attachment.Length));
            Assert.True(item.ContentAvailability.IsPotentiallyPartial);
            Assert.DoesNotContain(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_EMLX_PART_UNAVAILABLE");
            if (metadataOnly) {
                Assert.All(item.Document.Attachments, attachment => { Assert.Null(attachment.Content); Assert.Null(attachment.ContentSource); });
            } else {
                Assert.All(item.Document.Attachments, attachment => Assert.Null(attachment.Content));
                source = item.Document.Attachments[0].ContentSource!;
                outstanding = source.OpenRead();
                Assert.Equal(first, Read(item.Document.Attachments[0]));
                Assert.Equal(second, Read(item.Document.Attachments[1]));
                byte[] portable = new EmailDocumentWriter(new EmailWriterOptions(conversionLossPolicy: EmailConversionLossPolicy.Warn)).ToBytes(item.Document, EmailFileFormat.Eml);
                using EmailReadResult exported = new EmailDocumentReader().Read(portable);
                Assert.Equal(first, exported.Document.Attachments[0].Content);
                Assert.Equal(second, exported.Document.Attachments[1].Content);
                Assert.Contains(new EmailDocumentWriter().AnalyzeConversion(item.Document, EmailFileFormat.Eml).Diagnostics,
                    diagnostic => diagnostic.Code == "EMAIL_EMLX_PARTIAL_CONTENT");
            }
        }
        if (source != null) {
            Assert.Throws<ObjectDisposedException>(() => source.OpenRead());
            Assert.Throws<ObjectDisposedException>(() => outstanding!.ReadByte());
            outstanding!.Dispose();
        }
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("ambiguous")]
    [InlineData("encoding")]
    [InlineData("protected")]
    [InlineData("outside-root")]
    [InlineData("empty")]
    public void UnavailableOrUnsupportedPartsRemainIndeterminate(string scenario) {
        using var tree = new PartialTree();
        tree.WriteMessage(scenario == "encoding" ? "x-unsupported" : "base64", scenario == "protected");
        if (scenario != "missing") tree.Write("Attachments/123/2.1/payload.bin", scenario == "empty" ? Array.Empty<byte>() : new byte[] { 1, 2, 3 });
        if (scenario == "ambiguous") tree.Write("Attachments/123/2.1/other.bin", new byte[] { 4 });
        using EmailStoreSession session = EmailStoreSession.Open(scenario == "outside-root" ? Path.Combine(tree.Root, "Messages") : tree.Root);
        EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()),
            new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
        Assert.Equal(0, item.Document.Properties["Emlx:RecoveredPartCount"]);
        Assert.Equal(2, item.Document.Properties["Emlx:UnresolvedPartCount"]);
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(EmailStoreItemReadParts.Bodies));
        Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_EMLX_PART_UNAVAILABLE");
    }

    [Fact]
    public void MissingEmbeddedMessageRemainsIndeterminateDespiteAnEmptyParsedDocument() {
        using var tree = new PartialTree();
        tree.WriteMime("Subject: Missing embedded message\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=outer\r\n\r\n" +
            "--outer\r\nContent-Type: text/plain\r\n\r\nBody\r\n--outer\r\nContent-Type: message/rfc822\r\n" +
            "Content-Disposition: attachment; filename=missing.eml\r\nX-Apple-Content-Length: 100\r\n\r\n\r\n--outer--\r\n");
        using EmailStoreSession session = EmailStoreSession.Open(tree.Root);
        EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()));
        Assert.NotNull(Assert.Single(item.Document.Attachments).EmbeddedDocument);
        Assert.Equal(1, item.Document.Properties["Emlx:UnresolvedPartCount"]);
        Assert.False(item.ContentAvailability.AvailableParts.HasFlag(EmailStoreItemReadParts.EmbeddedItems));
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(EmailStoreItemReadParts.EmbeddedItems));
    }

    [Theory]
    [InlineData("application/pgp-encrypted")]
    [InlineData("application/x-encrypted")]
    public void EncryptedWrappersCannotBeReconstructedFromSiblingStorage(string protocol) {
        using var tree = new PartialTree();
        tree.WriteMessage(rootType: "multipart/encrypted; protocol=\"" + protocol + "\"");
        tree.Write("Attachments/123/2.1/payload.bin", new byte[] { 1, 2, 3 });
        using EmailStoreSession session = EmailStoreSession.Open(tree.Root);
        EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()));
        Assert.Equal(0, item.Document.Properties["Emlx:RecoveredPartCount"]);
        Assert.Equal(2, item.Document.Properties["Emlx:UnresolvedPartCount"]);
        Assert.DoesNotContain(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_EMLX_PART_RECOVERED");
    }

    [Theory]
    [InlineData("", "Content-Transfer-Encoding: x-unsupported\r\n", "EMAIL_MIME_SINGLETON_HEADER_DUPLICATE")]
    [InlineData("", "Content-Type: application/pkcs7-mime\r\n", "EMAIL_MIME_SINGLETON_HEADER_DUPLICATE")]
    [InlineData("Content-Type: multipart/signed; boundary=outer\r\n", "", "EMAIL_MIME_SINGLETON_HEADER_DUPLICATE")]
    [InlineData("Content-Transfer-Encoding: base64\r\nContent-Transfer-Encoding: binary\r\n", "", "EMAIL_MIME_SINGLETON_HEADER_DUPLICATE")]
    [InlineData("boundary", "", "EMAIL_MIME_PARAMETER_DUPLICATE")]
    public void AmbiguousMimeHeadersRemainUnresolvedAndRetainTheirDiagnostic(string rootHeaders, string partHeaders, string diagnostic) {
        using var tree = new PartialTree();
        tree.WriteMessage(rootType: rootHeaders == "boundary" ? "multipart/mixed; boundary=outer" : null,
            rootHeaders: rootHeaders == "boundary" ? string.Empty : rootHeaders, partHeaders: partHeaders);
        tree.Write("Attachments/123/2.1/payload.bin", new byte[] { 1, 2, 3 });
        tree.Write("Attachments/123/2.2/payload.bin", new byte[] { 4, 5, 6 });
        using EmailStoreSession session = EmailStoreSession.Open(tree.Root);
        EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()));
        Assert.Equal(0, item.Document.Properties["Emlx:RecoveredPartCount"]);
        Assert.Equal(2, item.Document.Properties["Emlx:UnresolvedPartCount"]);
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
        Assert.Contains(session.Diagnostics, entry => entry.Code == diagnostic);
        Assert.DoesNotContain(session.Diagnostics, entry => entry.Code == "EMAIL_STORE_EMLX_PART_RECOVERED");
    }

    [Fact]
    public void SiblingLengthMutationAndRecoveryLimitsRejectBeforeReturningAnItem() {
        using var tree = new PartialTree();
        tree.WriteMessage();
        tree.Write("Attachments/123/2.1/payload.bin", new byte[32]);
        using (EmailStoreSession session = EmailStoreSession.Open(tree.Root, new EmailStoreReaderOptions(maxAttachmentBytes: 16))) {
            EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(Assert.Single(session.EnumerateItems())));
            Assert.Equal(nameof(EmailStoreReaderOptions.MaxAttachmentBytes), error.LimitName);
        }
        using (EmailStoreSession session = EmailStoreSession.Open(tree.Root)) {
            tree.Write("Attachments/123/2.1/payload.bin", new byte[33]);
            Assert.Throws<InvalidDataException>(() => session.ReadItem(Assert.Single(session.EnumerateItems())));
        }
        Assert.Throws<EmailStoreLimitExceededException>(() => EmailStoreSession.Open(tree.Root,
            new EmailStoreReaderOptions(maxDirectoryFileCount: 1)));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void SiblingSymbolicLinksCannotSupplyAttachmentBytes() {
        if (OperatingSystem.IsWindows()) return;
        using var tree = new PartialTree();
        tree.WriteMessage();
        string outside = tree.Write("outside.bin", new byte[] { 42 });
        string link = Path.Combine(tree.Root, "Attachments/123/2.1/payload.bin");
        Directory.CreateDirectory(Path.GetDirectoryName(link)!);
        File.CreateSymbolicLink(link, outside);
        using EmailStoreSession session = EmailStoreSession.Open(tree.Root);
        EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()));
        Assert.Equal(0, item.Document.Properties["Emlx:RecoveredPartCount"]);
        Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_DIRECTORY_REPARSE_POINT_SKIPPED");
    }
#endif

    private static byte[] Read(EmailAttachment attachment) {
        using Stream input = attachment.OpenContentStream();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private sealed class PartialTree : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "OfficeIMO.Email.Partial." + Guid.NewGuid().ToString("N"), "Inbox.mbox");
        internal string Write(string relative, byte[] bytes) {
            string path = Path.Combine(Root, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, bytes);
            return path;
        }
        internal void WriteMessage(string encoding = "base64", bool signed = false, string? rootType = null,
            string rootHeaders = "", string partHeaders = "") {
            string message = "Subject: Partial sibling recovery\r\nMIME-Version: 1.0\r\nContent-Type: " + (rootType ?? "multipart/" + (signed ? "signed" : "mixed")) + "; boundary=outer\r\n" + rootHeaders + "\r\n" +
                "--outer\r\nContent-Type: text/plain\r\n\r\nBody\r\n" +
                "--outer\r\nContent-Type: multipart/mixed; boundary=inner\r\n\r\n" +
                Part(encoding, partHeaders) + Part(encoding, partHeaders) + "--inner--\r\n--outer--\r\n";
            WriteMime(message);
        }
        internal void WriteMime(string message) {
            byte[] bytes = Encoding.UTF8.GetBytes(message);
            byte[] prefix = Encoding.ASCII.GetBytes(bytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            Write("Messages/123.partial.emlx", prefix.Concat(bytes).ToArray());
        }
        private static string Part(string encoding, string extraHeaders) => "--inner\r\nContent-Type: application/octet-stream; name=../../payload.bin\r\n" +
            "Content-Disposition: attachment; filename=../../payload.bin\r\nContent-Transfer-Encoding: " + encoding + "\r\n" + extraHeaders + "X-Apple-Content-Length: 138412\r\n\r\n\r\n";
        public void Dispose() => Directory.Delete(Path.GetDirectoryName(Root)!, recursive: true);
    }
}
