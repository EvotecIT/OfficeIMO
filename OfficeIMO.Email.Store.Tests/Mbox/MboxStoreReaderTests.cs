using OfficeIMO.Email;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Email.Store.Tests;

public sealed class MboxStoreReaderTests {
    [Fact]
    public void SessionDetectsEnumeratesSearchesAndReadsMbox() {
        byte[] bytes = CreateMailboxBytes();
        using var stream = new MemoryStream(bytes);
        using EmailStoreSession session = EmailStoreSession.Open(stream, "archive.mbox");

        EmailStoreItemReference[] references = session.EnumerateItems().ToArray();
        EmailStoreSearchResult result = Assert.Single(session.Search(new EmailStoreQuery(subjectContains: "Second")));
        EmailStoreItem first = session.ReadItem(references[0]);

        Assert.Equal(EmailStoreFormat.Mbox, EmailStoreReader.DetectFormat(stream, "renamed.bin"));
        Assert.Equal(EmailStoreFormat.Mbox, session.Format);
        Assert.Equal("archive", session.DisplayName);
        Assert.Equal(bytes.LongLength, session.SourceLength);
        Assert.Equal(2, references.Length);
        Assert.Equal("Second message", result.Summary.Subject);
        Assert.Equal("first@example.com", first.Document.Properties["Mbox:EnvelopeSender"]);
        Assert.Equal(EmailStoreFormat.Mbox.ToString(), first.Document.Properties["EmailStore:Format"]);
    }

    [Fact]
    public void BomPrefixedMboxIsDetectedWithoutAFileExtension() {
        byte[] mailbox = CreateMailboxBytes();
        byte[] preamble = Encoding.UTF8.GetPreamble();
        var bytes = new byte[preamble.Length + mailbox.Length];
        Buffer.BlockCopy(preamble, 0, bytes, 0, preamble.Length);
        Buffer.BlockCopy(mailbox, 0, bytes, preamble.Length, mailbox.Length);
        using var stream = new MemoryStream(bytes);

        Assert.Equal(EmailStoreFormat.Mbox, EmailStoreReader.DetectFormat(stream, "renamed.bin"));
        using EmailStoreSession session = EmailStoreSession.Open(stream, "renamed.bin");
        Assert.Equal(2, session.EnumerateItems().Count());
    }

    [Fact]
    public void PerMessageLimitReportsTheStoreMaxMessageBytesOption() {
        using var stream = new MemoryStream(CreateMailboxBytes());
        var options = new EmailStoreReaderOptions(
            maxInputBytes: stream.Length + 1,
            maxMessageBytes: 64);

        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(() =>
            EmailStoreSession.Open(stream, "archive.mbox", options));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxMessageBytes), exception.LimitName);
    }

    [Fact]
    public void OpeningMboxStreamsWithoutRequestingTheWholeAggregate() {
        byte[] bytes = CreateLargeMailboxBytes();
        using var stream = new MaximumReadSizeStream(bytes, 32);

        using EmailStoreSession session = EmailStoreSession.Open(stream, "large.mbox");

        Assert.Equal(2, session.Folders.Single().ItemCount);
        Assert.Equal(2, session.EnumerateItems().Count());
    }

    [Fact]
    public void OpeningMboxEnforcesAggregateAttachmentBytesAcrossItems() {
        using var stream = new MemoryStream(CreateMailboxWithAttachments());
        var options = new EmailStoreReaderOptions(
            maxAttachmentBytes: 10,
            maxTotalAttachmentBytes: 10);

        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(() =>
            EmailStoreSession.Open(stream, "attachments.mbox", options));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), exception.LimitName);
        Assert.Equal(12, exception.Actual);
    }

    [Fact]
    public void SelectiveMboxReadsOnlyDecodeRequestedAttachmentContent() {
        using var stream = new MemoryStream(CreateMailboxWithAttachments());
        using EmailStoreSession session = EmailStoreSession.Open(stream, "attachments.mbox");
        EmailStoreItemReference reference = session.EnumerateItems().First();
        var metadataOnly = new EmailStoreItemReadOptions(EmailStoreItemReadParts.AttachmentMetadata);

        EmailStoreItem metadataItem = session.ReadItem(reference, metadataOnly);
        EmailStoreItem contentItem = session.ReadItem(reference,
            new EmailStoreItemReadOptions(EmailStoreItemReadParts.AttachmentContent));

        Assert.Equal(EmailStoreItemReadParts.All & ~EmailStoreItemReadParts.AttachmentContent,
            metadataItem.LoadedParts);
        Assert.Null(Assert.Single(metadataItem.Document.Attachments).Content);
        Assert.False(metadataItem.LoadedParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
        Assert.Equal(EmailStoreItemReadParts.All, contentItem.LoadedParts);
        Assert.True(contentItem.LoadedParts.HasFlag(EmailStoreItemReadParts.AttachmentContent));
        Assert.Equal("123456", Encoding.ASCII.GetString(
            Assert.Single(contentItem.Document.Attachments).Content!));
    }

    [Fact]
    public void SelectedMboxReadsDoNotDelegateThroughSingleByteBufferReads() {
        using var stream = new RejectSingleByteBufferReadStream(CreateMailboxBytes());
        using EmailStoreSession session = EmailStoreSession.Open(stream, "archive.mbox");
        EmailStoreItemReference reference = session.EnumerateItems().First();

        EmailStoreItem item = session.ReadItem(reference);

        Assert.Equal("First message", item.Document.Subject);
    }

    [Fact]
    public void RepeatedMboxReadsDoNotDuplicateIndexedDiagnostics() {
        byte[] bytes = Encoding.ASCII.GetBytes(
            "From sender@example.com Fri Jul 10 12:00:00 2026\nplain body\n");
        using var stream = new MemoryStream(bytes);
        using EmailStoreSession session = EmailStoreSession.Open(stream, "diagnostics.mbox");
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        int indexedCount = session.Diagnostics.Count;

        session.ReadItem(reference);
        session.ReadItem(reference);

        Assert.True(indexedCount > 0);
        Assert.Equal(indexedCount, session.Diagnostics.Count);
    }

    [Fact]
    public void MboxStoreCanExportThroughTheSharedSessionBoundary() {
        string destination = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".mbox");
        try {
            using var stream = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(stream, "source.mbox");

            EmailStoreMboxExportReport report = session.ExportToMbox(destination);
            EmailMailbox mailbox = EmailMailbox.Load(destination);

            Assert.False(report.WasTruncated);
            Assert.Equal(2, report.Entries.Count);
            Assert.Equal(new[] { "First message", "Second message" },
                mailbox.Messages.Select(entry => entry.Document.Subject));
            Assert.Equal(new[] { "first@example.com", "second@example.com" },
                mailbox.Messages.Select(entry => entry.EnvelopeSender));
            Assert.Equal(new[] {
                new DateTimeOffset(2026, 7, 17, 9, 0, 0, TimeSpan.Zero),
                new DateTimeOffset(2026, 7, 17, 10, 0, 0, TimeSpan.Zero)
            }, mailbox.Messages.Select(entry => entry.EnvelopeDate!.Value));
        } finally {
            if (File.Exists(destination)) File.Delete(destination);
        }
    }

    [Fact]
    public void MboxStoreCannotExportOverItsOwnFileBackedSource() {
        string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".mbox");
        byte[] original = CreateMailboxBytes();
        try {
            File.WriteAllBytes(source, original);
            using EmailStoreSession session = EmailStoreSession.Open(source);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                session.ExportToMbox(source,
                    new EmailStoreMboxExportOptions(overwriteExisting: true)));

            Assert.Contains("source store", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(original, File.ReadAllBytes(source));
        } finally {
            if (File.Exists(source)) File.Delete(source);
        }
    }

    [Fact]
    public void MboxExportContinuesAfterAMessageExceedsTheWriterLimit() {
        string destination = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".mbox");
        try {
            var sourceMailbox = new EmailMailbox();
            sourceMailbox.Messages.Add(new EmailMailboxEntry(new EmailDocument {
                Subject = "Oversized",
                Body = { Text = new string('x', 65_536) }
            }));
            sourceMailbox.Messages.Add(new EmailMailboxEntry(new EmailDocument {
                Subject = "Valid",
                Body = { Text = "Small body" }
            }));
            using var stream = new MemoryStream(sourceMailbox.ToBytes());
            using EmailStoreSession session = EmailStoreSession.Open(stream, "source.mbox");
            var writerOptions = new EmailMailboxWriterOptions(
                new EmailWriterOptions(maxOutputBytes: 4_096));

            EmailStoreMboxExportReport report = session.ExportToMbox(destination,
                new EmailStoreMboxExportOptions(
                    continueOnError: true,
                    writerOptions: writerOptions));

            Assert.Equal(2, report.Entries.Count);
            Assert.Equal(1, report.SucceededCount);
            Assert.Contains(report.Entries, entry => entry.Diagnostics.Any(diagnostic =>
                diagnostic.Code == "EMAIL_STORE_EXPORT_ITEM_LIMIT" &&
                diagnostic.Message.Contains(nameof(EmailWriterOptions.MaxOutputBytes),
                    StringComparison.Ordinal)));
            Assert.Equal("Valid", Assert.Single(EmailMailbox.Load(destination).Messages).Document.Subject);
        } finally {
            if (File.Exists(destination)) File.Delete(destination);
        }
    }

    private static byte[] CreateMailboxBytes() {
        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(new EmailDocument {
            Subject = "First message",
            From = new EmailAddress("sender@example.com")
        }) {
            EnvelopeSender = "first@example.com",
            EnvelopeDate = new DateTimeOffset(2026, 7, 17, 9, 0, 0, TimeSpan.Zero)
        });
        mailbox.Messages.Add(new EmailMailboxEntry(new EmailDocument {
            Subject = "Second message",
            From = new EmailAddress("sender@example.com")
        }) {
            EnvelopeSender = "second@example.com",
            EnvelopeDate = new DateTimeOffset(2026, 7, 17, 10, 0, 0, TimeSpan.Zero)
        });
        return mailbox.ToBytes();
    }

    private static byte[] CreateLargeMailboxBytes() {
        var mailbox = new EmailMailbox();
        for (int index = 0; index < 2; index++) {
            var document = new EmailDocument {
                Subject = "Large " + index.ToString(CultureInfo.InvariantCulture)
            };
            document.Body.Text = new string((char)('a' + index), 100_000);
            mailbox.Messages.Add(new EmailMailboxEntry(document));
        }
        return mailbox.ToBytes();
    }

    private static byte[] CreateMailboxWithAttachments() {
        string message =
            "Subject: Attachment\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=store-boundary\r\n\r\n" +
            "--store-boundary\r\nContent-Type: text/plain\r\n\r\nBody\r\n" +
            "--store-boundary\r\nContent-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename=payload.bin\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nMTIzNDU2\r\n" +
            "--store-boundary--\r\n";
        string mailbox =
            "From first@example.test Fri Jul 17 09:00:00 2026\r\n" + message +
            "From second@example.test Fri Jul 17 10:00:00 2026\r\n" + message;
        return Encoding.ASCII.GetBytes(mailbox);
    }

    private sealed class MaximumReadSizeStream : MemoryStream {
        private readonly int _maximumReadSize;

        internal MaximumReadSizeStream(byte[] bytes, int maximumReadSize) : base(bytes, writable: false) {
            _maximumReadSize = maximumReadSize;
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (count > _maximumReadSize) {
                throw new InvalidOperationException("The mbox reader requested an aggregate-sized buffer.");
            }
            return base.Read(buffer, offset, count);
        }
    }

    private sealed class RejectSingleByteBufferReadStream : MemoryStream {
        internal RejectSingleByteBufferReadStream(byte[] bytes) : base(bytes, writable: false) { }

        public override int Read(byte[] buffer, int offset, int count) {
            if (count == 1) {
                throw new InvalidOperationException(
                    "Selected mbox reads must use the source stream's ReadByte implementation.");
            }
            return base.Read(buffer, offset, count);
        }
    }
}
