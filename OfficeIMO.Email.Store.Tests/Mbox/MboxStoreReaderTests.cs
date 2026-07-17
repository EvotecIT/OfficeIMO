using OfficeIMO.Email;
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
}
