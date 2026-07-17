using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreNativeDirectoryExportTests {
    [Theory]
    [InlineData(EmailStoreNativeDirectoryFormat.Maildir)]
    [InlineData(EmailStoreNativeDirectoryFormat.Emlx)]
    public void NativeDirectoryExportCanBeReadBackThroughMailboxDirectory(
        EmailStoreNativeDirectoryFormat format) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            using var source = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(source, "source.mbox");

            EmailStoreExportReport report = session.ExportToNativeDirectory(root,
                new EmailStoreNativeDirectoryExportOptions(format));
            using EmailStoreSession reopened = EmailStoreSession.Open(root);
            string[] subjects = reopened.EnumerateItems().Select(reference =>
                reopened.ReadSummary(reference).Subject ?? string.Empty).OrderBy(value => value).ToArray();

            Assert.False(report.HasErrors);
            Assert.False(report.WasTruncated);
            Assert.Equal(2, report.SucceededCount);
            Assert.Equal(new[] { "Native first", "Native second" }, subjects);
            Assert.All(report.Entries, entry => Assert.True(File.Exists(entry.DestinationPath)));
            Assert.NotNull(report.ManifestPath);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    private static byte[] CreateMailboxBytes() {
        var mailbox = new EmailMailbox();
        var first = new EmailDocument { Subject = "Native first" };
        first.Body.Text = "First body";
        first.MessageMetadata.IsRead = true;
        mailbox.Messages.Add(new EmailMailboxEntry(first) { EnvelopeSender = "first@example.com" });
        var second = new EmailDocument { Subject = "Native second" };
        second.Body.Text = "Second body";
        mailbox.Messages.Add(new EmailMailboxEntry(second) { EnvelopeSender = "second@example.com" });
        return mailbox.ToBytes();
    }
}
