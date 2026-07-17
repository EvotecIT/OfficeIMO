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
            if (format == EmailStoreNativeDirectoryFormat.Maildir) {
                Assert.Equal("S", report.Entries[0].MaildirFlags);
                string manifest = File.ReadAllText(report.ManifestPath!);
                Assert.Contains("\tMaildirFlags\t", manifest);
                Assert.Contains("\tS\t", manifest);
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void EmlxExportDoesNotReplaceAnExistingDestinationByDefault() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            using var source = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(source, "source.mbox");
            var options = new EmailStoreNativeDirectoryExportOptions(
                EmailStoreNativeDirectoryFormat.Emlx, maxItems: 1);
            EmailStoreExportReport first = session.ExportToNativeDirectory(root, options);
            string path = Assert.Single(first.Entries).DestinationPath!;
            byte[] sentinel = Encoding.ASCII.GetBytes("existing destination");
            File.WriteAllBytes(path, sentinel);

            EmailStoreExportReport second = session.ExportToNativeDirectory(root, options);

            Assert.True(second.HasErrors);
            Assert.Equal(sentinel, File.ReadAllBytes(path));
            Assert.DoesNotContain(Directory.EnumerateFiles(Path.GetDirectoryName(path)!),
                candidate => candidate.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void NativeDirectoryExportDoesNotReplaceAnExistingManifestByDefault() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-native-export-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            string manifestPath = Path.Combine(root, "officeimo-email-store-export.tsv");
            byte[] sentinel = Encoding.ASCII.GetBytes("existing manifest");
            File.WriteAllBytes(manifestPath, sentinel);
            using var source = new MemoryStream(CreateMailboxBytes());
            using EmailStoreSession session = EmailStoreSession.Open(source, "source.mbox");

            EmailStoreExportReport report = session.ExportToNativeDirectory(root,
                new EmailStoreNativeDirectoryExportOptions(
                    EmailStoreNativeDirectoryFormat.Maildir, maxItems: 1));

            Assert.Null(report.ManifestPath);
            Assert.Equal(sentinel, File.ReadAllBytes(manifestPath));
            Assert.Contains(report.Diagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_MANIFEST_EXISTS");
            Assert.DoesNotContain(Directory.EnumerateFiles(root),
                path => Path.GetFileName(path).StartsWith(".officeimo-email-store-export.",
                    StringComparison.OrdinalIgnoreCase));
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
