using OfficeIMO.Email;
using OfficeIMO.Email.Store.Tests.Olm;
using System.Globalization;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstMergeTests {
    [Fact]
    public void MergesMultipleStoresWithSeparateRootsAndSemanticDeduplication() {
        string root = CreateTemporaryDirectory();
        string first = Path.Combine(root, "first.pst");
        string second = Path.Combine(root, "second.pst");
        string destination = Path.Combine(root, "merged.pst");
        try {
            CreateSource(first, "First", CreateDocument("one"), CreateDocument("shared"));
            CreateSource(second, "Second", CreateDocument("shared"), CreateDocument("two"));

            EmailStorePstMergeReport report = EmailStoreConverter.MergeToPst(new[] {
                new EmailStoreMergeSource(first, "Source One"),
                new EmailStoreMergeSource(second, "Source Two")
            }, destination);

            Assert.Equal(4, report.InspectedItems);
            Assert.Equal(3, report.WrittenItems);
            Assert.Equal(1, report.DuplicateItems);
            Assert.Equal(0, report.SkippedItems);
            Assert.Equal(3, report.WriteReport.ItemCount);
            Assert.All(report.Sources, source => Assert.True(source.Completed));
            using EmailStoreSession merged = EmailStoreSession.Open(destination);
            Assert.Equal(3, merged.EnumerateItems().Count());
            Assert.Contains(merged.Folders, folder => folder.Name == "Source One");
            Assert.Contains(merged.Folders, folder => folder.Name == "Source Two");
        } finally {
            DeleteDirectory(root);
        }
    }

    [Fact]
    public void MergeByFolderPathConsolidatesEquivalentHierarchies() {
        string root = CreateTemporaryDirectory();
        string first = Path.Combine(root, "first.pst");
        string second = Path.Combine(root, "second.pst");
        string destination = Path.Combine(root, "merged.pst");
        try {
            CreateSource(first, "First", CreateDocument("one"));
            CreateSource(second, "Second", CreateDocument("two"));

            EmailStorePstMergeReport report = EmailStoreConverter.MergeToPst(new[] {
                new EmailStoreMergeSource(first), new EmailStoreMergeSource(second)
            }, destination, new EmailStorePstMergeOptions(
                folderMode: EmailStoreMergeFolderMode.MergeByFolderPath));

            Assert.Equal(2, report.WrittenItems);
            using EmailStoreSession merged = EmailStoreSession.Open(destination);
            EmailStoreFolderInfo projects = Assert.Single(merged.Folders,
                folder => folder.Name == "Projects");
            Assert.Equal(2, merged.EnumerateItems(new EmailStoreEnumerationOptions(
                projects.Id, includeDescendants: false)).Count());
        } finally {
            DeleteDirectory(root);
        }
    }

    [Fact]
    public void UnreadableSourceIsDiagnosedWithoutDiscardingCompletedSources() {
        string root = CreateTemporaryDirectory();
        string valid = Path.Combine(root, "valid.pst");
        string destination = Path.Combine(root, "merged.pst");
        try {
            CreateSource(valid, "Valid", CreateDocument("kept"));

            EmailStorePstMergeReport report = EmailStoreConverter.MergeToPst(new[] {
                new EmailStoreMergeSource(valid),
                new EmailStoreMergeSource(Path.Combine(root, "missing.pst"))
            }, destination, new EmailStorePstMergeOptions(maxRetries: 1, retryDelay: TimeSpan.Zero));

            Assert.Equal(1, report.WrittenItems);
            Assert.True(report.HasIssues);
            Assert.Equal(2, report.Sources.Count);
            Assert.True(report.Sources[0].Completed);
            Assert.False(report.Sources[1].Completed);
            Assert.Contains(report.Diagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_MERGE_SOURCE_SKIPPED");
            using EmailStoreSession merged = EmailStoreSession.Open(destination);
            Assert.Single(merged.EnumerateItems());
        } finally {
            DeleteDirectory(root);
        }
    }

    [Fact]
    public void MergesPstOlmEmlxAndMailboxDirectoryThroughOneSurface() {
        string root = CreateTemporaryDirectory();
        string pst = Path.Combine(root, "source.pst");
        string olm = Path.Combine(root, "source.olm");
        string emlx = Path.Combine(root, "source.emlx");
        string mailbox = Path.Combine(root, "mailbox");
        string destination = Path.Combine(root, "merged.pst");
        try {
            CreateSource(pst, "PST", CreateDocument("pst"));
            const string olmXml = "<emails><email><OPFMessageCopySubject>olm</OPFMessageCopySubject>" +
                "<OPFMessageCopyBody>body-olm</OPFMessageCopyBody></email></emails>";
            using (var builder = new OlmTestArchiveBuilder()) {
                File.WriteAllBytes(olm, builder.AddText(
                    "Local/com.microsoft.__Messages/Inbox/message_00000.xml", olmXml).Build());
            }
            byte[] emlxMessage = Encoding.ASCII.GetBytes(
                "From: source@example.test\r\nSubject: emlx\r\n\r\nbody-emlx\r\n");
            File.WriteAllBytes(emlx, CreateEmlx(emlxMessage));
            Directory.CreateDirectory(Path.Combine(mailbox, "Inbox"));
            File.WriteAllText(Path.Combine(mailbox, "Inbox", "message.eml"),
                "From: source@example.test\r\nSubject: directory\r\n\r\nbody-directory\r\n");

            EmailStorePstMergeReport report = EmailStoreConverter.MergeToPst(new[] {
                new EmailStoreMergeSource(pst), new EmailStoreMergeSource(olm),
                new EmailStoreMergeSource(emlx), new EmailStoreMergeSource(mailbox)
            }, destination, new EmailStorePstMergeOptions(
                folderMode: EmailStoreMergeFolderMode.Flatten, deduplicate: false));

            Assert.Equal(4, report.WrittenItems);
            Assert.Equal(new[] { EmailStoreFormat.Pst, EmailStoreFormat.Olm,
                EmailStoreFormat.Emlx, EmailStoreFormat.MailboxDirectory },
                report.Sources.Select(source => source.Format));
            using EmailStoreSession merged = EmailStoreSession.Open(destination);
            Assert.Equal(4, merged.EnumerateItems().Count());
        } finally {
            DeleteDirectory(root);
        }
    }

    [Fact]
    public void Destination_cannot_be_created_inside_a_mailbox_directory_source() {
        string root = CreateTemporaryDirectory();
        string mailbox = Path.Combine(root, "mailbox");
        string destination = Path.Combine(mailbox, "merged.pst");
        try {
            Directory.CreateDirectory(mailbox);
            File.WriteAllText(Path.Combine(mailbox, "message.eml"),
                "From: source@example.test\r\nSubject: source\r\n\r\nbody\r\n");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                EmailStoreConverter.MergeToPst(new[] { new EmailStoreMergeSource(mailbox) },
                    destination));

            Assert.Contains("inside", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(destination));
        } finally {
            DeleteDirectory(root);
        }
    }

    private static void CreateSource(string path, string displayName, params EmailDocument[] documents) {
        using var writer = EmailStorePstWriter.Create(path, new EmailStorePstWriterOptions(displayName));
        string folder = writer.AddFolder("Projects");
        foreach (EmailDocument document in documents) writer.AddItem(folder, document);
        writer.Complete();
    }

    private static EmailDocument CreateDocument(string identity) {
        var document = new EmailDocument {
            Subject = identity,
            MessageClass = "IPM.Note",
            Date = new DateTimeOffset(2026, 7, 17, 0, 0, 0, TimeSpan.Zero),
            From = new EmailAddress("sender@example.test", "Sender")
        };
        document.Body.Text = string.Concat("body-", identity);
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("recipient@example.test", "Recipient")));
        document.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin",
            ContentType = "application/octet-stream",
            Content = Encoding.UTF8.GetBytes(string.Concat("attachment-", identity))
        });
        return document;
    }

    private static string CreateTemporaryDirectory() {
        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-merge-tests-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(path);
        return path;
    }

    private static byte[] CreateEmlx(byte[] message) {
        byte[] prefix = Encoding.ASCII.GetBytes(
            string.Concat(message.Length.ToString(CultureInfo.InvariantCulture), "\n"));
        var result = new byte[prefix.Length + message.Length];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        Buffer.BlockCopy(message, 0, result, prefix.Length, message.Length);
        return result;
    }

    private static void DeleteDirectory(string path) {
        try { if (Directory.Exists(path)) Directory.Delete(path, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
