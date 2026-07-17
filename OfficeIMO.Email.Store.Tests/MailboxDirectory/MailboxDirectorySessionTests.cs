using OfficeIMO.Email;
using System.Globalization;

namespace OfficeIMO.Email.Store.Tests;

public sealed class MailboxDirectorySessionTests {
    [Fact]
    public void Opens_apple_mail_and_maildir_trees_as_one_lazy_store() {
        string root = CreateMailboxDirectory();
        try {
            Assert.Equal(EmailStoreFormat.MailboxDirectory, EmailStoreReader.DetectFormat(root));
            using EmailStoreSession session = EmailStoreSession.Open(root);

            Assert.Equal(EmailStoreFormat.MailboxDirectory, session.Format);
            Assert.Equal(2, session.Folders.Count);
            EmailStoreFolderInfo inbox = Assert.Single(session.Folders, folder => folder.Name == "Inbox");
            EmailStoreFolderInfo sent = Assert.Single(session.Folders, folder => folder.Name == "Sent");
            Assert.Equal(1, inbox.ItemCount);
            Assert.Equal(1, sent.ItemCount);

            EmailStoreItemReference inboxReference = Assert.Single(session.EnumerateItems(
                new EmailStoreEnumerationOptions(folderId: inbox.Id)));
            EmailStoreItemReference sentReference = Assert.Single(session.EnumerateItems(
                new EmailStoreEnumerationOptions(folderId: sent.Id)));
            Assert.Contains("123.emlx", inboxReference.Id);
            Assert.Contains("maildir-message", sentReference.Id);
            Assert.Equal("Apple directory message", session.ReadItem(inboxReference).Document.Subject);
            Assert.Equal("Maildir message", session.ReadItem(sentReference).Document.Subject);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Materializes_and_searches_a_mailbox_directory_through_shared_store_contracts() {
        string root = CreateMailboxDirectory();
        try {
            using EmailStoreSession session = EmailStoreSession.Open(root);

            EmailStoreSearchResult search = Assert.Single(session.Search(
                new EmailStoreQuery(subjectContains: "maildir", maxResults: 1)));
            EmailStoreReadResult materialized = new EmailStoreReader().Read(root);

            Assert.Equal("Maildir message", search.Summary.Subject);
            Assert.Equal(2, materialized.Store.Folders.Sum(folder => folder.Items.Count));
            Assert.Equal(EmailStoreFormat.MailboxDirectory, materialized.Store.Format);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Enforces_mailbox_directory_file_bounds_during_indexing() {
        string root = CreateMailboxDirectory();
        try {
            var options = new EmailStoreReaderOptions(maxDirectoryFileCount: 1);

            EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(
                () => EmailStoreSession.Open(root, options));

            Assert.Equal(nameof(EmailStoreReaderOptions.MaxDirectoryFileCount), exception.LimitName);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Materialized_read_enforces_total_attachment_bytes_across_items() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-mailbox-attachment-budget-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            File.WriteAllText(Path.Combine(root, "first.eml"), CreateMessageWithAttachment("First"));
            File.WriteAllText(Path.Combine(root, "second.eml"), CreateMessageWithAttachment("Second"));
            var options = new EmailStoreReaderOptions(
                maxAttachmentBytes: 10,
                maxTotalAttachmentBytes: 10);

            EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(
                () => new EmailStoreReader(options).Read(root));

            Assert.Equal(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), exception.LimitName);
            Assert.Equal(12, exception.Actual);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void CanonicalMaildirFlagsProjectIntoMessageMetadataAndProperties() {
        string? flags = MailboxDirectoryStoreSessionBackend.ParseMaildirFlags(
            "maildir-message:2,DFPRST");
        var document = new EmailDocument();

        MailboxDirectoryStoreSessionBackend.ApplyMaildirFlags(document, flags);

        Assert.Equal("DFPRST", flags);
        Assert.True(document.MessageMetadata.IsDraft);
        Assert.True(document.MessageMetadata.IsRead);
        Assert.Equal(true, document.Properties["Emlx:Flag:Flagged"]);
        Assert.Equal(true, document.Properties["Emlx:Flag:Forwarded"]);
        Assert.Equal(true, document.Properties["Emlx:Flag:Answered"]);
        Assert.Equal(true, document.Properties["Emlx:Flag:Deleted"]);
    }

    [Fact]
    public void CaseDistinctFoldersRemainDistinctOnCaseSensitiveFileSystems() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-mailbox-directory-case-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            if (EmailStorePathIdentity.IsCaseInsensitiveFileSystem(root)) return;
            string upper = Path.Combine(root, "Inbox");
            string lower = Path.Combine(root, "inbox");
            Directory.CreateDirectory(upper);
            Directory.CreateDirectory(lower);
            File.WriteAllText(Path.Combine(upper, "upper.eml"),
                "Subject: Upper folder\r\n\r\nBody\r\n");
            File.WriteAllText(Path.Combine(lower, "lower.eml"),
                "Subject: Lower folder\r\n\r\nBody\r\n");

            using EmailStoreSession session = EmailStoreSession.Open(root);
            EmailStoreReadResult materialized = session.ReadAll();

            Assert.Equal(2, session.Folders.Count);
            Assert.Contains(session.Folders, folder => folder.Name == "Inbox");
            Assert.Contains(session.Folders, folder => folder.Name == "inbox");
            Assert.Equal(2, materialized.Store.Folders.Sum(folder => folder.Items.Count));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    private static string CreateMailboxDirectory() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mailbox-directory-" + Guid.NewGuid().ToString("N"));
        string apple = Path.Combine(root, "Inbox.mbox", "Messages");
        string maildir = Path.Combine(root, ".Sent", "cur");
        Directory.CreateDirectory(apple);
        Directory.CreateDirectory(maildir);

        byte[] appleMessage = Encoding.ASCII.GetBytes(
            "From: apple@example.test\r\nSubject: Apple directory message\r\n\r\nApple body\r\n");
        File.WriteAllBytes(Path.Combine(apple, "123.emlx"), CreateEmlx(appleMessage));
        File.WriteAllText(Path.Combine(maildir, "maildir-message"),
            "From: maildir@example.test\r\nSubject: Maildir message\r\n\r\nMaildir body\r\n");
        return root;
    }

    private static byte[] CreateEmlx(byte[] message) {
        byte[] prefix = Encoding.ASCII.GetBytes(message.Length.ToString(CultureInfo.InvariantCulture) + "\n");
        var result = new byte[prefix.Length + message.Length];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        Buffer.BlockCopy(message, 0, result, prefix.Length, message.Length);
        return result;
    }

    private static string CreateMessageWithAttachment(string subject) =>
        "Subject: " + subject + "\r\n" +
        "MIME-Version: 1.0\r\n" +
        "Content-Type: multipart/mixed; boundary=store-boundary\r\n\r\n" +
        "--store-boundary\r\nContent-Type: text/plain\r\n\r\nBody\r\n" +
        "--store-boundary\r\nContent-Type: application/octet-stream\r\n" +
        "Content-Disposition: attachment; filename=payload.bin\r\n" +
        "Content-Transfer-Encoding: base64\r\n\r\nMTIzNDU2\r\n" +
        "--store-boundary--\r\n";
}
