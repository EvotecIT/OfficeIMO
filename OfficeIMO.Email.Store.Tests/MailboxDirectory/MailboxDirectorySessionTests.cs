using OfficeIMO.Email;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;

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
    public void SummaryReadsDoNotMaterializeMailboxDirectoryAttachments() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-mailbox-summary-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            File.WriteAllText(Path.Combine(root, "message.eml"),
                CreateMessageWithAttachment("Bounded summary").Replace("MTIzNDU2", "%%%"));
            var options = new EmailStoreReaderOptions(
                maxAttachmentBytes: 10,
                maxTotalAttachmentBytes: 10,
                retainAttachmentContent: false);
            using EmailStoreSession session = EmailStoreSession.Open(root, options);
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());

            EmailStoreItemSummary summary = session.ReadSummary(reference);

            Assert.Equal("Bounded summary", summary.Subject);
            Assert.True(summary.HasAttachments);
            EmailStoreDiagnostic diagnostic = Assert.Single(session.Diagnostics,
                item => item.Code == "EMAIL_MIME_BASE64_INVALID");
            Assert.Equal("The invalid Base64 payload was preserved without decoding.", diagnostic.Message);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData("cur")]
    [InlineData("Cur")]
    [InlineData("CUR")]
    public void CanonicalMaildirFlagsProjectIntoMessageMetadataAndProperties(string parentDirectoryName) {
        string? flags = MailboxDirectoryStoreSessionBackend.ParseMaildirFlags(
            "maildir-message:2,DFPRST", parentDirectoryName);
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

    [Theory]
    [InlineData("archive:2,S.eml", "cur")]
    [InlineData("archive:2,S", "new")]
    [InlineData("archive:2,S", "Messages")]
    [InlineData(":2,S", "cur")]
    [InlineData("archive:2,s", "cur")]
    public void MaildirFlagsRequireACanonicalTerminalCurSuffix(string name, string parent) {
        Assert.Null(MailboxDirectoryStoreSessionBackend.ParseMaildirFlags(name, parent));
    }

    [Fact]
    public void MaildirTreatsOpaqueEmlxFileNamesAsRfcMessages() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-maildir-opaque-emlx-" + Guid.NewGuid().ToString("N"));
        try {
            string directory = Path.Combine(root, "new");
            Directory.CreateDirectory(directory);
            File.WriteAllText(Path.Combine(directory, "123.emlx"),
                "From: sender@example.test\r\nSubject: Opaque Maildir name\r\n\r\nBody\r\n");

            using EmailStoreSession session = EmailStoreSession.Open(root);
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            EmailStoreItem item = session.ReadItem(reference);

            Assert.Equal("Opaque Maildir name", item.Document.Subject);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void MailboxDirectoryDetectsEmlxContentInsideMaildirNamedFolders() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-maildir-real-emlx-" + Guid.NewGuid().ToString("N"));
        try {
            string directory = Path.Combine(root, "new");
            Directory.CreateDirectory(directory);
            byte[] message = Encoding.ASCII.GetBytes(
                "From: sender@example.test\r\nSubject: Real EMLX in new\r\n\r\nBody\r\n");
            File.WriteAllBytes(Path.Combine(directory, "123.emlx"), CreateEmlx(message));

            using EmailStoreSession session = EmailStoreSession.Open(root);
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            EmailStoreItem item = session.ReadItem(reference);

            Assert.Equal("Real EMLX in new", item.Document.Subject);
            Assert.Equal((long)message.Length, item.Document.Properties["Emlx:DeclaredMessageBytes"]);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
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

    [Fact]
    public void WindowsPerDirectoryCaseSensitivityKeepsDistinctFolders() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string root = Path.Combine(Path.GetTempPath(),
            "oims-caseflag-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        try {
            Directory.CreateDirectory(root);
            if (!TryEnableWindowsDirectoryCaseSensitivity(root)) return;
            string upper = Path.Combine(root, "Inbox");
            string lower = Path.Combine(root, "inbox");
            Directory.CreateDirectory(upper);
            Directory.CreateDirectory(lower);
            File.WriteAllText(Path.Combine(upper, "upper.eml"),
                "Subject: Upper folder\r\n\r\nBody\r\n");
            File.WriteAllText(Path.Combine(lower, "lower.eml"),
                "Subject: Lower folder\r\n\r\nBody\r\n");

            Assert.False(EmailStorePathIdentity.IsCaseInsensitiveFileSystem(root));
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

    [Fact]
    public void WindowsNestedCaseSensitiveDirectoryKeepsDistinctFolders() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string root = Path.Combine(Path.GetTempPath(),
            "oims-nested-caseflag-" + Guid.NewGuid().ToString("N").Substring(0, 8));
        try {
            Directory.CreateDirectory(root);
            if (!EmailStorePathIdentity.IsCaseInsensitiveFileSystem(root)) return;
            string caseSensitive = Path.Combine(root, "CaseSensitive");
            Directory.CreateDirectory(caseSensitive);
            if (!TryEnableWindowsDirectoryCaseSensitivity(caseSensitive)) return;
            string upper = Path.Combine(caseSensitive, "Inbox");
            string lower = Path.Combine(caseSensitive, "inbox");
            Directory.CreateDirectory(upper);
            Directory.CreateDirectory(lower);
            File.WriteAllText(Path.Combine(upper, "upper.eml"),
                "Subject: Upper nested folder\r\n\r\nBody\r\n");
            File.WriteAllText(Path.Combine(lower, "lower.eml"),
                "Subject: Lower nested folder\r\n\r\nBody\r\n");

            Assert.True(EmailStorePathIdentity.IsCaseInsensitiveFileSystem(root));
            Assert.False(EmailStorePathIdentity.IsCaseInsensitiveFileSystem(caseSensitive));
            using EmailStoreSession session = EmailStoreSession.Open(root);
            EmailStoreReadResult materialized = session.ReadAll();

            Assert.Equal(3, session.Folders.Count);
            EmailStoreFolderInfo upperFolder = Assert.Single(session.Folders,
                folder => folder.Name == "Inbox");
            EmailStoreFolderInfo lowerFolder = Assert.Single(session.Folders,
                folder => folder.Name == "inbox");
            Assert.Equal(upperFolder.ParentId, lowerFolder.ParentId);
            Assert.NotEqual(upperFolder.Id, lowerFolder.Id);
            Assert.Equal(2, materialized.Store.Folders.Sum(folder => folder.Items.Count));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void OpeningMailboxDirectoryDoesNotCreateCaseProbeArtifacts() {
        string root = CreateMailboxDirectory();
        try {
            string[] before = Directory.EnumerateFileSystemEntries(root)
                .Select(path => Path.GetFileName(path) ?? string.Empty)
                .OrderBy(value => value, StringComparer.Ordinal).ToArray();
            using var probeCreated = new ManualResetEventSlim();
            using var watcher = new FileSystemWatcher(root, ".officeimo-case-probe-*") {
                EnableRaisingEvents = true
            };
            watcher.Created += (_, _) => probeCreated.Set();

            using (EmailStoreSession session = EmailStoreSession.Open(root)) {
                Assert.NotEmpty(session.Folders);
            }

            string[] after = Directory.EnumerateFileSystemEntries(root)
                .Select(path => Path.GetFileName(path) ?? string.Empty)
                .OrderBy(value => value, StringComparer.Ordinal).ToArray();
            Assert.Equal(before, after);
            Assert.False(probeCreated.Wait(TimeSpan.FromMilliseconds(250)),
                "Opening the read-only mailbox session created a filesystem case probe in its source.");
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

    private static bool TryEnableWindowsDirectoryCaseSensitivity(string path) {
        var startInfo = new ProcessStartInfo {
            FileName = "fsutil.exe",
            Arguments = "file setCaseSensitiveInfo \"" + path + "\" enable",
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        using Process? process = Process.Start(startInfo);
        if (process == null) return false;
        if (!process.WaitForExit(10_000)) {
            process.Kill();
            return false;
        }
        return process.ExitCode == 0;
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
