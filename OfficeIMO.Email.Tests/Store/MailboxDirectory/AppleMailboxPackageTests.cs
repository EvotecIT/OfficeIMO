using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class AppleMailboxPackageTests {
    [Fact]
    public void ExportPackagesUseSharedMboxReaderAndPreserveEmptyNestedMailboxes() {
        using var tree = new MailboxTree();
        tree.Write("Exported.mbox/mbox", Mbox("First") + Mbox("Second"));
        tree.Write("Exported.mbox/Empty.mbox/mbox", string.Empty);
        tree.Write("Exported.mbox/Child.mbox/mbox", Mbox("Child"));
        using EmailStoreSession session = EmailStoreSession.Open(tree.Path);

        EmailStoreFolderInfo exported = Assert.Single(session.Folders, folder => folder.Name == "Exported");
        EmailStoreFolderInfo empty = Assert.Single(session.Folders, folder => folder.Name == "Empty");
        Assert.Equal(exported.Id, empty.ParentId);
        Assert.Equal(2, exported.ItemCount);
        Assert.Equal(0, empty.ItemCount);
        EmailStoreItemReference[] items = session.EnumerateItems().ToArray();
        Assert.Equal(new[] { "Child", "First", "Second" }, items.Select(item => session.ReadSummary(item).Subject));
        Assert.All(items, reference => {
            EmailStoreItem item = session.ReadItem(reference);
            Assert.Equal(reference.FolderId, item.FolderId);
            Assert.Equal(reference.Id, item.Document.Properties["EmailStore:ItemId"]);
            Assert.Contains("Body", item.Document.Body.Text);
        });

        // Opening the package itself must expose the same messages as opening its parent tree.
        using EmailStoreSession package = EmailStoreSession.Open(System.IO.Path.Combine(tree.Path, "Exported.mbox"));
        Assert.Equal(3, package.EnumerateItems().Count());
        EmailStoreFolderInfo packageRoot = Assert.Single(package.Folders, folder => folder.ParentId == null);
        Assert.Equal(3, package.EnumerateItems(new EmailStoreEnumerationOptions(packageRoot.Id, includeDescendants: true)).Count());
        Assert.DoesNotContain(session.Diagnostics, diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);
        Assert.DoesNotContain(package.Diagnostics, diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);
    }

    [Fact]
    public void EmptyMailboxIsValidInStandaloneAndPackageReaders() {
        using var tree = new MailboxTree();
        tree.Write("Empty.mbox/mbox", string.Empty);
        using EmailStoreSession package = EmailStoreSession.Open(tree.Path);
        Assert.Empty(package.EnumerateItems());
        Assert.Single(package.Folders);
        Assert.Empty(package.Diagnostics);
        var reader = new EmailMailboxReader();
        Assert.Empty(reader.Read(Array.Empty<byte>()).Diagnostics);
        using var empty = new MemoryStream();
        Assert.Empty(reader.ReadEntries(empty));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectedAggregateAttachmentsUseSessionOwnedStreamingSources(bool package) {
        using var tree = new MailboxTree();
        var document = new EmailDocument { Subject = "Aggregate streaming" };
        document.Body.Text = "Body";
        byte[] expected = Enumerable.Range(0, 65536).Select(index => (byte)index).ToArray();
        document.Attachments.Add(new EmailAttachment { FileName = "payload.bin", Content = expected });
        string path = System.IO.Path.Combine(tree.Path, package ? "Exported.mbox/mbox" : "source.mbox");
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(document));
        using (var output = File.Create(path)) new EmailMailboxWriter().Write(mailbox, output);
        IEmailContentSource content;
        Stream outstanding;
        using (EmailStoreSession session = EmailStoreSession.Open(package ? System.IO.Path.GetDirectoryName(path)! : path,
            new EmailStoreReaderOptions(maxAttachmentBytes: expected.Length, maxTotalAttachmentBytes: expected.Length))) {
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            var options = new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true);
            EmailAttachment attachment = Assert.Single(session.ReadItem(reference, options).Document.Attachments);
            Assert.Null(attachment.Content);
            content = Assert.IsAssignableFrom<IEmailContentSource>(attachment.ContentSource);
            outstanding = content.OpenRead();
            using Stream input = content.OpenRead();
            using var result = new MemoryStream();
            input.CopyTo(result);
            Assert.Equal(expected, result.ToArray());
            Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(reference, options));
        }
        Assert.Throws<ObjectDisposedException>(() => content.OpenRead());
        Assert.Throws<ObjectDisposedException>(() => outstanding.ReadByte());
        outstanding.Dispose();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StreamingSelectionPreservesHeaderlessBodiesInStandaloneAndPackageMailboxes(bool package) {
        using var tree = new MailboxTree();
        string relative = package ? "Exported.mbox/mbox" : "source.mbox";
        tree.Write(relative, "From synthetic@example.test Sat Sep 26 12:00:00 2026\r\nplain body\r\n>From escaped body\r\n");
        string path = System.IO.Path.Combine(tree.Path, relative);
        using EmailStoreSession session = EmailStoreSession.Open(package ? System.IO.Path.GetDirectoryName(path)! : path);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        EmailDocument normal = session.ReadItem(reference).Document;
        EmailDocument streaming = session.ReadItem(reference,
            new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true)).Document;

        Assert.Equal(normal.Body.Text, streaming.Body.Text);
        Assert.Equal("plain body\r\nFrom escaped body", streaming.Body.Text!.Trim());
        Assert.Equal(normal.Properties["Mbox:EnvelopeSender"], streaming.Properties["Mbox:EnvelopeSender"]);
        Assert.Equal(normal.Properties["Mbox:EnvelopeDate"], streaming.Properties["Mbox:EnvelopeDate"]);
        Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MBOX_MESSAGE_HEADERS_MISSING");
        Assert.DoesNotContain(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_FORMAT_UNKNOWN");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StreamingMboxSelectionKeepsEnvelopeAndEscapedBodySemantics(bool rd) {
        using var tree = new MailboxTree();
        string body = rd ? ">From one\r\n>>From two\r\n>>>From three\r\n>other\r\n>>>" : ">From one\r\n>other\r\n>>>";
        tree.Write("source.mbox", "From synthetic@example.test Sat Sep 26 12:00:00 2026\r\nSubject: Escapes\r\nContent-Type: text/plain\r\n\r\n" + body);
        using EmailStoreSession session = EmailStoreSession.Open(System.IO.Path.Combine(tree.Path, "source.mbox"));
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        EmailDocument normal = session.ReadItem(reference).Document;
        EmailDocument streaming = session.ReadItem(reference,
            new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true)).Document;
        Assert.Equal(normal.Body.Text, streaming.Body.Text);
        Assert.Contains("From one", streaming.Body.Text);
        Assert.EndsWith(">>>", streaming.Body.Text);
        Assert.Equal(normal.Properties["Mbox:EnvelopeSender"], streaming.Properties["Mbox:EnvelopeSender"]);
        Assert.Equal(normal.Properties["Mbox:EnvelopeDate"], streaming.Properties["Mbox:EnvelopeDate"]);
    }

    [Fact]
    public void EqualMailboxNamesInDifferentAccountsKeepSeparateIdentityAndScopes() {
        using var tree = new MailboxTree();
        tree.Write("Account-A/Inbox.mbox/Messages/1.eml", "Subject: A\r\n\r\nBody A");
        tree.Write("Account-B/Inbox.mbox/Messages/2.eml", "Subject: B\r\n\r\nBody B");
        using EmailStoreSession session = EmailStoreSession.Open(tree.Path);

        EmailStoreFolderInfo[] inboxes = session.Folders.Where(folder => folder.Name == "Inbox").ToArray();
        Assert.Equal(2, inboxes.Length);
        Assert.NotEqual(inboxes[0].Id, inboxes[1].Id);
        Assert.NotEqual(inboxes[0].ParentId, inboxes[1].ParentId);
        EmailStoreFolderInfo account = Assert.Single(session.Folders, folder => folder.Name == "Account-A");
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems(
            new EmailStoreEnumerationOptions(account.Id, includeDescendants: true)));
        Assert.Equal("A", session.ReadSummary(reference).Subject);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FolderBoundsIncludeParentsAndEmptyMailboxFolders(bool empty) {
        using var tree = new MailboxTree();
        if (empty) Directory.CreateDirectory(System.IO.Path.Combine(tree.Path, "Account", "Empty.mbox"));
        else tree.Write("Account/Inbox.mbox/Messages/1.eml", "Subject: Limit\r\n\r\nBody");

        EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() =>
            EmailStoreSession.Open(tree.Path, new EmailStoreReaderOptions(maxFolderCount: 1)));
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxFolderCount), error.LimitName);
        Assert.Equal(2, error.Actual);
    }

    [Fact]
    public void AggregateAndIndividualMessagesShareOneItemBudget() {
        using var tree = new MailboxTree();
        tree.Write("Exported.mbox/mbox", Mbox("First"));
        tree.Write("message.eml", "Subject: Second\r\n\r\nBody");

        EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() =>
            EmailStoreSession.Open(tree.Path, new EmailStoreReaderOptions(maxItemCount: 1)));
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxItemCount), error.LimitName);
    }

    [Fact]
    public void TraversalBoundsAlsoCountUnrelatedFilesAndEmptyDirectories() {
        using var tree = new MailboxTree();
        tree.Write("unrelated.txt", "not a message");
        Directory.CreateDirectory(System.IO.Path.Combine(tree.Path, "Empty"));
        EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() =>
            EmailStoreSession.Open(tree.Path, new EmailStoreReaderOptions(maxDirectoryEntryCount: 1)));
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxDirectoryEntryCount), error.LimitName);
        Assert.Equal(2, error.Actual);
    }

    [Fact]
    public void AppleDoubleSidecarsAreNotMessagesAndRepeatedReadsDoNotDuplicateDiagnostics() {
        using var tree = new MailboxTree();
        tree.Write("._message.emlx", "AppleDouble filesystem data");
        // The exact prefix is computed from bytes, independent of the property-list trailer.
        string message = "Subject: Test\r\n\r\nX\r\n";
        tree.Write("message.emlx", Encoding.UTF8.GetByteCount(message) + "\n" + message +
            "<plist><dict><key>broken</key></dict></plist>");
        using EmailStoreSession session = EmailStoreSession.Open(tree.Path);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        for (int index = 0; index < 100; index++) Assert.Equal("Test", session.ReadSummary(reference).Subject);

        Assert.Single(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_DIRECTORY_APPLEDOUBLE_SKIPPED");
        Assert.Single(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_EMLX_METADATA_INVALID");
    }

    [Theory]
    [InlineData("add")]
    [InlineData("rename")]
    [InlineData("remove")]
    [InlineData("empty-folder")]
    public void DurableOperationsRejectChangedCatalogMembership(string change) {
        using var tree = new MailboxTree();
        tree.Write("original.eml", "Subject: Original\r\n\r\nBody");
        using EmailStoreSession session = EmailStoreSession.Open(tree.Path);
        string fingerprint = session.GetDurableSourceFingerprint();
        if (change == "add") tree.Write("added.eml", "Subject: Added\r\n\r\nBody");
        else if (change == "rename") File.Move(System.IO.Path.Combine(tree.Path, "original.eml"),
            System.IO.Path.Combine(tree.Path, "renamed.eml"));
        else if (change == "remove") File.Delete(System.IO.Path.Combine(tree.Path, "original.eml"));
        else Directory.CreateDirectory(System.IO.Path.Combine(tree.Path, "Empty.mbox"));

        Assert.Throws<InvalidDataException>(() => session.GetDurableSourceFingerprint());
        Assert.Throws<InvalidDataException>(() => session.GetCatalogFingerprint());
        using EmailStoreSession fresh = EmailStoreSession.Open(tree.Path);
        Assert.NotEqual(fingerprint, fresh.GetDurableSourceFingerprint());
    }

    private static string Mbox(string subject) =>
        "From synthetic@example.test Sat Sep 26 12:00:00 2026\nSubject: " + subject + "\n\nBody\n\n";

    private sealed class MailboxTree : IDisposable {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "officeimo-apple-mailbox-" + Guid.NewGuid().ToString("N"));
        internal MailboxTree() => Directory.CreateDirectory(Path);
        internal void Write(string relativePath, string text) {
            string path = System.IO.Path.Combine(Path, relativePath.Replace('/', System.IO.Path.DirectorySeparatorChar));
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
            File.WriteAllText(path, text, new UTF8Encoding(false));
        }
        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
