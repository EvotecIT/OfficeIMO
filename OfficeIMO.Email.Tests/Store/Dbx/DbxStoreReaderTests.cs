using System.Buffers.Binary;
using OfficeIMO.Email.Store;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

namespace OfficeIMO.Email.Tests;

public sealed partial class DbxStoreReaderTests {
    private const int RootNode = 0x2800;
    private const int LeftNode = 0x2b00;
    private const int FirstRecord = 0x3100;
    private const int SecondRecord = 0x3140;
    private const int FirstFragment = 0x3300;
    private const int SecondFragment = 0x3510;
    private const int OtherMessage = 0x3720;

    [Fact]
    public void ReadsIndexedDbxMessagesAndRetainsSelectedAttachmentUntilDisposal() {
        byte[] database = CreateDatabase();
        using var input = new MemoryStream(database);
        input.Position = 7;
        Assert.Equal(EmailStoreFormat.Dbx, EmailStoreReader.DetectFormat(input, "renamed.bin"));
        Assert.Equal(7, input.Position);

        IEmailContentSource attachmentSource;
        using (EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx")) {
            Assert.Equal("Inbox", Assert.Single(session.Folders).Name);
            EmailStoreItemReference[] references = session.EnumerateItems().ToArray();
            Assert.Equal(new[] { "Archived receipt", "Other message" },
                references.Select(reference => session.ReadSummary(reference).Subject));
            EmailStoreItem first = session.ReadItem(references[0],
                new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
            Assert.Equal(true, first.Document.MessageMetadata.IsRead);
            Assert.Contains(new string('a', 600), first.Document.Body.Text);
            EmailAttachment attachment = Assert.Single(first.Document.Attachments);
            Assert.Equal("proof.txt", attachment.FileName);
            attachmentSource = Assert.IsAssignableFrom<IEmailContentSource>(attachment.ContentSource);
            using Stream opened = attachmentSource.OpenRead();
            using var text = new StreamReader(opened, Encoding.ASCII);
            Assert.Equal("proof bytes", text.ReadToEnd());
            Assert.True(session.ReadItem(references[1]).Document.MessageMetadata.IsDraft);
        }

        Assert.Throws<ObjectDisposedException>(() => attachmentSource.OpenRead());
        Assert.Equal(7, input.Position);
        Assert.Equal(database, input.ToArray());
    }

    [Fact]
    public void ProjectsDbxThroughReaderAndExportsReopenableEmlThroughStoreSession() {
        byte[] database = CreateDatabase();
        using var input = new MemoryStream(database);
        var reader = new OfficeDocumentReaderBuilder().AddEmailStoreHandler().Build();
        Assert.Contains(reader.ReadDocument(input, "Inbox.dbx").Chunks,
            chunk => chunk.Text.Contains("Archived receipt", StringComparison.Ordinal));

        string destination = Path.Combine(Path.GetTempPath(), "officeimo-dbx-" + Guid.NewGuid().ToString("N"));
        try {
            using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
            EmailStoreExportReport report = session.ExportToDirectory(destination);
            Assert.Equal(2, report.Entries.Count);
            Assert.All(report.Entries, entry => Assert.True(entry.Succeeded));
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(report.Entries[0].DestinationPath!);
            Assert.Equal("Archived receipt", exported.Document.Subject);
            Assert.Equal("proof bytes", Encoding.ASCII.GetString(Assert.Single(exported.Document.Attachments).Content!));
        } finally {
            if (Directory.Exists(destination)) Directory.Delete(destination, recursive: true);
        }
    }

    [Theory]
    [InlineData(FirstFragment)]
    [InlineData(0x4800)]
    public void RejectsCyclicOrOutOfBoundsMessageLinksWhenSelected(int nextFragment) {
        byte[] database = CreateDatabase();
        UInt32(database, FirstFragment + 12, nextFragment);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        EmailStoreItemReference[] references = session.EnumerateItems().ToArray();
        Assert.Equal("Other message", session.ReadItem(references[1]).Document.Subject);
        EmailStoreItemSummary summary = session.ReadSummary(references[0]);
        Assert.Equal("Archived receipt", summary.Subject);
        Assert.Equal(true, summary.HasAttachments);
        Assert.Throws<InvalidDataException>(() => session.ReadItem(references[0]));
    }

    [Fact]
    public void ReadsDbxSummaryAcrossAFragmentBoundaryInTheHeaderSeparator() {
        const int continuation = 0x3930;
        byte[] database = CreateDatabase();
        byte[] message = Encoding.ASCII.GetBytes("Subject: Other message\r\n\r\nBody\r\n");
        int split = Encoding.ASCII.GetByteCount("Subject: Other message\r\n\r");
        WriteFragment(database, OtherMessage, message, 0, split, continuation);
        WriteFragment(database, continuation, message, split, message.Length - split, 0x4800);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        EmailStoreItemReference other = session.EnumerateItems().Last();

        Assert.Equal("Other message", session.ReadSummary(other).Subject);
        Assert.Throws<InvalidDataException>(() => session.ReadItem(other));
    }

    [Fact]
    public void RejectsRepeatedTreeNodeBeforeExposingItems() {
        byte[] database = CreateDatabase();
        UInt32(database, RootNode + 0x1c, LeftNode);
        using var input = new MemoryStream(database);
        Assert.Throws<InvalidDataException>(() => EmailStoreSession.Open(input, "Inbox.dbx"));
    }

    [Fact]
    public void EnforcesMessageBudgetAcrossValidFragments() {
        using var input = new MemoryStream(CreateDatabase());
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx",
            new EmailStoreReaderOptions(maxMessageBytes: 600));
        EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(
            () => session.ReadItem(session.EnumerateItems().First()));
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxMessageBytes), error.LimitName);
    }

    [Fact]
    public void RejectsFolderCatalogAsAStandaloneMessageDatabase() {
        byte[] database = CreateDatabase();
        database[4] = 0xc6;
        using var input = new MemoryStream(database);
        Assert.Equal(EmailStoreFormat.Unknown, EmailStoreReader.DetectFormat(input, "Folders.dbx"));
        Assert.Throws<InvalidDataException>(() => EmailStoreSession.Open(input, "Folders.dbx"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpensLocalDbxCatalogAndExpiresContentAfterSiblingChanges(bool removeMailbox) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-catalog-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string catalogPath = Path.Combine(root, "Folders.dbx");
            string mailboxPath = Path.Combine(root, "Archive.dbx");
            byte[] catalog = CreateFolderCatalog("Archive.dbx");
            byte[] mailbox = CreateDatabase();
            File.WriteAllBytes(catalogPath, catalog);
            File.WriteAllBytes(mailboxPath, mailbox);
            Assert.Equal(EmailStoreFormat.Dbx, EmailStoreReader.DetectFormat(catalogPath));

            using EmailStoreSession session = EmailStoreSession.Open(catalogPath);
            EmailStoreFolderInfo local = Assert.Single(session.FolderCatalog.Roots);
            Assert.Equal("Local Folders", local.Name);
            EmailStoreFolderInfo archive = Assert.Single(session.FolderCatalog.GetChildren(local.Key));
            Assert.Equal("Client archive", archive.Name);
            Assert.Equal(EmailStoreSpecialFolderKind.Inbox, archive.SpecialFolderKind);
            Assert.Equal(EmailStoreFolderClassificationSource.SourceIdentifier, archive.ClassificationSource);
            Assert.Equal(catalog.LongLength + mailbox.LongLength, session.SourceLength);
            EmailStoreItemReference[] items = session.EnumerateItems(local.Key, includeDescendants: true).ToArray();
            Assert.Equal(2, items.Length);
            Assert.All(items, item => Assert.Equal(archive.Id, item.FolderId));
            EmailStoreItem selected = session.ReadItem(items[0],
                new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
            Assert.Equal("Archived receipt", selected.Document.Subject);
            IEmailContentSource content = Assert.IsAssignableFrom<IEmailContentSource>(
                Assert.Single(selected.Document.Attachments).ContentSource);
            Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_DBX_NONLOCAL_FOLDER_SKIPPED");
            Assert.Throws<InvalidOperationException>(() => session.ExportToMbox(mailboxPath));

            string fingerprint = session.GetDurableSourceFingerprint();
            if (removeMailbox) {
                File.Delete(mailboxPath);
                Assert.Throws<InvalidDataException>(() => session.GetDurableSourceFingerprint());
            } else {
                mailbox[mailbox.Length - 1] ^= 1;
                File.WriteAllBytes(mailboxPath, mailbox);
                Assert.NotEqual(fingerprint, session.GetDurableSourceFingerprint());
            }
            Assert.Throws<InvalidDataException>(() => session.ReadItem(items[0]));
            Assert.Throws<ObjectDisposedException>(() => content.OpenRead());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData("../Outside.dbx")]
    [InlineData("sub\\Outside.dbx")]
    [InlineData("NUL.dbx")]
    public void RejectsUnsafeMailboxNamesInDbxFolderCatalog(string fileName) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "Folders.dbx");
        try {
            File.WriteAllBytes(path, CreateFolderCatalog(fileName));
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => EmailStoreSession.Open(path));
            Assert.Contains("unsafe or repeated mailbox file name", error.Message);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ReportsMissingReferencedDbxMailboxWithoutCallingItEmpty() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "Folders.dbx");
        try {
            File.WriteAllBytes(path, CreateFolderCatalog(Guid.NewGuid().ToString("N") + ".dbx"));
            using EmailStoreSession session = EmailStoreSession.Open(path);
            Assert.Null(Assert.Single(session.FolderCatalog.FindSpecialFolders(EmailStoreSpecialFolderKind.Inbox)).ItemCount);
            Assert.Empty(session.EnumerateItems());
            Assert.Contains(session.Diagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_DBX_MAILBOX_MISSING" &&
                diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RejectsAggregateDbxFileLimitsBeforeOpeningTheSiblingIndex(bool bytesLimit) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-bounds-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string catalogPath = Path.Combine(root, "Folders.dbx");
            string mailboxPath = Path.Combine(root, "Archive.dbx");
            byte[] catalog = CreateFolderCatalog("Archive.dbx");
            byte[] mailbox = CreateDatabase();
            mailbox[4] = 0;
            File.WriteAllBytes(catalogPath, catalog);
            File.WriteAllBytes(mailboxPath, mailbox);
            var options = bytesLimit
                ? new EmailStoreReaderOptions(maxInputBytes: catalog.LongLength + mailbox.LongLength - 1)
                : new EmailStoreReaderOptions(maxDirectoryFileCount: 1);

            EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(
                () => EmailStoreSession.Open(catalogPath, options));

            Assert.Equal(bytesLimit ? nameof(EmailStoreReaderOptions.MaxInputBytes) :
                nameof(EmailStoreReaderOptions.MaxDirectoryFileCount), error.LimitName);
            Assert.Equal(bytesLimit ? catalog.LongLength + mailbox.LongLength : 2L, error.Actual);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RejectsAggregateDbxItemLimitBeforeWalkingTheNextSiblingTree() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-items-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string catalogPath = Path.Combine(root, "Folders.dbx");
            File.WriteAllBytes(catalogPath, CreateFolderCatalog("Archive.dbx", "Second.dbx"));
            File.WriteAllBytes(Path.Combine(root, "Archive.dbx"), CreateDatabase());
            byte[] second = CreateDatabase();
            UInt32(second, 0xe4, second.Length + 8);
            File.WriteAllBytes(Path.Combine(root, "Second.dbx"), second);

            EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(
                () => EmailStoreSession.Open(catalogPath, new EmailStoreReaderOptions(maxItemCount: 3)));

            Assert.Equal(nameof(EmailStoreReaderOptions.MaxItemCount), error.LimitName);
            Assert.Equal(4, error.Actual);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ReaderSourceHashIncludesSameLengthChangesToDbxSiblings() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-reader-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string catalogPath = Path.Combine(root, "Folders.dbx");
            string siblingPath = Path.Combine(root, "Archive.dbx");
            byte[] catalog = CreateFolderCatalog("Archive.dbx");
            byte[] sibling = CreateDatabase();
            File.WriteAllBytes(catalogPath, catalog);
            File.WriteAllBytes(siblingPath, sibling);
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddEmailStoreHandler(new ReaderEmailStoreOptions { ComputeSourceHash = true }).Build();

            OfficeDocumentReadResult first = reader.ReadDocument(catalogPath);
            using (EmailStoreSession session = EmailStoreSession.Open(catalogPath)) {
                Assert.Equal(session.GetDurableSourceFingerprint(), first.Source.SourceHash);
            }
            Assert.Equal(catalog.LongLength + sibling.LongLength, first.Source.LengthBytes);
            Assert.NotEmpty(first.Chunks);
            Assert.All(first.Chunks, chunk => {
                Assert.Equal(first.Source.SourceHash, chunk.SourceHash);
                Assert.Equal(first.Source.LengthBytes, chunk.SourceLengthBytes);
            });

            DateTime timestamp = File.GetLastWriteTimeUtc(siblingPath);
            sibling[sibling.Length - 1] ^= 1;
            File.WriteAllBytes(siblingPath, sibling);
            File.SetLastWriteTimeUtc(siblingPath, timestamp);
            OfficeDocumentReadResult second = reader.ReadDocument(catalogPath);
            Assert.Equal(first.Source.SourceId, second.Source.SourceId);
            Assert.Equal(first.Source.LengthBytes, second.Source.LengthBytes);
            Assert.NotEqual(first.Source.SourceHash, second.Source.SourceHash);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static byte[] CreateFolderCatalog(string mailboxName, string? secondMailbox = null) {
        const int localRoot = 0x2b00;
        const int localFolder = 0x2c00;
        const int remoteRoot = 0x2d00;
        byte[] bytes = CreateDatabase();
        Array.Clear(bytes, 0x100, bytes.Length - 0x100);
        bytes[4] = 0xc6;
        UInt32(bytes, 0xc4, 0);
        UInt32(bytes, 0xcc, 3);
        UInt32(bytes, 0xe4, 0);
        UInt32(bytes, 0xec, RootNode);
        UInt32(bytes, RootNode, RootNode);
        bytes[RootNode + 17] = 3;
        UInt32(bytes, RootNode + 24, localRoot);
        UInt32(bytes, RootNode + 36, localFolder);
        UInt32(bytes, RootNode + 48, remoteRoot);
        WriteFolderRecord(bytes, localRoot, 1, 0, 3, "Local Folders", null);
        WriteFolderRecord(bytes, localFolder, 4, 1, 3, "Client archive", mailboxName, 2, 1);
        if (secondMailbox == null) {
            WriteFolderRecord(bytes, remoteRoot, 9, 0, 2, "Remote account", null);
        } else {
            WriteFolderRecord(bytes, remoteRoot, 9, 1, 3, "Second archive", secondMailbox, 2);
        }
        return bytes;
    }

    private static void WriteFolderRecord(byte[] bytes, int address, int id, int parent, int sourceType,
        string name, string? fileName, int count = 0, int role = 0) {
        byte[] encodedName = Encoding.ASCII.GetBytes(name + "\0");
        byte[] encodedFileName = fileName == null ? Array.Empty<byte>() : Encoding.ASCII.GetBytes(fileName + "\0");
        int numberOfFields = fileName == null ? 6 : 7;
        int tableOffset = address + 12;
        UInt32(bytes, address, address);
        UInt32(bytes, address + 4, 4 * numberOfFields + encodedName.Length + encodedFileName.Length);
        bytes[address + 10] = (byte)numberOfFields;
        UInt32(bytes, tableOffset, (id << 8) | 0x80);
        UInt32(bytes, tableOffset + 4, (parent << 8) | 0x81);
        UInt32(bytes, tableOffset + 8, 0x02);
        int field = tableOffset + 12;
        if (fileName != null) {
            UInt32(bytes, field, (encodedName.Length << 8) | 0x03);
            field += 4;
        }
        UInt32(bytes, field, (count << 8) | 0x87);
        UInt32(bytes, field + 4, (role << 8) | 0x89);
        UInt32(bytes, field + 8, (sourceType << 8) | 0x8a);
        int dataOffset = tableOffset + 4 * numberOfFields;
        encodedName.CopyTo(bytes, dataOffset);
        encodedFileName.CopyTo(bytes, dataOffset + encodedName.Length);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void RejectsRemoteDbxMessageDatabases(int sourceType) {
        byte[] database = CreateDatabase();
        database[0x24c0] = (byte)sourceType;
        using var input = new MemoryStream(database);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => EmailStoreSession.Open(input, "Inbox.dbx"));
        Assert.Contains("outside the local mail profile", error.Message);
    }

    [Fact]
    public void ParentOnlyEmailProjectionKeepsAttachedDbxMessagesOpaque() {
        var parent = new EmailDocument();
        parent.Body.Text = "Parent body";
        parent.Attachments.Add(new EmailAttachment {
            FileName = "Inbox.dbx", ContentType = "application/octet-stream", Content = CreateDatabase()
        });
        byte[] bytes = parent.ToBytes();
        var normal = new OfficeDocumentReaderBuilder().AddEmailHandler().AddEmailStoreHandler().Build();
        using (var source = new MemoryStream(bytes)) {
            Assert.Contains(normal.ReadDocument(source, "parent.eml").Chunks,
                chunk => chunk.Text.Contains("Archived receipt", StringComparison.Ordinal));
        }

        var parentOnly = new OfficeDocumentReaderBuilder().AddEmailHandler(new ReaderEmailOptions {
            MessageOptions = new EmailReaderOptions(includeEmbeddedMessages: false)
        }).AddEmailStoreHandler().Build();
        using var restricted = new MemoryStream(bytes);
        IReadOnlyList<ReaderChunk> chunks = parentOnly.ReadDocument(restricted, "parent.eml").Chunks;
        Assert.Contains(chunks, chunk => chunk.Text.Contains("Parent body", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Text.Contains("Inbox.dbx", StringComparison.Ordinal));
        Assert.DoesNotContain(chunks, chunk => chunk.Text.Contains("Archived receipt", StringComparison.Ordinal));
    }

    private static byte[] CreateDatabase() {
        var bytes = new byte[0x4000];
        byte[] headerSignature = {
            0xcf, 0xad, 0x12, 0xfe, 0xc5, 0xfd, 0x74, 0x6f, 0x66, 0xe3, 0xd1, 0x11,
            0x9a, 0x4e, 0x00, 0xc0, 0x4f, 0xa3, 0x09, 0xd4, 0x05, 0x00, 0x00, 0x00
        };
        headerSignature.CopyTo(bytes, 0);
        UInt32(bytes, 0x1c, 5);
        UInt32(bytes, 0x24bc, 1);
        bytes[0x24c0] = 3;
        UInt32(bytes, 0xc4, 2);
        UInt32(bytes, 0xe4, RootNode);
        UInt32(bytes, RootNode, RootNode);
        UInt32(bytes, RootNode + 8, LeftNode);
        UInt32(bytes, RootNode + 20, 1);
        bytes[RootNode + 17] = 1;
        UInt32(bytes, RootNode + 24, SecondRecord);
        UInt32(bytes, LeftNode, LeftNode);
        UInt32(bytes, LeftNode + 12, RootNode);
        bytes[LeftNode + 17] = 1;
        UInt32(bytes, LeftNode + 24, FirstRecord);

        UInt32(bytes, FirstRecord, FirstRecord);
        UInt32(bytes, FirstRecord + 4, 8);
        bytes[FirstRecord + 10] = 2;
        UInt32(bytes, FirstRecord + 12, (0x4081 << 8) | 0x81);
        UInt32(bytes, FirstRecord + 16, (FirstFragment << 8) | 0x84);
        UInt32(bytes, SecondRecord, SecondRecord);
        UInt32(bytes, SecondRecord + 4, 12);
        bytes[SecondRecord + 10] = 2;
        UInt32(bytes, SecondRecord + 12, (0x09 << 8) | 0x81);
        UInt32(bytes, SecondRecord + 16, 0x04);
        UInt32(bytes, SecondRecord + 20, OtherMessage);

        byte[] first = Encoding.ASCII.GetBytes(
            "From: sender@example.test\r\nSubject: Archived receipt\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=part\r\n\r\n--part\r\nContent-Type: text/plain\r\n\r\n" +
            new string('a', 600) + "\r\n--part\r\nContent-Type: text/plain\r\n" +
            "Content-Disposition: attachment; filename=proof.txt\r\nContent-Transfer-Encoding: base64\r\n\r\n" +
            "cHJvb2YgYnl0ZXM=\r\n--part--\r\n");
        WriteFragment(bytes, FirstFragment, first, 0, 512, SecondFragment);
        WriteFragment(bytes, SecondFragment, first, 512, first.Length - 512, 0);
        byte[] other = Encoding.ASCII.GetBytes("Subject: Other message\r\n\r\nBody\r\n");
        WriteFragment(bytes, OtherMessage, other, 0, other.Length, 0);
        return bytes;
    }

    private static void WriteFragment(byte[] destination, int position, byte[] source, int offset, int count,
        int next) {
        UInt32(destination, position, position);
        UInt32(destination, position + 4, 512);
        UInt32(destination, position + 8, count);
        UInt32(destination, position + 12, next);
        Buffer.BlockCopy(source, offset, destination, position + 16, count);
    }

    private static void UInt32(byte[] bytes, int offset, int value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset, 4), (uint)value);
}
