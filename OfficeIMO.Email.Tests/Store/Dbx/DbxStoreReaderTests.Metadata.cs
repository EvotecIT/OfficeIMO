using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed partial class DbxStoreReaderTests {
    [Theory]
    [InlineData(4095, 0, false)]
    [InlineData(4096, 0, false)]
    [InlineData(12288, 0, false)]
    [InlineData(4096, 1, false)]
    [InlineData(4096, 2, false)]
    [InlineData(4096, 3, false)]
    [InlineData(4096, 0, true)]
    public void NativeSubjectReadsPastLegacyProbeWithinActualPropertyBudget(int subjectBytes, int limitKind,
        bool unterminated) {
        string subject = new string('s', subjectBytes);
        byte[] database = CreateDatabase();
        int record = database.Length;
        int stringBytes = subjectBytes + (unterminated ? 0 : 1);
        Array.Resize(ref database, record + 12 + 12 + stringBytes);
        UInt32(database, RootNode + 24, record);
        UInt32(database, record, record);
        UInt32(database, record + 4, 12 + stringBytes);
        database[record + 10] = 3;
        UInt32(database, record + 12, (0x09 << 8) | 0x81);
        UInt32(database, record + 16, (OtherMessage << 8) | 0x84);
        UInt32(database, record + 20, 0x05);
        Buffer.BlockCopy(Encoding.ASCII.GetBytes(subject), 0, database, record + 24, subjectBytes);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx",
            new EmailStoreReaderOptions(maxDecodedPropertyBytesPerItem: limitKind == 1 ? 4096 : 128L * 1024 * 1024));
        EmailStoreItemReference reference = session.EnumerateItems().Last();
        if (limitKind == 1) {
            EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadSummary(reference));
            Assert.Equal(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), error.LimitName);
            Assert.Equal(4096, error.Maximum);
            return;
        }
        if (unterminated) {
            Assert.Throws<InvalidDataException>(() => session.ReadSummary(reference));
            return;
        }

        Assert.Equal(subject, session.ReadSummary(reference).NormalizedSubject);
        // The single-part MIME body is six decoded bytes; native string storage also includes its terminator.
        long? itemLimit = limitKind is 2 or 3 ? subjectBytes + 1L + 6 - (limitKind == 2 ? 1 : 0) : null;
        var options = new EmailStoreItemReadOptions(maxDecodedPropertyBytes: itemLimit);
        if (limitKind == 2) {
            EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(reference, options));
            Assert.Equal(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), error.LimitName);
            Assert.Equal(itemLimit, error.Maximum);
            Assert.Equal(itemLimit + 1, error.Actual);
            return;
        }
        EmailDocument document = session.ReadItem(reference, options).Document;
        Assert.Equal(subject, document.MessageMetadata.NormalizedSubject);
        Assert.Equal("Other message", document.Subject);
        Assert.Equal("Body\r\n", document.Body.Text);
        if (subjectBytes == 12288) {
            Assert.Equal(subject, session.ReadAll().Store.Folders.Single().Items.Last().Document.MessageMetadata.NormalizedSubject);
            Assert.Equal(subject, Assert.Single(session.SearchContent(new EmailStoreContentQuery(new[] { "Other message" },
                fields: EmailStoreContentSearchFields.Subject)).Results).Summary.NormalizedSubject);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(new EmailDocumentWriter().ToBytes(document));
            Assert.Equal(document.Subject, exported.Document.Subject);
            Assert.Equal(document.Body.Text, exported.Document.Body.Text);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IndependentDbxIndexedMetadataSurvivesCatalogProjectionAndImportanceExport(bool catalog) {
        string fixtureRoot = Path.Combine(EmailTestRepository.FindRoot(), "OfficeIMO.Email.Tests", "Corpora",
            "dbx-metadata-unknown-producer");
        string sourcePath = Path.Combine(fixtureRoot, catalog ? "Folders.dbx" : "Inbox.dbx");
        using EmailStoreSession session = EmailStoreSession.Open(sourcePath);
        if (catalog) {
            Assert.Equal(new[] { "dbx:folder:00000001", "dbx:folder:00000004", "dbx:folder:00000005",
                "dbx:folder:00000006", "dbx:folder:00000007", "dbx:folder:00000008" },
                session.Folders.Select(folder => folder.Id));
            Assert.Equal(new[] { "Local Folders", "Inbox", "Outbox", "Sent Items", "Deleted Items", "Drafts" },
                session.Folders.Select(folder => folder.Name));
            Assert.All(session.Folders.Skip(1), folder => Assert.Equal("dbx:folder:00000001", folder.ParentId));
        }
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        Assert.Equal(catalog ? "dbx:folder:00000004:dbx:item:00002e10" : "dbx:item:00002e10", reference.Id);
        EmailStoreItemSummary summary = session.ReadSummary(reference);
        Assert.Equal("Welcome to Outlook Express 6", summary.NormalizedSubject);
        Assert.Equal(10139, summary.DeclaredSize);
        EmailDocument document = session.ReadItem(reference).Document;
        Assert.Equal(summary.NormalizedSubject, document.MessageMetadata.NormalizedSubject);
        Assert.Equal(summary.DeclaredSize, document.MessageMetadata.DeclaredSize);
        Assert.Equal(EmailMessageImportance.Normal, document.MessageMetadata.Importance);
        Assert.DoesNotContain(document.Headers, header => header.Name.Equals("Importance", StringComparison.OrdinalIgnoreCase) ||
            header.Name.Equals("X-Priority", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("b1c7745401bac15851730c5d2bb0e754d253af0e799e50fff936b3ad16290968", DbxMetadataHash(document.Body.Html!));
        Assert.Empty(document.Attachments);

        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-native-metadata-" + Guid.NewGuid().ToString("N"));
        try {
            var reports = new List<EmailStoreExportReport> { session.ExportToDirectory(Path.Combine(root, "eml")) };
            foreach (EmailStoreNativeDirectoryFormat format in new[] {
                         EmailStoreNativeDirectoryFormat.Maildir, EmailStoreNativeDirectoryFormat.Emlx }) {
                reports.Add(session.ExportToNativeDirectory(Path.Combine(root, format.ToString()),
                    new EmailStoreNativeDirectoryExportOptions(format)));
            }
            foreach (EmailStoreExportReport report in reports) {
                EmailStoreExportEntry entry = Assert.Single(report.Entries);
                Assert.True(entry.Succeeded);
                EmailDocument reopened;
                if (entry.DestinationPath!.EndsWith(".emlx", StringComparison.OrdinalIgnoreCase)) {
                    using EmailStoreSession exported = EmailStoreSession.Open(entry.DestinationPath);
                    reopened = exported.ReadItem(Assert.Single(exported.EnumerateItems())).Document;
                } else {
                    using EmailReadResult exported = new EmailDocumentReader().ReadDocument(entry.DestinationPath);
                    reopened = exported.Document;
                }
                Assert.Equal(EmailMessageImportance.Normal, reopened.MessageMetadata.Importance);
                Assert.Equal(document.Subject, reopened.Subject);
                Assert.Equal(document.Body.Html, reopened.Body.Html);
                Assert.Empty(reopened.Attachments);
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(1, false, EmailMessageImportance.High)]
    [InlineData(3, true, EmailMessageImportance.Normal)]
    [InlineData(5, false, EmailMessageImportance.Low)]
    [InlineData(0, true, null)]
    [InlineData(7, false, null)]
    public void NativeIndexedMetadataPreservesWindows1252SubjectAndAttachmentBytes(int priority, bool indirectPriority,
        EmailMessageImportance? expectedImportance) {
        byte[] database = CreateDatabase();
        SetIndexedMetadata(database, (uint)priority, 1, indirectPriority);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        EmailStoreItemReference reference = session.EnumerateItems().First();

        Assert.Equal("Native café", session.ReadSummary(reference).NormalizedSubject);
        EmailDocument document = session.ReadItem(reference).Document;

        Assert.Equal("Archived receipt", document.Subject);
        Assert.Equal("Native café", document.MessageMetadata.NormalizedSubject);
        Assert.Equal(expectedImportance, document.MessageMetadata.Importance);
        Assert.Equal(1, document.MessageMetadata.DeclaredSize);
        Assert.Equal("proof bytes", Encoding.ASCII.GetString(Assert.Single(document.Attachments).Content!));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NativeIndexedMetadataRejectsUnterminatedStringsAndTruncatedIntegers(bool unterminatedSubject) {
        byte[] database = CreateDatabase();
        SetIndexedMetadata(database, 3, 1, indirectPriority: false);
        if (unterminatedSubject) database[FirstRecord + 12 + 5 * 4 + 11] = 1;
        else {
            int length = (int)BinaryPrimitives.ReadUInt32LittleEndian(database.AsSpan(FirstRecord + 4, 4));
            UInt32(database, FirstRecord + 4, length - 1);
        }
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");

        Assert.Throws<InvalidDataException>(() => session.ReadSummary(session.EnumerateItems().First()));
    }

    [Fact]
    public void MimeImportanceWinsOverNativeIndexedPriority() {
        byte[] database = CreateDatabase();
        SetIndexedMetadata(database, 1, 1, indirectPriority: true);
        int firstLength = (int)BinaryPrimitives.ReadUInt32LittleEndian(database.AsSpan(FirstFragment + 8, 4));
        int secondLength = (int)BinaryPrimitives.ReadUInt32LittleEndian(database.AsSpan(SecondFragment + 8, 4));
        byte[] first = database.Skip(FirstFragment + 16).Take(firstLength)
            .Concat(database.Skip(SecondFragment + 16).Take(secondLength)).ToArray();
        byte[] message = Encoding.ASCII.GetBytes("Importance: low\r\n").Concat(first).ToArray();
        WriteFragment(database, FirstFragment, message, 0, 512, SecondFragment);
        WriteFragment(database, SecondFragment, message, 512, message.Length - 512, 0);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");

        EmailDocument document = session.ReadItem(session.EnumerateItems().First()).Document;

        Assert.Equal(EmailMessageImportance.Low, document.MessageMetadata.Importance);
        byte[] exported = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        using EmailReadResult reopened = new EmailDocumentReader().ReadDocument(exported);
        Assert.Equal(EmailMessageImportance.Low, reopened.Document.MessageMetadata.Importance);
        Assert.Equal("proof bytes", Encoding.ASCII.GetString(Assert.Single(reopened.Document.Attachments).Content!));
    }

    [Theory]
    [InlineData(1L, 600, true, 1)]
    [InlineData(2147483647L, 2048, false, 2147483647)]
    [InlineData(2147483648L, 2048, false, null)]
    public void NativeDeclaredSizeNeverOverridesFragmentAdmission(long declaredSize, int maximumMessageBytes,
        bool exceedsMessageBudget, int? expectedSize) {
        byte[] database = CreateDatabase();
        SetIndexedMetadata(database, 3, (uint)declaredSize, indirectPriority: false);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx",
            new EmailStoreReaderOptions(maxMessageBytes: maximumMessageBytes));
        EmailStoreItemReference reference = session.EnumerateItems().First();

        Assert.Equal(expectedSize, session.ReadSummary(reference).DeclaredSize);
        if (exceedsMessageBudget) {
            EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(reference));
            Assert.Equal(nameof(EmailStoreReaderOptions.MaxMessageBytes), error.LimitName);
        } else {
            EmailDocument document = session.ReadItem(reference).Document;
            Assert.Equal(expectedSize, document.MessageMetadata.DeclaredSize);
            Assert.Equal("proof bytes", Encoding.ASCII.GetString(Assert.Single(document.Attachments).Content!));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeReceivedDateFillsMissingMimeMetadataAndSurvivesExportManifest(bool hasMimeReceived) {
        var nativeReceived = new DateTimeOffset(2026, 11, 6, 10, 0, 0, TimeSpan.Zero).AddMilliseconds(123);
        var mimeReceived = new DateTimeOffset(2025, 1, 2, 3, 4, 5, TimeSpan.Zero);
        byte[] database = CreateDatabase();
        SetNativeReceivedTime(database, (ulong)nativeReceived.ToFileTime());
        SetOtherMessageDates(database, hasMimeReceived);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        EmailStoreItemReference reference = session.EnumerateItems().Last();
        DateTimeOffset expected = hasMimeReceived ? mimeReceived : nativeReceived;

        Assert.Equal(expected, session.ReadSummary(reference).ReceivedAt);
        EmailStoreItem item = session.ReadItem(reference);
        Assert.Equal(expected, item.Document.ReceivedDate);
        Assert.Equal(new DateTimeOffset(2020, 2, 1, 2, 3, 4, TimeSpan.Zero), item.Document.Date);
        EmailStoreSearchResult[] recent = session.Search(new EmailStoreQuery(since: nativeReceived.AddDays(-1))).ToArray();
        Assert.Equal(hasMimeReceived ? 0 : 1, recent.Length);

        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-metadata-" + Guid.NewGuid().ToString("N"));
        try {
            EmailStoreExportReport export = session.ExportToDirectory(Path.Combine(root, "eml"));
            AssertReceivedManifest(export, reference.Id, expected);
            using EmailReadResult reopened = new EmailDocumentReader().ReadDocument(export.Entries.Last().DestinationPath!);
            Assert.Equal(hasMimeReceived ? mimeReceived : null, reopened.Document.ReceivedDate);
            Assert.True(reopened.Document.MessageMetadata.IsDraft);
            Assert.False(reopened.Document.MessageMetadata.IsRead);
            foreach (EmailStoreNativeDirectoryFormat format in new[] {
                         EmailStoreNativeDirectoryFormat.Maildir, EmailStoreNativeDirectoryFormat.Emlx }) {
                EmailStoreExportReport native = session.ExportToNativeDirectory(Path.Combine(root, format.ToString()),
                    new EmailStoreNativeDirectoryExportOptions(format));
                AssertReceivedManifest(native, reference.Id, expected);
            }
            EmailStoreExportReport withoutManifest = session.ExportToDirectory(Path.Combine(root, "no-manifest"),
                new EmailStoreExportOptions(writeManifest: false));
            Assert.Equal(!hasMimeReceived, withoutManifest.Entries.SelectMany(entry => entry.Diagnostics)
                .Any(diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_RECEIVED_TIME_OMITTED"));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(false, false, false, 1234567)]
    [InlineData(false, true, false, 1234567)]
    [InlineData(false, false, true, 1234567)]
    [InlineData(true, false, false, 1234567)]
    [InlineData(true, true, false, 1234567)]
    [InlineData(true, false, true, 1234567)]
    [InlineData(true, false, false, 0)]
    public void EmlxExportReportsArrivalLossWhenAvailableCarriersCannotPreserveIt(
        bool includeMetadata, bool writeManifest, bool hasMimeReceived, int fractionalTicks) {
        var wholeNativeReceived = new DateTimeOffset(2026, 11, 6, 10, 0, 0, TimeSpan.Zero);
        DateTimeOffset nativeReceived = wholeNativeReceived.AddTicks(fractionalTicks);
        var mimeReceived = new DateTimeOffset(2025, 1, 2, 3, 4, 5, TimeSpan.Zero);
        byte[] database = CreateDatabase();
        SetNativeReceivedTime(database, (ulong)nativeReceived.ToFileTime());
        SetOtherMessageDates(database, hasMimeReceived);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-emlx-carriers-" + Guid.NewGuid().ToString("N"));
        try {
            EmailStoreExportReport export = session.ExportToNativeDirectory(root,
                new EmailStoreNativeDirectoryExportOptions(EmailStoreNativeDirectoryFormat.Emlx,
                    writeManifest: writeManifest,
                    emlxOptions: new EmailStoreEmlxWriterOptions(includeMetadata: includeMetadata)));
            EmailStoreExportEntry entry = export.Entries.Last();
            Assert.True(entry.Succeeded);
            EmailDocument reopened = new EmailStoreReader().Read(entry.DestinationPath!).Store.Folders.Single().Items.Single().Document;
            DateTimeOffset? artifactReceived = includeMetadata ? wholeNativeReceived : null;
            if (hasMimeReceived) artifactReceived = mimeReceived;
            Assert.Equal(artifactReceived, reopened.ReceivedDate);
            bool omitted = !writeManifest && !hasMimeReceived && !includeMetadata;
            bool rounded = !writeManifest && !hasMimeReceived && includeMetadata && fractionalTicks != 0;
            Assert.Equal(omitted, entry.Diagnostics.Any(diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_RECEIVED_TIME_OMITTED"));
            Assert.Equal(rounded, entry.Diagnostics.Any(diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_RECEIVED_TIME_PRECISION"));
            if (writeManifest)
                AssertReceivedManifest(export, session.EnumerateItems().Last().Id, hasMimeReceived ? mimeReceived : nativeReceived);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void InvalidNativeReceivedTimeIsDiagnosedWithoutDroppingTheMessage() {
        byte[] database = CreateDatabase();
        SetNativeReceivedTime(database, ulong.MaxValue);
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");

        EmailStoreItem item = session.ReadItem(session.EnumerateItems().Last());

        Assert.Equal("Other message", item.Document.Subject);
        Assert.Null(item.Document.ReceivedDate);
        Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_DBX_RECEIVED_TIME_INVALID");
    }

    private static void SetNativeReceivedTime(byte[] database, ulong fileTime) {
        UInt32(database, SecondRecord + 4, 24);
        database[SecondRecord + 10] = 3;
        UInt32(database, SecondRecord + 20, (4 << 8) | 0x12);
        UInt32(database, SecondRecord + 24, OtherMessage);
        BinaryPrimitives.WriteUInt64LittleEndian(database.AsSpan(SecondRecord + 28, 8), fileTime);
    }

    private static void SetOtherMessageDates(byte[] database, bool hasMimeReceived) {
        string receivedHeader = hasMimeReceived ? "Received: by example.test; Thu, 02 Jan 2025 03:04:05 +0000\r\n" : string.Empty;
        byte[] message = Encoding.ASCII.GetBytes("Subject: Other message\r\nDate: Sat, 01 Feb 2020 02:03:04 +0000\r\n" +
            receivedHeader + "\r\nBody\r\n");
        WriteFragment(database, OtherMessage, message, 0, message.Length, 0);
    }

    private static void AssertReceivedManifest(EmailStoreExportReport export, string itemId, DateTimeOffset expected) {
        Assert.All(export.Entries, entry => Assert.True(entry.Succeeded));
        string[] lines = File.ReadAllLines(export.ManifestPath!);
        Assert.Equal(new[] { "ItemId", "FolderId", "Associated", "Recovered", "Succeeded", "Bytes", "RelativePath",
            "MaildirFlags", "DiagnosticCount", "ReceivedAtUtc" }, lines[0].Split('\t'));
        string[] row = Assert.Single(lines.Skip(1), line => line.Split('\t')[0] == itemId).Split('\t');
        Assert.Equal(expected.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture), row[9]);
    }

    private static void SetIndexedMetadata(byte[] database, uint priority, uint declaredSize, bool indirectPriority) {
        byte[] subject = Encoding.ASCII.GetBytes("Native caf?\0");
        subject[10] = 0xe9;
        const int propertyCount = 5;
        int data = FirstRecord + 12 + propertyCount * 4;
        int sizeOffset = subject.Length + (indirectPriority ? 4 : 0);
        UInt32(database, FirstRecord + 4, propertyCount * 4 + sizeOffset + 4);
        database[FirstRecord + 10] = propertyCount;
        UInt32(database, FirstRecord + 20, 0x05);
        UInt32(database, FirstRecord + 24, indirectPriority ? (subject.Length << 8) | 0x10 : ((int)priority << 8) | 0x90);
        UInt32(database, FirstRecord + 28, (sizeOffset << 8) | 0x11);
        Buffer.BlockCopy(subject, 0, database, data, subject.Length);
        if (indirectPriority) BinaryPrimitives.WriteUInt32LittleEndian(database.AsSpan(data + subject.Length, 4), priority);
        BinaryPrimitives.WriteUInt32LittleEndian(database.AsSpan(data + sizeOffset, 4), declaredSize);
    }

    private static string DbxMetadataHash(string text) {
        using SHA256 algorithm = SHA256.Create();
        return string.Concat(algorithm.ComputeHash(Encoding.UTF8.GetBytes(text)).Select(value => value.ToString("x2")));
    }
}
