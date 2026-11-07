using System.Buffers.Binary;
using System.Globalization;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed partial class DbxStoreReaderTests {
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
}
