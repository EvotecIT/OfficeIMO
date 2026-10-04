using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailArchiveAnalysisTests {
    [Theory]
    [InlineData("directory")]
    [InlineData("mbox")]
    [InlineData("emlx")]
    public void AttachedMessagesRemainParentMetadataCandidates(string format) {
        string root = CreateDirectory();
        string eml = "Subject: Parent\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain\r\n\r\nparent\r\n" +
            "--x\r\nContent-Type: message/rfc822; name=child.eml\r\nContent-Disposition: attachment; filename=child.eml\r\n\r\n" +
            "Subject: Child\r\nContent-Type: text/plain\r\n\r\n" + new string('z', 2000) + "\r\n--x--\r\n";
        try {
            string path = root;
            if (format == "mbox") {
                path = Path.Combine(root, "mail.mbox");
                File.WriteAllText(path, "From a@example.test Thu Jan 01 09:00:00 2026\n" + eml +
                    "From a@example.test Thu Jan 01 09:00:00 2026\n" + eml, new UTF8Encoding(false));
            } else if (format == "emlx") {
                path = Path.Combine(root, "mail.emlx");
                File.WriteAllText(path, Encoding.UTF8.GetByteCount(eml) + "\n" + eml, new UTF8Encoding(false));
            } else {
                File.WriteAllText(Path.Combine(root, "a.eml"), eml, new UTF8Encoding(false));
                File.WriteAllText(Path.Combine(root, "b.eml"), eml, new UTF8Encoding(false));
            }
            using var session = EmailStoreSession.Open(path);
            if (format != "emlx") {
                var reference = session.EnumerateItems().First();
                var parent = session.ReadItem(reference, new EmailStoreItemReadOptions(
                    EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.Bodies | EmailStoreItemReadParts.AttachmentMetadata,
                    maxDecodedPropertyBytes: 64));
                Assert.Null(Assert.Single(parent.Document.Attachments).EmbeddedDocument);
                Assert.Null(parent.Document.Attachments[0].Content);
                Assert.False(parent.LoadedParts.HasFlag(EmailStoreItemReadParts.EmbeddedItems));
            }
            var report = session.AnalyzeArchive(new EmailArchiveAnalysisOptions(largeAttachmentThresholdBytes: 1));
            Assert.Equal(0, report.ItemsFailed);
            Assert.Equal(format == "emlx" ? 1 : 2, report.ItemsProjected);
            if (format != "emlx") Assert.Equal(2, Assert.Single(report.DuplicateCandidates).ItemCount);
            Assert.All(report.LargeAttachments, attachment => Assert.True(attachment.IsEmbeddedItem));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void ExplicitParentOnlyReadRetainsMetadataAndDefaultReadsRetainEmbeddedMessages(EmailFileFormat format) {
        var child = new EmailDocument { Subject = "child" }; child.Body.Text = "nested";
        var source = new EmailDocument { Subject = "parent" }; source.Body.Text = "parent";
        source.Attachments.Add(new EmailAttachment { FileName = "child.eml", EmbeddedDocument = child });
        byte[] bytes = new EmailDocumentWriter().ToBytes(source, format);
        var parent = new EmailDocumentReader(new EmailReaderOptions(includeAttachmentContent: false,
            includeEmbeddedMessages: false)).Read(bytes).Document;
        var attachment = Assert.Single(parent.Attachments);
        Assert.Null(attachment.EmbeddedDocument); Assert.Null(attachment.Content);
        Assert.Empty(attachment.StructuredStorageStreams); Assert.True(attachment.Length > 0);
        Assert.All(attachment.MapiProperties.Where(MapiKnownProperties.PidTag.AttachData.MatchesIdentity), property => {
            Assert.Null(property.Value); Assert.Null(property.RawData);
        });
        var full = new EmailDocumentReader().Read(bytes).Document;
        Assert.Equal("child", Assert.Single(full.Attachments).EmbeddedDocument!.Subject);
        var policy = new EmailSemanticComparisonOptions(includeAttachmentContent: false,
            maxEmbeddedMessageDepth: 0, includeEmbeddedMessageContent: false);
        var first = EmailSemanticComparer.CreateFingerprint(full, policy).HexDigest;
        full.Attachments[0].EmbeddedDocument!.Subject = "different nested content";
        Assert.Equal(first, EmailSemanticComparer.CreateFingerprint(full, policy).HexDigest);
    }

    [Fact]
    public void ReportsMetadataCandidatesAndLargestAttachmentsWithoutClaimingPayloadEquality() {
        string root = CreateDirectory(); string path = Path.Combine(root, "analysis.pst");
        try {
            using (var writer = EmailStorePstWriter.Create(path)) {
                string a = writer.AddFolder("A"), b = writer.AddFolder("B");
                writer.AddItem(a, Message("same", 1, new byte[] { 1, 2, 3 }));
                writer.AddItem(a, Message("same", 1, new byte[] { 4, 5, 6 }));
                writer.AddItem(b, Message("other", 2, new byte[] { 7, 8, 9, 10, 11 }));
                writer.AddItem(b, Message("other", 2, new byte[] { 12, 13, 14, 15, 16 }));
                writer.Complete();
            }
            byte[] before = File.ReadAllBytes(path);
            using var session = EmailStoreSession.Open(path, new EmailStoreReaderOptions(retainAttachmentContent: false));
            var report = session.AnalyzeArchive(new EmailArchiveAnalysisOptions(maxDuplicateGroups: 1,
                maxItemsPerDuplicateGroup: 1, maxLargeAttachments: 1, largeAttachmentThresholdBytes: 1));
            Assert.Equal(4, report.ItemsScanned); Assert.Equal(4, report.ItemsProjected); Assert.Equal(0, report.ItemsFailed);
            Assert.True(report.ExhaustedSelectedReferences);
            Assert.Equal(2, report.DuplicateCandidateGroupCount); Assert.True(report.DuplicateGroupsTruncated);
            var candidates = Assert.Single(report.DuplicateCandidates);
            Assert.Equal(2, candidates.ItemCount); Assert.Single(candidates.ItemIds); Assert.True(candidates.ItemIdsTruncated);
            Assert.True(report.EstimatedCandidateDeclaredBytes > 0);
            Assert.Equal(4, report.LargeAttachmentCount);
            Assert.Equal(5, Assert.Single(report.LargeAttachments).DeclaredBytes);
            Assert.Equal(4, report.Folders.Sum(bucket => bucket.Count));
            Assert.Equal(new[] { "2026-01", "2026-02" }, report.UtcMonths.Select(bucket => bucket.Key));
            using var hash = SHA256.Create();
            Assert.Equal(BitConverter.ToString(hash.ComputeHash(before)).Replace("-", "").ToLowerInvariant(), report.SourceFingerprint);
            Assert.Equal(before, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void ReportsScanAndDistributionBoundsWithoutLosingTotals() {
        string root = CreateDirectory(); string path = Path.Combine(root, "analysis.pst");
        try {
            using (var writer = EmailStorePstWriter.Create(path)) {
                string a = writer.AddFolder("A"), b = writer.AddFolder("B");
                writer.AddItem(a, Message("A", 1)); writer.AddItem(b, Message("B", 2)); writer.Complete();
            }
            using var session = EmailStoreSession.Open(path);
            var bounded = session.AnalyzeArchive(new EmailArchiveAnalysisOptions(maxItems: 1));
            Assert.Equal(1, bounded.ItemsScanned); Assert.True(bounded.StoppedAtItemLimit); Assert.False(bounded.ExhaustedSelectedReferences);
            var buckets = session.AnalyzeArchive(new EmailArchiveAnalysisOptions(maxDistributionBuckets: 1));
            Assert.Equal(2, buckets.Folders.Sum(bucket => bucket.Count) + buckets.ItemsInOtherFolders);
            Assert.Equal(1, buckets.ItemsInOtherFolders);
            Assert.Equal(2, buckets.UtcMonths.Sum(bucket => bucket.Count) + buckets.ItemsInOtherMonths);
            Assert.Equal(1, buckets.ItemsInOtherMonths);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void ExcludesPartialLocalItemsAndSamplesBoundedProjectionFailures() {
        string root = CreateDirectory();
        try {
            byte[] eml = Encoding.UTF8.GetBytes("Subject: Same\r\nContent-Type: text/plain\r\n\r\n" + new string('x', 1000));
            File.WriteAllBytes(Path.Combine(root, "01.eml"), eml);
            byte[] prefix = Encoding.ASCII.GetBytes(eml.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            File.WriteAllBytes(Path.Combine(root, "02.partial.emlx"), prefix.Concat(eml).ToArray());
            using var session = EmailStoreSession.Open(root);
            var complete = session.AnalyzeArchive();
            Assert.Equal(2, complete.ItemsProjected); Assert.Equal(1, complete.PartialItemsExcludedFromCandidates);
            Assert.Empty(complete.DuplicateCandidates);
            var failed = session.AnalyzeArchive(new EmailArchiveAnalysisOptions(maxDecodedPropertyBytesPerItem: 10, maxDiagnostics: 1));
            Assert.Equal(2, failed.ItemsScanned); Assert.Equal(2, failed.ItemsFailed);
            Assert.Equal(2, failed.ItemsWithoutDate); Assert.Empty(failed.DuplicateCandidates);
            Assert.Single(failed.Diagnostics); Assert.True(failed.DiagnosticCount >= 2);
            Assert.DoesNotContain(new string('x', 100), failed.Diagnostics[0].Message);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void RejectsSourceMutationAndCancellation() {
        byte[] mbox = Encoding.ASCII.GetBytes("From sender@example.test Thu Jan 01 09:00:00 2026\nSubject: Same\n\nbody\n");
        using var stream = new MutatingStream(mbox);
        using var session = EmailStoreSession.Open(stream, "archive.mbox");
        stream.Armed = true;
        Assert.Throws<InvalidDataException>(() => session.AnalyzeArchive());
        using var stable = new MemoryStream(mbox);
        using var stableSession = EmailStoreSession.Open(stable, "archive.mbox");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => stableSession.AnalyzeArchive(cancellationToken: cancellation.Token));
    }

    private static EmailDocument Message(string subject, int month, byte[]? attachment = null) {
        var message = new EmailDocument { Subject = subject, Date = new DateTimeOffset(2026, month, 1, 9, 0, 0, TimeSpan.Zero), From = new EmailAddress("a@example.test") };
        message.Body.Text = subject + " body";
        if (attachment != null) message.Attachments.Add(new EmailAttachment { FileName = "data.bin", ContentType = "application/octet-stream", Content = attachment, Length = attachment.Length });
        return message;
    }
    private static string CreateDirectory() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-archive-analysis-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path); return path;
    }
    private sealed class MutatingStream : MemoryStream {
        internal bool Armed;
        private bool _changed;
        internal MutatingStream(byte[] bytes) : base((byte[])bytes.Clone(), 0, bytes.Length, true, true) { }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (Armed && !_changed && read > 0 && Position == Length) {
                GetBuffer()[checked((int)Length - 2)] ^= 1; _changed = true;
            }
            return read;
        }
    }
}
