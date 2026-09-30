using System.Security.Cryptography;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreSnapshotTests {
    [Fact]
    public void SnapshotRemainsConsistentAfterCallerMutatesSourceAndRestoresItsPosition() {
        byte[] bytes = Corpus();
        using var source = new MemoryStream(bytes, writable: true);
        source.Position = 17;
        using var snapshot = EmailStoreSession.OpenSnapshot(source, "mail.mbox");
        Assert.True(snapshot.IsSnapshot);
        Assert.Equal(17, source.Position);
        using var hash = SHA256.Create();
        string expected = string.Concat(hash.ComputeHash((byte[])bytes.Clone()).Select(value => value.ToString("x2")));
        Assert.Equal(expected, snapshot.GetDurableSourceFingerprint());
        int subjectPosition = Encoding.ASCII.GetString(bytes).IndexOf("First", StringComparison.Ordinal);
        bytes[subjectPosition] = (byte)'X';
        var page = snapshot.SearchPage(new EmailStoreTableQuery(pageSize: 1, maxItemsScanned: 2));
        Assert.Single(page.Rows);
        var found = snapshot.SearchContent(new EmailStoreContentQuery(new[] { "First" }, fields: EmailStoreContentSearchFields.Subject));
        Assert.Single(found.Results);
        Assert.Equal(expected, snapshot.GetDurableSourceFingerprint());
        Assert.Throws<OperationCanceledException>(() => snapshot.GetDurableSourceFingerprint(new CancellationToken(true)));
    }

    [Fact]
    public void CheckpointsResumeAcrossSnapshotsAndRejectSameLengthSourceMutation() {
        byte[] bytes = Corpus();
        EmailStoreContentSearchCheckpoint checkpoint;
        using (var source = new MemoryStream(bytes))
        using (var snapshot = EmailStoreSession.OpenSnapshot(source, "mail.mbox")) {
            var report = snapshot.SearchContent(Query());
            Assert.Equal(1, report.ItemsScanned);
            checkpoint = EmailStoreContentSearchCheckpoint.Parse(report.NextCheckpoint!.Value);
        }
        using (var source = new MemoryStream(bytes))
        using (var snapshot = EmailStoreSession.OpenSnapshot(source, "mail.mbox")) {
            var resumed = snapshot.SearchContent(Query(checkpoint));
            Assert.Equal("Second", Assert.Single(resumed.Results).Summary.Subject);
        }
        bytes[Encoding.ASCII.GetString(bytes).IndexOf("First", StringComparison.Ordinal)] = (byte)'X';
        using var changed = new MemoryStream(bytes);
        using var changedSnapshot = EmailStoreSession.OpenSnapshot(changed, "mail.mbox");
        Assert.Throws<ArgumentException>(() => changedSnapshot.SearchContent(Query(checkpoint)));
    }

    [Fact]
    public void SnapshotConsumesForwardOnlyInputOnceAndPreservesCallerOwnership() {
        using var source = new ForwardSource(Corpus());
        using (var snapshot = EmailStoreSession.OpenSnapshot(source, "mail.mbox")) {
            Assert.Equal(source.Bytes.Length, source.BytesRead);
            Assert.Equal(2, snapshot.EnumerateItems().Count());
            snapshot.SearchContent(Query());
            snapshot.SearchContent(Query());
            Assert.Equal(source.Bytes.Length, source.BytesRead);
        }
        Assert.True(source.CanRead);
        using var owned = new MemoryStream(Corpus());
        using (EmailStoreSession.OpenSnapshot(owned, "mail.mbox", leaveOpen: false)) { }
        Assert.False(owned.CanRead);
    }

    [Fact]
    public void CopyRejectsOversizedForwardOnlySourceAndRestoresSeekableSourceAfterFailure() {
        using var forward = new ForwardSource(Corpus());
        Assert.Throws<EmailStoreLimitExceededException>(() => EmailStoreSession.OpenSnapshot(forward, "mail.mbox",
            new EmailStoreReaderOptions(maxInputBytes: 20)));
        Assert.Equal(21, forward.BytesRead);
        Assert.True(forward.CanRead);
        using var seekable = new MemoryStream(Encoding.ASCII.GetBytes("unsupported artifact"));
        seekable.Position = 3;
        Assert.Throws<InvalidDataException>(() => EmailStoreSession.OpenSnapshot(seekable, "file.bin"));
        Assert.Equal(3, seekable.Position);
        Assert.Throws<OperationCanceledException>(() => EmailStoreSession.OpenSnapshot(seekable, "file.bin", cancellationToken: new CancellationToken(true)));
        Assert.Equal(3, seekable.Position);
    }

    private static EmailStoreContentQuery Query(EmailStoreContentSearchCheckpoint? checkpoint = null) =>
        new EmailStoreContentQuery(new[] { "First", "Second" }, fields: EmailStoreContentSearchFields.Subject,
            matchMode: EmailStoreContentMatchMode.AnyTerm, maxItemsScanned: 1, maxResults: 2, resumeFrom: checkpoint);

    private static byte[] Corpus() => Encoding.ASCII.GetBytes(
        "From a@example.test Wed Sep 30 12:00:00 2026\nSubject: First\nContent-Type: text/plain\n\nBody\n\n" +
        "From b@example.test Wed Sep 30 12:01:00 2026\nSubject: Second\nContent-Type: text/plain\n\nBody\n");

    private sealed class ForwardSource : Stream {
        private readonly MemoryStream _source;
        internal ForwardSource(byte[] bytes) { Bytes = bytes; _source = new MemoryStream(bytes); }
        internal byte[] Bytes { get; }
        internal int BytesRead { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) { int read = _source.Read(buffer, offset, count); BytesRead += read; return read; }
        public override bool CanRead => _source.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _source.Dispose(); base.Dispose(disposing); }
    }
}
