namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreSessionTests {
    [Fact]
    public void Opens_catalog_enumerates_references_and_reads_only_selected_item() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        source.Position = 7;

        EmailStoreItemReference reference;
        using (EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst")) {
            Assert.Equal(EmailStoreFormat.Pst, session.Format);
            Assert.Equal("Test Store", session.DisplayName);
            EmailStoreFolderInfo inbox = Assert.Single(session.Folders, folder => folder.Name == "Inbox");

            reference = Assert.Single(session.EnumerateItems(
                new EmailStoreEnumerationOptions(folderId: inbox.Id, maxItems: 1)));
            EmailStoreItem item = session.ReadItem(reference);

            Assert.Equal("Synthetic PST message", item.Document.Subject);
            Assert.Equal(inbox.Id, item.FolderId);
            Assert.False(item.IsAssociated);
        }

        Assert.Equal(7, source.Position);
    }

    [Fact]
    public void Folder_descendant_enumeration_is_explicit() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");
        EmailStoreFolderInfo root = Assert.Single(session.Folders, folder => folder.Name == "Root");

        Assert.Empty(session.EnumerateItems(new EmailStoreEnumerationOptions(folderId: root.Id)));
        Assert.Single(session.EnumerateItems(new EmailStoreEnumerationOptions(
            folderId: root.Id, includeDescendants: true)));
    }

    [Fact]
    public void Default_session_accepts_virtual_sources_larger_than_legacy_eight_gigabyte_limit() {
        const long virtualLength = 64L * 1024 * 1024 * 1024;
        using var source = new VirtualLengthStream(PstTestFileBuilder.Create(), virtualLength);
        using EmailStoreSession session = EmailStoreSession.Open(source, "large.pst");

        Assert.Equal(virtualLength, session.SourceLength);
        Assert.Single(session.EnumerateItems(new EmailStoreEnumerationOptions(maxItems: 1)));
    }

    [Fact]
    public void Explicit_source_limit_still_rejects_oversized_store() {
        const long virtualLength = 64L * 1024 * 1024 * 1024;
        using var source = new VirtualLengthStream(PstTestFileBuilder.Create(), virtualLength);
        var options = new EmailStoreReaderOptions(maxInputBytes: 32L * 1024 * 1024 * 1024);

        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(
            () => EmailStoreSession.Open(source, "large.pst", options));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxInputBytes), exception.LimitName);
    }

    [Fact]
    public void Searches_selective_pst_summaries_without_decoding_message_bodies() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        var options = new EmailStoreReaderOptions(maxPropertiesPerItem: 2);
        using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst", options);

        EmailStoreSearchResult result = Assert.Single(session.Search(new EmailStoreQuery(
            subjectContains: "synthetic pst",
            itemKind: OfficeIMO.Email.OutlookItemKind.Message,
            maxItemsScanned: 1,
            maxResults: 1)));

        Assert.Equal("Synthetic PST message", result.Summary.Subject);
        Assert.Equal("IPM.Note", result.Summary.MessageClass);
        Assert.Null(result.Reference.Summary);
        Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(result.Reference));
    }

    [Fact]
    public void Search_does_not_match_unknown_values_for_explicit_boolean_filters() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");

        Assert.Empty(session.Search(new EmailStoreQuery(hasAttachments: false)));
        Assert.Empty(session.Search(new EmailStoreQuery(isRead: true)));
    }

    [Fact]
    public void Query_rejects_unbounded_or_inverted_ranges() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailStoreQuery(maxItemsScanned: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailStoreQuery(maxResults: 0));
        Assert.Throws<ArgumentException>(() => new EmailStoreQuery(
            since: new DateTimeOffset(2026, 2, 1, 0, 0, 0, TimeSpan.Zero),
            before: new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)));
    }

    private sealed class VirtualLengthStream : Stream {
        private readonly byte[] _data;
        private readonly long _length;
        private long _position;

        internal VirtualLengthStream(byte[] data, long length) {
            _data = data;
            _length = length;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position {
            get => _position;
            set {
                if (value < 0 || value > _length) throw new ArgumentOutOfRangeException(nameof(value));
                _position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (_position >= _data.LongLength) return 0;
            int available = checked((int)Math.Min(count, _data.LongLength - _position));
            Buffer.BlockCopy(_data, checked((int)_position), buffer, offset, available);
            _position += available;
            return available;
        }

        public override long Seek(long offset, SeekOrigin origin) {
            long target = origin == SeekOrigin.Begin ? offset
                : origin == SeekOrigin.Current ? checked(_position + offset)
                : checked(_length + offset);
            Position = target;
            return _position;
        }

        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
