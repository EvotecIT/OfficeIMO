using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed partial class DbxStoreReaderTests {
    [Fact]
    public void ReadAllUsesBoundedSourceValidation() {
        byte[] database = CreateDatabase();
        Array.Resize(ref database, 1024 * 1024);
        using var input = new ObservedDbxStream(database);

        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        input.ResetBytesRead();
        EmailStoreReadResult result = session.ReadAll();
        Assert.Equal(2, result.Store.Folders.Sum(folder => folder.Items.Count));
        Assert.InRange(input.BytesRead, 1, 3L * database.Length);
    }

    [Fact]
    public void ReadAllRejectsSameLengthChangesBeforeReturningAndInvalidatesRetainedAttachments() {
        byte[] database = CreateDatabase();
        using var input = new ObservedDbxStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        EmailStoreItem first = session.ReadItem(session.EnumerateItems().First(),
            new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
        IEmailContentSource retained = Assert.IsAssignableFrom<IEmailContentSource>(
            Assert.Single(first.Document.Attachments).ContentSource);
        input.BeforeRead = position => {
            if (position != OtherMessage + 16) return;
            input.BeforeRead = null;
            database[database.Length - 1] ^= 1;
        };

        Assert.Throws<InvalidDataException>(() => session.ReadAll());
        Assert.Throws<ObjectDisposedException>(() => retained.OpenRead());
        Assert.Throws<InvalidDataException>(() => session.ReadAll());
    }

    [Fact]
    public void CanceledMaterializationReleasesTheBulkScopeForAnotherRead() {
        byte[] database = CreateDatabase();
        using var input = new ObservedDbxStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        using var cancellation = new CancellationTokenSource();
        input.BeforeRead = position => {
            if (position != OtherMessage + 16) return;
            input.BeforeRead = null;
            cancellation.Cancel();
        };

        Assert.Throws<OperationCanceledException>(() => session.ReadAll(cancellation.Token));
        Assert.Equal(2, session.ReadAll().Store.Folders.Sum(folder => folder.Items.Count));
    }

    [Fact]
    public void LiveSearchRejectsSameLengthChangesBetweenYieldedResults() {
        byte[] database = CreateDatabase();
        using var input = new MemoryStream(database);
        using EmailStoreSession session = EmailStoreSession.Open(input, "Inbox.dbx");
        using IEnumerator<EmailStoreSearchResult> results = session.Search(new EmailStoreQuery()).GetEnumerator();
        Assert.True(results.MoveNext());
        Assert.Equal("Archived receipt", results.Current.Summary.Subject);
        database[database.Length - 1] ^= 1;

        Assert.Throws<InvalidDataException>(() => results.MoveNext());
    }

    [Fact]
    public void CatalogBulkReadIncludesEmptySiblingsAndReleasesCanceledScopes() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-bulk-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string catalogPath = Path.Combine(root, "Folders.dbx");
            string emptyPath = Path.Combine(root, "Empty.dbx");
            byte[] catalog = CreateFolderCatalog("Archive.dbx", "Empty.dbx");
            WriteFolderRecord(catalog, 0x2d00, 9, 1, 3, "Empty folder", "Empty.dbx");
            byte[] empty = CreateDatabase();
            UInt32(empty, 0xc4, 0);
            UInt32(empty, 0xe4, 0);
            File.WriteAllBytes(catalogPath, catalog);
            File.WriteAllBytes(Path.Combine(root, "Archive.dbx"), CreateDatabase());
            File.WriteAllBytes(emptyPath, empty);
            using EmailStoreSession session = EmailStoreSession.Open(catalogPath);
            Assert.Contains(session.Folders, folder => folder.Name == "Empty folder" && folder.ItemCount == 0);
            using (var cancellation = new CancellationTokenSource()) {
                using IDisposable canceledScope = Assert.IsAssignableFrom<IDisposable>(
                    session.BeginBulkRead(cancellation.Token));
                cancellation.Cancel();
                Assert.Throws<OperationCanceledException>(() => canceledScope.Dispose());
            }
            Assert.Equal(2, session.ReadAll().Store.Folders.Sum(folder => folder.Items.Count));
            EmailStoreItem first = session.ReadItem(session.EnumerateItems().First(),
                new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
            IEmailContentSource retained = Assert.IsAssignableFrom<IEmailContentSource>(
                Assert.Single(first.Document.Attachments).ContentSource);
            using IDisposable scope = Assert.IsAssignableFrom<IDisposable>(session.BeginBulkRead(CancellationToken.None));
            empty[empty.Length - 1] ^= 1;
            File.WriteAllBytes(emptyPath, empty);

            Assert.Throws<InvalidDataException>(() => scope.Dispose());
            Assert.Throws<ObjectDisposedException>(() => retained.OpenRead());
            Assert.Throws<InvalidDataException>(() => session.ReadAll());
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void ContentSearchValidatesTerminalSourceAndReleasesDbxScopes(bool catalog, bool cancel,
        bool exhaustedResume) {
        byte[] database = CreateDatabase();
        byte[] empty = CreateDatabase();
        using var input = new ObservedDbxStream(database);
        string root = Path.Combine(Path.GetTempPath(), "officeimo-dbx-content-search-" + Guid.NewGuid().ToString("N"));
        string catalogPath = Path.Combine(root, "Folders.dbx");
        string emptyPath = Path.Combine(root, "Empty.dbx");
        try {
            if (catalog) {
                Directory.CreateDirectory(root);
                byte[] folders = CreateFolderCatalog("Archive.dbx", "Empty.dbx");
                WriteFolderRecord(folders, 0x2d00, 9, 1, 3, "Empty folder", "Empty.dbx");
                UInt32(empty, 0xc4, 0);
                UInt32(empty, 0xe4, 0);
                File.WriteAllBytes(catalogPath, folders);
                File.WriteAllBytes(Path.Combine(root, "Archive.dbx"), database);
                File.WriteAllBytes(emptyPath, empty);
            }
            using EmailStoreSession session = catalog
                ? EmailStoreSession.Open(catalogPath) : EmailStoreSession.Open(input, "Inbox.dbx");
            EmailStoreItem first = session.ReadItem(session.EnumerateItems().First(),
                new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
            IEmailContentSource retained = Assert.IsAssignableFrom<IEmailContentSource>(
                Assert.Single(first.Document.Attachments).ContentSource);
            EmailStoreContentQuery Query(EmailStoreContentSearchCheckpoint? resume = null) => new(
                new[] { "aaa", "Body" }, fields: EmailStoreContentSearchFields.TextBody,
                matchMode: EmailStoreContentMatchMode.AnyTerm, maxItemsScanned: 10, maxResults: 10,
                progressInterval: 1, resumeFrom: resume);
            EmailStoreContentQuery query = Query();
            if (exhaustedResume) query = Query(EmailStoreContentSearchCheckpoint.Create(
                3, session.GetDurableSourceFingerprint(), query.Signature));
            using var cancellation = new CancellationTokenSource();
            bool triggered = false;
            var progress = new DbxContentSearchProgress(value => {
                if (triggered || !exhaustedResume && value.ItemsScanned != 2) return;
                triggered = true;
                if (cancel) cancellation.Cancel();
                else if (catalog) {
                    empty[empty.Length - 1] ^= 1;
                    File.WriteAllBytes(emptyPath, empty);
                } else database[database.Length - 1] ^= 1;
            });

            if (cancel) {
                Assert.Throws<OperationCanceledException>(() => session.SearchContent(query, progress, cancellation.Token));
                Assert.True(triggered);
                EmailStoreContentSearchReport retry = session.SearchContent(Query());
                Assert.True(retry.IsComplete);
                Assert.Equal(new[] { "Archived receipt", "Other message" }, retry.Results.Select(result => result.Summary.Subject));
                using Stream content = retained.OpenRead();
                using var text = new StreamReader(content, Encoding.ASCII);
                Assert.Equal("proof bytes", text.ReadToEnd());
            } else {
                Assert.Throws<InvalidDataException>(() => session.SearchContent(query, progress));
                Assert.True(triggered);
                Assert.Throws<ObjectDisposedException>(() => retained.OpenRead());
                Assert.Throws<InvalidDataException>(() => session.ReadAll());
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    private sealed class DbxContentSearchProgress : IProgress<EmailStoreContentSearchProgress> {
        private readonly Action<EmailStoreContentSearchProgress> _report;
        internal DbxContentSearchProgress(Action<EmailStoreContentSearchProgress> report) => _report = report;
        public void Report(EmailStoreContentSearchProgress value) => _report(value);
    }

    private sealed class ObservedDbxStream : Stream {
        private readonly MemoryStream _source;

        internal ObservedDbxStream(byte[] bytes) { _source = new MemoryStream(bytes, writable: false); }
        internal long BytesRead { get; private set; }
        internal Action<long>? BeforeRead { get; set; }
        internal void ResetBytesRead() => BytesRead = 0;

        public override int Read(byte[] buffer, int offset, int count) {
            BeforeRead?.Invoke(Position);
            int read = _source.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }

#if !NETFRAMEWORK
        public override int Read(Span<byte> buffer) {
            BeforeRead?.Invoke(Position);
            int read = _source.Read(buffer);
            BytesRead += read;
            return read;
        }
#endif

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _source.Length;
        public override long Position { get => _source.Position; set => _source.Position = value; }
        public override long Seek(long offset, SeekOrigin origin) => _source.Seek(offset, origin);
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _source.Dispose();
            base.Dispose(disposing);
        }
    }
}
