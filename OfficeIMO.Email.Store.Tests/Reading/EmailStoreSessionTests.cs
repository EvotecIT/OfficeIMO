using OfficeIMO.Email;

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

    [Fact]
    public void Inspects_catalog_without_enumerating_item_payloads() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");

        EmailStoreInspectionReport report = session.Inspect();

        Assert.Equal(EmailStoreFormat.Pst, report.Format);
        Assert.Equal("Test Store", report.DisplayName);
        Assert.Equal(2, report.FolderCount);
        Assert.Equal(2, report.FoldersWithUnknownItemCount);
        Assert.False(report.HasCompleteDeclaredItemCount);
        Assert.False(report.HasErrors);
    }

    [Fact]
    public void Validates_summaries_and_reports_bounded_full_item_failures() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        var readerOptions = new EmailStoreReaderOptions(maxPropertiesPerItem: 2);
        using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst", readerOptions);

        EmailStoreValidationReport summaries = session.Validate(
            new EmailStoreValidationOptions(mode: EmailStoreValidationMode.Summaries));
        EmailStoreValidationReport full = session.Validate(
            new EmailStoreValidationOptions(mode: EmailStoreValidationMode.FullItems));

        Assert.True(summaries.IsComplete);
        Assert.True(summaries.IsValid);
        Assert.Equal(1, summaries.ItemsExamined);
        Assert.False(full.IsComplete);
        Assert.True(full.IsValid);
        Assert.Equal(1, full.ItemsFailed);
        Assert.Contains(full.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_VALIDATION_ITEM_LIMIT");
    }

    [Fact]
    public void Recovery_discovery_does_not_misclassify_index_fallback_as_orphaned() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");

        EmailStoreRecoveryReport report = session.DiscoverRecoverableItems();

        Assert.Equal(1, report.ItemsScanned);
        Assert.Empty(report.RecoveredItems);
        Assert.False(report.StoppedAtLimit);
    }

    [Fact]
    public void Validation_and_recovery_require_positive_bounds() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailStoreValidationOptions(maxItems: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailStoreRecoveryOptions(maxItemsScanned: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailStoreRecoveryOptions(maxRecoveredItems: 0));
    }

    [Fact]
    public void Exports_selected_items_through_officeimo_email_with_a_manifest() {
        string destination = Path.Combine(Path.GetTempPath(), "officeimo-store-export-" + Guid.NewGuid().ToString("N"));
        try {
            using var source = new MemoryStream(PstTestFileBuilder.Create());
            using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");

            EmailStoreExportReport report = session.ExportToDirectory(destination);

            EmailStoreExportEntry entry = Assert.Single(report.Entries);
            Assert.True(entry.Succeeded);
            Assert.Equal(1, report.SucceededCount);
            Assert.NotNull(entry.DestinationPath);
            Assert.EndsWith(".eml", entry.DestinationPath!, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Root", entry.DestinationPath!);
            Assert.Contains("Inbox", entry.DestinationPath!);
            Assert.NotNull(report.ManifestPath);
            Assert.True(File.Exists(report.ManifestPath!));
            Assert.Contains(entry.Reference.Id, File.ReadAllText(report.ManifestPath!));

            EmailDocument exported = EmailDocument.Load(entry.DestinationPath!);
            Assert.Equal("Synthetic PST message", exported.Subject);
            Assert.Equal("Body from the PST property context", exported.Body.Text);
        } finally {
            if (Directory.Exists(destination)) Directory.Delete(destination, recursive: true);
        }
    }

    [Fact]
    public void Export_does_not_replace_existing_artifacts_without_explicit_policy() {
        string destination = Path.Combine(Path.GetTempPath(), "officeimo-store-export-" + Guid.NewGuid().ToString("N"));
        try {
            using var source = new MemoryStream(PstTestFileBuilder.Create());
            using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");
            EmailStoreExportReport first = session.ExportToDirectory(destination);

            EmailStoreExportReport second = session.ExportToDirectory(destination);

            Assert.Equal(1, first.SucceededCount);
            EmailStoreExportEntry failed = Assert.Single(second.Entries);
            Assert.False(failed.Succeeded);
            Assert.Contains(failed.Diagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_DESTINATION_EXISTS");
            Assert.Contains(second.Diagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_MANIFEST_EXISTS");
        } finally {
            if (Directory.Exists(destination)) Directory.Delete(destination, recursive: true);
        }
    }

    [Fact]
    public void Streams_store_items_to_an_atomically_committed_mbox() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-store-mbox-" + Guid.NewGuid().ToString("N"));
        string destination = Path.Combine(directory, "archive.mbox");
        try {
            using var source = new MemoryStream(PstTestFileBuilder.Create());
            using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");

            EmailStoreMboxExportReport report = session.ExportToMbox(destination);

            Assert.Equal(Path.GetFullPath(destination), report.DestinationPath);
            Assert.Equal(1, report.SucceededCount);
            Assert.False(report.HasErrors);
            Assert.True(report.BytesWritten > 0);
            Assert.False(report.WasTruncated);
            EmailMailbox mailbox = EmailMailbox.Load(destination);
            Assert.Equal("Synthetic PST message", Assert.Single(mailbox.Messages).Document.Subject);
            Assert.Empty(Directory.GetFiles(directory, "*.tmp"));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Mbox_export_refuses_existing_destination_by_default() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-store-mbox-" + Guid.NewGuid().ToString("N"));
        string destination = Path.Combine(directory, "archive.mbox");
        try {
            Directory.CreateDirectory(directory);
            File.WriteAllText(destination, "existing");
            using var source = new MemoryStream(PstTestFileBuilder.Create());
            using EmailStoreSession session = EmailStoreSession.Open(source, "archive.pst");

            EmailStoreMboxExportReport report = session.ExportToMbox(destination);

            Assert.Null(report.DestinationPath);
            Assert.True(report.HasErrors);
            Assert.Equal("existing", File.ReadAllText(destination));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
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
