using OfficeIMO.Email;
using System.Threading;
using System.Threading.Tasks;

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
        Assert.True(source.TotalBytesRead < 1024 * 1024,
            $"Opening and enumerating a virtual 64 GiB store read {source.TotalBytesRead} bytes.");
    }

    [Fact]
    public void Selective_workflows_stay_bounded_on_a_virtual_sixty_four_gibibyte_store() {
        const long virtualLength = 64L * 1024 * 1024 * 1024;
        byte[] attachmentContent = Enumerable.Repeat((byte)0x5A, 4096).ToArray();
        using var source = new VirtualLengthStream(
            PstTestFileBuilder.Create(attachmentContent: attachmentContent), virtualLength);
        var readerOptions = new EmailStoreReaderOptions(
            maxCachedBTreePages: 4,
            retainAttachmentContent: false);
        using EmailStoreSession session = EmailStoreSession.Open(source, "large.pst", readerOptions);

        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems(
            new EmailStoreEnumerationOptions(maxItems: 1)));
        EmailStoreItemSummary summary = session.ReadSummary(reference);
        EmailStoreItem item = session.ReadItem(reference, new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.Metadata |
            EmailStoreItemReadParts.Bodies |
            EmailStoreItemReadParts.AttachmentMetadata |
            EmailStoreItemReadParts.AttachmentContent,
            maxDecodedPropertyBytes: 1024 * 1024,
            preferStreamingAttachmentContent: true));
        EmailAttachment attachment = Assert.Single(item.Document.Attachments);

        Assert.Equal("Synthetic PST message", summary.Subject);
        Assert.Null(attachment.Content);
        Assert.NotNull(attachment.ContentSource);
        long beforeOpen = source.TotalBytesRead;
        using Stream content = attachment.OpenContentStream();
        Assert.Equal(beforeOpen, source.TotalBytesRead);
        Assert.Equal(0x5A, content.ReadByte());
        Assert.True(source.TotalBytesRead > beforeOpen);

        EmailStoreContentSearchReport search = session.SearchContent(
            new EmailStoreContentQuery(
                new[] { "PST property context" },
                fields: EmailStoreContentSearchFields.TextBody,
                maxItemsScanned: 1,
                maxResults: 1,
                maxDecodedPropertyBytesPerItem: 1024 * 1024,
                maxSearchableCharactersPerItem: 4096));
        EmailStoreValidationReport validation = session.Validate(
            new EmailStoreValidationOptions(
                mode: EmailStoreValidationMode.Shallow,
                maxItems: 1,
                verifyStructuralIntegrity: true,
                maxStructuralPages: 4,
                maxStructuralBlocks: 4,
                maxStructuralBytes: 64 * 1024));

        Assert.Single(search.Results);
        Assert.True(validation.StructuralIntegritySupported);
        Assert.Equal(0, validation.StructuralFailures);
        Assert.True(source.TotalBytesRead < 4 * 1024 * 1024,
            $"Selective workflows over a virtual 64 GiB store read {source.TotalBytesRead} bytes.");
        Assert.True(source.MaxSingleRead <= 128 * 1024,
            $"A single source read requested {source.MaxSingleRead} bytes.");
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
        using var source = new MemoryStream(PstTestFileBuilder.Create(attachmentContent: new byte[128]));
        var options = new EmailStoreReaderOptions(
            maxPropertiesPerItem: 3,
            maxAttachmentBytes: 64);
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
        using var source = new MemoryStream(PstTestFileBuilder.Create(attachmentContent: new byte[128]));
        var readerOptions = new EmailStoreReaderOptions(
            maxPropertiesPerItem: 3,
            maxAttachmentBytes: 64);
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
            string path = Assert.Single(first.Entries).DestinationPath!;
            byte[] sentinel = Encoding.ASCII.GetBytes("existing artifact");
            File.WriteAllBytes(path, sentinel);

            EmailStoreExportReport second = session.ExportToDirectory(destination);

            Assert.Equal(1, first.SucceededCount);
            EmailStoreExportEntry failed = Assert.Single(second.Entries);
            Assert.False(failed.Succeeded);
            Assert.Contains(failed.Diagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_DESTINATION_EXISTS");
            Assert.Equal(sentinel, File.ReadAllBytes(path));
            Assert.Contains(second.Diagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_MANIFEST_EXISTS");
        } finally {
            if (Directory.Exists(destination)) Directory.Delete(destination, recursive: true);
        }
    }

    [Theory]
    [InlineData("directory")]
    [InlineData("mbox")]
    [InlineData("pst")]
    public void MailboxDirectoryExportsRejectDestinationsInsideTheirSourceTree(string exportKind) {
        string sourceRoot = Path.Combine(Path.GetTempPath(),
            "oims-export-" + Guid.NewGuid().ToString("N").Substring(0, 12));
        try {
            Directory.CreateDirectory(sourceRoot);
            string sourceMessage = Path.Combine(sourceRoot, "source.eml");
            File.WriteAllText(sourceMessage, "Subject: Source\r\n\r\nBody\r\n");
            using EmailStoreSession session = EmailStoreSession.Open(sourceRoot);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => {
                if (exportKind == "directory") {
                    session.ExportToDirectory(Path.Combine(sourceRoot, "export"));
                } else if (exportKind == "mbox") {
                    session.ExportToMbox(Path.Combine(sourceRoot, "export.mbox"));
                } else {
                    session.ExportToPst(Path.Combine(sourceRoot, "export.pst"));
                }
            });

            Assert.Contains("source tree", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(new[] { sourceMessage }, Directory.EnumerateFiles(
                sourceRoot, "*", SearchOption.AllDirectories).ToArray());
        } finally {
            if (Directory.Exists(sourceRoot)) Directory.Delete(sourceRoot, recursive: true);
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

    [Fact]
    public void Mbox_export_continues_after_attachment_content_access_is_denied() {
        string destination = Path.Combine(Path.GetTempPath(),
            "officeimo-store-denied-" + Guid.NewGuid().ToString("N") + ".mbox");
        try {
            using EmailStoreSession session = CreateSession(new AttachmentAccessDeniedBackend());

            EmailStoreMboxExportReport report = session.ExportToMbox(destination,
                new EmailStoreMboxExportOptions(continueOnError: true));

            Assert.Equal(2, report.Entries.Count);
            Assert.Equal(1, report.SucceededCount);
            Assert.Contains(report.Entries, entry => entry.Diagnostics.Any(diagnostic =>
                diagnostic.Code == "EMAIL_STORE_EXPORT_ITEM_FAILED" &&
                diagnostic.Message.Contains("denied", StringComparison.OrdinalIgnoreCase)));
            Assert.Equal("Valid", Assert.Single(EmailMailbox.Load(destination).Messages).Document.Subject);
        } finally {
            if (File.Exists(destination)) File.Delete(destination);
        }
    }

    private static EmailStoreSession CreateSession(IEmailStoreSessionBackend backend) {
        System.Reflection.ConstructorInfo? constructor = typeof(EmailStoreSession).GetConstructor(
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            binder: null,
            new[] {
                typeof(Stream), typeof(bool), typeof(long), typeof(EmailStoreReaderOptions),
                typeof(IEmailStoreSessionBackend)
            },
            modifiers: null);
        Assert.NotNull(constructor);
        return (EmailStoreSession)constructor!.Invoke(new object[] {
            Stream.Null, true, 0L, EmailStoreReaderOptions.Default, backend
        });
    }

    private sealed class AttachmentAccessDeniedBackend : IEmailStoreSessionBackend {
        private const string FolderId = "folder";
        private readonly EmailStoreItemReference[] _references = {
            new EmailStoreItemReference("denied", FolderId, false, false),
            new EmailStoreItemReference("valid", FolderId, false, false)
        };
        private readonly EmailStoreFolderInfo[] _folders = {
            new EmailStoreFolderInfo(FolderId, null, "Inbox")
        };

        public EmailStoreFormat Format => EmailStoreFormat.Mbox;
        public string? DisplayName => "Test";
        public long SourceLength => 0;
        public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
        public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => Array.Empty<EmailStoreDiagnostic>();

        public IEnumerable<EmailStoreItemReference> EnumerateItems(
            EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
            foreach (EmailStoreItemReference reference in _references.Take(options.MaxItems)) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return reference;
            }
        }

        public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
            CancellationToken cancellationToken) =>
            EmailStoreItemSummary.FromItem(ReadItem(
                reference, EmailStoreItemReadOptions.Default, cancellationToken));

        public EmailStoreItem ReadItem(EmailStoreItemReference reference,
            EmailStoreItemReadOptions options, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            var document = new EmailDocument {
                Subject = reference.Id == "denied" ? "Denied" : "Valid",
                Body = { Text = "Body" }
            };
            if (reference.Id == "denied") {
                document.Attachments.Add(new EmailAttachment {
                    FileName = "denied.bin",
                    ContentType = "application/octet-stream",
                    Length = 1,
                    ContentSource = new DeniedContentSource()
                });
            }
            return new EmailStoreItem(reference.Id, reference.FolderId, document,
                format: EmailStoreFormat.Mbox);
        }

        public void Dispose() { }
    }

    private sealed class DeniedContentSource : IEmailContentSource {
        public long? Length => 1;

        public Stream OpenRead() =>
            throw new UnauthorizedAccessException("Attachment content access was denied.");

        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) =>
            Task.FromException<Stream>(
                new UnauthorizedAccessException("Attachment content access was denied."));
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
        internal long TotalBytesRead { get; private set; }
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
            TotalBytesRead += available;
            MaxSingleRead = Math.Max(MaxSingleRead, available);
            return available;
        }

        internal int MaxSingleRead { get; private set; }

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
