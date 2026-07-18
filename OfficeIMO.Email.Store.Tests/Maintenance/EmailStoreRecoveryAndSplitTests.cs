using OfficeIMO.Email;
using System.Threading;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreRecoveryAndSplitTests {
    [Fact]
    public void Recovery_export_attempts_only_index_orphans_and_isolates_corrupt_items_with_manifest_evidence() {
        string root = TemporaryDirectory();
        try {
            var backend = new RecoveryBackend();
            using EmailStoreSession session = CreateSession(backend);

            EmailStoreRecoveryExportReport report = session.ExportRecoverableItemsToDirectory(
                root,
                new EmailStoreRecoveryExportOptions(
                    format: EmailFileFormat.Eml,
                    continueOnItemError: true,
                    writeManifest: true));

            Assert.Equal(2, report.Discovery.RecoveredItems.Count);
            Assert.Equal(2, report.Entries.Count);
            Assert.Equal(1, report.RecoveredCount);
            Assert.Equal(1, report.FailedCount);
            Assert.False(report.IsComplete);
            Assert.NotNull(report.ManifestPath);
            string manifest = File.ReadAllText(report.ManifestPath!);
            Assert.Contains("orphan-good\tinbox\tfalse\ttrue\ttrue", manifest, StringComparison.Ordinal);
            Assert.Contains("orphan-bad\tinbox\tfalse\ttrue\tfalse", manifest, StringComparison.Ordinal);
            Assert.DoesNotContain("normal\tinbox", manifest, StringComparison.Ordinal);
            Assert.Contains(report.Entries.Single(entry => entry.Reference.Id == "orphan-bad").Diagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_EXPORT_ITEM_FAILED");
        } finally {
            DeleteDirectory(root);
        }
    }

    [Fact]
    public void Query_and_estimated_size_split_dry_runs_then_commits_only_verified_parts() {
        string root = TemporaryDirectory();
        string source = Path.Combine(root, "source.pst");
        try {
            CreateSourcePst(source);
            using EmailStoreSession session = EmailStoreSession.Open(source);
            var query = new EmailStoreTableQuery(
                filter: EmailStoreFields.Subject.Contains("Keep"),
                sorts: new[] { EmailStoreFields.Subject.Ascending() },
                maxItemsScanned: 20,
                pageSize: 20);
            var options = new EmailStorePstSplitOptions(
                query,
                maxEstimatedBytesPerPart: 1,
                failOnDataLoss: true,
                includeSearchFolders: false);

            EmailStorePstSplitPlan plan = session.PlanPstSplit(
                Path.Combine(root, "selected.pst"), options);

            Assert.True(plan.IsExecutable);
            Assert.Equal(2, plan.MatchedItems);
            Assert.Equal(2, plan.Parts.Count);
            Assert.All(plan.Parts, part => Assert.Single(part.Items));
            Assert.All(plan.Parts, part => Assert.True(part.ContainsOversizedItem));
            Assert.DoesNotContain(plan.Parts, part => File.Exists(part.DestinationPath));

            EmailStorePstSplitReport report = session.ExecutePstSplit(plan);

            Assert.True(report.IsSuccessful);
            Assert.Equal(2, report.WrittenItems);
            Assert.All(report.Parts, part => {
                Assert.True(File.Exists(part.WriteReport.DestinationPath));
                Assert.True(part.Verification.IsSuccessful);
                Assert.Equal(1, part.Verification.MatchedItems);
                using EmailStoreSession output = EmailStoreSession.Open(part.WriteReport.DestinationPath);
                EmailStoreItemReference item = Assert.Single(output.EnumerateItems());
                Assert.Contains("Keep", output.ReadSummary(item).Subject, StringComparison.OrdinalIgnoreCase);
            });
        } finally {
            DeleteDirectory(root);
        }
    }

    [Fact]
    public void Split_blocks_execution_when_query_scan_or_part_bound_omits_matches() {
        string root = TemporaryDirectory();
        string source = Path.Combine(root, "source.pst");
        try {
            CreateSourcePst(source);
            using EmailStoreSession session = EmailStoreSession.Open(source);
            var query = new EmailStoreTableQuery(
                filter: EmailStoreFilter.All,
                sorts: new[] { EmailStoreFields.ItemId.Ascending() },
                maxItemsScanned: 2,
                pageSize: 2);
            EmailStorePstSplitPlan plan = session.PlanPstSplit(
                Path.Combine(root, "bounded.pst"),
                new EmailStorePstSplitOptions(query, maxEstimatedBytesPerPart: 1, maxParts: 1));

            Assert.False(plan.IsExecutable);
            Assert.True(plan.ScanLimitReached);
            Assert.Contains(plan.Diagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_PST_SPLIT_QUERY_LIMIT");
            Assert.Throws<InvalidOperationException>(() => session.ExecutePstSplit(plan));
            Assert.Empty(Directory.GetFiles(root, "bounded.part*.pst"));
        } finally {
            DeleteDirectory(root);
        }
    }

    [Fact]
    public void Compaction_plans_then_rewrites_to_a_distinct_verified_pst_without_touching_source() {
        string root = TemporaryDirectory();
        string source = Path.Combine(root, "source.pst");
        string destination = Path.Combine(root, "compacted.pst");
        try {
            CreateSourcePst(source);
            byte[] before = File.ReadAllBytes(source);
            using EmailStoreSession session = EmailStoreSession.Open(source);

            EmailStorePstCompactionPlan plan = session.PlanPstCompaction(destination);

            Assert.True(plan.IsExecutable);
            Assert.Equal(3, plan.SelectedItems);
            Assert.Equal(before.LongLength, plan.SourceBytes);
            Assert.False(File.Exists(destination));

            EmailStorePstCompactionReport report = session.CompactToPst(destination);

            Assert.True(report.IsVerified);
            Assert.False(report.HasDataLoss);
            Assert.Equal(3, report.Conversion.Verification!.MatchedItems);
            Assert.Equal(before, File.ReadAllBytes(source));
            using EmailStoreSession compacted = EmailStoreSession.Open(destination);
            Assert.Equal(3, compacted.EnumerateItems().Count());
        } finally {
            DeleteDirectory(root);
        }
    }

    private static void CreateSourcePst(string path) {
        using EmailStorePstWriter writer = EmailStorePstWriter.Create(path);
        string inbox = writer.AddFolder("Inbox", EmailStoreSpecialFolderKind.Inbox);
        writer.AddItem(inbox, Document("Keep A", "keep-a@example.test", 1));
        writer.AddItem(inbox, Document("Drop", "drop@example.test", 2));
        writer.AddItem(inbox, Document("Keep B", "keep-b@example.test", 3));
        writer.Complete();
    }

    private static EmailDocument Document(string subject, string messageId, int day) => new EmailDocument {
        Subject = subject,
        MessageId = messageId,
        From = new EmailAddress("sender@example.test", "Sender"),
        Date = new DateTimeOffset(2026, 2, day, 9, 0, 0, TimeSpan.Zero),
        ReceivedDate = new DateTimeOffset(2026, 2, day, 10, 0, 0, TimeSpan.Zero),
        Body = { Text = string.Concat("Body for ", subject) }
    };

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

    private static string TemporaryDirectory() {
        string path = Path.Combine(Path.GetTempPath(),
            "officeimo-recovery-split-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static void DeleteDirectory(string path) {
        if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
    }

    private sealed class RecoveryBackend : IEmailStoreSessionBackend {
        private readonly EmailStoreItemReference[] _references;
        private readonly EmailDocument _document = Document("Recovered", "recovered@example.test", 1);
        internal RecoveryBackend() {
            _references = new[] {
                Reference("normal", orphaned: false),
                Reference("orphan-good", orphaned: true),
                Reference("orphan-bad", orphaned: true)
            };
        }
        public EmailStoreFormat Format => EmailStoreFormat.Pst;
        public string? DisplayName => "Recovery";
        public long SourceLength => 0;
        public IReadOnlyList<EmailStoreFolderInfo> Folders { get; } =
            new[] { new EmailStoreFolderInfo("inbox", null, "Inbox") };
        public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => Array.Empty<EmailStoreDiagnostic>();

        public IEnumerable<EmailStoreItemReference> EnumerateItems(
            EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
            foreach (EmailStoreItemReference reference in _references.Take(options.MaxItems)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!options.IncludeOrphanedItems && reference.IsOrphaned) continue;
                yield return reference;
            }
        }

        public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return reference.Summary!;
        }

        public EmailStoreItem ReadItem(EmailStoreItemReference reference,
            EmailStoreItemReadOptions options, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (reference.Id == "orphan-bad") throw new InvalidDataException("Synthetic corrupt item.");
            return new EmailStoreItem(reference.Id, reference.FolderId, _document,
                reference.IsAssociated, reference.IsOrphaned,
                options.Parts, EmailStoreFormat.Pst, reference.Summary);
        }

        public void Dispose() { }

        private EmailStoreItemReference Reference(string id, bool orphaned) =>
            new EmailStoreItemReference(id, "inbox", false, orphaned,
                new EmailStoreItemSummary(_document, false, false));
    }
}
