using OfficeIMO.Email;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreTableQueryTests {
    [Fact]
    public void Typed_identifiers_reject_empty_values_and_compare_ordinally() {
        Assert.Throws<ArgumentException>(() => new EmailStoreFolderId(" "));
        Assert.Throws<ArgumentException>(() => new EmailStoreItemId(string.Empty));
        Assert.True(default(EmailStoreFolderId).IsEmpty);
        Assert.Throws<InvalidOperationException>(() => default(EmailStoreFolderId).Value);

        Assert.Equal(new EmailStoreFolderId("Inbox"), EmailStoreFolderId.Parse("Inbox"));
        Assert.NotEqual(new EmailStoreFolderId("Inbox"), new EmailStoreFolderId("inbox"));
        Assert.True(EmailStoreItemId.TryParse("42", out EmailStoreItemId itemId));
        Assert.Equal("42", itemId.Value);
    }

    [Fact]
    public void Folder_catalog_navigates_paths_descendants_and_special_folders() {
        using EmailStoreSession session = CreateSession(new TableBackend());
        EmailStoreFolderCatalog catalog = session.FolderCatalog;
        EmailStoreFolderInfo root = catalog.Get(new EmailStoreFolderId("root"));
        EmailStoreFolderInfo inbox = catalog.Get(new EmailStoreFolderId("inbox"));

        Assert.Equal(new[] { "root" }, catalog.Roots.Select(folder => folder.Id));
        Assert.Equal(new[] { "inbox", "archive" }, catalog.GetChildren(root.Key).Select(folder => folder.Id));
        Assert.Equal(new[] { "root", "inbox", "projects" },
            catalog.GetPath(new EmailStoreFolderId("projects")).Select(folder => folder.Id));
        Assert.Equal(new[] { "inbox", "archive", "projects" },
            catalog.GetDescendants(root.Key).Select(folder => folder.Id));
        Assert.Same(inbox, catalog.FindSpecialFolder(EmailStoreSpecialFolderKind.Inbox));
        Assert.Empty(catalog.FindSpecialFolders(EmailStoreSpecialFolderKind.Drafts));
    }

    [Fact]
    public void Table_query_composes_typed_filter_projection_sort_and_keyset_pages() {
        var backend = new TableBackend();
        using EmailStoreSession session = CreateSession(backend);
        EmailStoreFilter filter = EmailStoreFields.Subject.Contains("alpha") &
            EmailStoreFields.IsRead.EqualTo(true);
        var query = new EmailStoreTableQuery(
            filter: filter,
            sorts: new[] {
                EmailStoreFields.ReceivedAt.Descending(),
                EmailStoreFields.Subject.Ascending()
            },
            projection: new EmailStoreProjection(EmailStoreFields.ItemId, EmailStoreFields.Subject),
            pageSize: 2,
            maxItemsScanned: 20);

        EmailStoreQueryPlan plan = query.Explain();
        EmailStoreTablePage first = session.SearchPage(query);

        Assert.Equal(EmailStoreFilterKind.And, plan.Filter.Kind);
        Assert.Equal(new[] { "item.subject", "item.isRead" },
            plan.Filter.Children.Select(child => child.Field!.Key));
        Assert.False(plan.ReadsItemPayloads);
        Assert.True(plan.MaterializesMatchesForSort);
        Assert.Equal(new[] { "item.receivedAt", "item.subject", "folder.id", "item.id" },
            plan.EffectiveSorts.Select(sort => sort.Field.Key));
        Assert.Equal(new[] { "alpha newest", "Alpha middle" },
            first.Rows.Select(row => row.Get(EmailStoreFields.Subject)));
        Assert.All(first.Rows, row => Assert.Equal(2, row.Values.Count));
        Assert.Throws<KeyNotFoundException>(() => first.Rows[0][EmailStoreFields.SentAt]);
        Assert.NotNull(first.NextToken);
        Assert.True(EmailStoreContinuationToken.TryParse(first.NextToken!.Value,
            out EmailStoreContinuationToken? parsed));

        backend.InsertNewerMatchingItem();
        EmailStoreTablePage second = session.SearchPage(query.ContinueFrom(parsed));

        Assert.Equal(new[] { "alpha oldest" }, second.Rows.Select(row => row.Get(EmailStoreFields.Subject)));
        Assert.Null(second.NextToken);
        Assert.DoesNotContain(second.Rows, row => row.Reference.Id == "inserted-before-token");
        Assert.Equal(3, first.MatchesInScan);
        Assert.Equal(4, second.MatchesInScan);
    }

    [Fact]
    public void Continuation_is_bound_to_scope_filter_order_and_scan_bound() {
        using EmailStoreSession session = CreateSession(new TableBackend());
        var query = new EmailStoreTableQuery(
            filter: EmailStoreFields.Subject.Contains("alpha"),
            sorts: new[] { EmailStoreFields.Subject.Ascending() },
            pageSize: 1,
            maxItemsScanned: 10);
        EmailStoreContinuationToken token = Assert.IsType<EmailStoreContinuationToken>(session.SearchPage(query).NextToken);

        var differentFilter = new EmailStoreTableQuery(
            filter: EmailStoreFields.Subject.Contains("beta"),
            sorts: new[] { EmailStoreFields.Subject.Ascending() },
            continuationToken: token,
            pageSize: 1,
            maxItemsScanned: 10);

        Assert.Throws<ArgumentException>(() => session.SearchPage(differentFilter));
        Assert.False(EmailStoreContinuationToken.TryParse("not-a-token", out _));
    }

    [Fact]
    public void Table_query_reports_a_reached_scan_bound_without_claiming_a_next_page() {
        using EmailStoreSession session = CreateSession(new TableBackend());
        var query = new EmailStoreTableQuery(
            filter: EmailStoreFilter.All,
            sorts: new[] { EmailStoreFields.ItemId.Ascending() },
            maxItemsScanned: 2,
            pageSize: 10);

        EmailStoreTablePage page = session.SearchPage(query);

        Assert.Equal(2, page.ItemsScanned);
        Assert.Equal(2, page.Rows.Count);
        Assert.True(page.ScanLimitReached);
        Assert.Null(page.NextToken);
    }

    [Fact]
    public async Task Dependency_free_async_sequence_supports_await_foreach_and_cancellation() {
        using EmailStoreSession session = CreateSession(new TableBackend());
        var ids = new List<string>();

        await foreach (EmailStoreItemReference reference in session.EnumerateItemsAsync()) {
            ids.Add(reference.Id);
        }

        Assert.Equal(6, ids.Count);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => {
            await foreach (EmailStoreItemReference ignored in session.EnumerateItemsAsync(
                cancellationToken: cancellation.Token)) {
            }
        });
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

    private sealed class TableBackend : IEmailStoreSessionBackend {
        private readonly List<EmailStoreItemReference> _references;
        private readonly EmailStoreFolderInfo[] _folders = {
            new EmailStoreFolderInfo("root", null, "Root"),
            new EmailStoreFolderInfo("inbox", "root", "Inbox",
                specialFolderKind: EmailStoreSpecialFolderKind.Inbox,
                classificationSource: EmailStoreFolderClassificationSource.SourceIdentifier),
            new EmailStoreFolderInfo("archive", "root", "Archive"),
            new EmailStoreFolderInfo("projects", "inbox", "Projects")
        };

        internal TableBackend() {
            _references = new List<EmailStoreItemReference> {
                Reference("a", "inbox", "alpha oldest", 1, true),
                Reference("b", "inbox", "Beta", 6, false),
                Reference("c", "inbox", "Alpha middle", 3, true),
                Reference("d", "archive", "Gamma", 2, true),
                Reference("e", "projects", "alpha newest", 5, true),
                Reference("f", "archive", "Delta", 4, false)
            };
        }

        public EmailStoreFormat Format => EmailStoreFormat.Pst;
        public string? DisplayName => "Table test";
        public long SourceLength => 0;
        public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
        public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => Array.Empty<EmailStoreDiagnostic>();

        internal void InsertNewerMatchingItem() => _references.Insert(0,
            Reference("inserted-before-token", "inbox", "alpha inserted", 10, true));

        public IEnumerable<EmailStoreItemReference> EnumerateItems(
            EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
            IEnumerable<EmailStoreItemReference> selected = _references;
            if (options.FolderId != null) {
                var ids = new HashSet<string>(StringComparer.Ordinal) { options.FolderId };
                if (options.IncludeDescendants) {
                    bool changed;
                    do {
                        changed = false;
                        foreach (EmailStoreFolderInfo folder in _folders) {
                            if (folder.ParentId != null && ids.Contains(folder.ParentId) && ids.Add(folder.Id)) changed = true;
                        }
                    } while (changed);
                }
                selected = selected.Where(reference => ids.Contains(reference.FolderId));
            }
            foreach (EmailStoreItemReference reference in selected.Take(options.MaxItems)) {
                cancellationToken.ThrowIfCancellationRequested();
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
            throw new InvalidOperationException("A table query must not read complete items.");
        }

        public void Dispose() { }

        private static EmailStoreItemReference Reference(string id, string folderId,
            string subject, int day, bool isRead) {
            var document = new EmailDocument {
                MessageClass = "IPM.Note",
                Subject = subject,
                From = new EmailAddress(string.Concat(id, "@example.test"), id.ToUpperInvariant()),
                Date = new DateTimeOffset(2026, 1, day, 9, 0, 0, TimeSpan.Zero),
                ReceivedDate = new DateTimeOffset(2026, 1, day, 10, 0, 0, TimeSpan.Zero)
            };
            var summary = new EmailStoreItemSummary(document, hasAttachments: false, isRead: isRead);
            return new EmailStoreItemReference(id, folderId, false, false, summary);
        }
    }
}
