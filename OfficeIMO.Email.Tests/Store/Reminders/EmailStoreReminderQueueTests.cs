using OfficeIMO.Email;
using System.Reflection;
using System.Threading;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreReminderQueueTests {
    [Fact]
    public void ExcludedFoldersDoNotConsumeTheEligibleItemScanBound() {
        var backend = new ExcludedFoldersFirstBackend();
        using EmailStoreSession session = CreateSession(backend);
        var options = new EmailStoreReminderQueryOptions(
            asOf: new DateTimeOffset(2026, 7, 19, 0, 0, 0, TimeSpan.Zero),
            maxItemsScanned: 1);

        EmailStoreReminderQueue queue = session.GetReminders(options);

        EmailStoreReminderQueueItem item = Assert.Single(queue.Items);
        Assert.Equal("included", item.Reference.Id);
        Assert.Equal(1, queue.ScannedItems);
        Assert.Equal(2, queue.ExcludedFolders);
        Assert.True(queue.IsComplete);
        Assert.Equal(new[] { "included" }, backend.ReadItemIds);
    }

    [Fact]
    public void ExplicitlyIncludedReminderFoldersConsumeTheSharedScanBound() {
        var backend = new ExcludedFoldersFirstBackend();
        using EmailStoreSession session = CreateSession(backend);
        var options = new EmailStoreReminderQueryOptions(
            includeExcludedFolders: true,
            asOf: new DateTimeOffset(2026, 7, 19, 0, 0, 0, TimeSpan.Zero),
            maxItemsScanned: 1);

        EmailStoreReminderQueue queue = session.GetReminders(options);

        Assert.Equal("excluded-draft", Assert.Single(queue.Items).Reference.Id);
        Assert.Equal(1, queue.ScannedItems);
        Assert.Equal(0, queue.ExcludedFolders);
        Assert.False(queue.IsComplete);
        Assert.Contains(queue.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_REMINDER_SCAN_LIMIT");
        Assert.Equal(new[] { "excluded-draft" }, backend.ReadItemIds);
    }

    private static EmailStoreSession CreateSession(IEmailStoreSessionBackend backend) {
        ConstructorInfo? constructor = typeof(EmailStoreSession).GetConstructor(
            BindingFlags.Instance | BindingFlags.NonPublic,
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

    private sealed class ExcludedFoldersFirstBackend : IEmailStoreSessionBackend {
        private readonly EmailStoreFolderInfo[] _folders = {
            new EmailStoreFolderInfo("drafts", null, "Drafts",
                specialFolderKind: EmailStoreSpecialFolderKind.Drafts),
            new EmailStoreFolderInfo("junk", null, "Junk Email",
                specialFolderKind: EmailStoreSpecialFolderKind.JunkEmail),
            new EmailStoreFolderInfo("inbox", null, "Inbox",
                specialFolderKind: EmailStoreSpecialFolderKind.Inbox)
        };
        private readonly EmailStoreItemReference[] _references = {
            new EmailStoreItemReference("excluded-draft", "drafts", false, false),
            new EmailStoreItemReference("excluded-junk", "junk", false, false),
            new EmailStoreItemReference("included", "inbox", false, false)
        };
        private readonly List<string> _readItemIds = new List<string>();

        public IReadOnlyList<string> ReadItemIds => _readItemIds;
        public EmailStoreFormat Format => EmailStoreFormat.Pst;
        public string? DisplayName => "Reminder test";
        public long SourceLength => 0;
        public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
        public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => Array.Empty<EmailStoreDiagnostic>();

        public IEnumerable<EmailStoreItemReference> EnumerateItems(
            EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
            IEnumerable<EmailStoreItemReference> selected = options.FolderId != null
                ? _references.Where(reference => string.Equals(
                    reference.FolderId, options.FolderId, StringComparison.Ordinal))
                : _references;
            foreach (EmailStoreItemReference reference in selected.Take(options.MaxItems)) {
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
            _readItemIds.Add(reference.Id);
            var signal = new DateTimeOffset(2026, 7, 20, 9, 0, 0, TimeSpan.Zero);
            var document = new EmailDocument { Subject = reference.Id + " reminder" };
            document.MessageMetadata.Reminder.IsSet = true;
            document.MessageMetadata.Reminder.SignalTime = signal;
            document.Mapi.Set(MapiKnownProperties.PidLid.ReminderSet, true);
            document.Mapi.Set(MapiKnownProperties.PidLid.ReminderSignalTime, signal);
            return new EmailStoreItem(reference.Id, reference.FolderId, document,
                loadedParts: options.Parts, format: Format);
        }

        public void Dispose() { }
    }
}
