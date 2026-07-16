using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.GoogleWorkspace.Sync {
    public enum GoogleWorkspaceSyncItemKind { SourceChange = 0, RemoteChange = 1, Conflict = 2, LossyAction = 3 }

    /// <summary>One caller-owned operation classified before any mutation is attempted.</summary>
    public sealed class GoogleWorkspaceSyncItem {
        public GoogleWorkspaceSyncItem(string id, GoogleWorkspaceSyncItemKind kind, string path, string message, string? sourceId = null, string? googleFileId = null) {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A stable plan item id is required.", nameof(id));
            Id = id; Kind = kind; Path = path ?? string.Empty; Message = message ?? string.Empty; SourceId = sourceId; GoogleFileId = googleFileId;
        }
        public string Id { get; }
        public GoogleWorkspaceSyncItemKind Kind { get; }
        public string Path { get; }
        public string Message { get; }
        public string? SourceId { get; }
        public string? GoogleFileId { get; }
        public bool RequiresApproval => Kind == GoogleWorkspaceSyncItemKind.LossyAction;
    }

    /// <summary>Immutable mutation plan suitable for review, dry-run, approval, and apply.</summary>
    public sealed class GoogleWorkspaceSyncPlan {
        private GoogleWorkspaceSyncPlan(IReadOnlyList<GoogleWorkspaceSyncItem> items, TranslationReport report) { Items = items; Report = report; }
        public IReadOnlyList<GoogleWorkspaceSyncItem> Items { get; }
        public TranslationReport Report { get; }
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleWorkspaceSyncItemKind.Conflict);
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleWorkspaceSyncItemKind.LossyAction);
        public bool CanApply => !HasConflicts && !Report.HasErrors;

        public static GoogleWorkspaceSyncPlan Create(IEnumerable<GoogleWorkspaceSyncItem> items, TranslationReport? report = null) {
            if (items == null) throw new ArgumentNullException(nameof(items));
            GoogleWorkspaceSyncItem[] materialized = items.ToArray();
            string? duplicate = materialized.GroupBy(item => item.Id, StringComparer.Ordinal).Where(group => group.Count() > 1).Select(group => group.Key).FirstOrDefault();
            if (duplicate != null) throw new ArgumentException($"Synchronization plan item id '{duplicate}' is duplicated.", nameof(items));
            return new GoogleWorkspaceSyncPlan(materialized, report ?? new TranslationReport());
        }
    }
}
