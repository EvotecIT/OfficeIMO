using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.GoogleWorkspace.Sync {
    /// <summary>Classifies an item in a caller-built synchronization plan.</summary>
    public enum GoogleWorkspaceSyncItemKind {
        /// <summary>A change originating in the caller's source.</summary>
        SourceChange = 0,
        /// <summary>A change observed in the remote copy.</summary>
        RemoteChange = 1,
        /// <summary>A change the caller has classified as a conflict; the executor will not apply it.</summary>
        Conflict = 2,
        /// <summary>An action that needs explicit item approval and an accepting data-loss policy.</summary>
        LossyAction = 3,
    }

    /// <summary>One caller-owned operation classified before any mutation is attempted.</summary>
    public sealed class GoogleWorkspaceSyncItem {
        /// <summary>Creates an item with a stable identifier, target, and expected revision decision.</summary>
        /// <remarks>The caller supplies the classification and descriptive fields; this constructor does not inspect either copy.</remarks>
        public GoogleWorkspaceSyncItem(string id, GoogleWorkspaceSyncItemKind kind, string path, string message,
            string targetResource, string expectedRevision, string? sourceId = null, string? googleFileId = null) {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A stable plan item id is required.", nameof(id));
            if (string.IsNullOrWhiteSpace(targetResource)) throw new ArgumentException("The target resource is required.", nameof(targetResource));
            if (string.IsNullOrWhiteSpace(expectedRevision)) throw new ArgumentException("The expected revision decision is required.", nameof(expectedRevision));
            Id = id; Kind = kind; Path = path ?? string.Empty; Message = message ?? string.Empty;
            TargetResource = targetResource; ExpectedRevision = expectedRevision;
            SourceId = sourceId; GoogleFileId = googleFileId;
        }
        /// <summary>Gets the caller-assigned identifier used for duplicate detection and lossy-action approval.</summary>
        public string Id { get; }
        /// <summary>Gets the caller-assigned change classification.</summary>
        public GoogleWorkspaceSyncItemKind Kind { get; }
        /// <summary>Gets the caller-assigned location of the change, or an empty string when null was supplied.</summary>
        public string Path { get; }
        /// <summary>Gets the caller-assigned description, or an empty string when null was supplied.</summary>
        public string Message { get; }
        /// <summary>Gets an optional identifier for the local source item.</summary>
        public string? SourceId { get; }
        /// <summary>Gets an optional Google Drive file identifier.</summary>
        public string? GoogleFileId { get; }
        /// <summary>Gets the required resource target carried into each decision receipt's policy.</summary>
        public string TargetResource { get; }
        /// <summary>Gets the required expected-revision decision carried into each decision receipt's policy.</summary>
        /// <remarks>The executor passes this value through; the mutation callback must enforce any revision precondition.</remarks>
        public string ExpectedRevision { get; }
        /// <summary>Gets whether this item is classified as a lossy action requiring explicit approval.</summary>
        public bool RequiresApproval => Kind == GoogleWorkspaceSyncItemKind.LossyAction;
    }

    /// <summary>Immutable mutation plan suitable for review, dry-run, approval, and apply.</summary>
    public sealed class GoogleWorkspaceSyncPlan {
        private GoogleWorkspaceSyncPlan(IReadOnlyList<GoogleWorkspaceSyncItem> items,
            GoogleWorkspaceOperationPolicy policy, TranslationReport report) { Items = items; Policy = policy; Report = report; }
        /// <summary>Gets a read-only snapshot of the supplied item sequence.</summary>
        public IReadOnlyList<GoogleWorkspaceSyncItem> Items { get; }
        /// <summary>Gets the caller-supplied operation policy.</summary>
        public GoogleWorkspaceOperationPolicy Policy { get; }
        /// <summary>Gets a read-only snapshot of the supplied translation report.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets whether any item was classified as a conflict.</summary>
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleWorkspaceSyncItemKind.Conflict);
        /// <summary>Gets whether any item was classified as a lossy action.</summary>
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleWorkspaceSyncItemKind.LossyAction);
        /// <summary>Gets whether the plan has neither classified conflicts nor report errors.</summary>
        /// <remarks>This is advisory; <see cref="GoogleWorkspaceSyncExecutor.ApplyAsync"/> still processes individual items and approval rules.</remarks>
        public bool CanApply => !HasConflicts && !Report.HasErrors;

        /// <summary>Snapshots the item sequence and report, rejecting duplicate item identifiers.</summary>
        /// <remarks>The operation policy is retained by reference. Item objects are immutable after construction.</remarks>
        public static GoogleWorkspaceSyncPlan Create(IEnumerable<GoogleWorkspaceSyncItem> items,
            GoogleWorkspaceOperationPolicy policy, TranslationReport? report = null) {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (policy == null) throw new ArgumentNullException(nameof(policy));
            GoogleWorkspaceSyncItem[] materialized = items.ToArray();
            string? duplicate = materialized.GroupBy(item => item.Id, StringComparer.Ordinal).Where(group => group.Count() > 1).Select(group => group.Key).FirstOrDefault();
            if (duplicate != null) throw new ArgumentException($"Synchronization plan item id '{duplicate}' is duplicated.", nameof(items));
            return new GoogleWorkspaceSyncPlan(Array.AsReadOnly(materialized), policy,
                (report ?? new TranslationReport()).CreateReadOnlySnapshot());
        }
    }
}
