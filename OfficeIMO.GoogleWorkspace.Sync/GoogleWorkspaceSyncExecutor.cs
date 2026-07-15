namespace OfficeIMO.GoogleWorkspace.Sync {
    public enum GoogleWorkspaceSyncApplyStatus { Planned = 0, Applied = 1, Skipped = 2, Conflict = 3, ApprovalRequired = 4, Failed = 5, Canceled = 6 }

    public delegate Task GoogleWorkspaceSyncOperation(GoogleWorkspaceSyncItem item, CancellationToken cancellationToken);

    public sealed class GoogleWorkspaceSyncApplyOptions {
        public bool DryRun { get; set; } = true;
        public bool ContinueOnError { get; set; } = true;
        public bool ReturnPartialResultOnCancellation { get; set; } = true;
        public IList<string> ApprovedLossyItemIds { get; } = new List<string>();
    }

    public sealed class GoogleWorkspaceSyncItemResult {
        internal GoogleWorkspaceSyncItemResult(GoogleWorkspaceSyncItem item, GoogleWorkspaceSyncApplyStatus status, Exception? exception = null) { Item = item; Status = status; Exception = exception; }
        public GoogleWorkspaceSyncItem Item { get; }
        public GoogleWorkspaceSyncApplyStatus Status { get; }
        public Exception? Exception { get; }
    }

    public sealed class GoogleWorkspaceSyncApplyResult {
        internal GoogleWorkspaceSyncApplyResult(IReadOnlyList<GoogleWorkspaceSyncItemResult> items, bool wasCanceled) { Items = items; WasCanceled = wasCanceled; }
        public IReadOnlyList<GoogleWorkspaceSyncItemResult> Items { get; }
        public bool WasCanceled { get; }
        public bool HasFailures => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Failed);
        public bool HasConflicts => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Conflict);
        public bool NeedsApproval => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.ApprovalRequired);
        public bool IsPartial => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Applied) && (WasCanceled || HasFailures || Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Skipped));
    }

    /// <summary>Applies only reviewed plan items and returns an outcome for every item, including cancellation and partial failure.</summary>
    public static class GoogleWorkspaceSyncExecutor {
        public static async Task<GoogleWorkspaceSyncApplyResult> ApplyAsync(GoogleWorkspaceSyncPlan plan, GoogleWorkspaceSyncOperation operation, GoogleWorkspaceSyncApplyOptions? options = null, CancellationToken cancellationToken = default) {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (operation == null) throw new ArgumentNullException(nameof(operation));
            options ??= new GoogleWorkspaceSyncApplyOptions();
            var approved = new HashSet<string>(options.ApprovedLossyItemIds, StringComparer.Ordinal);
            var results = new List<GoogleWorkspaceSyncItemResult>(plan.Items.Count);

            for (int index = 0; index < plan.Items.Count; index++) {
                GoogleWorkspaceSyncItem item = plan.Items[index];
                if (cancellationToken.IsCancellationRequested) return Cancel(plan, results, index, options, cancellationToken);
                if (item.Kind == GoogleWorkspaceSyncItemKind.Conflict) {
                    results.Add(new GoogleWorkspaceSyncItemResult(item, GoogleWorkspaceSyncApplyStatus.Conflict));
                    continue;
                }
                if (item.RequiresApproval && !approved.Contains(item.Id)) {
                    results.Add(new GoogleWorkspaceSyncItemResult(item, GoogleWorkspaceSyncApplyStatus.ApprovalRequired));
                    continue;
                }
                if (options.DryRun) {
                    results.Add(new GoogleWorkspaceSyncItemResult(item, GoogleWorkspaceSyncApplyStatus.Planned));
                    continue;
                }
                try {
                    await operation(item, cancellationToken).ConfigureAwait(false);
                    results.Add(new GoogleWorkspaceSyncItemResult(item, GoogleWorkspaceSyncApplyStatus.Applied));
                } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                    results.Add(new GoogleWorkspaceSyncItemResult(item, GoogleWorkspaceSyncApplyStatus.Canceled));
                    for (int remaining = index + 1; remaining < plan.Items.Count; remaining++) results.Add(new GoogleWorkspaceSyncItemResult(plan.Items[remaining], GoogleWorkspaceSyncApplyStatus.Canceled));
                    if (!options.ReturnPartialResultOnCancellation) cancellationToken.ThrowIfCancellationRequested();
                    return new GoogleWorkspaceSyncApplyResult(results, true);
                } catch (Exception exception) {
                    results.Add(new GoogleWorkspaceSyncItemResult(item, GoogleWorkspaceSyncApplyStatus.Failed, exception));
                    if (!options.ContinueOnError) {
                        for (int remaining = index + 1; remaining < plan.Items.Count; remaining++) results.Add(new GoogleWorkspaceSyncItemResult(plan.Items[remaining], GoogleWorkspaceSyncApplyStatus.Skipped));
                        return new GoogleWorkspaceSyncApplyResult(results, false);
                    }
                }
            }
            return new GoogleWorkspaceSyncApplyResult(results, false);
        }

        private static GoogleWorkspaceSyncApplyResult Cancel(GoogleWorkspaceSyncPlan plan, List<GoogleWorkspaceSyncItemResult> results, int index, GoogleWorkspaceSyncApplyOptions options, CancellationToken token) {
            for (int remaining = index; remaining < plan.Items.Count; remaining++) results.Add(new GoogleWorkspaceSyncItemResult(plan.Items[remaining], GoogleWorkspaceSyncApplyStatus.Canceled));
            if (!options.ReturnPartialResultOnCancellation) token.ThrowIfCancellationRequested();
            return new GoogleWorkspaceSyncApplyResult(results, true);
        }
    }
}
