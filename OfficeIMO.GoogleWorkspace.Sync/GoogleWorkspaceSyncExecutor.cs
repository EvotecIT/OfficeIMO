namespace OfficeIMO.GoogleWorkspace.Sync {
    /// <summary>Outcome assigned to one item by the synchronization executor.</summary>
    public enum GoogleWorkspaceSyncApplyStatus {
        /// <summary>Eligible in a dry run; the operation was not called.</summary>
        Planned = 0,
        /// <summary>The caller's operation completed successfully.</summary>
        Applied = 1,
        /// <summary>Not attempted after an earlier failure stopped execution.</summary>
        Skipped = 2,
        /// <summary>Classified as a conflict; the operation was not called.</summary>
        Conflict = 3,
        /// <summary>Blocked because a lossy action lacked an accepting policy or item approval.</summary>
        ApprovalRequired = 4,
        /// <summary>The caller's operation threw an exception other than cancellation requested by the supplied token.</summary>
        Failed = 5,
        /// <summary>Not completed because execution was canceled.</summary>
        Canceled = 6,
    }

    /// <summary>Caller-owned operation that applies a single approved plan item.</summary>
    public delegate Task GoogleWorkspaceSyncOperation(GoogleWorkspaceSyncItem item, CancellationToken cancellationToken);

    /// <summary>Controls dry-run, failure, cancellation, and lossy-action approval behavior.</summary>
    public sealed class GoogleWorkspaceSyncApplyOptions {
        /// <summary>Gets or sets whether eligible items are reported as planned without invoking the operation; defaults to true.</summary>
        public bool DryRun { get; set; } = true;
        /// <summary>Gets or sets whether later items are attempted after an operation fails; defaults to true.</summary>
        public bool ContinueOnError { get; set; } = true;
        /// <summary>Gets or sets whether cancellation returns per-item outcomes instead of throwing; defaults to true.</summary>
        public bool ReturnPartialResultOnCancellation { get; set; } = true;
        /// <summary>Gets the mutable list of lossy item identifiers approved for this apply call.</summary>
        /// <remarks>Approval also requires <see cref="GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss"/> in the plan policy.</remarks>
        public IList<string> ApprovedLossyItemIds { get; } = new List<string>();
    }

    /// <summary>Outcome and decision evidence for one synchronization item.</summary>
    public sealed class GoogleWorkspaceSyncItemResult {
        internal GoogleWorkspaceSyncItemResult(GoogleWorkspaceSyncItem item, GoogleWorkspaceSyncApplyStatus status,
            GoogleWorkspaceSyncDecisionReceipt decisionReceipt, Exception? exception = null) { Item = item; Status = status; DecisionReceipt = decisionReceipt; Exception = exception; }
        /// <summary>Gets the plan item to which this outcome belongs.</summary>
        public GoogleWorkspaceSyncItem Item { get; }
        /// <summary>Gets the executor's outcome for the item.</summary>
        public GoogleWorkspaceSyncApplyStatus Status { get; }
        /// <summary>Gets the operation exception for a failed item, or null otherwise.</summary>
        public Exception? Exception { get; }
        /// <summary>Evidence of the sync executor's plan/apply decision. Actual HTTP mutation receipts come from the session receipt sink.</summary>
        public GoogleWorkspaceSyncDecisionReceipt DecisionReceipt { get; }
    }

    /// <summary>Caller-observable evidence for one synchronization decision; this is not an HTTP mutation receipt.</summary>
    public sealed class GoogleWorkspaceSyncDecisionReceipt {
        internal GoogleWorkspaceSyncDecisionReceipt(GoogleWorkspaceOperationPolicy policy, string target,
            GoogleWorkspaceSyncApplyStatus status) {
            Policy = policy; Target = target; Status = status; CompletedAt = DateTimeOffset.UtcNow;
        }
        /// <summary>Gets the item-specific policy carrying its target and expected revision.</summary>
        public GoogleWorkspaceOperationPolicy Policy { get; }
        /// <summary>Gets the item's target resource.</summary>
        public string Target { get; }
        /// <summary>Gets the executor's decision status.</summary>
        public GoogleWorkspaceSyncApplyStatus Status { get; }
        /// <summary>Gets when the decision receipt was created in UTC.</summary>
        public DateTimeOffset CompletedAt { get; }
    }

    /// <summary>Read-only per-item outcomes from one apply or dry-run call.</summary>
    public sealed class GoogleWorkspaceSyncApplyResult {
        internal GoogleWorkspaceSyncApplyResult(IReadOnlyList<GoogleWorkspaceSyncItemResult> items, bool wasCanceled) {
            Items = Array.AsReadOnly(items.ToArray()); WasCanceled = wasCanceled;
        }
        /// <summary>Gets an immutable sequence of outcomes in plan order.</summary>
        public IReadOnlyList<GoogleWorkspaceSyncItemResult> Items { get; }
        /// <summary>Gets whether cancellation stopped the call and returned outcomes.</summary>
        public bool WasCanceled { get; }
        /// <summary>Gets whether any operation exception was recorded as a failure.</summary>
        public bool HasFailures => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Failed);
        /// <summary>Gets whether any plan item was blocked as a conflict.</summary>
        public bool HasConflicts => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Conflict);
        /// <summary>Gets whether any lossy item was blocked for missing approval or policy.</summary>
        public bool NeedsApproval => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.ApprovalRequired);
        /// <summary>Gets whether at least one item applied while another was canceled, failed, conflicted, blocked, or skipped.</summary>
        public bool IsPartial => Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Applied)
            && (WasCanceled || HasFailures || HasConflicts || NeedsApproval || Items.Any(item => item.Status == GoogleWorkspaceSyncApplyStatus.Skipped));
    }

    /// <summary>Evaluates each plan item and invokes a caller-owned operation for eligible items when dry-run is disabled.</summary>
    public static class GoogleWorkspaceSyncExecutor {
        /// <summary>Returns an outcome for every item unless configured to throw on cancellation or a precondition fails.</summary>
        /// <remarks>Conflicts never invoke the operation. Lossy items require both approval by ID and an accepting plan policy. The executor does not enforce the target's expected revision inside the caller's operation.</remarks>
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
                    results.Add(Result(plan, item, GoogleWorkspaceSyncApplyStatus.Conflict));
                    continue;
                }
                if (item.RequiresApproval && (!approved.Contains(item.Id) ||
                    plan.Policy.DataLossDecision != GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss)) {
                    results.Add(Result(plan, item, GoogleWorkspaceSyncApplyStatus.ApprovalRequired));
                    continue;
                }
                if (options.DryRun) {
                    results.Add(Result(plan, item, GoogleWorkspaceSyncApplyStatus.Planned));
                    continue;
                }
                try {
                    await operation(item, cancellationToken).ConfigureAwait(false);
                    results.Add(Result(plan, item, GoogleWorkspaceSyncApplyStatus.Applied));
                } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                    results.Add(Result(plan, item, GoogleWorkspaceSyncApplyStatus.Canceled));
                    for (int remaining = index + 1; remaining < plan.Items.Count; remaining++) results.Add(Result(plan, plan.Items[remaining], GoogleWorkspaceSyncApplyStatus.Canceled));
                    if (!options.ReturnPartialResultOnCancellation) cancellationToken.ThrowIfCancellationRequested();
                    return new GoogleWorkspaceSyncApplyResult(results, true);
                } catch (Exception exception) {
                    results.Add(Result(plan, item, GoogleWorkspaceSyncApplyStatus.Failed, exception));
                    if (!options.ContinueOnError) {
                        for (int remaining = index + 1; remaining < plan.Items.Count; remaining++) results.Add(Result(plan, plan.Items[remaining], GoogleWorkspaceSyncApplyStatus.Skipped));
                        return new GoogleWorkspaceSyncApplyResult(results, false);
                    }
                }
            }
            return new GoogleWorkspaceSyncApplyResult(results, false);
        }

        private static GoogleWorkspaceSyncApplyResult Cancel(GoogleWorkspaceSyncPlan plan, List<GoogleWorkspaceSyncItemResult> results, int index, GoogleWorkspaceSyncApplyOptions options, CancellationToken token) {
            for (int remaining = index; remaining < plan.Items.Count; remaining++) results.Add(Result(plan, plan.Items[remaining], GoogleWorkspaceSyncApplyStatus.Canceled));
            if (!options.ReturnPartialResultOnCancellation) token.ThrowIfCancellationRequested();
            return new GoogleWorkspaceSyncApplyResult(results, true);
        }

        private static GoogleWorkspaceSyncItemResult Result(GoogleWorkspaceSyncPlan plan,
            GoogleWorkspaceSyncItem item, GoogleWorkspaceSyncApplyStatus status, Exception? exception = null) {
            var itemPolicy = new GoogleWorkspaceOperationPolicy(plan.Policy.Account, plan.Policy.Scopes,
                item.TargetResource, item.ExpectedRevision, plan.Policy.MaxRetryCount,
                plan.Policy.MaxRetryElapsedTime, plan.Policy.RateLimitPolicy,
                plan.Policy.DataLossDecision, plan.Policy.AcceptedLoss);
            var receipt = new GoogleWorkspaceSyncDecisionReceipt(itemPolicy, item.TargetResource, status);
            return new GoogleWorkspaceSyncItemResult(item, status, receipt, exception);
        }
    }
}
