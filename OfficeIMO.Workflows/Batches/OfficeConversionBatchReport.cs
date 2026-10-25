using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

/// <summary>A bounded observation of a batch attempt. It does not authorize checkpoint reuse or output replacement.</summary>
public sealed record OfficeConversionBatchReport {
    /// <summary>The versioned batch report contract.</summary>
    public const string CurrentSchema = "officeimo.conversion.batch.v1";
    /// <summary>Versioned contract identifier.</summary>
    [JsonRequired]
    public string Schema { get; init; } = CurrentSchema;
    /// <summary>Identifies a complete report rather than an individual event.</summary>
    [JsonRequired]
    public string Type { get; init; } = "report";
    /// <summary>Whether the batch satisfies the caller's empty-selection policy without failures or cancellation.</summary>
    public bool Succeeded { get; init; }
    /// <summary>The full aggregate outcome, independent of retained item observations.</summary>
    public required OfficeConversionBatchResult Summary { get; init; }
    /// <summary>Bounded item observations in callback order.</summary>
    public IReadOnlyList<OfficeConversionBatchItemResult> Items { get; init; } = [];
    /// <summary>Whether some observations were omitted to enforce the report limits.</summary>
    public bool ItemsTruncated { get; init; }
}

/// <summary>Thread-safe bounded collection for CLI, PowerShell, desktop and service batch observers.</summary>
public sealed class OfficeConversionBatchReportCollector : IProgress<OfficeConversionBatchItemResult> {
    /// <summary>Maximum retained item observations.</summary>
    public const int MaximumItems = 10_000;
    /// <summary>Maximum combined UTF-8 size of retained item observations.</summary>
    public const int MaximumItemBytes = 16 * 1024 * 1024;
    private readonly object _gate = new();
    private readonly List<OfficeConversionBatchItemResult> _items = [];
    private long _reported;
    private int _retainedBytes;
    private long _selected, _completed, _reused, _failed, _skipped;

    /// <summary>Retains a bounded observation; oversized items are omitted without interrupting conversion.</summary>
    public void Report(OfficeConversionBatchItemResult item) {
        ArgumentNullException.ThrowIfNull(item);
        lock (_gate) {
            _reported++;
            if (item.Skipped) _skipped++;
            else {
                _selected++;
                if (item.Status == OfficeWorkflowStatus.Completed) { _completed++; if (item.Reused) _reused++; }
                else if (item.Status != OfficeWorkflowStatus.Cancelled) _failed++;
            }
            if (_items.Count == MaximumItems) return;
            int remaining = MaximumItemBytes - _retainedBytes;
            if (OfficeConversionBatchReportSerializer.TryGetItemSize(item, remaining, out int bytes)) {
                _items.Add(item);
                _retainedBytes += bytes;
            }
        }
    }

    /// <summary>Creates an independent report snapshot while preserving aggregate counts even when observations were omitted.</summary>
    public OfficeConversionBatchReport Complete(OfficeConversionBatchResult summary, bool allowEmpty = false) {
        ArgumentNullException.ThrowIfNull(summary);
        lock (_gate) return new OfficeConversionBatchReport {
            Summary = summary, Items = Array.AsReadOnly(_items.ToArray()), ItemsTruncated = _reported > _items.Count,
            Succeeded = !summary.Cancelled && summary.FailureSummary == null && summary.Failed == 0 && (allowEmpty || summary.Selected > 0)
        };
    }

    internal OfficeConversionBatchResult ObservedSummary(bool cancelled, string? failureSummary = null) {
        lock (_gate) return new(_selected, _completed, _reused, _failed, cancelled, _skipped) { FailureSummary = failureSummary };
    }
}
