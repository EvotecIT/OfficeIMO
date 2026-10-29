namespace OfficeIMO;

/// <summary>Thread-safe bounded collection for CLI, PowerShell, desktop and service batch observers.</summary>
public sealed class OfficeBatchReportCollector : IProgress<OfficeBatchItemReport> {
    /// <summary>Maximum retained item observations.</summary>
    public const int MaximumItems = 10_000;
    /// <summary>Maximum combined UTF-8 size of retained item observations.</summary>
    public const int MaximumItemBytes = 16 * 1024 * 1024;
    private readonly object _gate = new();
    private readonly List<OfficeBatchItemReport> _items = [];
    private long _reported;
    private int _retainedBytes;
    private long _selected, _completed, _reused, _failed, _skipped;

    /// <summary>Retains a bounded observation; oversized items are omitted without interrupting conversion.</summary>
    public void Report(OfficeBatchItemReport item) {
        if (item == null) throw new ArgumentNullException(nameof(item));
        lock (_gate) {
            _reported++;
            if (item.Skipped) {
                _skipped++;
                if (item.SkippedReason is OfficeBatchOutputSelection.OutputExists or OfficeBatchOutputSelection.UpToDate) _selected++;
            }
            else {
                _selected++;
                if (item.Status == OfficeBatchItemStatus.Completed) { _completed++; if (item.Reused) _reused++; }
                else if (item.Status != OfficeBatchItemStatus.Cancelled) _failed++;
            }
            if (_items.Count == MaximumItems) return;
            int remaining = MaximumItemBytes - _retainedBytes;
            if (OfficeBatchReportSerializer.TryGetItemSize(item, remaining, out int bytes)) {
                _items.Add(item);
                _retainedBytes += bytes;
            }
        }
    }

    /// <summary>Creates an independent report snapshot while preserving aggregate counts even when observations were omitted.</summary>
    public OfficeBatchReport Complete(OfficeBatchSummary summary, bool allowEmpty = false) {
        if (summary == null) throw new ArgumentNullException(nameof(summary));
        lock (_gate) return new OfficeBatchReport {
            Summary = summary, Items = Array.AsReadOnly(_items.ToArray()), ItemsTruncated = _reported > _items.Count,
            Succeeded = !summary.Cancelled && summary.FailureSummary == null && summary.Failed == 0 && (allowEmpty || summary.Selected > 0)
        };
    }

    internal OfficeBatchSummary ObservedSummary(bool cancelled, string? failureSummary = null) {
        lock (_gate) return new(_selected, _completed, _reused, _failed, cancelled, _skipped) { FailureSummary = failureSummary };
    }
}
