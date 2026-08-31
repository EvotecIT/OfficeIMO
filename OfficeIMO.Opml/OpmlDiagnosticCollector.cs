using System;
using System.Collections.Generic;

namespace OfficeIMO.Opml;

/// <summary>Keeps repeated diagnostics bounded while retaining severity and occurrence totals.</summary>
internal sealed class OpmlDiagnosticCollector {
    private readonly int _maxDetailedPerCode;
    private readonly List<OpmlDiagnostic> _diagnostics = new List<OpmlDiagnostic>();
    private readonly Dictionary<string, int> _detailedCounts = new Dictionary<string, int>(StringComparer.Ordinal);
    private readonly Dictionary<string, SuppressedDiagnostic> _suppressed = new Dictionary<string, SuppressedDiagnostic>(StringComparer.Ordinal);
    private readonly List<string> _suppressedOrder = new List<string>();

    internal OpmlDiagnosticCollector(int maxDetailedPerCode) {
        if (maxDetailedPerCode < 1) throw new ArgumentOutOfRangeException(nameof(maxDetailedPerCode));
        _maxDetailedPerCode = maxDetailedPerCode;
    }

    internal void Add(OpmlDiagnostic diagnostic) {
        if (diagnostic == null) throw new ArgumentNullException(nameof(diagnostic));
        _detailedCounts.TryGetValue(diagnostic.Code, out int count);
        if (count < _maxDetailedPerCode) {
            _diagnostics.Add(diagnostic);
            _detailedCounts[diagnostic.Code] = count + 1;
            return;
        }

        if (!_suppressed.TryGetValue(diagnostic.Code, out SuppressedDiagnostic? summary)) {
            summary = new SuppressedDiagnostic(diagnostic.Severity);
            _suppressed.Add(diagnostic.Code, summary);
            _suppressedOrder.Add(diagnostic.Code);
        }
        summary.Count++;
        if (diagnostic.Severity > summary.Severity) summary.Severity = diagnostic.Severity;
    }

    internal IReadOnlyList<OpmlDiagnostic> ToArray() {
        var result = new List<OpmlDiagnostic>(_diagnostics.Count + _suppressedOrder.Count);
        result.AddRange(_diagnostics);
        foreach (string code in _suppressedOrder) {
            SuppressedDiagnostic summary = _suppressed[code];
            result.Add(new OpmlDiagnostic(code, summary.Severity,
                $"{summary.Count} additional '{code}' diagnostics were summarized after the per-code detail limit of {_maxDetailedPerCode} was reached."));
        }
        return result.ToArray();
    }

    private sealed class SuppressedDiagnostic {
        internal SuppressedDiagnostic(OpmlDiagnosticSeverity severity) {
            Severity = severity;
        }

        internal int Count { get; set; }
        internal OpmlDiagnosticSeverity Severity { get; set; }
    }
}
