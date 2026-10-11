using System;
using System.Collections.Generic;

namespace OfficeIMO {

    /// <summary>Bounds detailed diagnostics per code and summarizes suppressed severity and loss.</summary>
    internal abstract class OfficeDiagnosticCollector<TDiagnostic> where TDiagnostic : class {
        private readonly int _maxDetailedPerCode;
        private readonly List<TDiagnostic> _diagnostics = new List<TDiagnostic>();
        private readonly Dictionary<string, int> _detailedCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly Dictionary<string, SuppressedDiagnostic> _suppressed = new Dictionary<string, SuppressedDiagnostic>(StringComparer.Ordinal);
        private readonly List<string> _suppressedOrder = new List<string>();

        protected OfficeDiagnosticCollector(int maxDetailedPerCode) {
            if (maxDetailedPerCode < 1) throw new ArgumentOutOfRangeException(nameof(maxDetailedPerCode));
            _maxDetailedPerCode = maxDetailedPerCode;
        }

        protected abstract string CodeOf(TDiagnostic diagnostic);
        protected abstract int SeverityOf(TDiagnostic diagnostic);
        protected abstract OfficeConversionLossKind LossKindOf(TDiagnostic diagnostic);
        protected abstract TDiagnostic CreateSummary(string code, int severity, string message, OfficeConversionLossKind lossKind);

        protected void Collect(TDiagnostic diagnostic) {
            if (diagnostic == null) throw new ArgumentNullException(nameof(diagnostic));
            string code = CodeOf(diagnostic);
            _detailedCounts.TryGetValue(code, out int count);
            if (count < _maxDetailedPerCode) {
                _diagnostics.Add(diagnostic);
                _detailedCounts[code] = count + 1;
                return;
            }

            if (!_suppressed.TryGetValue(code, out SuppressedDiagnostic? summary)) {
                summary = new SuppressedDiagnostic(SeverityOf(diagnostic), LossKindOf(diagnostic));
                _suppressed.Add(code, summary);
                _suppressedOrder.Add(code);
            }
            summary.Count++;
            summary.Severity = Math.Max(summary.Severity, SeverityOf(diagnostic));
            OfficeConversionLossKind lossKind = LossKindOf(diagnostic);
            if (lossKind > summary.LossKind) summary.LossKind = lossKind;
        }

        protected IReadOnlyList<TDiagnostic> GetDiagnostics() {
            var result = new List<TDiagnostic>(_diagnostics.Count + _suppressedOrder.Count);
            result.AddRange(_diagnostics);
            foreach (string code in _suppressedOrder) {
                SuppressedDiagnostic summary = _suppressed[code];
                result.Add(CreateSummary(code, summary.Severity,
                    $"{summary.Count} additional '{code}' diagnostics were summarized after the per-code detail limit of {_maxDetailedPerCode} was reached.",
                    summary.LossKind));
            }
            return result.ToArray();
        }

        private sealed class SuppressedDiagnostic {
            internal SuppressedDiagnostic(int severity, OfficeConversionLossKind lossKind) {
                Severity = severity;
                LossKind = lossKind;
            }
            internal int Count { get; set; }
            internal int Severity { get; set; }
            internal OfficeConversionLossKind LossKind { get; set; }
        }
    }
}
