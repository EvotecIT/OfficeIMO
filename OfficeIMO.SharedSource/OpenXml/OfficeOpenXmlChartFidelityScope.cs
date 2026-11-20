using OfficeIMO.Drawing;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>
/// Carries the requested chart fidelity through the native chart readers for one projection and
/// collects what was approximated. Readers that run outside a scope behave strictly.
/// </summary>
internal sealed class OfficeOpenXmlChartFidelityScope : IDisposable {
    [ThreadStatic] private static OfficeOpenXmlChartFidelityScope? _current;
    private readonly OfficeOpenXmlChartFidelityScope? _previous;
    private readonly List<string> _approximations = new();
    private bool _disposed;

    private OfficeOpenXmlChartFidelityScope(OfficeChartFidelity fidelity) {
        Fidelity = fidelity;
        _previous = _current;
        _current = this;
    }

    /// <summary>Starts a projection scope on the current thread.</summary>
    internal static OfficeOpenXmlChartFidelityScope Enter(OfficeChartFidelity fidelity) => new(fidelity);

    internal OfficeChartFidelity Fidelity { get; }

    /// <summary>Distinct approximations recorded while this scope was current, in first-seen order.</summary>
    internal IReadOnlyList<string> Approximations => _approximations;

    /// <summary>Whether the current projection may draw an approximation instead of skipping the chart.</summary>
    internal static bool AllowsApproximation => _current?.Fidelity == OfficeChartFidelity.BestEffort;

    /// <summary>
    /// Accepts an approximation in best-effort mode and records it; strict mode throws
    /// <see cref="NotSupportedException"/> with <paramref name="strictReason"/>.
    /// </summary>
    internal static void Approximate(string strictReason, string approximation) {
        if (!AllowsApproximation) throw new NotSupportedException(strictReason);
        Record(approximation);
    }

    /// <summary>Records a substitution that every fidelity draws, such as a General number-format fallback.</summary>
    internal static void Record(string approximation) {
        OfficeOpenXmlChartFidelityScope? scope = _current;
        if (scope == null || string.IsNullOrWhiteSpace(approximation)) return;
        if (!scope._approximations.Contains(approximation)) scope._approximations.Add(approximation);
    }

    /// <summary>The number of approximations recorded so far, for <see cref="RollBack"/>.</summary>
    internal static int Checkpoint => _current?._approximations.Count ?? 0;

    /// <summary>Discards approximations recorded by a read that was abandoned for a fallback.</summary>
    internal static void RollBack(int checkpoint) {
        List<string>? approximations = _current?._approximations;
        if (approximations != null && checkpoint < approximations.Count)
            approximations.RemoveRange(checkpoint, approximations.Count - checkpoint);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        if (ReferenceEquals(_current, this)) _current = _previous;
    }
}
