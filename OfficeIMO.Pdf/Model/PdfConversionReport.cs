namespace OfficeIMO.Pdf;

/// <summary>
/// Collects shared PDF conversion warnings from one adapter run.
/// </summary>
public sealed class PdfConversionReport {
    private readonly List<PdfConversionWarning> _warnings = new();

    /// <summary>Warnings recorded by the converter in production order.</summary>
    public IReadOnlyList<PdfConversionWarning> Warnings => _warnings;

    /// <summary>True when at least one warning was recorded.</summary>
    public bool HasWarnings => _warnings.Count > 0;

    /// <summary>Adds one warning to the report.</summary>
    public void Add(PdfConversionWarning warning) {
        Guard.NotNull(warning, nameof(warning));
        _warnings.Add(warning);
    }

    /// <summary>Adds all warnings from another report.</summary>
    public void AddRange(IEnumerable<PdfConversionWarning> warnings) {
        Guard.NotNull(warnings, nameof(warnings));
        foreach (PdfConversionWarning warning in warnings) {
            Add(warning);
        }
    }

    /// <summary>Clears warnings from a previous conversion run.</summary>
    public void Clear() {
        _warnings.Clear();
    }
}
