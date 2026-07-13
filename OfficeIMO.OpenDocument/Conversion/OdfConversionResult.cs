namespace OfficeIMO.OpenDocument;

/// <summary>Converted document together with the feature mapping evidence for that conversion.</summary>
public sealed class OdfConversionResult<TDocument> where TDocument : class {
    /// <summary>Creates a conversion result.</summary>
    public OdfConversionResult(TDocument value, OdfConversionReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>The converted in-memory document.</summary>
    public TDocument Value { get; }
    /// <summary>Feature-level conversion report.</summary>
    public OdfConversionReport Report { get; }
    /// <summary>True when at least one feature was approximated, skipped, or unsupported.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the converted document.</summary>
    public TDocument RequireValue() => Value;

    /// <summary>Returns the converted document or throws when the conversion was lossy.</summary>
    public TDocument RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
