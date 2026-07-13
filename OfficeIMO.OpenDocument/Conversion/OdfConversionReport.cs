namespace OfficeIMO.OpenDocument;

/// <summary>Feature mapping report for one explicit conversion between OpenDocument and another Office format.</summary>
public sealed class OdfConversionReport {
    private readonly List<OdfConversionMapping> _mappings = new List<OdfConversionMapping>();

    /// <summary>Creates an empty conversion report.</summary>
    public OdfConversionReport(string sourceFormat, string targetFormat) {
        if (string.IsNullOrWhiteSpace(sourceFormat)) throw new ArgumentException("Source format cannot be empty.", nameof(sourceFormat));
        if (string.IsNullOrWhiteSpace(targetFormat)) throw new ArgumentException("Target format cannot be empty.", nameof(targetFormat));
        SourceFormat = sourceFormat;
        TargetFormat = targetFormat;
    }

    /// <summary>Source format identifier.</summary>
    public string SourceFormat { get; }
    /// <summary>Target format identifier.</summary>
    public string TargetFormat { get; }
    /// <summary>Feature results in the order reported by the adapter.</summary>
    public IReadOnlyList<OdfConversionMapping> Mappings => _mappings;
    /// <summary>True when at least one feature was approximated, skipped, or unsupported.</summary>
    public bool HasLoss => _mappings.Any(mapping => mapping.Status != OdfConversionMappingStatus.Converted);

    /// <summary>Throws when any feature was approximated, skipped, or unsupported.</summary>
    public void RequireNoLoss() {
        if (HasLoss) {
            throw new InvalidOperationException(
                $"Conversion from {SourceFormat} to {TargetFormat} was lossy. Inspect the conversion report for details.");
        }
    }

    /// <summary>Adds one feature-level result and returns this report.</summary>
    public OdfConversionReport Add(string feature, OdfConversionMappingStatus status, int count = 1, string? message = null) {
        _mappings.Add(new OdfConversionMapping(feature, status, count, message));
        return this;
    }

    /// <summary>Returns all results for one stable feature identifier.</summary>
    public IReadOnlyList<OdfConversionMapping> ForFeature(string feature) {
        if (string.IsNullOrWhiteSpace(feature)) throw new ArgumentException("Feature name cannot be empty.", nameof(feature));
        return _mappings.Where(mapping => string.Equals(mapping.Feature, feature, StringComparison.Ordinal)).ToList();
    }
}
