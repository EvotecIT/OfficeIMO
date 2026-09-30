namespace OfficeIMO.CSV;

/// <summary>Exposes the complete delimiter selected by a CSV reader.</summary>
/// <remarks>The inherited character delimiter is the first character of <see cref="DelimiterText"/>.</remarks>
public interface ICsvDataReaderDialectMetadata : ICsvDataReaderMetadata {
    /// <summary>Gets the complete single- or multi-character input delimiter.</summary>
    string DelimiterText { get; }
}
