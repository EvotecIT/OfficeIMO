namespace OfficeIMO.CSV;

/// <summary>
/// Exposes parsing metadata selected by an OfficeIMO CSV data reader.
/// </summary>
public interface ICsvDataReaderMetadata {
    /// <summary>
    /// Gets the delimiter used to parse the reader input, including a delimiter selected by detection.
    /// </summary>
    char Delimiter { get; }
}
