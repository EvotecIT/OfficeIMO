namespace OfficeIMO.OpenDocument;

/// <summary>Converted document together with the feature mapping evidence for that conversion.</summary>
public sealed class OdfConversionResult<TDocument> where TDocument : class {
    /// <summary>Creates a conversion result.</summary>
    public OdfConversionResult(TDocument document, OdfConversionReport report) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>The converted in-memory document.</summary>
    public TDocument Document { get; }
    /// <summary>Feature-level conversion report.</summary>
    public OdfConversionReport Report { get; }
}
