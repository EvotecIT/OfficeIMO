namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one executable row in a PDF rewrite-preservation proof matrix.
/// </summary>
public sealed class PdfRewritePreservationMatrixScenario {
    private readonly List<string> _sourceFeatures = new List<string>();

    /// <summary>
    /// Creates a rewrite-preservation matrix scenario.
    /// </summary>
    public PdfRewritePreservationMatrixScenario(string id, string operation, byte[] sourcePdf, Func<byte[], byte[]> rewrite) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Scenario id cannot be empty or whitespace.", nameof(id));
        }

        if (string.IsNullOrWhiteSpace(operation)) {
            throw new ArgumentException("Operation name cannot be empty or whitespace.", nameof(operation));
        }

        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        Guard.NotNull(rewrite, nameof(rewrite));

        Id = id;
        Operation = operation;
        SourcePdf = (byte[])sourcePdf.Clone();
        Rewrite = rewrite;
    }

    /// <summary>Stable scenario id for reports and proof manifests.</summary>
    public string Id { get; }

    /// <summary>Human-readable operation name, such as MetadataUpdate or PageExtraction.</summary>
    public string Operation { get; }

    /// <summary>Original PDF bytes used by the scenario.</summary>
    public byte[] SourcePdf { get; }

    /// <summary>Rewrite operation to execute against <see cref="SourcePdf"/>.</summary>
    public Func<byte[], byte[]> Rewrite { get; }

    /// <summary>Expected matrix classification. Defaults to rewrite-safe.</summary>
    public PdfRewritePreservationMatrixClassification ExpectedClassification { get; set; } = PdfRewritePreservationMatrixClassification.RewriteSafe;

    /// <summary>Preservation options used when the rewrite completes.</summary>
    public PdfRewritePreservationOptions? PreservationOptions { get; set; }

    /// <summary>Feature labels describing the source fixture.</summary>
    public IList<string> SourceFeatures => _sourceFeatures;

    /// <summary>Adds source feature labels and returns this scenario for fluent setup.</summary>
    public PdfRewritePreservationMatrixScenario WithSourceFeatures(params string[] features) {
        Guard.NotNull(features, nameof(features));
        for (int i = 0; i < features.Length; i++) {
            if (!string.IsNullOrWhiteSpace(features[i])) {
                _sourceFeatures.Add(features[i]);
            }
        }

        return this;
    }
}
