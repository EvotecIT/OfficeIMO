namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Extracts image and font files, ICC profiles, page SVGs and JSON and HTML reports into a new directory.</summary>
    /// <returns>The relative paths of the files written, using forward slashes.</returns>
    public IReadOnlyList<string> ExtractAssets(string outputDirectory, PdfAssetExtractionOptions? options = null,
        CancellationToken cancellationToken = default) =>
        PdfAssetOutput.WriteDirectory(outputDirectory, create => ExtractAssetsTo(create, options, cancellationToken),
            cancellationToken: cancellationToken);

    /// <summary>Extracts a PDF asset bundle through a factory that opens one writable stream per relative path.</summary>
    /// <remarks>The engine disposes each stream before opening the next. A failure can leave entries already written by the caller's factory.</remarks>
    /// <returns>The relative paths of the files written, using forward slashes.</returns>
    public IReadOnlyList<string> ExtractAssetsTo(Func<string, Stream> createOutput, PdfAssetExtractionOptions? options = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(createOutput, nameof(createOutput));
        PdfAssetExtractionOptions effective = options ?? new PdfAssetExtractionOptions();
        effective.Validate();
        return PdfAssetExtractor.Extract(GetReadSnapshot(cancellationToken: cancellationToken).Document,
            createOutput, effective, cancellationToken);
    }

    /// <summary>Saves distinct decoded embedded font programs from page resources and reachable appearances into a new directory.</summary>
    /// <remarks>Supported standalone TrueType and OpenType fonts with unusable names receive a name table. PDF subsets retain their original glyph coverage.</remarks>
    /// <returns>The relative paths of the saved font files, using forward slashes.</returns>
    public IReadOnlyList<string> ExtractEmbeddedFonts(string outputDirectory, CancellationToken cancellationToken = default) =>
        PdfAssetOutput.WriteDirectory(outputDirectory, create => {
            var output = new PdfAssetOutput(create, new PdfAssetExtractionOptions().MaximumOutputBytes, cancellationToken);
            PdfAssetExtractor.ExtractStandaloneFonts(GetReadSnapshot(cancellationToken: cancellationToken).Document, output, cancellationToken);
            return output.Paths;
        }, cancellationToken: cancellationToken);
}
