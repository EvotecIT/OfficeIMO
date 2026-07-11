namespace OfficeIMO.Html;

/// <summary>
/// Shared result contract for HTML conversions that produce a native target artifact.
/// </summary>
/// <typeparam name="TArtifact">Native artifact type produced by the adapter.</typeparam>
public abstract class HtmlConversionResult<TArtifact> {
    /// <summary>Creates a conversion result for the supplied artifact.</summary>
    protected HtmlConversionResult(TArtifact artifact) {
        Artifact = artifact ?? throw new ArgumentNullException(nameof(artifact));
    }

    /// <summary>Native artifact produced by the conversion.</summary>
    public TArtifact Artifact { get; }

    /// <summary>Structured conversion diagnostics in emission order.</summary>
    public HtmlDiagnosticReport Diagnostics { get; } = new HtmlDiagnosticReport();

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !Diagnostics.HasErrors;

    /// <summary>
    /// Returns the native artifact when conversion succeeded, or throws a structured conversion exception.
    /// </summary>
    public TArtifact GetArtifactOrThrow() {
        if (!Succeeded) {
            throw new HtmlConversionException(Diagnostics.Diagnostics);
        }

        return Artifact;
    }
}
