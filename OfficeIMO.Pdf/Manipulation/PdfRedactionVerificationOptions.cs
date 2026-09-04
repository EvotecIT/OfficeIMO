using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Configures post-redaction checks for removed and retained PDF text markers.
/// </summary>
public sealed class PdfRedactionVerificationOptions {
    private readonly List<string> _removedTextMarkers = new List<string>();
    private readonly List<string> _retainedTextMarkers = new List<string>();
    private readonly List<IPdfRedactionExternalValidator> _externalValidators = new List<IPdfRedactionExternalValidator>();

    /// <summary>Text markers that must not remain extractable after redaction.</summary>
    public IList<string> RemovedTextMarkers => _removedTextMarkers;

    /// <summary>Text markers that must remain extractable after redaction.</summary>
    public IList<string> RetainedTextMarkers => _retainedTextMarkers;

    /// <summary>True when removed markers should also be checked against raw rewritten PDF bytes decoded as Latin-1.</summary>
    public bool CheckRawPdfBytes { get; set; } = true;

    /// <summary>True when removed markers should also be searched in common PDF string byte encodings and hex strings.</summary>
    public bool CheckEncodedPdfStrings { get; set; } = true;

    /// <summary>True when decoded PDF stream content should be searched for removed markers.</summary>
    public bool CheckDecodedPdfStreams { get; set; } = true;

    /// <summary>True when redaction verification should fail if a PDF stream cannot be decoded while decoded stream checks are enabled.</summary>
    public bool FailOnUndecodablePdfStreams { get; set; } = true;

    /// <summary>
    /// True when every semantic PDF stream must be inspectable even when decoded-stream marker checks or undecodable-stream failures were disabled.
    /// Opaque image-codec payloads are inspected through image placement and mutation evidence instead of generic stream decoding.
    /// Use this for plan-aware verification where image or annotation removal also requires fail-closed proof.
    /// </summary>
    public bool RequireCompleteStreamInspection { get; set; }

    /// <summary>True to compare removed and retained text markers with ordinal casing; false to use ordinal case-insensitive comparison.</summary>
    public bool MatchCase { get; set; } = true;

    /// <summary>Render every page through the managed renderer and fail when a page cannot produce output.</summary>
    public bool CheckManagedRendering { get; set; }

    /// <summary>Cooperatively cancels parsing, marker inspection, managed rendering, and external validation boundaries.</summary>
    public CancellationToken CancellationToken { get; set; }

    /// <summary>
    /// Optional independent validators. Implement <see cref="IPdfRedactionCancellationAwareExternalValidator"/>
    /// when the validator starts a process or another potentially long-running operation.
    /// </summary>
    public IList<IPdfRedactionExternalValidator> ExternalValidators => _externalValidators;

    /// <summary>Adds text markers that must be removed and returns this options object for fluent setup.</summary>
    public PdfRedactionVerificationOptions RequireRemovedText(params string[] markers) {
        AddMarkers(_removedTextMarkers, markers);
        return this;
    }

    /// <summary>Adds text markers that must remain readable and returns this options object for fluent setup.</summary>
    public PdfRedactionVerificationOptions RequireRetainedText(params string[] markers) {
        AddMarkers(_retainedTextMarkers, markers);
        return this;
    }

    private static void AddMarkers(List<string> target, string[] markers) {
        Guard.NotNull(markers, nameof(markers));
        for (int i = 0; i < markers.Length; i++) {
            if (!string.IsNullOrEmpty(markers[i])) {
                target.Add(markers[i]);
            }
        }
    }
}
