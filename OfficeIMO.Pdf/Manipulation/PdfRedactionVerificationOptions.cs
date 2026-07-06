namespace OfficeIMO.Pdf;

/// <summary>
/// Configures post-redaction checks for removed and retained PDF text markers.
/// </summary>
public sealed class PdfRedactionVerificationOptions {
    private readonly List<string> _removedTextMarkers = new List<string>();
    private readonly List<string> _retainedTextMarkers = new List<string>();

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
