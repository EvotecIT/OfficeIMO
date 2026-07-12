namespace OfficeIMO.Pdf;

/// <summary>Controls bounded PDF debugger projections.</summary>
public sealed class PdfDebuggerOptions {
    /// <summary>Includes decoded text previews for streams whose filters are supported.</summary>
    public bool IncludeDecodedStreamPreviews { get; set; }

    /// <summary>Maximum decoded bytes retained per stream preview.</summary>
    public int MaxDecodedStreamPreviewBytes { get; set; } = 4096;

    /// <summary>Maximum content operators retained per page.</summary>
    public int MaxContentOperatorsPerPage { get; set; } = 4096;

    internal void Validate() {
        if (MaxDecodedStreamPreviewBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxDecodedStreamPreviewBytes), MaxDecodedStreamPreviewBytes, "Decoded stream preview limit must be positive.");
        }

        if (MaxContentOperatorsPerPage <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxContentOperatorsPerPage), MaxContentOperatorsPerPage, "Content operator limit must be positive.");
        }
    }
}
