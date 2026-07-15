namespace OfficeIMO.Reader.Image;

/// <summary>Controls standalone image projection into the Reader document model.</summary>
public sealed class ReaderImageOptions {
    /// <summary>
    /// Gets or sets whether the source bytes are retained as a materializable document asset. Default: true.
    /// </summary>
    public bool IncludePayload { get; set; } = true;

    /// <summary>
    /// Gets or sets whether an OCR candidate is emitted for a retained image payload. Default: true.
    /// This does not execute OCR or configure an OCR engine.
    /// </summary>
    public bool CreateOcrCandidate { get; set; } = true;

    internal ReaderImageOptions CloneValidated() {
        if (CreateOcrCandidate && !IncludePayload) {
            throw new ArgumentException(
                nameof(CreateOcrCandidate) + " requires " + nameof(IncludePayload) + " so an OCR engine can access the image bytes.");
        }
        return new ReaderImageOptions {
            IncludePayload = IncludePayload,
            CreateOcrCandidate = CreateOcrCandidate
        };
    }
}
