namespace OfficeIMO.Pdf;

/// <summary>Controls where a newly added or rewritten image is placed relative to existing page content.</summary>
public enum PdfImageEditLayer {
    /// <summary>Places the image after existing page content.</summary>
    AboveExistingContent,
    /// <summary>Places the image before existing page content.</summary>
    BehindExistingContent
}

/// <summary>Options for adding, replacing, or moving an image on an existing PDF page.</summary>
public sealed class PdfImageEditOptions {
    private PdfImageEditLayer _layer = PdfImageEditLayer.AboveExistingContent;

    /// <summary>
    /// Placement layer for the newly written image. Existing exact paint order cannot be preserved by a portable
    /// page-content rewrite, so callers choose explicitly between the front and back of existing content.
    /// </summary>
    public PdfImageEditLayer Layer {
        get => _layer;
        set {
            if (value != PdfImageEditLayer.AboveExistingContent && value != PdfImageEditLayer.BehindExistingContent) {
                throw new ArgumentOutOfRangeException(nameof(Layer), "Image edit layer is not supported.");
            }
            _layer = value;
        }
    }

    internal PdfImageEditOptions Snapshot() => new PdfImageEditOptions { Layer = Layer };
}

/// <summary>Result of an existing-page image edit.</summary>
public sealed class PdfImageEditResult {
    internal PdfImageEditResult(PdfDocument document, int affectedCount) {
        Document = document;
        AffectedCount = affectedCount;
    }

    /// <summary>Edited immutable document.</summary>
    public PdfDocument Document { get; }

    /// <summary>Number of source placements affected, or one for a successful add.</summary>
    public int AffectedCount { get; }
}
