namespace OfficeIMO.Studio.Features.Editor;

using OfficeIMO.Pdf;

internal readonly record struct PdfEditorVisualPoint(double X, double Y);

internal sealed record PdfEditorGesture(
    int PageNumber,
    double Left,
    double Top,
    double Right,
    double Bottom,
    IReadOnlyList<PdfEditorVisualPoint> Path);

/// <summary>Kinds of existing page objects that can be selected in the editor.</summary>
public enum PdfEditorSelectionKind {
    /// <summary>Existing page text inside a visual selection.</summary>
    Text,

    /// <summary>One exact image placement.</summary>
    Image,

    /// <summary>One indirect annotation object.</summary>
    Annotation
}

/// <summary>Object-selection policy projected by the active document mode.</summary>
public enum PdfEditorSelectionMode {
    /// <summary>Retain reader text selection and link activation without selecting editable objects.</summary>
    None,

    /// <summary>Select annotations while retaining reader behavior for other page content.</summary>
    Annotations,

    /// <summary>Select existing text, images, and annotations.</summary>
    PageContent
}

/// <summary>Visual top-left bounds for one editor selection.</summary>
public readonly record struct PdfEditorVisualBounds(double Left, double Top, double Right, double Bottom) {
    /// <summary>Selection width.</summary>
    public double Width => Right - Left;

    /// <summary>Selection height.</summary>
    public double Height => Bottom - Top;
}

/// <summary>Revision-scoped selection of existing PDF page content.</summary>
public sealed record PdfEditorSelection(
    PdfEditorSelectionKind Kind,
    int PageNumber,
    PdfEditorVisualBounds Bounds,
    string? Text = null,
    int? ObjectNumber = null,
    string? Subtype = null,
    PdfImagePlacement? ImagePlacement = null);
