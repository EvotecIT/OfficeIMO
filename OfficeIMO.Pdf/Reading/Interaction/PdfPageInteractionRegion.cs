namespace OfficeIMO.Pdf;

/// <summary>Selectable text or an interactive PDF page region.</summary>
public sealed class PdfPageInteractionRegion {
    internal PdfPageInteractionRegion(
        PdfInteractionKind kind,
        PdfSelectionQuad quad,
        string? text = null,
        string? target = null,
        string? subtype = null,
        string? fieldName = null,
        int? objectNumber = null,
        int textIndex = -1) {
        Kind = kind;
        Quad = quad;
        Text = text;
        Target = target;
        Subtype = subtype;
        FieldName = fieldName;
        ObjectNumber = objectNumber;
        TextIndex = textIndex;
    }

    /// <summary>Region kind.</summary>
    public PdfInteractionKind Kind { get; }

    /// <summary>Visual top-left page-coordinate quad.</summary>
    public PdfSelectionQuad Quad { get; }

    /// <summary>Unicode text element for text regions.</summary>
    public string? Text { get; }

    /// <summary>Link URI, destination, named action, or remote target summary.</summary>
    public string? Target { get; }

    /// <summary>Annotation subtype when applicable.</summary>
    public string? Subtype { get; }

    /// <summary>Fully qualified form field name when applicable.</summary>
    public string? FieldName { get; }

    /// <summary>Indirect annotation or widget object number when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Zero-based text-element index on the page, or -1 for non-text regions.</summary>
    public int TextIndex { get; }
}
