namespace OfficeIMO.Pdf;

/// <summary>Kind of content intersecting a redaction planning area.</summary>
public enum PdfRedactionMatchKind {
    /// <summary>Line-level logical text block.</summary>
    TextBlock = 0,
    /// <summary>Page annotation rectangle.</summary>
    Annotation = 1,
    /// <summary>Image XObject placement rectangle.</summary>
    ImagePlacement = 2
}

/// <summary>Content intersecting a requested redaction area.</summary>
public sealed class PdfRedactionMatch {
    internal PdfRedactionMatch(
        PdfRedactionMatchKind kind,
        PdfRedactionArea area,
        int pageNumber,
        double x,
        double y,
        double width,
        double height,
        string? text,
        string? subtype,
        int? objectNumber,
        string? resourceName = null) {
        Kind = kind;
        Area = area;
        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Text = text;
        Subtype = subtype;
        ObjectNumber = objectNumber;
        ResourceName = resourceName;
    }

    /// <summary>Matched content kind.</summary>
    public PdfRedactionMatchKind Kind { get; }

    /// <summary>Requested redaction area that produced the match.</summary>
    public PdfRedactionArea Area { get; }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Matched content left coordinate in PDF points.</summary>
    public double X { get; }

    /// <summary>Matched content bottom coordinate in PDF points.</summary>
    public double Y { get; }

    /// <summary>Matched content width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Matched content height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Matched text, when available.</summary>
    public string? Text { get; }

    /// <summary>Annotation subtype, when matching an annotation.</summary>
    public string? Subtype { get; }

    /// <summary>Related PDF object number, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Related PDF resource name, when matching an image placement.</summary>
    public string? ResourceName { get; }
}
