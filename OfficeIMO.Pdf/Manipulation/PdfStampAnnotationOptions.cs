namespace OfficeIMO.Pdf;

/// <summary>Placement, appearance, and mutation policy for a stamp annotation added to an existing PDF page.</summary>
public sealed class PdfStampAnnotationOptions {
    /// <summary>One-based target page number.</summary>
    public int PageNumber { get; set; } = 1;

    /// <summary>Left edge in PDF points.</summary>
    public double X { get; set; } = 36D;

    /// <summary>Bottom edge in PDF points.</summary>
    public double Y { get; set; } = 36D;

    /// <summary>Annotation width in PDF points.</summary>
    public double Width { get; set; } = 144D;

    /// <summary>Annotation height in PDF points.</summary>
    public double Height { get; set; } = 48D;

    /// <summary>Standard or custom PDF stamp name, for example Approved, Draft, or TopSecret.</summary>
    public string StampName { get; set; } = "Approved";

    /// <summary>Optional human-readable annotation contents.</summary>
    public string? Contents { get; set; }

    /// <summary>Optional annotation author/title.</summary>
    public string? Title { get; set; }

    /// <summary>Optional stable annotation name stored in /NM.</summary>
    public string? Name { get; set; }

    /// <summary>Annotation flags. The default value 4 enables printing.</summary>
    public int Flags { get; set; } = 4;

    /// <summary>Border and label color. The default is dark red.</summary>
    public PdfColor StrokeColor { get; set; } = new PdfColor(0.7D, 0.05D, 0.05D);

    /// <summary>Optional appearance fill color.</summary>
    public PdfColor? FillColor { get; set; }

    /// <summary>Border width in PDF points.</summary>
    public double BorderWidth { get; set; } = 2D;

    /// <summary>Preferred mutation mode. Automatic uses append-only when signature policy requires it.</summary>
    public PdfMutationExecutionPreference ExecutionPreference { get; set; } = PdfMutationExecutionPreference.Automatic;
}
