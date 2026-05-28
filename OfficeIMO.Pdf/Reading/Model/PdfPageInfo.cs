namespace OfficeIMO.Pdf;

/// <summary>
/// Basic geometry and identity information for a page in a PDF document.
/// </summary>
public sealed class PdfPageInfo {
    internal PdfPageInfo(int pageNumber, double width, double height, int rotationDegrees = 0, IReadOnlyList<PdfLinkAnnotation>? linkAnnotations = null, IReadOnlyList<PdfFormWidget>? formWidgets = null) {
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        RotationDegrees = rotationDegrees;
        LinkAnnotations = linkAnnotations ?? Array.Empty<PdfLinkAnnotation>();
        FormWidgets = formWidgets ?? Array.Empty<PdfFormWidget>();
    }

    /// <summary>One-based page number in document order.</summary>
    public int PageNumber { get; }

    /// <summary>Page width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Page height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Inherited page rotation in degrees.</summary>
    public int RotationDegrees { get; }

    /// <summary>Simple URI link annotations on this page.</summary>
    public IReadOnlyList<PdfLinkAnnotation> LinkAnnotations { get; }

    /// <summary>Simple AcroForm widget annotations on this page.</summary>
    public IReadOnlyList<PdfFormWidget> FormWidgets { get; }

    /// <summary>True when at least one simple AcroForm widget annotation was read on this page.</summary>
    public bool HasFormWidgets => FormWidgets.Count > 0;

    /// <summary>Page size in PDF points.</summary>
    public PageSize Size => new PageSize(Width, Height);

    /// <summary>True when the page is wider than it is tall.</summary>
    public bool IsLandscape => Width > Height;
}
