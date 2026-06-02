namespace OfficeIMO.Pdf;

/// <summary>
/// Image XObject entry in the logical page model.
/// </summary>
public sealed class PdfLogicalImage : IPdfLogicalElement {
    internal PdfLogicalImage(PdfExtractedImage image) {
        SourceImage = image;
    }

    /// <inheritdoc />
    public int PageNumber => SourceImage.PageNumber;

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.Image;

    /// <summary>Underlying extracted image payload and metadata.</summary>
    public PdfExtractedImage SourceImage { get; }

    /// <summary>PDF image resource name.</summary>
    public string ResourceName => SourceImage.ResourceName;

    /// <summary>Image width in pixels.</summary>
    public int Width => SourceImage.Width;

    /// <summary>Image height in pixels.</summary>
    public int Height => SourceImage.Height;

    /// <summary>Suggested MIME type when bytes are a complete image file.</summary>
    public string? MimeType => SourceImage.MimeType;
}

/// <summary>
/// Link annotation entry in the logical page model.
/// </summary>
public sealed class PdfLogicalLinkAnnotation : IPdfLogicalElement {
    internal PdfLogicalLinkAnnotation(int pageNumber, PdfLinkAnnotation link) {
        PageNumber = pageNumber;
        SourceLink = link.PageNumber == pageNumber ? link : link.WithPageNumber(pageNumber);
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.LinkAnnotation;

    /// <summary>Underlying parsed link annotation.</summary>
    public PdfLinkAnnotation SourceLink { get; }

    /// <summary>Absolute URI opened by the link annotation, or null for an internal named-destination link.</summary>
    public string? Uri => SourceLink.Uri;

    /// <summary>Named destination opened by the link annotation, or null for a URI link.</summary>
    public string? DestinationName => SourceLink.DestinationName;

    /// <summary>True when the link annotation opens an absolute URI.</summary>
    public bool IsUriLink => SourceLink.IsUriLink;

    /// <summary>True when the link annotation opens an internal named destination.</summary>
    public bool IsNamedDestinationLink => SourceLink.IsNamedDestinationLink;

    /// <summary>Optional annotation contents metadata.</summary>
    public string? Contents => SourceLink.Contents;

    /// <summary>Left edge of the annotation rectangle in PDF points.</summary>
    public double X1 => SourceLink.X1;

    /// <summary>Bottom edge of the annotation rectangle in PDF points.</summary>
    public double Y1 => SourceLink.Y1;

    /// <summary>Right edge of the annotation rectangle in PDF points.</summary>
    public double X2 => SourceLink.X2;

    /// <summary>Top edge of the annotation rectangle in PDF points.</summary>
    public double Y2 => SourceLink.Y2;

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => SourceLink.Width;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => SourceLink.Height;
}

/// <summary>
/// AcroForm widget annotation entry in the logical page model.
/// </summary>
public sealed class PdfLogicalFormWidget : IPdfLogicalElement {
    internal PdfLogicalFormWidget(int pageNumber, PdfFormField field, PdfFormWidget widget) {
        PageNumber = pageNumber;
        Field = field;
        SourceWidget = widget;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.FormWidget;

    /// <summary>Field represented by this widget annotation.</summary>
    public PdfFormField Field { get; }

    /// <summary>Underlying parsed widget annotation.</summary>
    public PdfFormWidget SourceWidget { get; }

    /// <summary>Fully qualified field name when a name can be read.</summary>
    public string? FieldName => Field.Name;

    /// <summary>Field type name, for example Tx, Btn, Ch, or Sig, when present or inherited.</summary>
    public string? FieldType => Field.FieldType;

    /// <summary>Simple field value formatted for wrapper display, when present.</summary>
    public string? Value => Field.Value;

    /// <summary>Indirect object number for the widget annotation, when known.</summary>
    public int? ObjectNumber => SourceWidget.ObjectNumber;

    /// <summary>Left edge of the widget rectangle in PDF points.</summary>
    public double X1 => SourceWidget.X1;

    /// <summary>Bottom edge of the widget rectangle in PDF points.</summary>
    public double Y1 => SourceWidget.Y1;

    /// <summary>Right edge of the widget rectangle in PDF points.</summary>
    public double X2 => SourceWidget.X2;

    /// <summary>Top edge of the widget rectangle in PDF points.</summary>
    public double Y2 => SourceWidget.Y2;

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => SourceWidget.Width;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => SourceWidget.Height;

    /// <summary>Current widget appearance state name from /AS, when present.</summary>
    public string? AppearanceState => SourceWidget.AppearanceState;

    /// <summary>Raw widget annotation flags from /F, when present.</summary>
    public int? Flags => SourceWidget.Flags;

    /// <summary>True when the widget has the PDF annotation Invisible flag.</summary>
    public bool IsInvisible => SourceWidget.IsInvisible;

    /// <summary>True when the widget has the PDF annotation Hidden flag.</summary>
    public bool IsHidden => SourceWidget.IsHidden;

    /// <summary>True when the widget has the PDF annotation Print flag.</summary>
    public bool IsPrint => SourceWidget.IsPrint;

    /// <summary>True when the widget has the PDF annotation NoZoom flag.</summary>
    public bool IsNoZoom => SourceWidget.IsNoZoom;

    /// <summary>True when the widget has the PDF annotation NoRotate flag.</summary>
    public bool IsNoRotate => SourceWidget.IsNoRotate;

    /// <summary>True when the widget has the PDF annotation NoView flag.</summary>
    public bool IsNoView => SourceWidget.IsNoView;

    /// <summary>True when the widget has the PDF annotation ReadOnly flag.</summary>
    public bool IsReadOnly => SourceWidget.IsReadOnly;

    /// <summary>True when the widget has the PDF annotation Locked flag.</summary>
    public bool IsLocked => SourceWidget.IsLocked;

    /// <summary>True when the widget has the PDF annotation ToggleNoView flag.</summary>
    public bool IsToggleNoView => SourceWidget.IsToggleNoView;

    /// <summary>True when the widget has the PDF annotation LockedContents flag.</summary>
    public bool IsLockedContents => SourceWidget.IsLockedContents;

    /// <summary>Normal appearance state names from /AP /N, when the widget exposes named appearance streams.</summary>
    public IReadOnlyList<string> NormalAppearanceStates => SourceWidget.NormalAppearanceStates;

    /// <summary>Number of readable normal appearance states.</summary>
    public int NormalAppearanceStateCount => SourceWidget.NormalAppearanceStateCount;

    /// <summary>True when at least one normal appearance state was readable.</summary>
    public bool HasNormalAppearanceStates => SourceWidget.HasNormalAppearanceStates;

    /// <summary>Returns true when the widget exposes a matching normal appearance state name.</summary>
    public bool HasNormalAppearanceState(string state) {
        return SourceWidget.HasNormalAppearanceState(state);
    }
}
