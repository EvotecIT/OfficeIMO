namespace OfficeIMO.Pdf;

/// <summary>
/// Simple AcroForm field information read from a PDF document.
/// </summary>
public sealed class PdfFormField {
    internal PdfFormField(int? objectNumber, string? name, string? partialName, string? fieldType, string? value, string? alternateName, string? mappingName, int? flags, IReadOnlyList<PdfFormWidget>? widgets = null) {
        ObjectNumber = objectNumber;
        Name = name;
        PartialName = partialName;
        FieldType = fieldType;
        Value = value;
        AlternateName = alternateName;
        MappingName = mappingName;
        Flags = flags;
        Widgets = widgets ?? Array.Empty<PdfFormWidget>();
    }

    /// <summary>Indirect object number for the field dictionary, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Fully qualified field name when a name can be read.</summary>
    public string? Name { get; }

    /// <summary>Partial field name from the field dictionary.</summary>
    public string? PartialName { get; }

    /// <summary>Field type name, for example Tx, Btn, Ch, or Sig, when present or inherited.</summary>
    public string? FieldType { get; }

    /// <summary>Simple field value formatted for wrapper display, when present.</summary>
    public string? Value { get; }

    /// <summary>Alternate field name used as a user-facing label, when present.</summary>
    public string? AlternateName { get; }

    /// <summary>Mapping name used for export workflows, when present.</summary>
    public string? MappingName { get; }

    /// <summary>Raw field flags from /Ff, when present.</summary>
    public int? Flags { get; }

    /// <summary>Simple widget annotations that visually represent this field, when readable.</summary>
    public IReadOnlyList<PdfFormWidget> Widgets { get; }

    /// <summary>Number of readable widget annotations associated with this field.</summary>
    public int WidgetCount => Widgets.Count;

    /// <summary>True when at least one widget annotation was associated with this field.</summary>
    public bool HasWidgets => Widgets.Count > 0;
}

/// <summary>
/// Simple AcroForm widget annotation geometry read from a PDF document.
/// </summary>
public sealed class PdfFormWidget {
    internal PdfFormWidget(int? objectNumber, int? pageNumber, double x1, double y1, double x2, double y2, string? appearanceState, int? flags) {
        ObjectNumber = objectNumber;
        PageNumber = pageNumber;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        AppearanceState = appearanceState;
        Flags = flags;
    }

    /// <summary>Indirect object number for the widget annotation, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>One-based page number containing the widget annotation, when known.</summary>
    public int? PageNumber { get; }

    /// <summary>Left edge of the widget rectangle in PDF points.</summary>
    public double X1 { get; }

    /// <summary>Bottom edge of the widget rectangle in PDF points.</summary>
    public double Y1 { get; }

    /// <summary>Right edge of the widget rectangle in PDF points.</summary>
    public double X2 { get; }

    /// <summary>Top edge of the widget rectangle in PDF points.</summary>
    public double Y2 { get; }

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => X2 - X1;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => Y2 - Y1;

    /// <summary>Current widget appearance state name from /AS, when present.</summary>
    public string? AppearanceState { get; }

    /// <summary>Raw widget annotation flags from /F, when present.</summary>
    public int? Flags { get; }
}
