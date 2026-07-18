namespace OfficeIMO.Pdf;

/// <summary>Field kinds supported for existing-document creation.</summary>
public enum PdfFormFieldCreationKind {
    /// <summary>Text field.</summary>
    Text,
    /// <summary>Check box button field.</summary>
    CheckBox,
    /// <summary>Single-value choice field.</summary>
    Choice,
    /// <summary>Empty signature field.</summary>
    Signature
}

/// <summary>Page annotation tab-order hint stored in /Tabs.</summary>
public enum PdfPageTabOrder {
    /// <summary>Row order (/R).</summary>
    Row,
    /// <summary>Column order (/C).</summary>
    Column,
    /// <summary>Structure-tree order (/S).</summary>
    Structure,
    /// <summary>Annotation-array order (/A).</summary>
    Annotations
}

/// <summary>Creates one flat AcroForm field and widget on an existing page.</summary>
public sealed class PdfFormFieldCreateOptions {
    /// <summary>Unique fully qualified field name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Field kind.</summary>
    public PdfFormFieldCreationKind Kind { get; set; } = PdfFormFieldCreationKind.Text;
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; set; } = 1;
    /// <summary>Widget left coordinate.</summary>
    public double X { get; set; } = 36D;
    /// <summary>Widget bottom coordinate.</summary>
    public double Y { get; set; } = 36D;
    /// <summary>Widget width.</summary>
    public double Width { get; set; } = 180D;
    /// <summary>Widget height.</summary>
    public double Height { get; set; } = 22D;
    /// <summary>Initial scalar value.</summary>
    public string Value { get; set; } = string.Empty;
    /// <summary>Optional default value.</summary>
    public string? DefaultValue { get; set; }
    /// <summary>Raw field /Ff flags.</summary>
    public int FieldFlags { get; set; }
    /// <summary>Raw widget annotation /F flags. Default enables printing.</summary>
    public int WidgetFlags { get; set; } = 4;
    /// <summary>Choice options for choice fields.</summary>
    public IReadOnlyList<string> ChoiceOptions { get; set; } = Array.Empty<string>();
    /// <summary>Export name used for a checked checkbox.</summary>
    public string CheckedValueName { get; set; } = "Yes";
}

/// <summary>Result and proof for a transactional AcroForm edit.</summary>
public sealed class PdfAcroFormEditResult {
    private readonly byte[] _pdf;
    internal PdfAcroFormEditResult(byte[] pdf, PdfMutationPlan plan, PdfRewritePreservationReport preservation, IReadOnlyList<PdfFormField> fields, IReadOnlyList<string> calculationOrder, IReadOnlyList<string> operations) { _pdf = (byte[])pdf.Clone(); MutationPlan = plan; PreservationReport = preservation; Fields = fields; CalculationOrder = calculationOrder; Operations = operations; }
    /// <summary>Shared mutation plan.</summary>
    public PdfMutationPlan MutationPlan { get; }
    /// <summary>Non-form preservation proof.</summary>
    public PdfRewritePreservationReport PreservationReport { get; }
    /// <summary>Fields read back from the saved artifact.</summary>
    public IReadOnlyList<PdfFormField> Fields { get; }
    /// <summary>Fully qualified field names read back from AcroForm /CO in order.</summary>
    public IReadOnlyList<string> CalculationOrder { get; }
    /// <summary>Stable operation descriptions applied in transaction order.</summary>
    public IReadOnlyList<string> Operations { get; }
    /// <summary>Returns edited PDF bytes.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the edited artifact.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdf);
}
