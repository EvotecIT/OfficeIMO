using OfficeIMO.Data;

namespace OfficeIMO.Pdf;

/// <summary>Common typed table settings with PDF placement, native styling and bounded object projection.</summary>
public sealed class PdfTableOptions : OfficeTableOptions {
    /// <summary>Gets or sets the table placement alignment.</summary>
    public PdfAlign Align { get; set; } = PdfAlign.Left;
    /// <summary>Gets or sets the native style, copied before common table options are applied.</summary>
    public PdfTableStyle? Style { get; set; }
    /// <summary>Gets or sets nested paths, collection expansion and materialization limits. Common table options take precedence.</summary>
    public ObjectFlattenerOptions? Projection { get; set; }
}
