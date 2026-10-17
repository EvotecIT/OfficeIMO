namespace OfficeIMO.Html;

public partial class HtmlRenderOptions {
    internal global::OfficeIMO.Pdf.PdfImageExportOptions ToPdfImageOptions() {
        var options = CopyImageExportOptionsTo(new global::OfficeIMO.Pdf.PdfImageExportOptions());
        // HTML dimensions are CSS pixels; PDF page dimensions are points.
        // TargetDpi is resolved by the PDF owner against its own logical units.
        if (!TargetDpi.HasValue) options.Scale = Scale * CssPixelsPerInch / 72D;
        return options;
    }
}
