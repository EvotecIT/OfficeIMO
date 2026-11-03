using OfficeIMO.Html;

namespace OfficeIMO.Html.Pdf;

internal static class HtmlRenderOptionsPdfImages {
    internal static global::OfficeIMO.Pdf.PdfImageExportOptions ToPdfImageOptions(this HtmlRenderOptions source) {
        var options = source.CopyImageExportOptionsTo(new global::OfficeIMO.Pdf.PdfImageExportOptions());
        // HTML dimensions are CSS pixels; PDF page dimensions are points.
        // TargetDpi is resolved by the PDF owner against its own logical units.
        if (!source.TargetDpi.HasValue) options.Scale = source.Scale * HtmlRenderOptions.CssPixelsPerInch / 72D;
        return options;
    }

    internal static global::OfficeIMO.Pdf.PdfPageSelection? ToPdfPageSelection(this HtmlRenderOptions source) =>
        source.PageSelection == null ? null : global::OfficeIMO.Pdf.PdfPageSelection.From(source.PageSelection.ToArray());
}
