namespace OfficeIMO.Workflows;

internal sealed partial class OfficeWorkflowRenderingOptions {
    internal static OfficeWorkflowRenderingOptions Capture(OfficeConversionOptions options) {
        OfficeConversionOptions source = options.Clone();
        OfficeWorkflowRenderingOptions copy = (source.RegisteredSettings as OfficeWorkflowRouteSettings ?? new()).Capture(false);
        copy.SourcePassword = source.SourcePassword;
        copy.Word = source.Word; copy.Excel = source.Excel; copy.PowerPoint = source.PowerPoint;
        copy.Html = source.Html ?? copy.Html; copy.Markdown = source.Markdown; copy.PlainText = source.PlainText;
        copy.RequireNoLoss = source.FailOnLoss;
        copy.CompressPdfOutput = source.CompressPdfOutput;
        copy.LegacyDocLossPolicy = source.WordLoad?.FailOnLoss == true ? OfficeConversionLossPolicy.Block : OfficeConversionLossPolicy.Allow;
        copy.PageRanges = source.PdfLoad?.PageSelection?.ToString();
        copy.WordMode = source.PdfWord?.Mode; copy.PowerPointMode = source.PdfPowerPoint?.Mode;
        copy.RasterDpi = source.PdfWord?.Dpi ?? source.PdfPowerPoint?.Dpi;
        copy.WorksheetLayout = source.Excel?.WorksheetLayout; copy.HtmlProfile = source.PdfHtml?.Profile;
        return copy;
    }
}
