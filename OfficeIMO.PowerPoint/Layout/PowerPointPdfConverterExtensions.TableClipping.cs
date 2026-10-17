using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static bool ReportTableSlideOverflow(double x, double y, double width, double height,
        double? slideWidth, double? slideHeight, int slideNumber, PowerPointToPdfOptions options) {
        if (!slideWidth.HasValue || !slideHeight.HasValue
            || x >= 0D && y >= 0D && x + width <= slideWidth.Value + .001D && y + height <= slideHeight.Value + .001D) return false;
        AddLayoutWarning(options, slideNumber, "table-cell-overflow",
            "PowerPoint table text grows beyond the slide boundary and is clipped to the slide.",
            PdfCore.PdfLayoutDiagnosticKind.ClippedContent, "PowerPointTableCell",
            "The fitted table exceeds the slide bounds; content outside the slide is clipped.",
            x, y, width, height);
        return true;
    }
}
