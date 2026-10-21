namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Layout capture and painting enclose the same authored run background.
    private static (double Horizontal, double Vertical) GetRichTextBackgroundPadding(double fontSize) =>
        (Math.Max(1.4D, fontSize * .14D), Math.Max(.45D, fontSize * .05D));
}
