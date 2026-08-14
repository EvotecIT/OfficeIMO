namespace OfficeIMO.Pdf;

internal static class PdfRadioButtonLayout {
    internal static double GetLabelGap(double buttonSize) => Math.Max(4D, buttonSize * 0.4D);

    internal static double GetLabelFontSize(double requestedFontSize, double buttonSize) =>
        Math.Min(Math.Max(8D, requestedFontSize), Math.Max(8D, buttonSize));
}
