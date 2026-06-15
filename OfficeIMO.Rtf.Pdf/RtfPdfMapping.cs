using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class RtfPdfMapping {
    internal static PdfCore.PdfAlign ToPdfAlign(RtfTextAlignment alignment) {
        switch (alignment) {
            case RtfTextAlignment.Center:
                return PdfCore.PdfAlign.Center;
            case RtfTextAlignment.Right:
                return PdfCore.PdfAlign.Right;
            case RtfTextAlignment.Justify:
                return PdfCore.PdfAlign.Justify;
            default:
                return PdfCore.PdfAlign.Left;
        }
    }

    internal static PdfCore.PdfTextBaseline ToPdfBaseline(RtfVerticalPosition position) {
        switch (position) {
            case RtfVerticalPosition.Superscript:
                return PdfCore.PdfTextBaseline.Superscript;
            case RtfVerticalPosition.Subscript:
                return PdfCore.PdfTextBaseline.Subscript;
            default:
                return PdfCore.PdfTextBaseline.Normal;
        }
    }

    internal static PdfCore.PdfColor? ToPdfColor(RtfDocument document, int? oneBasedColorIndex) {
        if (!oneBasedColorIndex.HasValue || oneBasedColorIndex.Value <= 0) {
            return null;
        }

        int index = oneBasedColorIndex.Value - 1;
        if (index < 0 || index >= document.Colors.Count) {
            return null;
        }

        RtfColor color = document.Colors[index];
        return PdfCore.PdfColor.FromRgb(color.Red, color.Green, color.Blue);
    }

    internal static PdfCore.PdfStandardFont? ToPdfFont(RtfDocument document, int? fontId, bool bold, bool italic) {
        if (!fontId.HasValue) {
            return null;
        }

        RtfFont? font = document.Fonts.FirstOrDefault(item => item.Id == fontId.Value);
        if (font == null) {
            return null;
        }

        return PdfCore.PdfStandardFontMapper.TryMapFontFamily(font.Name, bold, italic, out PdfCore.PdfStandardFont mapped)
            ? mapped
            : (PdfCore.PdfStandardFont?)null;
    }

    internal static double TwipsToPoints(int twips) => twips / 20D;
}
