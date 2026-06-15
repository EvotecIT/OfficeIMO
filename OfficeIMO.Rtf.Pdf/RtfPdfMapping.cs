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

    internal static PdfCore.PdfPageNumberStyle ToPdfPageNumberStyle(RtfPageNumberFormat format) {
        switch (format) {
            case RtfPageNumberFormat.UpperRoman:
                return PdfCore.PdfPageNumberStyle.UpperRoman;
            case RtfPageNumberFormat.LowerRoman:
                return PdfCore.PdfPageNumberStyle.LowerRoman;
            case RtfPageNumberFormat.UpperLetter:
                return PdfCore.PdfPageNumberStyle.UpperLetter;
            case RtfPageNumberFormat.LowerLetter:
                return PdfCore.PdfPageNumberStyle.LowerLetter;
            default:
                return PdfCore.PdfPageNumberStyle.Arabic;
        }
    }

    internal static PdfCore.PdfPageBorder? ToPdfPageBorder(RtfDocument document, RtfPageBorders borders) {
        RtfPageBorder? source = GetFirstRenderablePageBorder(borders);
        if (source == null) {
            return null;
        }

        PdfCore.PdfPageBorder border = new PdfCore.PdfPageBorder {
            DashStyle = ToPdfDashStyle(source.Style),
            Width = source.Width.HasValue && source.Width.Value > 0 ? source.Width.Value / 8D : 1D,
            Inset = source.Space.HasValue && source.Space.Value >= 0 ? source.Space.Value : 36D
        };

        PdfCore.PdfColor? color = ToPdfColor(document, source.ColorIndex);
        if (color.HasValue) {
            border.Color = color.Value;
        }

        return border;
    }

    private static RtfPageBorder? GetFirstRenderablePageBorder(RtfPageBorders borders) {
        if (IsRenderablePageBorder(borders.Top)) return borders.Top;
        if (IsRenderablePageBorder(borders.Bottom)) return borders.Bottom;
        if (IsRenderablePageBorder(borders.Left)) return borders.Left;
        if (IsRenderablePageBorder(borders.Right)) return borders.Right;
        return null;
    }

    private static bool IsRenderablePageBorder(RtfPageBorder border) => border.Style != RtfPageBorderStyle.None;

    private static OfficeIMO.Drawing.OfficeStrokeDashStyle ToPdfDashStyle(RtfPageBorderStyle style) {
        switch (style) {
            case RtfPageBorderStyle.Dashed:
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash;
            case RtfPageBorderStyle.Dotted:
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dot;
            default:
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
        }
    }

    internal static double TwipsToPoints(int twips) => twips / 20D;
}
