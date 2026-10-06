namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    private static PdfResolvedTextStyle ResolveStyle(PdfTextEditOptions options, PdfRegionText? detected,
        byte[]? pdf = null, PdfLoadOptions? readOptions = null, int pageNumber = 0) {
        PdfSourceTextFont? sourceFont = null;
        if (!options.Font.HasValue && detected?.Spans.Count > 0 && pdf is not null && pageNumber > 0 &&
            !string.Equals(detected.SourceFont, detected.SuggestedFont.ToBaseFontName(), StringComparison.Ordinal)) {
            PdfTextSpan dominant = detected.Spans.First(span => string.Equals(span.BaseFont, detected.SourceFont, StringComparison.Ordinal));
            sourceFont = PdfReadDocument.Open(pdf, readOptions).Pages[pageNumber - 1].GetSourceTextFont(dominant);
        }
        return new PdfResolvedTextStyle(
            options.Font ?? detected?.SuggestedFont ?? PdfStandardFont.Helvetica,
            options.FontSize ?? detected?.FontSize ?? 12D,
            options.Color ?? detected?.Color ?? PdfColor.Black,
            options.RotationDegrees ?? detected?.RotationDegrees ?? 0D,
            detected?.Spans.Count > 0 && detected.Spans.All(static span => span.TextRenderingMode == 3) ? 3 : 0,
            sourceFont);
    }

    private static PdfResolvedTextStyle FitStyleToRegion(
        PdfResolvedTextStyle style,
        string text,
        PdfPageRegion region,
        PdfTextEditOptions options,
        out string? warning) {
        double radians = style.RotationDegrees * Math.PI / 180D;
        double availableWidth = Math.Abs(Math.Cos(radians)) * region.Width + Math.Abs(Math.Sin(radians)) * region.Height;
        return FitStyleToBaselineExtent(style, text, availableWidth, options, out warning);
    }

    private static PdfResolvedTextStyle FitStyleToBaselineExtent(
        PdfResolvedTextStyle style, string text, double availableWidth,
        PdfTextEditOptions options, out string? warning) {
        warning = null;
        if (options.RegionWidthPolicy == PdfTextRegionWidthPolicy.PreserveFontSize || text.Length == 0) return style;
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        double widestLine = lines.Max(line => style.Measure(line));
        if (widestLine <= availableWidth + 0.01D) return style;
        if (options.RegionWidthPolicy == PdfTextRegionWidthPolicy.RejectOverflow) {
            throw new NotSupportedException("The replacement text exceeds the selected region's baseline extent under the RejectOverflow width policy.");
        }
        double fittedSize = style.FontSize * availableWidth / widestLine;
        if (fittedSize + 0.001D < options.MinimumFontSize) {
            throw new NotSupportedException("The replacement text cannot fit the selected region without reducing the font below MinimumFontSize.");
        }
        warning = "The replacement font size was reduced from " + style.FontSize.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " to " + fittedSize.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " points to fit the selected region.";
        return style.WithFontSize(fittedSize);
    }

    private static string[] BuildSubstitutionWarnings(PdfRegionText detected, PdfResolvedTextStyle targetStyle) {
        if (targetStyle.SourceFont is not null) return Array.Empty<string>();
        PdfStandardFont targetFont = targetStyle.Font;
        string source = StripSubsetPrefix(detected.SourceFont);
        if (source.Length == 0 || string.Equals(source, targetFont.ToBaseFontName(), StringComparison.OrdinalIgnoreCase)) return Array.Empty<string>();
        return new[] { "The source font '" + source + "' was substituted; replacement text uses '" + targetFont.ToBaseFontName() + "'. Metrics and letterforms can differ." };
    }

    private readonly struct PdfResolvedTextStyle : IEquatable<PdfResolvedTextStyle> {
        internal PdfResolvedTextStyle(PdfStandardFont font, double fontSize, PdfColor color, double rotationDegrees,
            int textRenderingMode, PdfSourceTextFont? sourceFont = null) {
            Font = font; FontSize = fontSize; Color = color; RotationDegrees = rotationDegrees;
            TextRenderingMode = textRenderingMode; SourceFont = sourceFont;
        }
        internal PdfStandardFont Font { get; }
        internal double FontSize { get; }
        internal PdfColor Color { get; }
        internal double RotationDegrees { get; }
        internal int TextRenderingMode { get; }
        internal PdfSourceTextFont? SourceFont { get; }
        internal double Measure(string text) => SourceFont?.Measure(text, FontSize) ?? PdfWriter.EstimateSimpleTextWidth(text, Font, FontSize);
        internal PdfResolvedTextStyle WithFontSize(double fontSize) => new PdfResolvedTextStyle(Font, fontSize, Color, RotationDegrees, TextRenderingMode, SourceFont);
        public bool Equals(PdfResolvedTextStyle other) => Font == other.Font && FontSize.Equals(other.FontSize) && Color.Equals(other.Color) &&
            RotationDegrees.Equals(other.RotationDegrees) && TextRenderingMode == other.TextRenderingMode &&
            string.Equals(SourceFont?.ResourceName, other.SourceFont?.ResourceName, StringComparison.Ordinal) &&
            string.Equals(SourceFont?.BaseFont, other.SourceFont?.BaseFont, StringComparison.Ordinal);
        public override bool Equals(object? obj) => obj is PdfResolvedTextStyle other && Equals(other);
        public override int GetHashCode() {
            unchecked {
                int hash = (int)Font; hash = (hash * 397) ^ FontSize.GetHashCode(); hash = (hash * 397) ^ Color.GetHashCode();
                hash = (hash * 397) ^ RotationDegrees.GetHashCode(); hash = (hash * 397) ^ TextRenderingMode;
                hash = (hash * 397) ^ (SourceFont?.ResourceName.GetHashCode() ?? 0);
                return (hash * 397) ^ (SourceFont?.BaseFont.GetHashCode() ?? 0);
            }
        }
    }
}
