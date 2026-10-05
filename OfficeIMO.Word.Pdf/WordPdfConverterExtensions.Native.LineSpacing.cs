using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Retain authored units until the effective paragraph font is resolved. An
        // automatic value is a multiplier, not a height measured with the style font.
        private readonly record struct NativeLineSpacing(double? Value, W.LineSpacingRuleValues? Rule) {
            // Word ignores a rule without a value in the same declaration.
            // Inherit the complete pair so its numeric units do not change.
            public NativeLineSpacing Inherit(NativeLineSpacing inherited) =>
                Value.HasValue ? this : inherited;

            public OfficeIMO.Pdf.PdfLineSpacing? ToPdfLineSpacing(double naturalLineHeight) {
                if (!Value.HasValue) return OfficeIMO.Pdf.PdfLineSpacing.Multiple(naturalLineHeight)
                    .WithFontLineBoxBaseline(naturalLineHeight);
                if (Rule == null || Rule == W.LineSpacingRuleValues.Auto)
                    return OfficeIMO.Pdf.PdfLineSpacing.Multiple(Math.Max(0.01D, naturalLineHeight * Value.Value / 240D))
                        .WithFontLineBoxBaseline(naturalLineHeight);
                return Rule == W.LineSpacingRuleValues.AtLeast
                    ? OfficeIMO.Pdf.PdfLineSpacing.AtLeast(Value.Value / 20D, naturalLineHeight).WithFontLineBoxBaseline(naturalLineHeight)
                    // Word's exact-height baseline is four fifths of the line
                    // box in body, column and table exports, regardless of font.
                    : OfficeIMO.Pdf.PdfLineSpacing.Exactly(Value.Value / 20D)
                        .WithFixedLineBoxBaseline(Value.Value / 20D * .8D);
            }

            public double? Resolve(double fontSize, double naturalLineHeight) {
                if (!Value.HasValue) return null;
                if (Rule == null || Rule == W.LineSpacingRuleValues.Auto) {
                    return Math.Max(0.01D, naturalLineHeight * (Value.Value / 240D));
                }
                return fontSize > 0D
                    ? ResolveNativeLineSpacingHeight(Value.Value / 20D, Rule, fontSize, naturalLineHeight)
                    : null;
            }
        }

        private static NativeLineSpacing ReadNativeLineSpacing(W.SpacingBetweenLines? spacing) {
            double? value = double.TryParse(spacing?.Line?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double line) &&
                line >= 0D && !double.IsNaN(line) && !double.IsInfinity(line) ? line : null;
            return new NativeLineSpacing(value, spacing?.LineRule?.Value ?? (value.HasValue ? W.LineSpacingRuleValues.Auto : null));
        }

        private static NativeLineSpacing ResolveNativeParagraphLineSpacing(
            WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults,
            NativeDocumentDefaults documentDefaults, NativeLineSpacing tableSpacing = default) =>
            ReadNativeLineSpacing(paragraph._paragraph.ParagraphProperties?.SpacingBetweenLines)
                .Inherit(styleDefaults.LineSpacing.Inherit(tableSpacing.Inherit(documentDefaults.LineSpacing)));

        // A paragraph mark or a large run on another line must not impose a
        // paragraph-wide minimum on every visible line. Runs retain their own sizes.
        private static double ResolveNativeParagraphLayoutFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults,
            NativeParagraphStyleDefaults styleDefaults, NativeTableRunStyleDefaults tableRunStyleDefaults = default) {
            double? minimum = null;
            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) continue;
                double size = ResolveNativeTextRunStyle(run, paragraph, tableRunStyleDefaults, nativeDefaults).FontSize ?? nativeDefaults.FontSize;
                if (size > 0D) minimum = minimum.HasValue ? Math.Min(minimum.Value, size) : size;
            }
            if (minimum.HasValue) return minimum.Value;
            string? markSize = paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.FontSize>()?.Val?.Value;
            if (int.TryParse(markSize, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints) && halfPoints > 0)
                return halfPoints / 2D;
            return Math.Max(ResolveNativeParagraphEffectiveFontSize(paragraph, nativeDefaults, styleDefaults, tableRunStyleDefaults),
                tableRunStyleDefaults.FontSize ?? 0D);
        }
    }
}
