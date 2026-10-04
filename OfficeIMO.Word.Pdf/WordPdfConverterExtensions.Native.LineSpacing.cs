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
    }
}
