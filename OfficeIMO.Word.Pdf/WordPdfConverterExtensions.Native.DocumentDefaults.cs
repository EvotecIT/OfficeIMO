using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeDocumentDefaults(double FontSize, double ParagraphLineHeight, double ParagraphSpacingAfter, bool ParagraphSpacingAfterDeclared, bool ParagraphWidowControl) {
            public static NativeDocumentDefaults WordDefault { get; } = new(11D, NativeDefaultParagraphLineHeight, NativeDefaultParagraphSpacingAfter, false, true);
        }

        private static NativeDocumentDefaults GetNativeDocumentDefaults(WordDocument? document) {
            W.DocDefaults? defaults = document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
            if (defaults == null) {
                return NativeDocumentDefaults.WordDefault;
            }

            double fontSize = GetNativeDefaultFontSize(defaults) ?? NativeDocumentDefaults.WordDefault.FontSize;
            W.SpacingBetweenLines? spacing = defaults
                .GetFirstChild<W.ParagraphPropertiesDefault>()?
                .GetFirstChild<W.ParagraphPropertiesBaseStyle>()?
                .GetFirstChild<W.SpacingBetweenLines>();

            double lineHeight = GetNativeDefaultParagraphLineHeight(spacing) ?? NativeDocumentDefaults.WordDefault.ParagraphLineHeight;
            bool spacingAfterDeclared = spacing?.After != null;
            double spacingAfter = ConvertNativeTwipsToPoints(spacing?.After?.Value) ?? NativeDocumentDefaults.WordDefault.ParagraphSpacingAfter;
            bool widowControl = ReadNativeOnOff(defaults
                .GetFirstChild<W.ParagraphPropertiesDefault>()?
                .GetFirstChild<W.ParagraphPropertiesBaseStyle>()?
                .GetFirstChild<W.WidowControl>()) ?? NativeDocumentDefaults.WordDefault.ParagraphWidowControl;
            return new NativeDocumentDefaults(fontSize, lineHeight, spacingAfter, spacingAfterDeclared, widowControl);
        }

        private static double? GetNativeDefaultFontSize(W.DocDefaults defaults) {
            string? value = defaults
                .GetFirstChild<W.RunPropertiesDefault>()?
                .GetFirstChild<W.RunPropertiesBaseStyle>()?
                .GetFirstChild<W.FontSize>()?
                .Val?
                .Value;
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) ||
                halfPoints <= 0D ||
                double.IsNaN(halfPoints) ||
                double.IsInfinity(halfPoints)) {
                return null;
            }

            return halfPoints / 2D;
        }

        private static double? GetNativeDefaultParagraphLineHeight(W.SpacingBetweenLines? spacing) {
            if (spacing?.Line == null) {
                return null;
            }

            if (!double.TryParse(spacing.Line.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double line) ||
                line <= 0D ||
                double.IsNaN(line) ||
                double.IsInfinity(line)) {
                return null;
            }

            if (spacing.LineRule?.Value == W.LineSpacingRuleValues.Auto) {
                return Math.Max(0.01D, NativeWordAutoLineSpacingHeight * (line / 240D));
            }

            return null;
        }
    }
}
