using System.Globalization;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeDocumentDefaults(string? FontFamily, string? Language, double FontSize, double ParagraphLineHeight, double ParagraphSpacingBefore, bool ParagraphSpacingBeforeDeclared, double ParagraphSpacingAfter, bool ParagraphSpacingAfterDeclared, bool ParagraphWidowControl, double? DefaultTabStopWidth) {
            public static NativeDocumentDefaults WordDefault { get; } = new(null, null, 11D, NativeDefaultParagraphLineHeight, 0D, false, NativeDefaultParagraphSpacingAfter, false, true, null);
        }

        private static NativeDocumentDefaults GetNativeDocumentDefaults(WordDocument? document) {
            W.DocDefaults? defaults = document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
            double? defaultTabStopWidth = GetNativeDefaultTabStopWidth(document);
            if (defaults == null) {
                return new NativeDocumentDefaults(
                    NativeDocumentDefaults.WordDefault.FontFamily,
                    GetNativeDocumentLanguage(document, null),
                    NativeDocumentDefaults.WordDefault.FontSize,
                    NativeDocumentDefaults.WordDefault.ParagraphLineHeight,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingBefore,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingBeforeDeclared,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingAfter,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingAfterDeclared,
                    NativeDocumentDefaults.WordDefault.ParagraphWidowControl,
                    defaultTabStopWidth);
            }

            string? fontFamily = ResolveNativeRunFontsFamily(document, defaults
                .GetFirstChild<W.RunPropertiesDefault>()?
                .GetFirstChild<W.RunPropertiesBaseStyle>()?
                .GetFirstChild<W.RunFonts>());
            string? language = GetNativeDocumentLanguage(document, defaults);
            double fontSize = GetNativeDefaultFontSize(defaults) ?? NativeDocumentDefaults.WordDefault.FontSize;
            W.SpacingBetweenLines? spacing = defaults
                .GetFirstChild<W.ParagraphPropertiesDefault>()?
                .GetFirstChild<W.ParagraphPropertiesBaseStyle>()?
                .GetFirstChild<W.SpacingBetweenLines>();

            double lineHeight = GetNativeDefaultParagraphLineHeight(spacing) ?? NativeDocumentDefaults.WordDefault.ParagraphLineHeight;
            bool spacingBeforeDeclared = IsNativeSpacingBeforeDeclared(spacing);
            bool spacingAfterDeclared = IsNativeSpacingAfterDeclared(spacing);
            double spacingBefore = GetNativeSpacingBeforePoints(spacing, fontSize, lineHeight) ?? NativeDocumentDefaults.WordDefault.ParagraphSpacingBefore;
            double spacingAfter = GetNativeSpacingAfterPoints(spacing, fontSize, lineHeight) ?? NativeDocumentDefaults.WordDefault.ParagraphSpacingAfter;
            bool widowControl = ReadNativeOnOff(defaults
                .GetFirstChild<W.ParagraphPropertiesDefault>()?
                .GetFirstChild<W.ParagraphPropertiesBaseStyle>()?
                .GetFirstChild<W.WidowControl>()) ?? NativeDocumentDefaults.WordDefault.ParagraphWidowControl;
            return new NativeDocumentDefaults(fontFamily, language, fontSize, lineHeight, spacingBefore, spacingBeforeDeclared, spacingAfter, spacingAfterDeclared, widowControl, defaultTabStopWidth);
        }

        private static string? GetNativeDocumentLanguage(WordDocument? document, W.DocDefaults? defaults) {
            string? language = defaults?
                .GetFirstChild<W.RunPropertiesDefault>()?
                .GetFirstChild<W.RunPropertiesBaseStyle>()?
                .GetFirstChild<W.Languages>()?
                .Val?
                .Value;
            if (!string.IsNullOrWhiteSpace(language)) {
                return language;
            }

            return document?._wordprocessingDocument?
                .MainDocumentPart?
                .DocumentSettingsPart?
                .Settings?
                .GetFirstChild<W.ThemeFontLanguages>()?
                .Val?
                .Value;
        }

        private static double? GetNativeDefaultTabStopWidth(WordDocument? document) {
            int defaultTabStopTwips = document?.Settings?.DefaultTabStop ?? 0;
            if (defaultTabStopTwips <= 0 || defaultTabStopTwips == 720) {
                return null;
            }

            return ConvertNativeTwipsToPoints(defaultTabStopTwips);
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

        private static bool IsNativeSpacingBeforeDeclared(W.SpacingBetweenLines? spacing) =>
            spacing?.Before != null || spacing?.BeforeLines != null || spacing?.BeforeAutoSpacing != null;

        private static bool IsNativeSpacingAfterDeclared(W.SpacingBetweenLines? spacing) =>
            spacing?.After != null || spacing?.AfterLines != null || spacing?.AfterAutoSpacing != null;

        private static double? GetNativeSpacingBeforePoints(W.SpacingBetweenLines? spacing, double fontSize, double lineHeight) {
            if (spacing == null || IsNativeOnOffTrue(spacing.BeforeAutoSpacing)) {
                return null;
            }

            return ConvertNativeLineHundredthsToPoints(spacing.BeforeLines, fontSize, lineHeight) ??
                ConvertNativeTwipsToPoints(spacing.Before?.Value);
        }

        private static double? GetNativeSpacingAfterPoints(W.SpacingBetweenLines? spacing, double fontSize, double lineHeight) {
            if (spacing == null || IsNativeOnOffTrue(spacing.AfterAutoSpacing)) {
                return null;
            }

            return ConvertNativeLineHundredthsToPoints(spacing.AfterLines, fontSize, lineHeight) ??
                ConvertNativeTwipsToPoints(spacing.After?.Value);
        }

        private static double? ConvertNativeLineHundredthsToPoints(Int32Value? value, double fontSize, double lineHeight) {
            if (value == null ||
                !int.TryParse(value.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int lineHundredths) ||
                lineHundredths < 0 ||
                fontSize <= 0D ||
                lineHeight <= 0D ||
                double.IsNaN(fontSize) ||
                double.IsInfinity(fontSize) ||
                double.IsNaN(lineHeight) ||
                double.IsInfinity(lineHeight)) {
                return null;
            }

            return (lineHundredths / 100D) * fontSize * lineHeight;
        }

        private static bool IsNativeOnOffTrue(OnOffValue? value) =>
            value?.Value == true;
    }
}
