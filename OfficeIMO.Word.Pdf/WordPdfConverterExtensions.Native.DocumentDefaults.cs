using System.Globalization;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeDocumentDefaults(string? FontFamily, string? Language, double FontSize, double ParagraphLineHeight, double ParagraphSpacingBefore, bool ParagraphSpacingBeforeDeclared, double ParagraphSpacingAfter, bool ParagraphSpacingAfterDeclared, bool ParagraphWidowControl, double? DefaultTabStopWidth) {
            public NativeNoteNumbering? NoteNumbering { get; init; }
            public NativeBookmarkTargets? BookmarkTargets { get; init; }
            public WordSection? NoteContentSection { get; init; }
            public string? NoteBodyLabel { get; init; }
            public NativeLatinFontFamilies FontFamilies { get; init; }
            public NativeComplexScriptDefaults ComplexScript { get; init; }
            public NativeTextSpacing TextSpacing { get; init; }
            public NativeCharacterStyleDefaults RunStyle { get; init; }
            public bool OpenTypeFeaturesEnabled { get; init; }
            public NativeLineSpacing LineSpacing { get; init; }
            public NativeParagraphSpacing ParagraphSpacing { get; init; }
            public OfficeIMO.Pdf.PdfRunningContentContext? RunningContentContext { get; init; }
            // An inherited story uses the selecting section's page and margin geometry.
            public WordSection? RunningContentSection { get; init; }
            // Computed once by the read operation with its effective body styles and fonts.
            // Immutable per-part visibility/permission inventory captured by this read operation.
            public IReadOnlyDictionary<OfficeXmlNode, NativeRunningTextBoxStoryPeers>? RunningTextBoxStoryPeers { get; init; }
            public bool RunningStoryHasDisallowedTextBoxPeers { get; init; }
            // A missing docDefaults element follows Word's 12pt application fallback.
            public static NativeDocumentDefaults WordDefault { get; } = new(null, null, 12D, NativeDefaultParagraphLineHeight, 0D, false, NativeDefaultParagraphSpacingAfter, false, true, null) { OpenTypeFeaturesEnabled = true };
        }

        private static NativeDocumentDefaults GetNativeDocumentDefaults(
            WordDocument? document,
            NativeFontMap? nativeFontMap = null) {
            OfficeXmlNode defaults = document == null ? default : WordStyleCatalog.For(document).DocumentDefaultsNode();
            double? defaultTabStopWidth = GetNativeDefaultTabStopWidth(document);
            if (defaults.IsNone) {
                return new NativeDocumentDefaults(
                    NativeDocumentDefaults.WordDefault.FontFamily,
                    GetNativeDocumentLanguage(document, default),
                    NativeDocumentDefaults.WordDefault.FontSize,
                    NativeDocumentDefaults.WordDefault.ParagraphLineHeight,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingBefore,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingBeforeDeclared,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingAfter,
                    NativeDocumentDefaults.WordDefault.ParagraphSpacingAfterDeclared,
                    NativeDocumentDefaults.WordDefault.ParagraphWidowControl,
                    defaultTabStopWidth) { OpenTypeFeaturesEnabled = AreNativeOpenTypeFeaturesEnabled(document) };
            }

            NativeLatinFontFamilies fontFamilies = GetNativeRunFontFamilies(document, defaults.WChild("rPrDefault").WChild("rPr").WChild("rFonts"));
            string? fontFamily = fontFamilies.Primary;
            string? language = GetNativeDocumentLanguage(document, defaults);
            // Word treats an existing docDefaults with no declared size as 10pt.
            // OfficeIMO-created documents explicitly declare 11pt and retain it.
            double fontSize = GetNativeDefaultFontSize(defaults) ?? 10D;
            OfficeXmlNode spacing = defaults.WChild("pPrDefault").WChild("pPr").WChild("spacing");

            double lineHeight = GetNativeDefaultParagraphLineHeight(
                spacing,
                fontFamily,
                nativeFontMap) ?? NativeDocumentDefaults.WordDefault.ParagraphLineHeight;
            bool spacingBeforeDeclared = IsNativeSpacingBeforeDeclared(spacing);
            bool spacingAfterDeclared = IsNativeSpacingAfterDeclared(spacing);
            double spacingBefore = GetNativeSpacingBeforePoints(spacing, fontSize, lineHeight) ?? NativeDocumentDefaults.WordDefault.ParagraphSpacingBefore;
            double spacingAfter = GetNativeSpacingAfterPoints(spacing, fontSize, lineHeight) ?? NativeDocumentDefaults.WordDefault.ParagraphSpacingAfter;
            bool widowControl = ReadNativeOnOff(defaults.WChild("pPrDefault").WChild("pPr").WChild("widowControl")) ?? NativeDocumentDefaults.WordDefault.ParagraphWidowControl;
            OfficeXmlNode defaultRuns = defaults.WChild("rPrDefault").WChild("rPr");
            return new NativeDocumentDefaults(fontFamily, language, fontSize, lineHeight, spacingBefore, spacingBeforeDeclared, spacingAfter, spacingAfterDeclared, widowControl, defaultTabStopWidth) {
                FontFamilies = fontFamilies,
                ComplexScript = default(NativeComplexScriptDefaults).Merge(defaultRuns),
                TextSpacing = default(NativeTextSpacing).Merge(defaultRuns),
                RunStyle = new NativeCharacterStyleDefaults(null, null,
                    ReadNativeOnOff(defaultRuns.WChild("b")),
                    ReadNativeOnOff(defaultRuns.WChild("i")),
                    MapNativeUnderlineStyle(defaultRuns.WChild("u")),
                    MapNativeStrikeStyle(defaultRuns),
                    null,
                    ReadNativeOnOff(defaultRuns.WChild("caps")),
                    defaultRuns.WChild("vertAlign").WVal(),
                    WordConverterNodes.ResolveColorHex(defaultRuns.WChild("color"), document),
                    null) {
                    SmallCaps = ReadNativeOnOff(defaultRuns.WChild("smallCaps"))
                },
                OpenTypeFeaturesEnabled = AreNativeOpenTypeFeaturesEnabled(document),
                LineSpacing = ReadNativeLineSpacing(spacing, defaults.WChild("pPrDefault").WChild("pPr").WChild("textAlignment")),
                ParagraphSpacing = default(NativeParagraphSpacing).Merge(defaults.WChild("pPrDefault").WChild("pPr"))
            };
        }

        private static string? GetNativeDocumentLanguage(WordDocument? document, OfficeXmlNode defaults) {
            string? language = defaults.WChild("rPrDefault").WChild("rPr").WChild("lang").WVal();
            if (!string.IsNullOrWhiteSpace(language)) {
                return language;
            }

            return document.SettingsNode().WChild("themeFontLang").WVal();
        }

        private static global::OfficeIMO.OfficeColor? ResolveNativeInheritedRunColor(
            NativeCharacterStyleDefaults characterStyle,
            NativeParagraphStyleDefaults paragraphStyle,
            NativeTableRunStyleDefaults tableStyle,
            NativeDocumentDefaults documentDefaults) {
            // A declared automatic color stops inheritance just like an explicit RGB color.
            if (characterStyle.ColorHex != null) return ParseNativeColor(characterStyle.ColorHex);
            if (paragraphStyle.ColorHex != null) return ParseNativeColor(paragraphStyle.ColorHex);
            if (tableStyle.Color.HasValue) return tableStyle.Color;
            return ParseNativeColor(tableStyle.ColorHex ?? documentDefaults.RunStyle.ColorHex);
        }

        private static double? GetNativeDefaultTabStopWidth(WordDocument? document) {
            int defaultTabStopTwips = document?.Settings?.DefaultTabStop ?? 0;
            if (defaultTabStopTwips <= 0 || defaultTabStopTwips == 720) {
                return null;
            }

            return ConvertNativeTwipsToPoints(defaultTabStopTwips);
        }

        private static double? GetNativeDefaultFontSize(OfficeXmlNode defaults) {
            OfficeXmlNode runProperties = defaults.WChild("rPrDefault").WChild("rPr");
            return ParseNativeDefaultFontSize(runProperties.WChild("sz").WVal());
        }

        private static double? ParseNativeDefaultFontSize(string? value) {
            if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) ||
                halfPoints <= 0D ||
                double.IsNaN(halfPoints) ||
                double.IsInfinity(halfPoints)) {
                return null;
            }

            return halfPoints / 2D;
        }

        private static double? GetNativeDefaultParagraphLineHeight(
            OfficeXmlNode spacing,
            string? fontFamily,
            NativeFontMap? nativeFontMap = null) {
            if (ReadNativeWordLineSpacing(spacing)?.MultipleValue is not double multiple || multiple <= 0D) return null;
            return Math.Max(0.01D, ResolveNativeWordSingleLineHeight(nativeFontMap, fontFamily) * multiple);
        }

        private static bool IsNativeSpacingBeforeDeclared(OfficeXmlNode spacing) =>
            spacing.WAttr("before") != null || spacing.WAttr("beforeLines") != null || spacing.WAttr("beforeAutospacing") != null;

        private static bool IsNativeSpacingAfterDeclared(OfficeXmlNode spacing) =>
            spacing.WAttr("after") != null || spacing.WAttr("afterLines") != null || spacing.WAttr("afterAutospacing") != null;

        private static double? GetNativeSpacingBeforePoints(OfficeXmlNode spacing, double fontSize, double lineHeight) {
            if (spacing.IsNone || IsNativeOnOffTrue(spacing.WAttr("beforeAutospacing"))) {
                return null;
            }

            return ConvertNativeLineHundredthsToPoints(spacing.WAttr("beforeLines"), fontSize, lineHeight) ??
                ConvertNativeTwipsToPoints(spacing.WAttr("before"));
        }

        private static double? GetNativeSpacingAfterPoints(OfficeXmlNode spacing, double fontSize, double lineHeight) {
            if (spacing.IsNone || IsNativeOnOffTrue(spacing.WAttr("afterAutospacing"))) {
                return null;
            }

            return ConvertNativeLineHundredthsToPoints(spacing.WAttr("afterLines"), fontSize, lineHeight) ??
                ConvertNativeTwipsToPoints(spacing.WAttr("after"));
        }

        private static double? ConvertNativeLineHundredthsToPoints(string? value, double fontSize, double lineHeight) {
            if (value == null ||
                !int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int lineHundredths) ||
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


        private static bool IsNativeOnOffTrue(string? value) =>
            WordXml.ParseOnOff(value) == true;
    }
}
