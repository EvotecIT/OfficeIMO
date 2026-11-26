using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // The SDK side of run formatting: what the members in WordParagraph.RunProperties.cs do once the Open XML SDK holds
    // the document, after OpenXmlDocument or for packages the owned layer doesn't open.
    public partial class WordParagraph {
        internal RunProperties? ScopedRunProperties => _stdRun != null
            ? _stdRun.SdtContentRun?.Descendants<Run>().FirstOrDefault()?.GetFirstChild<RunProperties>()
            : IsHyperLink ? Hyperlink?._runProperties : _runProperties;
        private static bool IsOnOffPropertyEnabled(OnOffType? property) {
            return property != null && (property.Val == null || property.Val.Value);
        }

        private OfficeIMO.OfficeColor? SdkColor => WordRunColor.Read(ScopedRunProperties?.Color);

        private bool SdkBold {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.Bold);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                runProperties.Bold = new Bold { Val = value };
                runProperties.BoldComplexScript = new BoldComplexScript { Val = value };
            }
        }

        private bool SdkItalic {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.Italic);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                runProperties.Italic = new Italic { Val = value };
                runProperties.ItalicComplexScript = new ItalicComplexScript { Val = value };
            }
        }

        private WordUnderlineStyle? SdkUnderline {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.Underline != null) {
                    return runProperties.Underline.Val?.Value.ToOfficeEnum();
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    if (runProperties.Underline == null) {
                        runProperties.Underline = new Underline();
                    }

                    runProperties.Underline.Val = value.Value.ToOpenXml();
                } else {
                    if (runProperties.Underline != null) runProperties.Underline.Remove();
                }
            }
        }

        private bool SdkDoNotCheckSpellingOrGrammar {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.NoProof);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != true) {
                    if (runProperties.NoProof != null) runProperties.NoProof.Remove();
                } else {
                    runProperties.NoProof = new NoProof();
                }
            }
        }

        private OfficeLength? SdkSpacing {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.Spacing != null) {
                    return runProperties.Spacing.Val?.Value is int nativeSpacing
                        ? OfficeLength.Emus((long)nativeSpacing * 635) : (OfficeLength?)null;
                }
                return null;
            }
            set {
                int? nativeSpacing = value.HasValue ? checked((int)value.Value.ToTwips()) : (int?)null;
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    Spacing spacing = new Spacing();
                    spacing.Val = nativeSpacing;
                    runProperties.Spacing = spacing;
                } else {
                    if (runProperties.Spacing != null) runProperties.Spacing.Remove();
                }
            }
        }

        private bool SdkStrikethrough {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.Strike);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != true) {
                    if (runProperties.Strike != null) runProperties.Strike.Remove();
                } else {
                    runProperties.Strike = new Strike();
                }
            }
        }

        private bool SdkDoubleStrikethrough {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.DoubleStrike);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != true) {
                    if (runProperties.DoubleStrike != null) runProperties.DoubleStrike.Remove();
                } else {
                    runProperties.DoubleStrike = new DoubleStrike();
                }
            }
        }

        private bool SdkOutline {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.Outline);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != true) {
                    if (runProperties.Outline != null) runProperties.Outline.Remove();
                } else {
                    runProperties.Outline = new Outline();
                }
            }
        }

        private bool SdkShadow {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.Shadow);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != true) {
                    if (runProperties.Shadow != null) runProperties.Shadow.Remove();
                } else {
                    runProperties.Shadow = new Shadow();
                }
            }
        }

        private bool SdkEmboss {
            get {
                var runProperties = ScopedRunProperties;
                return IsOnOffPropertyEnabled(runProperties?.Emboss);
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != true) {
                    if (runProperties.Emboss != null) runProperties.Emboss.Remove();
                } else {
                    runProperties.Emboss = new Emboss();
                }
            }
        }

        private double? SdkNativeFontSizePoints {
            get {
                var runProperties = ScopedRunProperties;
                string? value = runProperties?.FontSize?.Val?.Value;
                return int.TryParse(value, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out int halfPoints)
                    ? halfPoints / 2D
                    : (double?)null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    if (double.IsNaN(value.Value) || double.IsInfinity(value.Value) ||
                        value.Value < 0D || value.Value > int.MaxValue / 2D) {
                        throw new ArgumentOutOfRangeException(nameof(value));
                    }
                    int halfPoints = checked((int)Math.Round(value.Value * 2D, MidpointRounding.AwayFromZero));
                    runProperties.FontSize = new FontSize {
                        Val = halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    };
                } else {
                    if (runProperties.FontSize != null) runProperties.FontSize.Remove();
                }
                UpdatePageNumberFormatting(runProperties.FontSize);
            }
        }

        private string SdkRunColorValue {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.Color != null) {
                    return runProperties.Color.Val?.Value ?? "";
                }
                return "";
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != "") {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color();
                    color.Val = Helpers.NormalizeOpenXmlColor(value);
                    runProperties.Color = color;
                } else {
                    if (runProperties.Color != null) runProperties.Color.Remove();
                }
                UpdatePageNumberFormatting(runProperties.Color);
            }
        }

        private WordThemeColor? SdkThemeColor {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.Color != null && runProperties.Color.ThemeColor != null) {
                    return runProperties.Color.ThemeColor?.Value.ToOfficeEnum();
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color {
                        ThemeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value.ToOpenXml()
                        }
                    };
                    // w:val is required; it caches the theme colour for applications that ignore the theme.
                    color.Val = WordRunFormatting.ResolveColorHex(color, _document) ?? "auto";
                    runProperties.Color = color;
                } else {
                    if (runProperties.Color != null) runProperties.Color.Remove();
                }
            }
        }

        private WordHighlightColor? SdkHighlight {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.Highlight != null) {
                    return runProperties.Highlight.Val?.Value.ToOfficeEnum();
                }
                return null;
            }
            set {
                RunProperties? existingRunProperties = ScopedRunProperties;
                if (!value.HasValue) {
                    existingRunProperties?.Highlight?.Remove();
                    return;
                }

                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                var highlight = new Highlight {
                    Val = value.Value.ToOpenXml()
                };
                runProperties.Highlight = highlight;
            }
        }

        private WordCapsStyle SdkCapsStyle {
            get {
                var runProperties = ScopedRunProperties;
                if (IsOnOffPropertyEnabled(runProperties?.Caps)) {
                    return WordCapsStyle.Caps;
                } else if (IsOnOffPropertyEnabled(runProperties?.SmallCaps)) {
                    return WordCapsStyle.SmallCaps;
                } else {
                    return WordCapsStyle.None;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value == WordCapsStyle.None) {
                    runProperties.Caps = null;
                    runProperties.SmallCaps = null;
                } else if (value == WordCapsStyle.Caps) {
                    runProperties.Caps = new Caps();
                    runProperties.SmallCaps = null;
                } else if (value == WordCapsStyle.SmallCaps) {
                    runProperties.SmallCaps = new SmallCaps();
                    runProperties.Caps = null;
                }
            }
        }

        private string? SdkFontFamily {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.Ascii;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }

                if (runProperties.RunFonts == null) {
                    runProperties.RunFonts = new RunFonts { };
                }

                if (string.IsNullOrEmpty(value)) {
                    runProperties.RunFonts.Ascii = null;
                } else {
                    runProperties.RunFonts.Ascii = value;
                    // we set the same font for HighAnsi as well, because in 90% cases it's required for special characters
                    // and it should be the same
                    runProperties.RunFonts.HighAnsi = value;
                    runProperties.RunFonts.ComplexScript = value;
                    runProperties.RunFonts.EastAsia = value;
                }
                UpdatePageNumberFormatting(runProperties.RunFonts);
            }
        }

        private string? SdkFontFamilyHighAnsi {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.HighAnsi;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (runProperties.RunFonts == null) {
                    runProperties.RunFonts = new RunFonts { };
                }

                if (string.IsNullOrEmpty(value)) {
                    runProperties.RunFonts.HighAnsi = null;
                } else {
                    runProperties.RunFonts.HighAnsi = value;
                }
            }
        }

        private string? SdkFontFamilyEastAsia {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.EastAsia;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (runProperties.RunFonts == null) {
                    runProperties.RunFonts = new RunFonts { };
                }

                if (string.IsNullOrEmpty(value)) {
                    runProperties.RunFonts.EastAsia = null;
                } else {
                    runProperties.RunFonts.EastAsia = value;
                }
            }
        }

        private string? SdkFontFamilyComplexScript {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.ComplexScript;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (runProperties.RunFonts == null) {
                    runProperties.RunFonts = new RunFonts { };
                }

                if (string.IsNullOrEmpty(value)) {
                    runProperties.RunFonts.ComplexScript = null;
                } else {
                    runProperties.RunFonts.ComplexScript = value;
                }
            }
        }

        private string? SdkLanguage {
            get {
                var runProperties = ScopedRunProperties;
                return runProperties?.Languages?.Val;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }

                if (string.IsNullOrWhiteSpace(value)) {
                    runProperties.Languages?.Remove();
                } else {
                    var language = value!.Trim();
                    var languages = runProperties.Languages ?? new Languages();
                    runProperties.Languages = languages;
                    languages.Val = language;
                    languages.EastAsia = language;
                    languages.Bidi = language;
                }
            }
        }

        private WordCharacterStyles? SdkCharacterStyle {
            get {
                var runProperties = ScopedRunProperties;
                if (runProperties != null && runProperties.RunStyle != null) {
                    var styleId = runProperties.RunStyle.Val;
                    if (!string.IsNullOrEmpty(styleId)) {
                        return WordCharacterStyle.GetStyle(styleId!);
                    }
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }

                if (value == null) {
                    if (runProperties.RunStyle != null) runProperties.RunStyle.Remove();
                } else {
                    if (runProperties.RunStyle == null) runProperties.RunStyle = new RunStyle();
                    string styleId = WordCharacterStyle.ToStringStyle(value.Value);
                    _document?.EnsureStyleDefinition(styleId);
                    runProperties.RunStyle.Val = styleId;
                }
            }
        }

        private string? SdkCharacterStyleId {
            get {
                var runProperties = ScopedRunProperties;
                return runProperties?.RunStyle?.Val;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }

                if (string.IsNullOrEmpty(value)) {
                    if (runProperties.RunStyle != null) runProperties.RunStyle.Remove();
                } else {
                    if (runProperties.RunStyle == null) runProperties.RunStyle = new RunStyle();
                    _document?.EnsureStyleDefinition(value!);
                    runProperties.RunStyle.Val = value;
                }
            }
        }
    }
}
