using System;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Manages run property settings.
    /// </summary>
    public partial class WordParagraph {

        /// <summary>
        /// Gets or sets a value indicating whether the run is bold.
        /// </summary>
        public bool Bold {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Bold != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value == true) {
                    runProperties.Bold = new Bold();
                    runProperties.BoldComplexScript = new BoldComplexScript();
                } else {
                    if (runProperties.BoldComplexScript != null) {
                        runProperties.BoldComplexScript.Remove();
                    }

                    if (runProperties.Bold != null) {
                        runProperties.Bold.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is italic.
        /// </summary>
        public bool Italic {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Italic != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != true) {
                    runProperties.Italic = null;
                } else {
                    runProperties.Italic = new Italic { };
                }
            }
        }

        /// <summary>
        /// Gets or sets the underline style for the run.
        /// </summary>
        public UnderlineValues? Underline {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Underline != null) {
                    return runProperties.Underline.Val?.Value;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    if (runProperties.Underline == null) {
                        runProperties.Underline = new Underline();
                    }

                    runProperties.Underline.Val = value;
                } else {
                    if (runProperties.Underline != null) runProperties.Underline.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether spelling and grammar checks are disabled.
        /// </summary>
        public bool DoNotCheckSpellingOrGrammar {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.NoProof != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
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

        /// <summary>
        /// Gets or sets the character spacing value in twentieths of a point.
        /// </summary>
        public int? Spacing {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Spacing != null) {
                    return runProperties.Spacing.Val?.Value;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    Spacing spacing = new Spacing();
                    spacing.Val = value;
                    runProperties.Spacing = spacing;
                } else {
                    if (runProperties.Spacing != null) runProperties.Spacing.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is struck through.
        /// </summary>
        public bool Strike {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Strike != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
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

        /// <summary>
        /// Gets or sets a value indicating whether the run has a double strikethrough.
        /// </summary>
        public bool DoubleStrike {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.DoubleStrike != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
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
        /// <summary>
        /// Gets or sets the font size in points.
        /// </summary>
        public int? FontSize {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.FontSize != null) {
                    var val = runProperties.FontSize.Val;
                    if (!string.IsNullOrEmpty(val) && int.TryParse(val, out var fontSizeInHalfPoint)) {
                        return fontSizeInHalfPoint / 2;
                    }
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    FontSize fontSize = new FontSize();
                    fontSize.Val = (value * 2).ToString();
                    runProperties.FontSize = fontSize;
                } else {
                    if (runProperties.FontSize != null) runProperties.FontSize.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the text color using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public SixLabors.ImageSharp.Color? Color {
            get {
                if (ColorHex == "") {
                    return null;
                }
                return Helpers.ParseColor(ColorHex);
            }
            set {
                if (value != null) {
                    this.ColorHex = value.Value.ToHexColor();
                }
            }
        }

        /// <summary>
        /// Gets or sets the text color as a hexadecimal string.
        /// </summary>
        public string ColorHex {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Color != null) {
                    return runProperties.Color.Val?.Value ?? "";
                }
                return "";
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != "") {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color();
                    color.Val = value.Replace("#", "").ToLowerInvariant();
                    runProperties.Color = color;
                } else {
                    if (runProperties.Color != null) runProperties.Color.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the theme color applied to the run.
        /// </summary>
        public ThemeColorValues? ThemeColor {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Color != null && runProperties.Color.ThemeColor != null) {
                    return runProperties.Color.ThemeColor?.Value;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color {
                        ThemeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value
                        }
                    };
                    runProperties.Color = color;
                } else {
                    if (runProperties.Color != null) runProperties.Color.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the highlight color applied to the run.
        /// </summary>
        public HighlightColorValues? Highlight {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Highlight != null) {
                    return runProperties.Highlight.Val?.Value;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                var highlight = new Highlight {
                    Val = value
                };
                runProperties.Highlight = highlight;
            }
        }

        /// <summary>
        /// Gets or sets the capitalization style for the run.
        /// </summary>
        public CapsStyle CapsStyle {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.Caps != null) {
                    return CapsStyle.Caps;
                } else if (runProperties != null && runProperties.SmallCaps != null) {
                    return CapsStyle.SmallCaps;
                } else {
                    return CapsStyle.None;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value == CapsStyle.None) {
                    runProperties.Caps = null;
                    runProperties.SmallCaps = null;
                } else if (value == CapsStyle.Caps) {
                    runProperties.Caps = new Caps();
                    runProperties.SmallCaps = null;
                } else if (value == CapsStyle.SmallCaps) {
                    runProperties.SmallCaps = new SmallCaps();
                    runProperties.Caps = null;
                }
            }
        }

        /// <summary>
        /// FontFamily gets and sets the FontFamily of a WordParagraph.
        /// To make sure that FontFamily works correctly on special characters
        /// we change the RunFonts.Ascii and RunFonts.HighAnsi, EastAsia and ComplexScript properties
        /// If you want to set different FontFamily for HighAnsi, EastAsia and ComplexScript
        /// please use FontFamilyHighAnsi, FontFamilyEastAsia or FontFamilyComplexScript
        /// in proper order (to overwrite given FontFamily)
        /// </summary>
        public string? FontFamily {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.Ascii;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
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
            }
        }

        /// <summary>
        /// Gets or sets the HighAnsi font family.
        /// </summary>
        public string? FontFamilyHighAnsi {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.HighAnsi;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
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

        /// <summary>
        /// Gets or sets the East Asia font family.
        /// </summary>
        public string? FontFamilyEastAsia {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.EastAsia;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
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

        /// <summary>
        /// Gets or sets the complex script font family.
        /// </summary>
        public string? FontFamilyComplexScript {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                if (runProperties != null && runProperties.RunFonts != null) {
                    return runProperties.RunFonts.ComplexScript;
                }
                return null;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
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

        /// <summary>
        /// Gets or sets the character style applied to the run.
        /// </summary>
        public WordCharacterStyles? CharacterStyle {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
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
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }

                if (value == null) {
                    if (runProperties.RunStyle != null) runProperties.RunStyle.Remove();
                } else {
                    if (runProperties.RunStyle == null) runProperties.RunStyle = new RunStyle();
                    runProperties.RunStyle.Val = WordCharacterStyle.ToStringStyle(value.Value);
                }
            }
        }

        /// <summary>
        /// Gets or sets the style identifier applied to the run.
        /// </summary>
        public string? CharacterStyleId {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink?._runProperties : _runProperties;
                return runProperties?.RunStyle?.Val;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }

                if (string.IsNullOrEmpty(value)) {
                    if (runProperties.RunStyle != null) runProperties.RunStyle.Remove();
                } else {
                    if (runProperties.RunStyle == null) runProperties.RunStyle = new RunStyle();
                    runProperties.RunStyle.Val = value;
                }
            }
        }
    }
}