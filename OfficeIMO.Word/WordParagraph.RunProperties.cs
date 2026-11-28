using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using static OfficeIMO.Word.OpenXml.WordXmlEdits;

namespace OfficeIMO.Word {
    // Manages run property settings. Each member reads and edits the run properties through their element handle, so it
    // works the same whether OfficeIMO's own package model or the Open XML SDK holds the document.
    public partial class WordParagraph {
        /// <summary>
        /// Gets or sets a value indicating whether the run is bold. Setting false explicitly overrides inherited bold.
        /// </summary>
        public bool Bold {
            get => ReadOnOff(RunPropertiesNode.Child(W, "b"));
            set => SetRunOnOffPair("b", "bCs", value);
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is italic. Setting false explicitly overrides inherited italic.
        /// </summary>
        public bool Italic {
            get => ReadOnOff(RunPropertiesNode.Child(W, "i"));
            set => SetRunOnOffPair("i", "iCs", value);
        }

        /// <summary>
        /// Gets or sets the underline style for the run.
        /// </summary>
        public WordUnderlineStyle? Underline {
            get {
                string? token = ChildValue(RunPropertiesNode, "u");
                return token == null ? null : Token(WordUnderlineStyleTokens.TryParse(token, out WordUnderlineStyle style), style, token);
            }
            set {
                OfficeXmlNode properties = EditRunProperties();
                if (value == null) {
                    RemoveChild(properties, "u");
                    return;
                }
                // An existing underline keeps its other attributes, such as its colour.
                properties.GetOrAddChild(W, "u", WRPrView.Order).SetAttribute(W, "val", WordUnderlineStyleTokens.ToToken(value.Value));
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether spelling and grammar checks are disabled.
        /// </summary>
        public bool DoNotCheckSpellingOrGrammar {
            get => ReadOnOff(RunPropertiesNode.Child(W, "noProof"));
            set => SetRunToggle("noProof", value);
        }

        /// <summary>
        /// Gets or sets the signed physical character spacing. Null removes the local value.
        /// </summary>
        public OfficeLength? Spacing {
            get => int.TryParse(ChildValue(RunPropertiesNode, "spacing"), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out int twips) ? OfficeLength.Emus((long)twips * 635) : null;
            set {
                int? nativeSpacing = value.HasValue ? checked((int)value.Value.ToTwips()) : (int?)null;
                SetRunValue("spacing", nativeSpacing?.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is struck through.
        /// </summary>
        public bool Strikethrough {
            get => ReadOnOff(RunPropertiesNode.Child(W, "strike"));
            set => SetRunToggle("strike", value);
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run has a double strikethrough.
        /// </summary>
        public bool DoubleStrikethrough {
            get => ReadOnOff(RunPropertiesNode.Child(W, "dstrike"));
            set => SetRunToggle("dstrike", value);
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run text is outlined.
        /// </summary>
        public bool Outline {
            get => ReadOnOff(RunPropertiesNode.Child(W, "outline"));
            set => SetRunToggle("outline", value);
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run has a shadow effect.
        /// </summary>
        public bool Shadow {
            get => ReadOnOff(RunPropertiesNode.Child(W, "shadow"));
            set => SetRunToggle("shadow", value);
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is embossed.
        /// </summary>
        public bool Emboss {
            get => ReadOnOff(RunPropertiesNode.Child(W, "emboss"));
            set => SetRunToggle("emboss", value);
        }

        /// <summary>
        /// Gets or sets the font size in points with Word's native half-point precision.
        /// Values finer than half a point are rounded away from zero to the nearest representable size.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FontSize {
            get => NativeFontSizePoints is double points ? global::OfficeIMO.OfficeLength.Points(points) : null;
            set => NativeFontSizePoints = value?.ToPoints();
        }

        /// <summary>Gets or sets native point values without a whole-EMU round trip.</summary>
        internal double? NativeFontSizePoints {
            get => int.TryParse(ChildValue(RunPropertiesNode, "sz"), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out int halfPoints) ? halfPoints / 2D : null;
            set {
                string? halfPoints = null;
                if (value != null) {
                    if (double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value.Value < 0D || value.Value > int.MaxValue / 2D) {
                        throw new ArgumentOutOfRangeException(nameof(value));
                    }
                    halfPoints = checked((int)Math.Round(value.Value * 2D, MidpointRounding.AwayFromZero)).ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
                SyncPageNumberFormatting(SetRunValue("sz", halfPoints), "sz");
            }
        }

        /// <summary>
        /// Gets or sets the text color using <see cref="OfficeIMO.OfficeColor"/>.
        /// </summary>
        public OfficeIMO.OfficeColor? Color {
            get {
                OfficeXmlNode color = RunPropertiesNode.Child(W, "color");
                return color.IsNone ? null : WordRunColor.Read(color.Attribute(W, "themeColor"), color.HasAttribute(W, "themeTint"),
                    color.HasAttribute(W, "themeShade"), color.Attribute(W, "val"));
            }
            set {
                if (value != null && value.Value.IsThemeRole) {
                    SetRunColor(value.Value);
                } else {
                    RunColorValue = value?.ToRgbHex() ?? string.Empty;
                }
            }
        }

        /// <summary>
        /// Gets or sets the native run colour value (hexadecimal or <c>auto</c>); empty removes it.
        /// </summary>
        internal string RunColorValue {
            get => ChildValue(RunPropertiesNode, "color") ?? string.Empty;
            set {
                OfficeXmlNode properties = EditRunProperties();
                if (value != string.Empty) {
                    OfficeXmlNode color = ReplaceChild(properties, "color", WRPrView.Order);
                    string? normalized = Helpers.NormalizeOpenXmlColor(value);
                    if (normalized != null) color.SetAttribute(W, "val", normalized);
                } else {
                    RemoveChild(properties, "color");
                }
                SyncPageNumberFormatting(properties, "color");
            }
        }

        /// <summary>
        /// Gets or sets the theme color applied to the run.
        /// </summary>
        public WordThemeColor? ThemeColor {
            get {
                OfficeXmlNode color = RunPropertiesNode.Child(W, "color");
                string? token = color.IsNone ? null : color.Attribute(W, "themeColor");
                return token == null ? null : Token(WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordThemeColor parsed), parsed, token);
            }
            // The colour value caches the theme's colour, which the SDK path reads from the theme.
            set => SdkThemeColor = value;
        }

        /// <summary>
        /// Gets or sets the highlight color applied to the run.
        /// </summary>
        public WordHighlightColor? Highlight {
            get {
                string? token = ChildValue(RunPropertiesNode, "highlight");
                return token == null ? null : Token(WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordHighlightColor color), color, token);
            }
            set {
                // Clearing the highlight leaves a run without properties as it is.
                if (value == null) {
                    if (!RunPropertiesNode.Child(W, "highlight").IsNone) RemoveChild(EditRunProperties(), "highlight");
                    return;
                }
                ReplaceChild(EditRunProperties(), "highlight", WRPrView.Order, value.Value.ToOpenXmlToken());
            }
        }

        /// <summary>
        /// Gets or sets the capitalization style for the run.
        /// </summary>
        public WordCapsStyle CapsStyle {
            get {
                OfficeXmlNode properties = RunPropertiesNode;
                return ReadOnOff(properties.Child(W, "caps")) ? WordCapsStyle.Caps
                    : ReadOnOff(properties.Child(W, "smallCaps")) ? WordCapsStyle.SmallCaps : WordCapsStyle.None;
            }
            set {
                OfficeXmlNode properties = EditRunProperties();
                if (value == WordCapsStyle.Caps) ReplaceChild(properties, "caps", WRPrView.Order);
                else RemoveChild(properties, "caps");
                if (value == WordCapsStyle.SmallCaps) ReplaceChild(properties, "smallCaps", WRPrView.Order);
                else RemoveChild(properties, "smallCaps");
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
            get => RunFont("ascii");
            set {
                OfficeXmlNode properties = EditRunProperties();
                OfficeXmlNode fonts = properties.GetOrAddChild(W, "rFonts", WRPrView.Order);
                if (string.IsNullOrEmpty(value)) {
                    fonts.RemoveAttribute(W, "ascii");
                } else {
                    // The other scripts get the same font, which special characters usually need.
                    fonts.SetAttribute(W, "ascii", value!);
                    fonts.SetAttribute(W, "hAnsi", value!);
                    fonts.SetAttribute(W, "eastAsia", value!);
                    fonts.SetAttribute(W, "cs", value!);
                }
                SyncPageNumberFormatting(properties, "rFonts");
            }
        }

        /// <summary>
        /// Gets or sets the HighAnsi font family.
        /// </summary>
        public string? FontFamilyHighAnsi {
            get => RunFont("hAnsi");
            set => SetRunFont("hAnsi", value);
        }

        /// <summary>
        /// Gets or sets the East Asia font family.
        /// </summary>
        public string? FontFamilyEastAsia {
            get => RunFont("eastAsia");
            set => SetRunFont("eastAsia", value);
        }

        /// <summary>
        /// Gets or sets the complex script font family.
        /// </summary>
        public string? FontFamilyComplexScript {
            get => RunFont("cs");
            set => SetRunFont("cs", value);
        }

        /// <summary>
        /// Gets or sets the language applied to the run.
        /// </summary>
        public string? Language {
            get => ChildValue(RunPropertiesNode, "lang");
            set {
                OfficeXmlNode properties = EditRunProperties();
                if (string.IsNullOrWhiteSpace(value)) {
                    RemoveChild(properties, "lang");
                    return;
                }
                // A language sets the main, East Asian and right-to-left languages of an existing or new element.
                string language = value!.Trim();
                OfficeXmlNode element = properties.GetOrAddChild(W, "lang", WRPrView.Order);
                element.SetAttribute(W, "val", language);
                element.SetAttribute(W, "eastAsia", language);
                element.SetAttribute(W, "bidi", language);
            }
        }

        /// <summary>
        /// Gets or sets the character style applied to the run.
        /// </summary>
        public WordCharacterStyles? CharacterStyle {
            get {
                string? styleId = ChildValue(RunPropertiesNode, "rStyle");
                return string.IsNullOrEmpty(styleId) ? null : WordCharacterStyle.GetStyle(styleId!);
            }
            set => ApplyCharacterStyleId(value == null ? null : WordCharacterStyle.ToStringStyle(value.Value));
        }

        /// <summary>
        /// Gets or sets the style identifier applied to the run.
        /// </summary>
        public string? CharacterStyleId {
            get => ChildValue(RunPropertiesNode, "rStyle");
            set => ApplyCharacterStyleId(value);
        }

        // ---- Shared edits ------------------------------------------------------------------------------------------

        private void SetRunOnOffPair(string name, string complexScriptName, bool value) {
            OfficeXmlNode properties = EditRunProperties();
            string token = value ? "true" : "false";
            ReplaceChild(properties, name, WRPrView.Order, token);
            ReplaceChild(properties, complexScriptName, WRPrView.Order, token);
        }

        // On adds a fresh element without a value; off removes it.
        private void SetRunToggle(string name, bool value) {
            OfficeXmlNode properties = EditRunProperties();
            if (value) ReplaceChild(properties, name, WRPrView.Order);
            else RemoveChild(properties, name);
        }

        // A value replaces the element; null removes it. Returns the run properties.
        private OfficeXmlNode SetRunValue(string name, string? value) {
            OfficeXmlNode properties = EditRunProperties();
            if (value == null) RemoveChild(properties, name);
            else ReplaceChild(properties, name, WRPrView.Order, value);
            return properties;
        }

        private string? RunFont(string attribute) {
            OfficeXmlNode properties = RunPropertiesNode;
            if (properties.IsNone) return null;
            OfficeXmlNode fonts = properties.Child(W, "rFonts");
            return fonts.IsNone ? null : fonts.Attribute(W, attribute);
        }

        private void SetRunFont(string attribute, string? value) {
            OfficeXmlNode fonts = EditRunProperties().GetOrAddChild(W, "rFonts", WRPrView.Order);
            if (string.IsNullOrEmpty(value)) fonts.RemoveAttribute(W, attribute);
            else fonts.SetAttribute(W, attribute, value!);
        }

        /// <summary>
        /// Applies a character style, adding its definition first when the document lacks it or what it depends on.
        /// Clearing removes the run's style.
        /// </summary>
        private void ApplyCharacterStyleId(string? styleId) {
            if (string.IsNullOrEmpty(styleId)) {
                RemoveChild(EditRunProperties(), "rStyle");
                return;
            }
            // Adding a definition can open the SDK package, so it runs before the run properties are taken to edit.
            WordPackageModel? model = OwnedModel;
            if (model == null) {
                _document?.EnsureStyleDefinition(styleId!);
            } else if (_document.FileOpenAccess != FileAccess.Read && !model.Styles.DefinesWithDependencies(styleId!)) {
                _document.EnsureStyleDefinition(styleId!);
            }
            EditRunProperties().GetOrAddChild(W, "rStyle", WRPrView.Order).SetAttribute(W, "val", styleId);
        }
    }
}
