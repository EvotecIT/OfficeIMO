namespace OfficeIMO.Word {
    // Manages run property settings. Each member reads and edits OfficeIMO's own package model while the document works
    // without the Open XML SDK; once the SDK holds the document, the same member goes through the bridge.
    public partial class WordParagraph {
        /// <summary>
        /// Gets or sets a value indicating whether the run is bold. Setting false explicitly overrides inherited bold.
        /// </summary>
        public bool Bold {
            get => _ownedTree != null && TryGetOwnedOnOff("b", out bool owned) ? owned : SdkBold;
            set {
                if (_ownedTree != null && TrySetOwnedOnOffPair("b", "bCs", value)) return;
                SdkBold = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is italic. Setting false explicitly overrides inherited italic.
        /// </summary>
        public bool Italic {
            get => _ownedTree != null && TryGetOwnedOnOff("i", out bool owned) ? owned : SdkItalic;
            set {
                if (_ownedTree != null && TrySetOwnedOnOffPair("i", "iCs", value)) return;
                SdkItalic = value;
            }
        }

        /// <summary>
        /// Gets or sets the underline style for the run.
        /// </summary>
        public WordUnderlineStyle? Underline {
            get => _ownedTree != null && TryGetOwnedUnderline(out WordUnderlineStyle? owned) ? owned : SdkUnderline;
            set {
                if (_ownedTree != null && TrySetOwnedUnderline(value)) return;
                SdkUnderline = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether spelling and grammar checks are disabled.
        /// </summary>
        public bool DoNotCheckSpellingOrGrammar {
            get => _ownedTree != null && TryGetOwnedOnOff("noProof", out bool owned) ? owned : SdkDoNotCheckSpellingOrGrammar;
            set {
                if (_ownedTree != null && TrySetOwnedToggle("noProof", value)) return;
                SdkDoNotCheckSpellingOrGrammar = value;
            }
        }

        /// <summary>
        /// Gets or sets the signed physical character spacing. Null removes the local value.
        /// </summary>
        public OfficeLength? Spacing {
            get => _ownedTree != null && TryGetOwnedCharacterSpacing(out OfficeLength? owned) ? owned : SdkSpacing;
            set {
                int? nativeSpacing = value.HasValue ? checked((int)value.Value.ToTwips()) : (int?)null;
                if (_ownedTree != null && TrySetOwnedValue("spacing", nativeSpacing?.ToString(System.Globalization.CultureInfo.InvariantCulture))) return;
                SdkSpacing = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is struck through.
        /// </summary>
        public bool Strikethrough {
            get => _ownedTree != null && TryGetOwnedOnOff("strike", out bool owned) ? owned : SdkStrikethrough;
            set {
                if (_ownedTree != null && TrySetOwnedToggle("strike", value)) return;
                SdkStrikethrough = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run has a double strikethrough.
        /// </summary>
        public bool DoubleStrikethrough {
            get => _ownedTree != null && TryGetOwnedOnOff("dstrike", out bool owned) ? owned : SdkDoubleStrikethrough;
            set {
                if (_ownedTree != null && TrySetOwnedToggle("dstrike", value)) return;
                SdkDoubleStrikethrough = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run text is outlined.
        /// </summary>
        public bool Outline {
            get => _ownedTree != null && TryGetOwnedOnOff("outline", out bool owned) ? owned : SdkOutline;
            set {
                if (_ownedTree != null && TrySetOwnedToggle("outline", value)) return;
                SdkOutline = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run has a shadow effect.
        /// </summary>
        public bool Shadow {
            get => _ownedTree != null && TryGetOwnedOnOff("shadow", out bool owned) ? owned : SdkShadow;
            set {
                if (_ownedTree != null && TrySetOwnedToggle("shadow", value)) return;
                SdkShadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is embossed.
        /// </summary>
        public bool Emboss {
            get => _ownedTree != null && TryGetOwnedOnOff("emboss", out bool owned) ? owned : SdkEmboss;
            set {
                if (_ownedTree != null && TrySetOwnedToggle("emboss", value)) return;
                SdkEmboss = value;
            }
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
            get => _ownedTree != null && TryGetOwnedFontSize(out double? owned) ? owned : SdkNativeFontSizePoints;
            set {
                if (_ownedTree != null && OwnedModel != null) {
                    if (value == null) {
                        if (TrySetOwnedValue("sz", null)) return;
                    } else if (!double.IsNaN(value.Value) && !double.IsInfinity(value.Value) && value.Value >= 0D && value.Value <= int.MaxValue / 2D) {
                        int ownedHalfPoints = checked((int)Math.Round(value.Value * 2D, MidpointRounding.AwayFromZero));
                        if (TrySetOwnedValue("sz", ownedHalfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture))) return;
                    }
                }
                SdkNativeFontSizePoints = value;
            }
        }

        /// <summary>
        /// Gets or sets the text color using <see cref="OfficeIMO.OfficeColor"/>.
        /// </summary>
        public OfficeIMO.OfficeColor? Color {
            get => _ownedTree != null && TryGetOwnedColor(out OfficeIMO.OfficeColor? owned) ? owned : SdkColor;
            set {
                if (_ownedTree != null && (value == null || !value.Value.IsThemeRole)
                    && TrySetOwnedValue("color", value?.ToRgbHex())) return;
                if (value != null) {
                    SetRunColor(value.Value);
                } else {
                    RunColorValue = string.Empty;
                }
            }
        }

        /// <summary>
        /// Gets or sets the native run colour value (hexadecimal or <c>auto</c>); empty removes it.
        /// </summary>
        internal string RunColorValue {
            get => _ownedTree != null && TryGetOwnedValue("color", out string? owned) ? owned ?? "" : SdkRunColorValue;
            set {
                if (_ownedTree != null && (value == "" || Helpers.NormalizeOpenXmlColor(value) != null)
                    && TrySetOwnedValue("color", value != "" ? Helpers.NormalizeOpenXmlColor(value) : null)) return;
                SdkRunColorValue = value;
            }
        }

        /// <summary>
        /// Gets or sets the theme color applied to the run.
        /// </summary>
        public WordThemeColor? ThemeColor {
            get => _ownedTree != null && TryGetOwnedThemeColor(out WordThemeColor? owned) ? owned : SdkThemeColor;
            // The colour value caches the theme's colour, which the SDK path reads from the theme.
            set => SdkThemeColor = value;
        }

        /// <summary>
        /// Gets or sets the highlight color applied to the run.
        /// </summary>
        public WordHighlightColor? Highlight {
            get => _ownedTree != null && TryGetOwnedHighlight(out WordHighlightColor? owned) ? owned : SdkHighlight;
            set {
                if (_ownedTree != null && TrySetOwnedHighlight(value)) return;
                SdkHighlight = value;
            }
        }

        /// <summary>
        /// Gets or sets the capitalization style for the run.
        /// </summary>
        public WordCapsStyle CapsStyle {
            get => _ownedTree != null && TryGetOwnedCaps(out WordCapsStyle owned) ? owned : SdkCapsStyle;
            set {
                if (_ownedTree != null && TrySetOwnedCaps(value)) return;
                SdkCapsStyle = value;
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
            get => _ownedTree != null && TryGetOwnedFont("ascii", out string? owned) ? owned : SdkFontFamily;
            set {
                if (_ownedTree != null && TrySetOwnedFontFamily(value)) return;
                SdkFontFamily = value;
            }
        }

        /// <summary>
        /// Gets or sets the HighAnsi font family.
        /// </summary>
        public string? FontFamilyHighAnsi {
            get => _ownedTree != null && TryGetOwnedFont("hAnsi", out string? owned) ? owned : SdkFontFamilyHighAnsi;
            set {
                if (_ownedTree != null && TrySetOwnedFont("hAnsi", value)) return;
                SdkFontFamilyHighAnsi = value;
            }
        }

        /// <summary>
        /// Gets or sets the East Asia font family.
        /// </summary>
        public string? FontFamilyEastAsia {
            get => _ownedTree != null && TryGetOwnedFont("eastAsia", out string? owned) ? owned : SdkFontFamilyEastAsia;
            set {
                if (_ownedTree != null && TrySetOwnedFont("eastAsia", value)) return;
                SdkFontFamilyEastAsia = value;
            }
        }

        /// <summary>
        /// Gets or sets the complex script font family.
        /// </summary>
        public string? FontFamilyComplexScript {
            get => _ownedTree != null && TryGetOwnedFont("cs", out string? owned) ? owned : SdkFontFamilyComplexScript;
            set {
                if (_ownedTree != null && TrySetOwnedFont("cs", value)) return;
                SdkFontFamilyComplexScript = value;
            }
        }

        /// <summary>
        /// Gets or sets the language applied to the run.
        /// </summary>
        public string? Language {
            get => _ownedTree != null && TryGetOwnedLanguage(out string? owned) ? owned : SdkLanguage;
            set {
                if (_ownedTree != null && TrySetOwnedLanguage(value)) return;
                SdkLanguage = value;
            }
        }

        /// <summary>
        /// Gets or sets the character style applied to the run.
        /// </summary>
        public WordCharacterStyles? CharacterStyle {
            get {
                if (_ownedTree == null || !TryGetOwnedValue("rStyle", out string? styleId)) return SdkCharacterStyle;
                return string.IsNullOrEmpty(styleId) ? null : WordCharacterStyle.GetStyle(styleId!);
            }
            set {
                if (_ownedTree != null && TrySetOwnedCharacterStyleId(value == null ? null : WordCharacterStyle.ToStringStyle(value.Value))) return;
                SdkCharacterStyle = value;
            }
        }

        /// <summary>
        /// Gets or sets the style identifier applied to the run.
        /// </summary>
        public string? CharacterStyleId {
            get => _ownedTree != null && TryGetOwnedValue("rStyle", out string? owned) ? owned : SdkCharacterStyleId;
            set {
                if (_ownedTree != null && TrySetOwnedCharacterStyleId(value)) return;
                SdkCharacterStyleId = value;
            }
        }
    }
}
