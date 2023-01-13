using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word {
    public partial class WordParagraph {

        public bool Bold {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink._runProperties : _runProperties;
                if (runProperties != null && runProperties.Bold != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    VerifyRunProperties(this.Hyperlink._hyperlink, this.Hyperlink._run, this.Hyperlink._runProperties);
                    runProperties = this.Hyperlink._runProperties;
                } else {
                    VerifyRunProperties();
                    runProperties = _runProperties;
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

        public bool Italic {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink._runProperties : _runProperties;
                if (runProperties != null && runProperties.Italic != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    VerifyRunProperties(this.Hyperlink._hyperlink, this.Hyperlink._run, this.Hyperlink._runProperties);
                    runProperties = this.Hyperlink._runProperties;
                } else {
                    VerifyRunProperties();
                    runProperties = _runProperties;
                }
                if (value != true) {
                    runProperties.Italic = null;
                } else {
                    runProperties.Italic = new Italic { };
                }
            }
        }

        public UnderlineValues? Underline {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink._runProperties : _runProperties;
                if (runProperties != null && runProperties.Underline != null) {
                    return runProperties.Underline.Val;
                } else {
                    return null;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    VerifyRunProperties(this.Hyperlink._hyperlink, this.Hyperlink._run, this.Hyperlink._runProperties);
                    runProperties = this.Hyperlink._runProperties;
                } else {
                    VerifyRunProperties();
                    runProperties = _runProperties;
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

        public bool DoNotCheckSpellingOrGrammar {
            get {
                var runProperties = IsHyperLink ? this.Hyperlink._runProperties : _runProperties;
                if (runProperties != null && runProperties.NoProof != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    VerifyRunProperties(this.Hyperlink._hyperlink, this.Hyperlink._run, this.Hyperlink._runProperties);
                    runProperties = this.Hyperlink._runProperties;
                } else {
                    VerifyRunProperties();
                    runProperties = _runProperties;
                }
                if (value != true) {
                    if (runProperties.NoProof != null) runProperties.NoProof.Remove();
                } else {
                    runProperties.NoProof = new NoProof();
                }
            }
        }

        public int? Spacing {
            get {
                if (_runProperties != null && _runProperties.Spacing != null) {
                    return _runProperties.Spacing.Val;
                } else {
                    return null;
                }
            }
            set {
                VerifyRunProperties();
                if (value != null) {
                    Spacing spacing = new Spacing();
                    spacing.Val = value;
                    _runProperties.Spacing = spacing;
                } else {
                    if (_runProperties.Spacing != null) _runProperties.Spacing.Remove();
                }
            }
        }

        public bool Strike {
            get {
                if (_runProperties != null && _runProperties.Strike != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                VerifyRunProperties();
                if (value != true) {
                    if (_runProperties.Strike != null) _runProperties.Strike.Remove();
                } else {
                    _runProperties.Strike = new Strike();
                }
            }
        }

        public bool DoubleStrike {
            get {
                if (_runProperties != null && _runProperties.DoubleStrike != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                VerifyRunProperties();
                if (value != true) {
                    if (_runProperties.DoubleStrike != null) _runProperties.DoubleStrike.Remove();
                } else {
                    _runProperties.DoubleStrike = new DoubleStrike();
                }
            }
        }
        public int? FontSize {
            get {
                if (_runProperties != null && _runProperties.FontSize != null) {
                    var fontSizeInHalfPoint = int.Parse(_runProperties.FontSize.Val);
                    return fontSizeInHalfPoint / 2;
                } else {
                    return null;
                }
            }
            set {
                VerifyRunProperties();
                if (value != null) {
                    FontSize fontSize = new FontSize();
                    fontSize.Val = (value * 2).ToString();
                    _runProperties.FontSize = fontSize;
                } else {
                    if (_runProperties.FontSize != null) _runProperties.FontSize.Remove();
                }
            }
        }

        public SixLabors.ImageSharp.Color Color {
            get { return SixLabors.ImageSharp.Color.Parse("#" + ColorHex); }
            set { this.ColorHex = value.ToHexColor(); }
        }

        public string ColorHex {
            get {
                if (_runProperties != null && _runProperties.Color != null) {
                    return _runProperties.Color.Val;
                } else {
                    return "";
                }
            }
            set {
                VerifyRunProperties();
                //string stringColor = value;
                // var color = SixLabors.ImageSharp.Color.FromArgb(Convert.ToInt32(stringColor.Substring(0, 2), 16), Convert.ToInt32(stringColor.Substring(2, 2), 16), Convert.ToInt32(stringColor.Substring(4, 2), 16));
                if (value != "") {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color();
                    color.Val = value.Replace("#", "");
                    _runProperties.Color = color;
                } else {
                    if (_runProperties.Color != null) _runProperties.Color.Remove();
                }
            }
        }

        public ThemeColorValues? ThemeColor {
            get {
                if (_runProperties != null && _runProperties.Color != null && _runProperties.Color.ThemeColor != null) {
                    return _runProperties.Color.ThemeColor.Value;
                } else {
                    return null;
                }
            }
            set {
                VerifyRunProperties();
                //string stringColor = value;
                // var color = SixLabors.ImageSharp.Color.FromArgb(Convert.ToInt32(stringColor.Substring(0, 2), 16), Convert.ToInt32(stringColor.Substring(2, 2), 16), Convert.ToInt32(stringColor.Substring(4, 2), 16));
                if (value != null) {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color {
                        ThemeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value
                        }
                    };
                    _runProperties.Color = color;
                } else {
                    if (_runProperties.Color != null) _runProperties.Color.Remove();
                }
            }
        }

        public HighlightColorValues? Highlight {
            get {
                if (_runProperties != null && _runProperties.Highlight != null) {
                    return _runProperties.Highlight.Val;
                } else {
                    return null;
                }
            }
            set {
                VerifyRunProperties();
                var highlight = new Highlight {
                    Val = value
                };
                _runProperties.Highlight = highlight;
            }
        }

        public CapsStyle CapsStyle {
            get {
                if (_runProperties != null && _runProperties.Caps != null) {
                    return CapsStyle.Caps;
                } else if (_runProperties != null && _runProperties.SmallCaps != null) {
                    return CapsStyle.SmallCaps;
                } else {
                    return CapsStyle.None;
                }
            }
            set {
                VerifyRunProperties();
                if (value == CapsStyle.None) {
                    _runProperties.Caps = null;
                    _runProperties.SmallCaps = null;
                } else if (value == CapsStyle.Caps) {
                    _runProperties.Caps = new Caps();
                } else if (value == CapsStyle.SmallCaps) {
                    _runProperties.SmallCaps = new SmallCaps();
                }
            }
        }

        public string FontFamily {
            get {
                if (_runProperties != null && _runProperties.RunFonts != null) {
                    return _runProperties.RunFonts.Ascii;
                } else {
                    return null;
                }
            }
            set {
                VerifyRunProperties();
                var runFonts = new RunFonts();
                runFonts.Ascii = value;
                _runProperties.RunFonts = runFonts;
            }
        }
    }
}
