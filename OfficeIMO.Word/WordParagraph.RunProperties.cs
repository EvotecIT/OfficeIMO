using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace OfficeIMO.Word {
    public partial class WordParagraph {

        public bool Bold {
            get {
                if (_runProperties != null && _runProperties.Bold != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                VerifyRunProperties();
                if (value == true) {
                    _runProperties.Bold = new Bold();
                    _runProperties.BoldComplexScript = new BoldComplexScript();
                } else {
                    if (_runProperties.BoldComplexScript != null) {
                        _runProperties.BoldComplexScript.Remove();
                    }
                    if (_runProperties.Bold != null) {
                        _runProperties.Bold.Remove();
                    }
                }
            }
        }

        public bool Italic {
            get {
                if (_runProperties != null && _runProperties.Italic != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                VerifyRunProperties();
                if (value != true) {
                    _runProperties.Italic = null;
                } else {
                    _runProperties.Italic = new Italic { };
                }
            }
        }

        public UnderlineValues? Underline {
            get {
                if (_runProperties != null && _runProperties.Underline != null) {
                    return _runProperties.Underline.Val;
                } else {
                    return null;
                }
            }
            set {
                VerifyRunProperties();
                if (value != null) {
                    if (_runProperties.Underline == null) {
                        _runProperties.Underline = new Underline();
                    }

                    _runProperties.Underline.Val = value;
                } else {
                    if (_runProperties.Underline != null) _runProperties.Underline.Remove();
                }
            }
        }

        public bool DoNotCheckSpellingOrGrammar {
            get {
                if (_runProperties != null && _runProperties.NoProof != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                VerifyRunProperties();
                if (value != true) {
                    if (_runProperties.NoProof != null) _runProperties.NoProof.Remove();
                } else {
                    _runProperties.NoProof = new NoProof();
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

        public string Color {
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
                // var color = System.Drawing.Color.FromArgb(Convert.ToInt32(stringColor.Substring(0, 2), 16), Convert.ToInt32(stringColor.Substring(2, 2), 16), Convert.ToInt32(stringColor.Substring(4, 2), 16));
                if (value != "") {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color();
                    color.Val = value;
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
                // var color = System.Drawing.Color.FromArgb(Convert.ToInt32(stringColor.Substring(0, 2), 16), Convert.ToInt32(stringColor.Substring(2, 2), 16), Convert.ToInt32(stringColor.Substring(4, 2), 16));
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
