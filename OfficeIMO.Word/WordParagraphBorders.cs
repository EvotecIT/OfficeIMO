using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    public enum WordParagraphBorder {
        None,
        Custom,
        Box,
        Shadow
    }

    public enum WordParagraphBorderType {
        Left,
        Right,
        Top,
        Bottom
    }

    public class WordParagraphBorders {
        private readonly WordDocument _document;
        private readonly WordParagraph _wordParagraph;

        internal WordParagraphBorders(WordDocument wordDocument, WordParagraph wordParagraph) {
            _document = wordDocument;
            _wordParagraph = wordParagraph;
        }

        public UInt32Value LeftSize {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.LeftBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Size = value;
            }
        }

        public string LeftColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.LeftBorder != null && pageBorder.LeftBorder.Color != null) {
                    return (pageBorder.LeftBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Color = value.Replace("#", "");
            }
        }

        public SixLabors.ImageSharp.Color? LeftColor {
            get {
                if (LeftColorHex == null || LeftColorHex == "auto") {
                    return null;
                }
                return SixLabors.ImageSharp.Color.Parse("#" + LeftColorHex);
            }
            set => LeftColorHex = value.Value.ToHexColor();
        }

        public ThemeColorValues? LeftThemeColor {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.LeftBorder.ThemeColor.Value;
                }
                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    if (value != null) {
                        var themeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value
                        };
                        pageBorder.LeftBorder.ThemeColor = themeColor;
                    } else {
                        if (pageBorder.LeftBorder.ThemeColor != null) pageBorder.LeftBorder.ThemeColor = null;
                    }
                }
            }
        }

        public BorderValues? LeftStyle {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.LeftBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Val = value;
            }
        }

        public UInt32Value LeftSpace {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.LeftBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Space = value;
            }
        }

        public bool? LeftShadow {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.LeftBorder.Shadow != null) {
                    return pageBorder.LeftBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Shadow = value;
            }
        }

        public bool? LeftFrame {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.LeftBorder.Frame != null) {
                    return pageBorder.LeftBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Frame = value;
            }
        }

        public UInt32Value RightSize {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.RightBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Size = value;
            }
        }

        public string RightColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.RightBorder != null && pageBorder.RightBorder.Color != null) {
                    return (pageBorder.RightBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Color = value.Replace("#", "");
            }
        }

        public SixLabors.ImageSharp.Color? RightColor {
            get {
                if (RightColorHex == null || RightColorHex == "auto") {
                    return null;
                }
                return SixLabors.ImageSharp.Color.Parse("#" + RightColorHex);
            }
            set => RightColorHex = value.Value.ToHexColor();
        }

        public ThemeColorValues? RightThemeColor {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.RightBorder.ThemeColor.Value;
                }
                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    if (value != null) {
                        var themeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value
                        };
                        pageBorder.RightBorder.ThemeColor = themeColor;
                    } else {
                        if (pageBorder.RightBorder.ThemeColor != null) pageBorder.RightBorder.ThemeColor = null;
                    }
                }
            }
        }

        public BorderValues? RightStyle {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.RightBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Val = value;
            }
        }

        public UInt32Value RightSpace {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.RightBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Space = value;
            }
        }

        public bool? RightShadow {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.RightBorder.Shadow != null) {
                    return pageBorder.RightBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Shadow = value;
            }
        }

        public bool? RightFrame {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.RightBorder.Frame != null) {
                    return pageBorder.RightBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Frame = value;
            }
        }

        public UInt32Value TopSize {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.TopBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Size = value;
            }
        }

        public string TopColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.TopBorder != null && pageBorder.TopBorder.Color != null) {
                    return (pageBorder.TopBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Color = value.Replace("#", "");
            }
        }

        public SixLabors.ImageSharp.Color? TopColor {
            get {
                if (TopColorHex == null || TopColorHex == "auto"
                    ) {
                    return null;
                }
                return SixLabors.ImageSharp.Color.Parse("#" + TopColorHex);
            }
            set { this.TopColorHex = value.Value.ToHexColor(); }
        }

        public ThemeColorValues? TopThemeColor {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.TopBorder.ThemeColor.Value;
                }
                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    if (value != null) {
                        var themeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value
                        };
                        pageBorder.TopBorder.ThemeColor = themeColor;
                    } else {
                        if (pageBorder.TopBorder.ThemeColor != null) pageBorder.TopBorder.ThemeColor = null;
                    }
                }
            }
        }

        public BorderValues? TopStyle {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.TopBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Val = value;
            }
        }

        public UInt32Value TopSpace {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.TopBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Space = value;
            }
        }

        public bool? TopShadow {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.TopBorder.Shadow != null) {
                    return pageBorder.TopBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Shadow = value;
            }
        }

        public bool? TopFrame {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.TopBorder.Frame != null) {
                    return pageBorder.TopBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Frame = value;
            }
        }


        public UInt32Value BottomSize {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.BottomBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Size = value;
            }
        }

        public string BottomColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.BottomBorder != null && pageBorder.BottomBorder.Color != null) {
                    return (pageBorder.BottomBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Color = value.Replace("#", "");
            }
        }

        public SixLabors.ImageSharp.Color? BottomColor {
            get {
                if (BottomColorHex == null || BottomColorHex == "auto") {
                    return null;
                }
                return SixLabors.ImageSharp.Color.Parse("#" + BottomColorHex);
            }
            set {
                if (value == null) {
                    this.BottomColorHex = null;
                    return;
                }
                this.BottomColorHex = value.Value.ToHexColor();
            }
        }

        public ThemeColorValues? BottomThemeColor {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.BottomBorder.ThemeColor.Value;
                }
                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    if (value != null) {
                        var themeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value
                        };
                        pageBorder.BottomBorder.ThemeColor = themeColor;
                    } else {
                        if (pageBorder.BottomBorder.ThemeColor != null) pageBorder.BottomBorder.ThemeColor = null;
                    }
                }
            }
        }

        public BorderValues? BottomStyle {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.BottomBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Val = value;
            }
        }

        public UInt32Value BottomSpace {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    return pageBorder.BottomBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Space = value;
            }
        }

        public bool? BottomShadow {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.BottomBorder.Shadow != null) {
                    return pageBorder.BottomBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Shadow = value;
            }
        }

        public bool? BottomFrame {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.BottomBorder.Frame != null) {
                    return pageBorder.BottomBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(Custom);
                    pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Frame = value;
            }
        }


        internal void SetBorder(WordBorder wordBorder) {
            var ParagraphBordersettings = GetDefault(wordBorder);
            if (ParagraphBordersettings == null) {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    pageBorder.Remove();
                }
            } else {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties.Append(ParagraphBordersettings);
                } else {
                    pageBorder.Remove();
                    _wordParagraph._paragraphProperties.Append(ParagraphBordersettings);
                }
            }
        }

        public WordBorder Type {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null) {
                    foreach (WordBorder wordBorder in Enum.GetValues(typeof(WordBorder))) {
                        if (wordBorder == WordBorder.None) {
                            continue;
                        }

                        var ParagraphBordersBuiltin = GetDefault(wordBorder);

                        if ((ParagraphBordersBuiltin.LeftBorder == null && pageBorder.LeftBorder == null) &&
                            (ParagraphBordersBuiltin.RightBorder == null && pageBorder.RightBorder == null) &&
                            (ParagraphBordersBuiltin.TopBorder == null && pageBorder.TopBorder == null) &&
                            (ParagraphBordersBuiltin.BottomBorder == null && pageBorder.BottomBorder == null)) {
                            return wordBorder;
                        }

                        if (ParagraphBordersBuiltin.LeftBorder != null && pageBorder.LeftBorder != null &&
                            ParagraphBordersBuiltin.RightBorder != null && pageBorder.RightBorder != null &&
                            ParagraphBordersBuiltin.TopBorder != null && pageBorder.TopBorder != null &&
                            ParagraphBordersBuiltin.BottomBorder != null && pageBorder.BottomBorder != null &&
                            ParagraphBordersBuiltin.LeftBorder.Shadow == pageBorder.LeftBorder.Shadow &&
                            ParagraphBordersBuiltin.RightBorder.Shadow == pageBorder.RightBorder.Shadow &&
                            ParagraphBordersBuiltin.TopBorder.Shadow == pageBorder.TopBorder.Shadow &&
                            ParagraphBordersBuiltin.BottomBorder.Shadow == pageBorder.BottomBorder.Shadow &&
                            ParagraphBordersBuiltin.LeftBorder.Color == pageBorder.LeftBorder.Color &&
                            ParagraphBordersBuiltin.RightBorder.Color == pageBorder.RightBorder.Color &&
                            ParagraphBordersBuiltin.TopBorder.Color == pageBorder.TopBorder.Color &&
                            ParagraphBordersBuiltin.BottomBorder.Color == pageBorder.BottomBorder.Color &&
                            ParagraphBordersBuiltin.LeftBorder.Size == pageBorder.LeftBorder.Size &&
                            ParagraphBordersBuiltin.RightBorder.Size == pageBorder.RightBorder.Size &&
                            ParagraphBordersBuiltin.TopBorder.Size == pageBorder.TopBorder.Size &&
                            ParagraphBordersBuiltin.BottomBorder.Size == pageBorder.BottomBorder.Size &&
                            ParagraphBordersBuiltin.LeftBorder.Space == pageBorder.LeftBorder.Space &&
                            ParagraphBordersBuiltin.RightBorder.Space == pageBorder.RightBorder.Space &&
                            ParagraphBordersBuiltin.TopBorder.Space == pageBorder.TopBorder.Space &&
                            ParagraphBordersBuiltin.BottomBorder.Space == pageBorder.BottomBorder.Space) {
                            return wordBorder;
                        }
                    }

                    return WordBorder.Custom;
                } else {
                    return WordBorder.None;
                }
            }
            set => SetBorder(value);
        }

        private static ParagraphBorders GetDefault(WordBorder border) {
            switch (border) {
                case WordBorder.Box: return Box;
                case WordBorder.Shadow: return Shadow;
                case WordBorder.None: return null;
                case WordBorder.Custom: return Custom;
            }

            throw new ArgumentOutOfRangeException(nameof(border));
        }

        private static ParagraphBorders Custom => new ParagraphBorders();

        private static ParagraphBorders Box {
            get {
                ParagraphBorders ParagraphBorders1 = new ParagraphBorders();
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };

                ParagraphBorders1.Append(topBorder1);
                ParagraphBorders1.Append(leftBorder1);
                ParagraphBorders1.Append(bottomBorder1);
                ParagraphBorders1.Append(rightBorder1);
                return ParagraphBorders1;
            }
        }

        private static ParagraphBorders Shadow {
            get {
                ParagraphBorders ParagraphBorders1 = new ParagraphBorders();
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };

                ParagraphBorders1.Append(topBorder1);
                ParagraphBorders1.Append(leftBorder1);
                ParagraphBorders1.Append(bottomBorder1);
                ParagraphBorders1.Append(rightBorder1);
                return ParagraphBorders1;
            }
        }

        public void SetBorder(WordParagraphBorderType type, BorderValues style, Color color, UInt32Value size, bool shadow) {
            if (type == WordParagraphBorderType.Left) {
                LeftStyle = style;
                LeftColor = color;
                LeftSize = (UInt32Value)size;
                LeftShadow = shadow;
            } else if (type == WordParagraphBorderType.Right) {
                RightStyle = style;
                RightColor = color;
                RightSize = (UInt32Value)size;
                RightShadow = shadow;
            } else if (type == WordParagraphBorderType.Top) {
                TopStyle = style;
                TopColor = color;
                TopSize = (UInt32Value)size;
                TopShadow = shadow;
            } else if (type == WordParagraphBorderType.Bottom) {
                BottomStyle = style;
                BottomColor = color;
                BottomSize = (UInt32Value)size;
                BottomShadow = shadow;
            }
        }
    }
}
