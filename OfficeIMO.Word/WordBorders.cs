using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum WordBorder {
        None,
        Custom,
        Box,
        B3D,
        Shadow
    }
    public class WordBorders {
        private readonly WordDocument _document;
        private readonly WordSection _section;



        public WordBorders(WordDocument wordDocument, WordSection wordSection) {
            _document = wordDocument;
            _section = wordSection;
        }


        public UInt32Value LeftSize {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.LeftBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }
                pageBorder.LeftBorder.Size = value;
            }
        }
        public string LeftColorHex {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return (pageBorder.LeftBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }
                pageBorder.LeftBorder.Color = value.Replace("#", "");
            }
        }
        public System.Drawing.Color LeftColor {
            get {
                return System.Drawing.ColorTranslator.FromHtml("#" + LeftColorHex);
            }
            set {
                this.LeftColorHex = value.ToHexColor();
            }
        }
        public BorderValues? LeftStyle {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.LeftBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Val = value;
            }
        }

        public UInt32Value RightSize {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.RightBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }
                pageBorder.RightBorder.Size = value;
            }
        }
        public string RightColorHex {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return (pageBorder.RightBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }
                pageBorder.RightBorder.Color = value.Replace("#", "");
            }
        }
        public System.Drawing.Color RightColor {
            get {
                return System.Drawing.ColorTranslator.FromHtml("#" + RightColorHex);
            }
            set {
                this.RightColorHex = value.ToHexColor();
            }
        }
        public BorderValues? RightStyle {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.RightBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Val = value;
            }
        }

        public UInt32Value TopSize {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.TopBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }
                pageBorder.TopBorder.Size = value;
            }
        }
        public string TopColorHex {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return (pageBorder.TopBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }
                pageBorder.TopBorder.Color = value.Replace("#", "");
            }
        }
        public System.Drawing.Color TopColor {
            get {
                return System.Drawing.ColorTranslator.FromHtml("#" + TopColorHex);
            }
            set {
                this.TopColorHex = value.ToHexColor();
            }
        }
        public BorderValues? TopStyle {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.TopBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Val = value;
            }
        }

        public UInt32Value BottomSize {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.BottomBorder.Size;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }
                pageBorder.BottomBorder.Size = value;
            }
        }
        public string BottomColorHex {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return (pageBorder.BottomBorder.Color).Value.Replace("#", "");
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }
                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }
                pageBorder.BottomBorder.Color = value.Replace("#", "");
            }
        }
        public System.Drawing.Color BottomColor {
            get {
                return System.Drawing.ColorTranslator.FromHtml("#" + BottomColorHex);
            }
            set {
                this.BottomColorHex = value.ToHexColor();
            }
        }
        public BorderValues? BottomStyle {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.BottomBorder.Val;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(WordBordersSettings.Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Val = value;
            }
        }

        public void SetBorder(WordBorder wordBorder) {
            var pageBorderSettings = WordBordersSettings.GetDefault(wordBorder);
            if (pageBorderSettings == null) {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    pageBorder.Remove();
                }
            } else {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(pageBorderSettings);
                } else {
                    pageBorder.Remove();
                    _section._sectionProperties.Append(pageBorderSettings);
                }
            }
        }

        public WordBorder Type {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    foreach (WordBorder wordBorder in Enum.GetValues(typeof(WordBorder))) {
                        var pageBordersBuiltin = WordBordersSettings.GetDefault(wordBorder);
                        if (pageBordersBuiltin == pageBorder) {
                            //if (pageBordersBuiltin == pageBorder) {
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
    }

    internal static class WordBordersSettings {

        public static PageBorders GetDefault(WordBorder border) {
            switch (border) {
                case WordBorder.Box: return Box;
                case WordBorder.B3D: return B3D;
                case WordBorder.Shadow: return Shadow;
                case WordBorder.None: return null;
                case WordBorder.Custom: return Custom;
            }

            throw new ArgumentOutOfRangeException(nameof(border));
        }

        internal static PageBorders Custom {
            get { return new PageBorders(); }
        }
        internal static PageBorders Box {
            get {
                PageBorders pageBorders1 = new PageBorders() { OffsetFrom = PageBorderOffsetValues.Page };
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };

                pageBorders1.Append(topBorder1);
                pageBorders1.Append(leftBorder1);
                pageBorders1.Append(bottomBorder1);
                pageBorders1.Append(rightBorder1);
                return pageBorders1;
            }
        }
        internal static PageBorders B3D {
            get {
                PageBorders pageBorders1 = new PageBorders() { OffsetFrom = PageBorderOffsetValues.Page };
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };

                pageBorders1.Append(topBorder1);
                pageBorders1.Append(leftBorder1);
                pageBorders1.Append(bottomBorder1);
                pageBorders1.Append(rightBorder1);
                return pageBorders1;
            }
        }
        internal static PageBorders Shadow {
            get {
                PageBorders pageBorders1 = new PageBorders() { OffsetFrom = PageBorderOffsetValues.Page };
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };

                pageBorders1.Append(topBorder1);
                pageBorders1.Append(leftBorder1);
                pageBorders1.Append(bottomBorder1);
                pageBorders1.Append(rightBorder1);
                return pageBorders1;
            }
        }
    }
}
