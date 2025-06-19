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
        Shadow
    }

    public class WordBorders {
        private readonly WordDocument _document;
        private readonly WordSection _section;

        internal WordBorders(WordDocument wordDocument, WordSection wordSection) {
            _document = wordDocument;
            _section = wordSection;
        }

        /// <summary>
        /// Gets or sets the LeftSize.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the LeftColorHex.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Gets or sets the LeftColor.
        /// </summary>
        public SixLabors.ImageSharp.Color LeftColor {
            get { return SixLabors.ImageSharp.Color.Parse("#" + LeftColorHex); }
            set { this.LeftColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the LeftStyle.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the LeftSpace.
        /// </summary>
        public UInt32Value LeftSpace {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.LeftBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets the LeftShadow.
        /// </summary>
        public bool? LeftShadow {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.LeftBorder.Shadow != null) {
                    return pageBorder.LeftBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets the LeftFrame.
        /// </summary>
        public bool? LeftFrame {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.LeftBorder.Frame != null) {
                    return pageBorder.LeftBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.LeftBorder == null) {
                    pageBorder.LeftBorder = new LeftBorder();
                }

                pageBorder.LeftBorder.Frame = value;
            }
        }

        /// <summary>
        /// Gets or sets the RightSize.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the RightColorHex.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Gets or sets the RightColor.
        /// </summary>
        public SixLabors.ImageSharp.Color RightColor {
            get { return SixLabors.ImageSharp.Color.Parse("#" + RightColorHex); }
            set { this.RightColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the RightStyle.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the RightSpace.
        /// </summary>
        public UInt32Value RightSpace {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.RightBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets the RightShadow.
        /// </summary>
        public bool? RightShadow {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.RightBorder.Shadow != null) {
                    return pageBorder.RightBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets the RightFrame.
        /// </summary>
        public bool? RightFrame {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.RightBorder.Frame != null) {
                    return pageBorder.RightBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.RightBorder == null) {
                    pageBorder.RightBorder = new RightBorder();
                }

                pageBorder.RightBorder.Frame = value;
            }
        }

        /// <summary>
        /// Gets or sets the TopSize.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the TopColorHex.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Gets or sets the TopColor.
        /// </summary>
        public SixLabors.ImageSharp.Color TopColor {
            get { return SixLabors.ImageSharp.Color.Parse("#" + TopColorHex); }
            set { this.TopColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the TopStyle.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the TopSpace.
        /// </summary>
        public UInt32Value TopSpace {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.TopBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets the TopShadow.
        /// </summary>
        public bool? TopShadow {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.TopBorder.Shadow != null) {
                    return pageBorder.TopBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets the TopFrame.
        /// </summary>
        public bool? TopFrame {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.TopBorder.Frame != null) {
                    return pageBorder.TopBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.TopBorder == null) {
                    pageBorder.TopBorder = new TopBorder();
                }

                pageBorder.TopBorder.Frame = value;
            }
        }


        /// <summary>
        /// Gets or sets the BottomSize.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the BottomColorHex.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Gets or sets the BottomColor.
        /// </summary>
        public SixLabors.ImageSharp.Color BottomColor {
            get { return SixLabors.ImageSharp.Color.Parse("#" + BottomColorHex); }
            set { this.BottomColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the BottomStyle.
        /// </summary>
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
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the BottomSpace.
        /// </summary>
        public UInt32Value BottomSpace {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    return pageBorder.BottomBorder.Space;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets the BottomShadow.
        /// </summary>
        public bool? BottomShadow {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.BottomBorder.Shadow != null) {
                    return pageBorder.BottomBorder.Shadow;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets the BottomFrame.
        /// </summary>
        public bool? BottomFrame {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null && pageBorder.BottomBorder.Frame != null) {
                    return pageBorder.BottomBorder.Frame;
                }

                return null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    _section._sectionProperties.Append(Custom);
                    pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                }

                if (pageBorder.BottomBorder == null) {
                    pageBorder.BottomBorder = new BottomBorder();
                }

                pageBorder.BottomBorder.Frame = value;
            }
        }


        internal void SetBorder(WordBorder wordBorder) {
            var pageBorderSettings = GetDefault(wordBorder);
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

        /// <summary>
        /// Gets or sets the Type.
        /// </summary>
        public WordBorder Type {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder != null) {
                    foreach (WordBorder wordBorder in Enum.GetValues(typeof(WordBorder))) {
                        if (wordBorder == WordBorder.None) {
                            continue;
                        }

                        var pageBordersBuiltin = GetDefault(wordBorder);

                        if ((pageBordersBuiltin.LeftBorder == null && pageBorder.LeftBorder == null) &&
                            (pageBordersBuiltin.RightBorder == null && pageBorder.RightBorder == null) &&
                            (pageBordersBuiltin.TopBorder == null && pageBorder.TopBorder == null) &&
                            (pageBordersBuiltin.BottomBorder == null && pageBorder.BottomBorder == null)) {
                            return wordBorder;
                        }

                        if (pageBordersBuiltin.LeftBorder != null && pageBorder.LeftBorder != null &&
                            pageBordersBuiltin.RightBorder != null && pageBorder.RightBorder != null &&
                            pageBordersBuiltin.TopBorder != null && pageBorder.TopBorder != null &&
                            pageBordersBuiltin.BottomBorder != null && pageBorder.BottomBorder != null &&
                            pageBordersBuiltin.LeftBorder.Shadow == pageBorder.LeftBorder.Shadow &&
                            pageBordersBuiltin.RightBorder.Shadow == pageBorder.RightBorder.Shadow &&
                            pageBordersBuiltin.TopBorder.Shadow == pageBorder.TopBorder.Shadow &&
                            pageBordersBuiltin.BottomBorder.Shadow == pageBorder.BottomBorder.Shadow &&
                            pageBordersBuiltin.LeftBorder.Color == pageBorder.LeftBorder.Color &&
                            pageBordersBuiltin.RightBorder.Color == pageBorder.RightBorder.Color &&
                            pageBordersBuiltin.TopBorder.Color == pageBorder.TopBorder.Color &&
                            pageBordersBuiltin.BottomBorder.Color == pageBorder.BottomBorder.Color &&
                            pageBordersBuiltin.LeftBorder.Size == pageBorder.LeftBorder.Size &&
                            pageBordersBuiltin.RightBorder.Size == pageBorder.RightBorder.Size &&
                            pageBordersBuiltin.TopBorder.Size == pageBorder.TopBorder.Size &&
                            pageBordersBuiltin.BottomBorder.Size == pageBorder.BottomBorder.Size &&
                            pageBordersBuiltin.LeftBorder.Space == pageBorder.LeftBorder.Space &&
                            pageBordersBuiltin.RightBorder.Space == pageBorder.RightBorder.Space &&
                            pageBordersBuiltin.TopBorder.Space == pageBorder.TopBorder.Space &&
                            pageBordersBuiltin.BottomBorder.Space == pageBorder.BottomBorder.Space) {
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

        private static PageBorders GetDefault(WordBorder border) {
            switch (border) {
                case WordBorder.Box: return Box;
                case WordBorder.Shadow: return Shadow;
                case WordBorder.None: return null;
                case WordBorder.Custom: return Custom;
            }

            throw new ArgumentOutOfRangeException(nameof(border));
        }

        private static PageBorders Custom => new PageBorders();

        private static PageBorders Box {
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

        private static PageBorders Shadow {
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