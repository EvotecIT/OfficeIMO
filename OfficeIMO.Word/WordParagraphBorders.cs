using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    /// <summary>
    /// Defines predefined paragraph border styles.
    /// </summary>
    public enum WordParagraphBorder {
        /// <summary>No borders are applied.</summary>
        None,
        /// <summary>Custom border configuration.</summary>
        Custom,
        /// <summary>Box border surrounding the paragraph.</summary>
        Box,
        /// <summary>Shadowed box border.</summary>
        Shadow
    }

    /// <summary>
    /// Specifies which side of the paragraph border is affected.
    /// </summary>
    public enum WordParagraphBorderType {
        /// <summary>Left border.</summary>
        Left,
        /// <summary>Right border.</summary>
        Right,
        /// <summary>Top border.</summary>
        Top,
        /// <summary>Bottom border.</summary>
        Bottom
    }

    /// <summary>
    /// Provides access to paragraph border properties.
    /// </summary>
    public class WordParagraphBorders {
        private readonly WordDocument _document;
        private readonly WordParagraph _wordParagraph;

        internal WordParagraphBorders(WordDocument wordDocument, WordParagraph wordParagraph) {
            _document = wordDocument;
            _wordParagraph = wordParagraph;
        }

        /// <summary>
        /// Gets or sets the left border width in points.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the left border color as a hex value.
        /// </summary>
        public string LeftColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.LeftBorder != null && pageBorder.LeftBorder.Color != null) {
                    return (pageBorder.LeftBorder.Color).Value.Replace("#", "").ToLowerInvariant();
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

                pageBorder.LeftBorder.Color = value.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the left border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? LeftColor {
            get {
                if (LeftColorHex == null || LeftColorHex == "auto") {
                    return null;
                }
                return Helpers.ParseColor(LeftColorHex);
            }
            set => LeftColorHex = value.Value.ToHexColor();
        }

        /// <summary>
        /// Gets or sets the left border theme color.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the left border style.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the left border spacing.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the left border has a shadow.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the left border is part of a frame.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the right border width in points.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the right border color as a hex value.
        /// </summary>
        public string RightColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.RightBorder != null && pageBorder.RightBorder.Color != null) {
                    return (pageBorder.RightBorder.Color).Value.Replace("#", "").ToLowerInvariant();
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

                pageBorder.RightBorder.Color = value.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the right border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? RightColor {
            get {
                if (RightColorHex == null || RightColorHex == "auto") {
                    return null;
                }
                return Helpers.ParseColor(RightColorHex);
            }
            set => RightColorHex = value.Value.ToHexColor();
        }

        /// <summary>
        /// Gets or sets the right border theme color.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the right border style.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the right border spacing.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the right border has a shadow.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the right border is part of a frame.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the top border width in points.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the top border color as a hex value.
        /// </summary>
        public string TopColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.TopBorder != null && pageBorder.TopBorder.Color != null) {
                    return (pageBorder.TopBorder.Color).Value.Replace("#", "").ToLowerInvariant();
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

                pageBorder.TopBorder.Color = value.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the top border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? TopColor {
            get {
                if (TopColorHex == null || TopColorHex == "auto"
                    ) {
                    return null;
                }
                return Helpers.ParseColor(TopColorHex);
            }
            set { this.TopColorHex = value.Value.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the top border theme color.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the top border style.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the top border spacing.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the top border has a shadow.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the top border is part of a frame.
        /// </summary>
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


        /// <summary>
        /// Gets or sets the bottom border width in points.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the bottom border color as a hex value.
        /// </summary>
        public string BottomColorHex {
            get {
                var pageBorder = _wordParagraph._paragraphProperties.GetFirstChild<ParagraphBorders>();
                if (pageBorder != null && pageBorder.BottomBorder != null && pageBorder.BottomBorder.Color != null) {
                    return (pageBorder.BottomBorder.Color).Value.Replace("#", "").ToLowerInvariant();
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

                pageBorder.BottomBorder.Color = value.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the bottom border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? BottomColor {
            get {
                if (BottomColorHex == null || BottomColorHex == "auto") {
                    return null;
                }
                return Helpers.ParseColor(BottomColorHex);
            }
            set {
                if (value == null) {
                    this.BottomColorHex = null;
                    return;
                }
                this.BottomColorHex = value.Value.ToHexColor();
            }
        }

        /// <summary>
        /// Gets or sets the bottom border theme color.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the bottom border style.
        /// </summary>
        public BorderValues? BottomStyle {
            get {
                var props = _wordParagraph._paragraphProperties;
                if (props != null) {
                    var pageBorder = props.GetFirstChild<ParagraphBorders>();
                    if (pageBorder != null && pageBorder.BottomBorder != null) {
                        return pageBorder.BottomBorder.Val;
                    }
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

        /// <summary>
        /// Gets or sets the bottom border spacing.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the bottom border has a shadow.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the bottom border is part of a frame.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the current border preset type.
        /// </summary>
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

        /// <summary>
        /// Applies border settings to a specific side of the paragraph.
        /// </summary>
        /// <param name="type">Side of the paragraph.</param>
        /// <param name="style">Border style.</param>
        /// <param name="color">Border color.</param>
        /// <param name="size">Border width.</param>
        /// <param name="shadow">Whether the border has a shadow.</param>
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
