using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Border presets that can be applied to a document section.
    /// </summary>
    public enum WordBorder {
        /// <summary>
        /// No preset border is applied.
        /// </summary>
        None,
        /// <summary>
        /// Custom border configured by the user.
        /// </summary>
        Custom,
        /// <summary>
        /// Box border around the page.
        /// </summary>
        Box,
        /// <summary>
        /// Border with a shadow effect.
        /// </summary>
        Shadow
    }

    /// <summary>
    /// Provides access to page border settings for a section.
    /// </summary>
    public class WordBorders {
        private readonly WordDocument _document;
        private readonly WordSection _section;

        internal WordBorders(WordDocument wordDocument, WordSection wordSection) {
            _document = wordDocument;
            _section = wordSection;
        }

        /// <summary>
        /// Gets or sets the width of the left border.
        /// </summary>
        public UInt32Value? LeftSize {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Size;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the left border color using a hexadecimal value.
        /// </summary>
        public string? LeftColorHex {
            get {
                var color = _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the left border color using a <see cref="SixLabors.ImageSharp.Color"/> value.
        /// </summary>
        public SixLabors.ImageSharp.Color LeftColor {
            get {
                var hex = LeftColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("LeftColorHex is null"));
            }
            set {
                LeftColorHex = value.ToHexColor();
            }
        }

        /// <summary>
        /// Gets or sets the style of the left border.
        /// </summary>
        public BorderValues? LeftStyle {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Val?.Value;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                if (value.HasValue) {
                    leftBorder.Val = value.Value;
                } else {
                    leftBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the space between the left border and page text.
        /// </summary>
        public UInt32Value? LeftSpace {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Space;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the left border has a shadow.
        /// </summary>
        public bool? LeftShadow {
            get {
                var shadow = _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Shadow;
                return shadow != null ? shadow.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the left border is part of a frame.
        /// </summary>
        public bool? LeftFrame {
            get {
                var frame = _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Frame;
                return frame != null ? frame.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Frame = value;
            }
        }

        /// <summary>
        /// Gets or sets the width of the right border.
        /// </summary>
        public UInt32Value? RightSize {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Size;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the right border color using a hexadecimal value.
        /// </summary>
        public string? RightColorHex {
            get {
                var color = _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the right border color using a <see cref="SixLabors.ImageSharp.Color"/> value.
        /// </summary>
        public SixLabors.ImageSharp.Color RightColor {
            get {
                var hex = RightColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("RightColorHex is null"));
            }
            set {
                RightColorHex = value.ToHexColor();
            }
        }

        /// <summary>
        /// Gets or sets the style of the right border.
        /// </summary>
        public BorderValues? RightStyle {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Val?.Value;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                if (value.HasValue) {
                    rightBorder.Val = value.Value;
                } else {
                    rightBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the space between the right border and page text.
        /// </summary>
        public UInt32Value? RightSpace {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Space;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the right border has a shadow.
        /// </summary>
        public bool? RightShadow {
            get {
                var shadow = _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Shadow;
                return shadow != null ? shadow.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the right border is part of a frame.
        /// </summary>
        public bool? RightFrame {
            get {
                var frame = _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Frame;
                return frame != null ? frame.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Frame = value;
            }
        }

        /// <summary>
        /// Gets or sets the width of the top border.
        /// </summary>
        public UInt32Value? TopSize {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Size;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the top border color using a hexadecimal value.
        /// </summary>
        public string? TopColorHex {
            get {
                var color = _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the top border color using a <see cref="SixLabors.ImageSharp.Color"/> value.
        /// </summary>
        public SixLabors.ImageSharp.Color TopColor {
            get {
                var hex = TopColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("TopColorHex is null"));
            }
            set {
                TopColorHex = value.ToHexColor();
            }
        }

        /// <summary>
        /// Gets or sets the style of the top border.
        /// </summary>
        public BorderValues? TopStyle {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Val?.Value;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                if (value.HasValue) {
                    topBorder.Val = value.Value;
                } else {
                    topBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the space between the top border and page text.
        /// </summary>
        public UInt32Value? TopSpace {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Space;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the top border has a shadow.
        /// </summary>
        public bool? TopShadow {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                var shadow = pageBorder?.TopBorder?.Shadow;
                return shadow != null ? shadow.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the top border is part of a frame.
        /// </summary>
        public bool? TopFrame {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                var frame = pageBorder?.TopBorder?.Frame;
                return frame != null ? frame.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Frame = value;
            }
        }


        /// <summary>
        /// Gets or sets the width of the bottom border.
        /// </summary>
        public UInt32Value? BottomSize {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Size;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the bottom border color using a hexadecimal value.
        /// </summary>
        public string? BottomColorHex {
            get {
                var color = _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the bottom border color using a <see cref="SixLabors.ImageSharp.Color"/> value.
        /// </summary>
        public SixLabors.ImageSharp.Color BottomColor {
            get {
                var hex = BottomColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("BottomColorHex is null"));
            }
            set {
                BottomColorHex = value.ToHexColor();
            }
        }

        /// <summary>
        /// Gets or sets the style of the bottom border.
        /// </summary>
        public BorderValues? BottomStyle {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Val?.Value;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                if (value.HasValue) {
                    bottomBorder.Val = value.Value;
                } else {
                    bottomBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the space between the bottom border and page text.
        /// </summary>
        public UInt32Value? BottomSpace {
            get {
                return _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Space;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the bottom border has a shadow.
        /// </summary>
        public bool? BottomShadow {
            get {
                var shadow = _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Shadow;
                return shadow != null ? shadow.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the bottom border is part of a frame.
        /// </summary>
        public bool? BottomFrame {
            get {
                var frame = _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Frame;
                return frame != null ? frame.Value : (bool?)null;
            }
            set {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder == null) {
                    pageBorder = Custom;
                    _section._sectionProperties.Append(pageBorder);
                }

                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Frame = value;
            }
        }


        internal void SetBorder(WordBorder wordBorder) {
            var pageBorderSettings = GetDefault(wordBorder);
            if (pageBorderSettings == null) {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                pageBorder?.Remove();
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
        /// Gets or sets the preset border configuration applied to the section.
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
                        if (pageBordersBuiltin == null) {
                            continue;
                        }

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

        private static PageBorders? GetDefault(WordBorder border) {
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