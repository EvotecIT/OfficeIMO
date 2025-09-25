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

        private ParagraphBorders? GetParagraphBorders() => _wordParagraph._paragraphProperties?.GetFirstChild<ParagraphBorders>();

        private ParagraphBorders GetOrCreateParagraphBorders() {
            var pageBorder = GetParagraphBorders();
            if (pageBorder == null) {
                pageBorder = Custom;
                _wordParagraph._paragraphProperties!.Append(pageBorder);
            }

            return pageBorder;
        }

        /// <summary>
        /// Gets or sets the left border width in points.
        /// </summary>
        public UInt32Value? LeftSize {
            get => GetParagraphBorders()?.LeftBorder?.Size;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the left border color as a hex value.
        /// </summary>
        public string? LeftColorHex {
            get {
                var color = GetParagraphBorders()?.LeftBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the left border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? LeftColor {
            get => LeftColorHex == null || LeftColorHex == "auto" ? null : Helpers.ParseColor(LeftColorHex);
            set => LeftColorHex = value?.ToHexColor();
        }

        /// <summary>
        /// Gets or sets the left border theme color.
        /// </summary>
        public ThemeColorValues? LeftThemeColor {
            get => GetParagraphBorders()?.LeftBorder?.ThemeColor?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.ThemeColor = value.HasValue ? new EnumValue<ThemeColorValues>(value.Value) : null;
            }
        }

        /// <summary>
        /// Gets or sets the left border style.
        /// </summary>
        public BorderValues? LeftStyle {
            get => GetParagraphBorders()?.LeftBorder?.Val?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                if (value.HasValue) {
                    leftBorder.Val = value.Value;
                } else {
                    leftBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the left border spacing.
        /// </summary>
        public UInt32Value? LeftSpace {
            get => GetParagraphBorders()?.LeftBorder?.Space;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the left border has a shadow.
        /// </summary>
        public bool? LeftShadow {
            get => GetParagraphBorders()?.LeftBorder?.Shadow?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the left border is part of a frame.
        /// </summary>
        public bool? LeftFrame {
            get => GetParagraphBorders()?.LeftBorder?.Frame?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var leftBorder = pageBorder.LeftBorder ?? (pageBorder.LeftBorder = new LeftBorder());
                leftBorder.Frame = value;
            }
        }

        /// <summary>
        /// Gets or sets the right border width in points.
        /// </summary>
        public UInt32Value? RightSize {
            get => GetParagraphBorders()?.RightBorder?.Size;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the right border color as a hex value.
        /// </summary>
        public string? RightColorHex {
            get {
                var color = GetParagraphBorders()?.RightBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the right border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? RightColor {
            get => RightColorHex == null || RightColorHex == "auto" ? null : Helpers.ParseColor(RightColorHex);
            set => RightColorHex = value?.ToHexColor();
        }

        /// <summary>
        /// Gets or sets the right border theme color.
        /// </summary>
        public ThemeColorValues? RightThemeColor {
            get => GetParagraphBorders()?.RightBorder?.ThemeColor?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.ThemeColor = value.HasValue ? new EnumValue<ThemeColorValues>(value.Value) : null;
            }
        }

        /// <summary>
        /// Gets or sets the right border style.
        /// </summary>
        public BorderValues? RightStyle {
            get => GetParagraphBorders()?.RightBorder?.Val?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                if (value.HasValue) {
                    rightBorder.Val = value.Value;
                } else {
                    rightBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the right border spacing.
        /// </summary>
        public UInt32Value? RightSpace {
            get => GetParagraphBorders()?.RightBorder?.Space;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the right border has a shadow.
        /// </summary>
        public bool? RightShadow {
            get => GetParagraphBorders()?.RightBorder?.Shadow?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the right border is part of a frame.
        /// </summary>
        public bool? RightFrame {
            get => GetParagraphBorders()?.RightBorder?.Frame?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var rightBorder = pageBorder.RightBorder ?? (pageBorder.RightBorder = new RightBorder());
                rightBorder.Frame = value;
            }
        }

        /// <summary>
        /// Gets or sets the top border width in points.
        /// </summary>
        public UInt32Value? TopSize {
            get => GetParagraphBorders()?.TopBorder?.Size;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the top border color as a hex value.
        /// </summary>
        public string? TopColorHex {
            get {
                var color = GetParagraphBorders()?.TopBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the top border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? TopColor {
            get => TopColorHex == null || TopColorHex == "auto" ? null : Helpers.ParseColor(TopColorHex);
            set => TopColorHex = value?.ToHexColor();
        }

        /// <summary>
        /// Gets or sets the top border theme color.
        /// </summary>
        public ThemeColorValues? TopThemeColor {
            get => GetParagraphBorders()?.TopBorder?.ThemeColor?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.ThemeColor = value.HasValue ? new EnumValue<ThemeColorValues>(value.Value) : null;
            }
        }

        /// <summary>
        /// Gets or sets the top border style.
        /// </summary>
        public BorderValues? TopStyle {
            get => GetParagraphBorders()?.TopBorder?.Val?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                if (value.HasValue) {
                    topBorder.Val = value.Value;
                } else {
                    topBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the top border spacing.
        /// </summary>
        public UInt32Value? TopSpace {
            get => GetParagraphBorders()?.TopBorder?.Space;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the top border has a shadow.
        /// </summary>
        public bool? TopShadow {
            get => GetParagraphBorders()?.TopBorder?.Shadow?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the top border is part of a frame.
        /// </summary>
        public bool? TopFrame {
            get => GetParagraphBorders()?.TopBorder?.Frame?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var topBorder = pageBorder.TopBorder ?? (pageBorder.TopBorder = new TopBorder());
                topBorder.Frame = value;
            }
        }


        /// <summary>
        /// Gets or sets the bottom border width in points.
        /// </summary>
        public UInt32Value? BottomSize {
            get => GetParagraphBorders()?.BottomBorder?.Size;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Size = value;
            }
        }

        /// <summary>
        /// Gets or sets the bottom border color as a hex value.
        /// </summary>
        public string? BottomColorHex {
            get {
                var color = GetParagraphBorders()?.BottomBorder?.Color?.Value;
                return color != null ? color.Replace("#", "").ToLowerInvariant() : null;
            }
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Gets or sets the bottom border color.
        /// </summary>
        public SixLabors.ImageSharp.Color? BottomColor {
            get => BottomColorHex == null || BottomColorHex == "auto" ? null : Helpers.ParseColor(BottomColorHex);
            set => BottomColorHex = value?.ToHexColor();
        }

        /// <summary>
        /// Gets or sets the bottom border theme color.
        /// </summary>
        public ThemeColorValues? BottomThemeColor {
            get => GetParagraphBorders()?.BottomBorder?.ThemeColor?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.ThemeColor = value.HasValue ? new EnumValue<ThemeColorValues>(value.Value) : null;
            }
        }

        /// <summary>
        /// Gets or sets the bottom border style.
        /// </summary>
        public BorderValues? BottomStyle {
            get => GetParagraphBorders()?.BottomBorder?.Val?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                if (value.HasValue) {
                    bottomBorder.Val = value.Value;
                } else {
                    bottomBorder.Val = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the bottom border spacing.
        /// </summary>
        public UInt32Value? BottomSpace {
            get => GetParagraphBorders()?.BottomBorder?.Space;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Space = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the bottom border has a shadow.
        /// </summary>
        public bool? BottomShadow {
            get => GetParagraphBorders()?.BottomBorder?.Shadow?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Shadow = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the bottom border is part of a frame.
        /// </summary>
        public bool? BottomFrame {
            get => GetParagraphBorders()?.BottomBorder?.Frame?.Value;
            set {
                var pageBorder = GetOrCreateParagraphBorders();
                var bottomBorder = pageBorder.BottomBorder ?? (pageBorder.BottomBorder = new BottomBorder());
                bottomBorder.Frame = value;
            }
        }


        internal void SetBorder(WordBorder wordBorder) {
            var ParagraphBordersettings = GetDefault(wordBorder);
            if (ParagraphBordersettings == null) {
                var pageBorder = GetParagraphBorders();
                pageBorder?.Remove();
            } else {
                var pageBorder = GetParagraphBorders();
                if (pageBorder == null) {
                    _wordParagraph._paragraphProperties!.Append(ParagraphBordersettings);
                } else {
                    pageBorder.Remove();
                    _wordParagraph._paragraphProperties!.Append(ParagraphBordersettings);
                }
            }
        }

        /// <summary>
        /// Gets or sets the current border preset type.
        /// </summary>
        public WordBorder Type {
            get {
                var pageBorder = GetParagraphBorders();
                if (pageBorder != null) {
                    foreach (WordBorder wordBorder in Enum.GetValues(typeof(WordBorder))) {
                        if (wordBorder == WordBorder.None) {
                            continue;
                        }

                        var ParagraphBordersBuiltin = GetDefault(wordBorder)!;

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

        private static ParagraphBorders? GetDefault(WordBorder border) {
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
