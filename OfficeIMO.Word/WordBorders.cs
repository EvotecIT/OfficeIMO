using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Border presets that can be applied to a document section.
    /// </summary>
    public enum WordBorder {
        /// <summary>No preset border is applied.</summary>
        None,
        /// <summary>Custom border configured by the user.</summary>
        Custom,
        /// <summary>Box border around the page.</summary>
        Box,
        /// <summary>Border with a shadow effect.</summary>
        Shadow
    }

    /// <summary>
    /// Provides access to page border settings for a section.
    /// </summary>
    public class WordBorders {
        private readonly WordSection _section;

        internal WordBorders(WordDocument wordDocument, WordSection wordSection) {
            _ = wordDocument;
            _section = wordSection;
        }

        private PageBorders EnsurePageBorders() {
            var pageBorders = _section._sectionProperties.GetFirstChild<PageBorders>();
            if (pageBorders is null) {
                pageBorders = Custom;
                _section._sectionProperties.Append(pageBorders);
            }
            return pageBorders;
        }

        private static string? NormalizeColor(string? value) => value?.Replace("#", "").ToUpperInvariant();

        /// <summary>Gets or sets the width of the left border.</summary>
        public UInt32Value? LeftSize {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Size;
            set => (EnsurePageBorders().LeftBorder ??= new LeftBorder()).Size = value;
        }

        /// <summary>Gets or sets the left border color using a hexadecimal value.</summary>
        public string? LeftColorHex {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Color?.Value?.Replace("#", "").ToUpperInvariant();
            set => (EnsurePageBorders().LeftBorder ??= new LeftBorder()).Color = NormalizeColor(value);
        }

        /// <summary>Gets or sets the left border color using a <see cref="SixLabors.ImageSharp.Color"/> value.</summary>
        public SixLabors.ImageSharp.Color LeftColor {
            get => Helpers.ParseColor(LeftColorHex ?? throw new InvalidOperationException("LeftColorHex is null"));
            set => LeftColorHex = value.ToHexColor();
        }

        /// <summary>Gets or sets the style of the left border.</summary>
        public BorderValues? LeftStyle {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Val?.Value;
            set => (EnsurePageBorders().LeftBorder ??= new LeftBorder()).Val = value;
        }

        /// <summary>Gets or sets the space between the left border and page text.</summary>
        public UInt32Value? LeftSpace {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Space;
            set => (EnsurePageBorders().LeftBorder ??= new LeftBorder()).Space = value;
        }

        /// <summary>Gets or sets a value indicating whether the left border has a shadow.</summary>
        public bool? LeftShadow {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Shadow?.Value;
            set => (EnsurePageBorders().LeftBorder ??= new LeftBorder()).Shadow = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        /// <summary>Gets or sets a value indicating whether the left border is part of a frame.</summary>
        public bool? LeftFrame {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.LeftBorder?.Frame?.Value;
            set => (EnsurePageBorders().LeftBorder ??= new LeftBorder()).Frame = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        /// <summary>Gets or sets the width of the right border.</summary>
        public UInt32Value? RightSize {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Size;
            set => (EnsurePageBorders().RightBorder ??= new RightBorder()).Size = value;
        }

        /// <summary>Gets or sets the right border color using a hexadecimal value.</summary>
        public string? RightColorHex {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Color?.Value?.Replace("#", "").ToUpperInvariant();
            set => (EnsurePageBorders().RightBorder ??= new RightBorder()).Color = NormalizeColor(value);
        }

        /// <summary>Gets or sets the right border color using a <see cref="SixLabors.ImageSharp.Color"/> value.</summary>
        public SixLabors.ImageSharp.Color RightColor {
            get => Helpers.ParseColor(RightColorHex ?? throw new InvalidOperationException("RightColorHex is null"));
            set => RightColorHex = value.ToHexColor();
        }

        /// <summary>Gets or sets the style of the right border.</summary>
        public BorderValues? RightStyle {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Val?.Value;
            set => (EnsurePageBorders().RightBorder ??= new RightBorder()).Val = value;
        }

        /// <summary>Gets or sets the space between the right border and page text.</summary>
        public UInt32Value? RightSpace {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Space;
            set => (EnsurePageBorders().RightBorder ??= new RightBorder()).Space = value;
        }

        /// <summary>Gets or sets a value indicating whether the right border has a shadow.</summary>
        public bool? RightShadow {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Shadow?.Value;
            set => (EnsurePageBorders().RightBorder ??= new RightBorder()).Shadow = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        /// <summary>Gets or sets a value indicating whether the right border is part of a frame.</summary>
        public bool? RightFrame {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.RightBorder?.Frame?.Value;
            set => (EnsurePageBorders().RightBorder ??= new RightBorder()).Frame = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        /// <summary>Gets or sets the width of the top border.</summary>
        public UInt32Value? TopSize {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Size;
            set => (EnsurePageBorders().TopBorder ??= new TopBorder()).Size = value;
        }

        /// <summary>Gets or sets the top border color using a hexadecimal value.</summary>
        public string? TopColorHex {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Color?.Value?.Replace("#", "").ToUpperInvariant();
            set => (EnsurePageBorders().TopBorder ??= new TopBorder()).Color = NormalizeColor(value);
        }

        /// <summary>Gets or sets the top border color using a <see cref="SixLabors.ImageSharp.Color"/> value.</summary>
        public SixLabors.ImageSharp.Color TopColor {
            get => Helpers.ParseColor(TopColorHex ?? throw new InvalidOperationException("TopColorHex is null"));
            set => TopColorHex = value.ToHexColor();
        }

        /// <summary>Gets or sets the style of the top border.</summary>
        public BorderValues? TopStyle {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Val?.Value;
            set => (EnsurePageBorders().TopBorder ??= new TopBorder()).Val = value;
        }

        /// <summary>Gets or sets the space between the top border and page text.</summary>
        public UInt32Value? TopSpace {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Space;
            set => (EnsurePageBorders().TopBorder ??= new TopBorder()).Space = value;
        }

        /// <summary>Gets or sets a value indicating whether the top border has a shadow.</summary>
        public bool? TopShadow {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Shadow?.Value;
            set => (EnsurePageBorders().TopBorder ??= new TopBorder()).Shadow = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        /// <summary>Gets or sets a value indicating whether the top border is part of a frame.</summary>
        public bool? TopFrame {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.TopBorder?.Frame?.Value;
            set => (EnsurePageBorders().TopBorder ??= new TopBorder()).Frame = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        /// <summary>Gets or sets the width of the bottom border.</summary>
        public UInt32Value? BottomSize {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Size;
            set => (EnsurePageBorders().BottomBorder ??= new BottomBorder()).Size = value;
        }

        /// <summary>Gets or sets the bottom border color using a hexadecimal value.</summary>
        public string? BottomColorHex {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Color?.Value?.Replace("#", "").ToUpperInvariant();
            set => (EnsurePageBorders().BottomBorder ??= new BottomBorder()).Color = NormalizeColor(value);
        }

        /// <summary>Gets or sets the bottom border color using a <see cref="SixLabors.ImageSharp.Color"/> value.</summary>
        public SixLabors.ImageSharp.Color BottomColor {
            get => Helpers.ParseColor(BottomColorHex ?? throw new InvalidOperationException("BottomColorHex is null"));
            set => BottomColorHex = value.ToHexColor();
        }

        /// <summary>Gets or sets the style of the bottom border.</summary>
        public BorderValues? BottomStyle {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Val?.Value;
            set => (EnsurePageBorders().BottomBorder ??= new BottomBorder()).Val = value;
        }

        /// <summary>Gets or sets the space between the bottom border and page text.</summary>
        public UInt32Value? BottomSpace {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Space;
            set => (EnsurePageBorders().BottomBorder ??= new BottomBorder()).Space = value;
        }

        /// <summary>Gets or sets a value indicating whether the bottom border has a shadow.</summary>
        public bool? BottomShadow {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Shadow?.Value;
            set => (EnsurePageBorders().BottomBorder ??= new BottomBorder()).Shadow = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        /// <summary>Gets or sets a value indicating whether the bottom border is part of a frame.</summary>
        public bool? BottomFrame {
            get => _section._sectionProperties.GetFirstChild<PageBorders>()?.BottomBorder?.Frame?.Value;
            set => (EnsurePageBorders().BottomBorder ??= new BottomBorder()).Frame = value.HasValue ? new OnOffValue(value.Value) : null;
        }

        internal void SetBorder(WordBorder wordBorder) {
            var pageBorderSettings = GetDefault(wordBorder);
            var existing = _section._sectionProperties.GetFirstChild<PageBorders>();

            if (pageBorderSettings is null) {
                existing?.Remove();
            } else if (existing is null) {
                _section._sectionProperties.Append(pageBorderSettings);
            } else {
                existing.Remove();
                _section._sectionProperties.Append(pageBorderSettings);
            }
        }

        /// <summary>Gets or sets the preset border configuration applied to the section.</summary>
        public WordBorder Type {
            get {
                var pageBorder = _section._sectionProperties.GetFirstChild<PageBorders>();
                if (pageBorder is null) {
                    return WordBorder.None;
                }

                foreach (WordBorder wordBorder in Enum.GetValues(typeof(WordBorder))) {
                    if (wordBorder is WordBorder.None) {
                        continue;
                    }

                    var builtin = GetDefault(wordBorder);
                    if (builtin is null) {
                        continue;
                    }

                    if ((builtin.LeftBorder == null && pageBorder.LeftBorder == null) &&
                        (builtin.RightBorder == null && pageBorder.RightBorder == null) &&
                        (builtin.TopBorder == null && pageBorder.TopBorder == null) &&
                        (builtin.BottomBorder == null && pageBorder.BottomBorder == null)) {
                        return wordBorder;
                    }

                    if (builtin.LeftBorder != null && pageBorder.LeftBorder != null &&
                        builtin.RightBorder != null && pageBorder.RightBorder != null &&
                        builtin.TopBorder != null && pageBorder.TopBorder != null &&
                        builtin.BottomBorder != null && pageBorder.BottomBorder != null &&
                        builtin.LeftBorder.Shadow == pageBorder.LeftBorder.Shadow &&
                        builtin.RightBorder.Shadow == pageBorder.RightBorder.Shadow &&
                        builtin.TopBorder.Shadow == pageBorder.TopBorder.Shadow &&
                        builtin.BottomBorder.Shadow == pageBorder.BottomBorder.Shadow &&
                        builtin.LeftBorder.Color == pageBorder.LeftBorder.Color &&
                        builtin.RightBorder.Color == pageBorder.RightBorder.Color &&
                        builtin.TopBorder.Color == pageBorder.TopBorder.Color &&
                        builtin.BottomBorder.Color == pageBorder.BottomBorder.Color &&
                        builtin.LeftBorder.Size == pageBorder.LeftBorder.Size &&
                        builtin.RightBorder.Size == pageBorder.RightBorder.Size &&
                        builtin.TopBorder.Size == pageBorder.TopBorder.Size &&
                        builtin.BottomBorder.Size == pageBorder.BottomBorder.Size &&
                        builtin.LeftBorder.Space == pageBorder.LeftBorder.Space &&
                        builtin.RightBorder.Space == pageBorder.RightBorder.Space &&
                        builtin.TopBorder.Space == pageBorder.TopBorder.Space &&
                        builtin.BottomBorder.Space == pageBorder.BottomBorder.Space) {
                        return wordBorder;
                    }
                }

                return WordBorder.Custom;
            }
            set => SetBorder(value);
        }

        private static PageBorders? GetDefault(WordBorder border) => border switch {
            WordBorder.Box => Box,
            WordBorder.Shadow => Shadow,
            WordBorder.None => null,
            WordBorder.Custom => Custom,
            _ => throw new ArgumentOutOfRangeException(nameof(border))
        };

        private static PageBorders Custom => new();

        private static PageBorders Box {
            get {
                PageBorders pageBorders1 = new() { OffsetFrom = PageBorderOffsetValues.Page };
                TopBorder topBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                LeftBorder leftBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                BottomBorder bottomBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };
                RightBorder rightBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U };

                pageBorders1.Append(topBorder1);
                pageBorders1.Append(leftBorder1);
                pageBorders1.Append(bottomBorder1);
                pageBorders1.Append(rightBorder1);
                return pageBorders1;
            }
        }

        private static PageBorders Shadow {
            get {
                PageBorders pageBorders1 = new() { OffsetFrom = PageBorderOffsetValues.Page };
                TopBorder topBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                LeftBorder leftBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                BottomBorder bottomBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };
                RightBorder rightBorder1 = new() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)24U, Shadow = true };

                pageBorders1.Append(topBorder1);
                pageBorders1.Append(leftBorder1);
                pageBorders1.Append(bottomBorder1);
                pageBorders1.Append(rightBorder1);
                return pageBorders1;
            }
        }
    }
}

