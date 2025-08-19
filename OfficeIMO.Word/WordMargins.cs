using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace OfficeIMO.Word {
    /// <summary>
    /// Predefined margin configurations available for a document section.
    /// </summary>
    public enum WordMargin {
        /// <summary>
        /// Standard one inch margins on all sides.
        /// </summary>
        Normal,
        /// <summary>
        /// Mirrored inside and outside margins for book binding.
        /// </summary>
        Mirrored,
        /// <summary>
        /// Slightly smaller left and right margins.
        /// </summary>
        Moderate,
        /// <summary>
        /// Minimal margins on all sides.
        /// </summary>
        Narrow,
        /// <summary>
        /// Extra wide margins on left and right.
        /// </summary>
        Wide,
        /// <summary>
        /// Margin preset used by Word 2003.
        /// </summary>
        Office2003Default,
        /// <summary>
        /// Custom margins not represented by other presets.
        /// </summary>
        Unknown
    }

    /// <summary>
    /// Provides access to page margin settings for a document section.
    /// </summary>
    public class WordMargins {
        private readonly WordDocument _document;
        private readonly WordSection _section;

        /// <summary>
        /// Initializes a new instance for managing margins within the specified section.
        /// </summary>
        /// <param name="wordDocument">Parent document.</param>
        /// <param name="wordSection">Section associated with the margins.</param>
        public WordMargins(WordDocument wordDocument, WordSection wordSection) {
            _document = wordDocument;
            _section = wordSection;
        }

        /// <summary>
        /// Get or set the left margin in Twips
        /// </summary>
        public UInt32Value Left {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin?.Left != null) {
                    return pageMargin.Left!;
                }

                return WordMargins.Normal.Left!;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(WordMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin!.Left = value;
            }
        }

        /// <summary>
        /// Get or set the right margin in Twips
        /// </summary>
        public UInt32Value Right {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin?.Right != null) {
                    return pageMargin.Right!;
                }

                return WordMargins.Normal.Right!;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(WordMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin!.Right = value;
            }
        }

        /// <summary>
        /// Get or set the top margin in Twips
        /// </summary>
        public int? Top {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                var top = pageMargin?.Top?.Value;
                return top ?? WordMargins.Normal.Top!.Value;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(WordMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin!.Top = value;
            }
        }

        /// <summary>
        /// Get or set the left margin in Twips
        /// </summary>
        public int? Bottom {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                var bottom = pageMargin?.Bottom?.Value;
                return bottom ?? WordMargins.Normal.Bottom!.Value;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(WordMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin!.Bottom = value;
            }
        }

        /// <summary>
        /// Get or set the top margin in centimeters
        /// </summary>
        public double? TopCentimeters {
            get {
                var top = Top;
                return top.HasValue ? Helpers.ConvertTwipsToCentimeters(top.Value) : (double?)null;
            }
            set {
                Top = value.HasValue ? Helpers.ConvertCentimetersToTwips(value.Value) : (int?)null;
            }
        }

        /// <summary>
        /// Get or set the bottom margin in centimeters
        /// </summary>
        public double? BottomCentimeters {
            get {
                var bottom = Bottom;
                return bottom.HasValue ? Helpers.ConvertTwipsToCentimeters(bottom.Value) : (double?)null;
            }
            set {
                Bottom = value.HasValue ? Helpers.ConvertCentimetersToTwips(value.Value) : (int?)null;
            }
        }

        /// <summary>
        /// Get or set the left margin in centimeters
        /// </summary>
        public double? LeftCentimeters {
            get {
                var left = Left;
                return left != null ? Helpers.ConvertTwipsToCentimeters(left.Value) : (double?)null;
            }
            set {
                if (value.HasValue) {
                    Left = Helpers.ConvertCentimetersToTwipsUInt32(value.Value);
                }
            }
        }

        /// <summary>
        /// Get or set the right margin in centimeters
        /// </summary>
        public double? RightCentimeters {
            get {
                var right = Right;
                return right != null ? Helpers.ConvertTwipsToCentimeters(right.Value) : (double?)null;
            }
            set {
                if (value.HasValue) {
                    Right = Helpers.ConvertCentimetersToTwipsUInt32(value.Value);
                }
            }
        }

        /// <summary>
        /// Get or set the header distance in Twips
        /// </summary>
        public UInt32Value HeaderDistance {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin?.Header != null) {
                    return pageMargin.Header!;
                }

                return WordMargins.Normal.Header!;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(WordMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin!.Header = value;
            }
        }

        /// <summary>
        /// Get or set the footer distance in Twips
        /// </summary>
        public UInt32Value FooterDistance {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin?.Footer != null) {
                    return pageMargin.Footer!;
                }

                return WordMargins.Normal.Footer!;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(WordMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin!.Footer = value;
            }
        }

        /// <summary>
        /// Gets or sets the gutter size in Twips.
        /// </summary>
        public UInt32Value Gutter {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin?.Gutter != null) {
                    return pageMargin.Gutter!;
                }

                return WordMargins.Normal.Gutter!;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(WordMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin!.Gutter = value;
            }
        }

        /// <summary>
        /// Gets or sets a predefined margin set that matches the current settings.
        /// </summary>
        public WordMargin? Type {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    foreach (WordMargin wordMargin in Enum.GetValues(typeof(WordMargin))) {
                        if (wordMargin == WordMargin.Unknown) {
                            continue;
                        }

                        var pageMarginBuiltin = GetDefault(wordMargin);
                        if (pageMarginBuiltin.Bottom == null &&
                            pageMarginBuiltin.Footer == null &&
                            pageMarginBuiltin.Gutter == null &&
                            pageMarginBuiltin.Header == null &&
                            pageMarginBuiltin.Left == null &&
                            pageMarginBuiltin.Right == null &&
                            pageMarginBuiltin.Top == null) {
                            return wordMargin;
                        }

                        if (pageMarginBuiltin.Bottom != null && pageMargin.Bottom != null &&
                            pageMarginBuiltin.Footer != null && pageMargin.Footer != null &&
                            pageMarginBuiltin.Gutter != null && pageMargin.Gutter != null &&
                            pageMarginBuiltin.Header != null && pageMargin.Header != null &&
                            pageMarginBuiltin.Left != null && pageMargin.Left != null &&
                            pageMarginBuiltin.Right != null && pageMargin.Right != null &&
                            pageMarginBuiltin.Top != null && pageMargin.Top != null &&


                            pageMarginBuiltin.Bottom == pageMargin.Bottom &&
                            pageMarginBuiltin.Footer == pageMargin.Footer &&
                            pageMarginBuiltin.Gutter == pageMargin.Gutter &&
                             pageMarginBuiltin.Header == pageMargin.Header &&
                            pageMarginBuiltin.Left == pageMargin.Left &&
                            pageMarginBuiltin.Right == pageMargin.Right &&
                             pageMarginBuiltin.Top == pageMargin.Top
                            ) {
                            return wordMargin;
                        }
                    }

                    return WordMargin.Unknown;
                }

                return WordMargin.Normal;
            }
            set => SetMargins(value);
        }

        private void SetMargins(WordMargin? wordMargin) {
            if (wordMargin == null) {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    pageMargin.Remove();
                }
            } else {
                var pageMarginData = GetDefault(wordMargin);
                if (pageMarginData != null) {
                    var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                    if (pageMargin == null) {
                        _section._sectionProperties.Append(pageMarginData);
                        // pageMargin = _sectionProperties.GetFirstChild<PageMargin>();
                    } else {
                        pageMargin.Remove();
                        _section._sectionProperties.Append(pageMarginData);
                    }
                }
            }
        }

        /// <summary>
        /// Applies a predefined margin configuration to the given section.
        /// </summary>
        /// <param name="wordSection">Target section to update.</param>
        /// <param name="pageMargins">Predefined margin set to apply.</param>
        /// <returns>The updated section.</returns>
        public static WordSection SetMargins(WordSection wordSection, WordMargin pageMargins) {
            var pageMarginData = GetDefault(pageMargins);

            var pageMargin = wordSection._sectionProperties.GetFirstChild<PageMargin>();
            if (pageMargin == null) {
                wordSection._sectionProperties.Append(pageMarginData);
                // pageMargin = _sectionProperties.GetFirstChild<PageMargin>();
            } else {
                pageMargin.Remove();
                wordSection._sectionProperties.Append(pageMarginData);
            }

            return wordSection;
        }


        private static PageMargin GetDefault(WordMargin? pageMargins) {
            switch (pageMargins) {
                case WordMargin.Mirrored: return Mirrored;
                case WordMargin.Normal: return Normal;
                case WordMargin.Moderate: return Moderate;
                case WordMargin.Narrow: return Narrow;
                case WordMargin.Office2003Default: return Office2003Default;
                case WordMargin.Wide: return Wide;
            }

            throw new ArgumentOutOfRangeException(nameof(pageMargins));
        }

        private static PageMargin Normal {
            get {
                return new PageMargin() {
                    Top = 1440,
                    Right = (UInt32Value)1440U,
                    Bottom = 1440,
                    Left = (UInt32Value)1440U,
                    Header = (UInt32Value)720U,
                    Footer = (UInt32Value)720U,
                    Gutter = (UInt32Value)0U
                };
            }
        }

        private static PageMargin Mirrored {
            get {
                return new PageMargin() {
                    Top = 1440,
                    Right = (UInt32Value)1440U,
                    Bottom = 1440,
                    Left = (UInt32Value)1800U,
                    Header = (UInt32Value)720U,
                    Footer = (UInt32Value)720U,
                    Gutter = (UInt32Value)0U
                };
            }
        }

        private static PageMargin Moderate => new PageMargin() {
            Top = 1440,
            Right = (UInt32Value)1080U,
            Bottom = 1440,
            Left = (UInt32Value)1080U,
            Header = (UInt32Value)720U,
            Footer = (UInt32Value)720U,
            Gutter = (UInt32Value)0U
        };

        private static PageMargin Narrow {
            get {
                return new PageMargin() {
                    Top = 720,
                    Right = (UInt32Value)720U,
                    Bottom = 720,
                    Left = (UInt32Value)720U,
                    Header = (UInt32Value)720U,
                    Footer = (UInt32Value)720U,
                    Gutter = (UInt32Value)0U
                };
            }
        }

        private static PageMargin Wide {
            get {
                return new PageMargin() {
                    Top = 1440,
                    Right = (UInt32Value)2880U,
                    Bottom = 1440,
                    Left = (UInt32Value)2880U,
                    Header = (UInt32Value)720U,
                    Footer = (UInt32Value)720U,
                    Gutter = (UInt32Value)0U
                };
            }
        }

        private static PageMargin Office2003Default {
            get {
                return new PageMargin() {
                    Top = 1440,
                    Right = (UInt32Value)1800U,
                    Bottom = 1440,
                    Left = (UInt32Value)1800U,
                    Header = (UInt32Value)720U,
                    Footer = (UInt32Value)720U,
                    Gutter = (UInt32Value)0U
                };
            }
        }
    }
}