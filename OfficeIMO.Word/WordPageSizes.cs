using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// PaperSizes used in Microsoft Office
    /// </summary>
    public enum WordPageSize {
        /// <summary>
        /// Custom/unknown paper size that is not defined within OfficeIMO
        /// </summary>
        Unknown,
        /// <summary>
        /// Letter is part of the North American loose paper size series.
        /// It is the standard for business and academic documents, and measures 216 × 279 mm or 8.5 × 11 inches.
        /// </summary>
        Letter,
        /// <summary>
        /// Legal is part of the North American loose paper size series.
        /// It is used to make legal pads, and measures 216 × 356 mm or 8.5 × 14 inches.
        /// </summary>
        Legal,
        /// <summary>
        /// Statement paper size 5.5 x 8.5 inches
        /// </summary>
        Statement,
        /// <summary>
        /// Executive paper size 7.25 x 10.50 inches.
        /// </summary>
        Executive,
        /// <summary>
        /// An A3 piece of paper measures 297 × 420 mm or 11.7 × 16.5 inches.
        /// </summary>
        A3,
        /// <summary>
        /// An A4 piece of paper measures 210 × 297 mm or 8.3 × 11.7 inches
        /// </summary>
        A4,
        /// <summary>
        /// An A5 piece of paper measures 148 × 210 mm or 5.8 × 8.3 inches.
        /// </summary>
        A5,
        /// <summary>
        /// An A6 piece of paper measures 105 × 148 mm or 4.1 × 5.8 inches.
        /// </summary>
        A6,
        /// <summary>
        /// A B5 piece of paper measures 176 × 250 mm or 6.9 × 9.8 inches.
        /// </summary>
        B5
    }

    /// <summary>Dimensions and Word paper code for a built-in page-size preset.</summary>
    public sealed class WordPageSizeDefinition {
        internal WordPageSizeDefinition(uint widthTwips, uint heightTwips, ushort paperCode) {
            WidthTwips = widthTwips;
            HeightTwips = heightTwips;
            PaperCode = paperCode;
        }

        /// <summary>Gets the page width in twips.</summary>
        public uint WidthTwips { get; }

        /// <summary>Gets the page height in twips.</summary>
        public uint HeightTwips { get; }

        /// <summary>Gets the Word paper-size code.</summary>
        public ushort PaperCode { get; }
    }

    /// <summary>
    /// Provides helpers for manipulating Word page size and orientation.
    /// </summary>
    public class WordPageSizes {
        private readonly WordSection _section;
        private readonly WordDocument _document;

        /// <summary>
        /// This element specifies the properties (size and orientation) for all pages in the current section.
        /// </summary>
        public WordPageSize? PageSize {
            get {
                var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();
                if (pageSize != null) {
                    foreach (WordPageSize wordPageSize in global::OfficeIMO.Internal.EnumCompat.GetValues<WordPageSize>()) {
                        if (wordPageSize == WordPageSize.Unknown) {
                            continue;
                        }

                        var pageSizeBuiltin = GetDefault(wordPageSize);
                        if (pageSizeBuiltin == null) {
                            continue;
                        }

                        if ((pageSizeBuiltin.Width == null && pageSize.Width == null) &&
                            (pageSizeBuiltin.Height == null && pageSize.Height == null) &&
                            (pageSizeBuiltin.Code == null && pageSize.Code == null)) {
                            return wordPageSize;
                        }

                        if (pageSizeBuiltin.Width != null && pageSize.Width != null &&
                            pageSizeBuiltin.Height != null && pageSize.Height != null &&
                            pageSizeBuiltin.Code != null && pageSize.Code != null &&
                            pageSizeBuiltin.Width == pageSize.Width &&
                            pageSizeBuiltin.Height == pageSize.Height &&
                            pageSizeBuiltin.Code == pageSize.Code) {
                            return wordPageSize;
                        }

                        if (pageSizeBuiltin.Width != null && pageSize.Width != null &&
                            pageSizeBuiltin.Height != null && pageSize.Height != null &&
                            pageSizeBuiltin.Code != null && pageSize.Code != null &&
                            pageSizeBuiltin.Width == pageSize.Height &&
                            pageSizeBuiltin.Height == pageSize.Width &&
                            pageSizeBuiltin.Code == pageSize.Code) {
                            return wordPageSize;
                        }
                    }
                    return WordPageSize.Unknown;
                } else {
                    return null;
                }
            }
            set => SetPageSize(value);
        }

        private void SetPageSize(WordPageSize? wordPageSize) {
            var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();

            if (wordPageSize == null) {
                pageSize?.Remove();
                return;
            }

            var pageSizeSettings = GetDefault(wordPageSize);
            if (pageSizeSettings == null) {
                pageSize?.Remove();
                return;
            }

            if (pageSize == null) {
                _section._sectionProperties.Append(pageSizeSettings);
                return;
            }

            bool requiresPageOrient = false;
            PageOrientationValues pageOrientation = PageOrientationValues.Portrait;
            if (pageSize.Orient != null && pageSize.Orient.Value != PageOrientationValues.Portrait) {
                pageOrientation = pageSize.Orient.Value;
                requiresPageOrient = true;
            }

            pageSize.Remove();
            _section._sectionProperties.Append(pageSizeSettings);

            if (requiresPageOrient) {
                SetOrientation(_section._sectionProperties, pageOrientation);
            }
        }

        private PageSize EnsurePageSize() {
            var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();
            if (pageSize == null) {
                pageSize = new PageSize();
                _section._sectionProperties.Append(pageSize);
            }
            return pageSize;
        }

        /// <summary>
        /// Get or Set section/page Width
        /// </summary>
        public uint? Width {
            get {
                var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();
                return pageSize?.Width?.Value;
            }
            set {
                var pageSize = EnsurePageSize();
                pageSize.Width = value;
            }
        }

        /// <summary>
        /// Get or Set section/page Height
        /// </summary>
        public uint? Height {
            get {
                var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();
                return pageSize?.Height?.Value;
            }
            set {
                var pageSize = EnsurePageSize();
                pageSize.Height = value;
            }
        }

        /// <summary>
        /// Get or Set section/page Code
        /// </summary>
        public ushort? Code {
            get {
                var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();
                return pageSize?.Code?.Value;
            }
            set {
                var pageSize = EnsurePageSize();
                pageSize.Code = value;
            }
        }

        internal static PageOrientationValues GetOrientation(SectionProperties sectionProperties) {
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            if (pageSize == null) {
                return PageOrientationValues.Portrait;
            }

            if (pageSize.Orient != null) {
                return pageSize.Orient.Value;
            }

            return PageOrientationValues.Portrait;
        }

        internal static void SetOrientation(SectionProperties sectionProperties, PageOrientationValues pageOrientationValue) {
            var pageSize = sectionProperties.Descendants<PageSize>().FirstOrDefault();
            if (pageSize == null) {
                // we need to setup default values for A4 
                pageSize = ToOpenXmlPageSize(WordPageSizes.A4);
                pageSize.Orient = PageOrientationValues.Portrait;
                sectionProperties.Append(pageSize);
            }
            if (pageSize.Orient == null) {
                pageSize.Orient = PageOrientationValues.Portrait;
            }
            if (pageSize.Orient != pageOrientationValue) {
                // changing orientation is not enough, we need to change width with height and vice versa
                var width = pageSize.Width;
                var height = pageSize.Height;
                pageSize.Width = height;
                pageSize.Height = width;

                pageSize.Orient = pageOrientationValue;
            }
        }

        /// <summary>
        /// Get or Set section/page Orientation
        /// </summary>
        public OfficePageOrientation Orientation {
            get => GetOrientation(_section._sectionProperties).ToOfficeEnum();
            set => SetOrientation(_section._sectionProperties, value.ToOpenXml());
        }

        /// <summary>
        /// Manipulate section/page settings
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <param name="wordSection"></param>
        public WordPageSizes(WordDocument wordDocument, WordSection wordSection) {
            _section = wordSection;
            _document = wordDocument;
        }

        private static PageSize? GetDefault(WordPageSize? pageSize) {
            if (pageSize == null) {
                return null;
            }

            switch (pageSize) {
                case WordPageSize.A3: return ToOpenXmlPageSize(A3);
                case WordPageSize.A4: return ToOpenXmlPageSize(A4);
                case WordPageSize.A5: return ToOpenXmlPageSize(A5);
                case WordPageSize.Executive: return ToOpenXmlPageSize(Executive);
                case WordPageSize.Unknown: return null;
                case WordPageSize.A6: return ToOpenXmlPageSize(A6);
                case WordPageSize.B5: return ToOpenXmlPageSize(B5);
                case WordPageSize.Letter: return ToOpenXmlPageSize(Letter);
                case WordPageSize.Statement: return ToOpenXmlPageSize(Statement);
                case WordPageSize.Legal: return ToOpenXmlPageSize(Legal);
            }

            throw new ArgumentOutOfRangeException(nameof(pageSize));
        }

        /// <summary>
        /// Gets the default A3 page size.
        /// </summary>
        public static WordPageSizeDefinition A3 { get; } = new WordPageSizeDefinition(16838U, 23811U, 8);

        /// <summary>
        /// Gets the default A4 page size.
        /// </summary>
        public static WordPageSizeDefinition A4 { get; } = new WordPageSizeDefinition(11906U, 16838U, 9);

        /// <summary>
        /// Gets the default A5 page size.
        /// </summary>
        public static WordPageSizeDefinition A5 { get; } = new WordPageSizeDefinition(8391U, 11906U, 11);

        /// <summary>
        /// Gets the default Executive page size.
        /// </summary>
        public static WordPageSizeDefinition Executive { get; } = new WordPageSizeDefinition(10440U, 15120U, 7);

        /// <summary>
        /// Gets the default A6 page size.
        /// </summary>
        public static WordPageSizeDefinition A6 { get; } = new WordPageSizeDefinition(5953U, 8391U, 70);
        /// <summary>
        /// Gets the default B5 page size.
        /// </summary>
        public static WordPageSizeDefinition B5 { get; } = new WordPageSizeDefinition(10318U, 14570U, 13);
        /// <summary>
        /// Gets the default Statement page size.
        /// </summary>
        public static WordPageSizeDefinition Statement { get; } = new WordPageSizeDefinition(7920U, 12240U, 6);
        /// <summary>
        /// Gets the default Legal page size.
        /// </summary>
        public static WordPageSizeDefinition Legal { get; } = new WordPageSizeDefinition(12240U, 20160U, 5);
        /// <summary>
        /// Gets the default Letter page size.
        /// </summary>
        public static WordPageSizeDefinition Letter { get; } = new WordPageSizeDefinition(12240U, 15840U, 1);

        private static PageSize ToOpenXmlPageSize(WordPageSizeDefinition definition) =>
            new PageSize {
                Width = definition.WidthTwips,
                Height = definition.HeightTwips,
                Code = definition.PaperCode
            };
    }
}
