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
        /// JIS B5 paper, 182 × 257 millimeters. This retains the established B5 preset.
        /// </summary>
        B5,
        /// <summary>US Tabloid, 11 × 17 inches.</summary>
        Tabloid,
        /// <summary>JIS B4, 257 × 364 millimeters.</summary>
        B4Jis,
        /// <summary>Number 9 envelope, 3.875 × 8.875 inches.</summary>
        Envelope9,
        /// <summary>Number 10 envelope, 4.125 × 9.5 inches.</summary>
        Envelope10,
        /// <summary>US C sheet, 17 × 22 inches.</summary>
        CSheet,
        /// <summary>DL envelope, 110 × 220 millimeters.</summary>
        EnvelopeDl,
        /// <summary>C5 envelope, 162 × 229 millimeters.</summary>
        EnvelopeC5,
        /// <summary>C4 envelope, 229 × 324 millimeters.</summary>
        EnvelopeC4,
        /// <summary>B5 envelope, 176 × 250 millimeters.</summary>
        EnvelopeB5,
        /// <summary>Monarch envelope, 3.875 × 7.5 inches.</summary>
        EnvelopeMonarch
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
    public partial class WordPageSizes {
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

                        // Printer codes are optional in producer DOCX files and native DOC
                        // imports. A one-twip tolerance accepts producer unit rounding.
                        if ((pageSize.Code == null || pageSizeBuiltin.Code == pageSize.Code) &&
                            ((PageDimensionMatches(pageSizeBuiltin.Width, pageSize.Width) &&
                              PageDimensionMatches(pageSizeBuiltin.Height, pageSize.Height)) ||
                             (PageDimensionMatches(pageSizeBuiltin.Width, pageSize.Height) &&
                              PageDimensionMatches(pageSizeBuiltin.Height, pageSize.Width)))) {
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

        private static bool PageDimensionMatches(DocumentFormat.OpenXml.UInt32Value? expected, DocumentFormat.OpenXml.UInt32Value? actual) =>
            expected != null && actual != null && Math.Abs((long)expected.Value - actual.Value) <= 1L;

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
                _section._sectionProperties.AddChild(pageSizeSettings, true);
                return;
            }

            bool requiresPageOrient = false;
            PageOrientationValues pageOrientation = PageOrientationValues.Portrait;
            if (pageSize.Orient != null && pageSize.Orient.Value != PageOrientationValues.Portrait) {
                pageOrientation = pageSize.Orient.Value;
                requiresPageOrient = true;
            }

            pageSize.Remove();
            _section._sectionProperties.AddChild(pageSizeSettings, true);

            if (requiresPageOrient) {
                SetOrientation(_section._sectionProperties, pageOrientation);
            }
        }

        private PageSize EnsurePageSize() {
            var pageSize = _section._sectionProperties.GetFirstChild<PageSize>();
            if (pageSize == null) {
                pageSize = new PageSize();
                _section._sectionProperties.AddChild(pageSize, true);
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
                sectionProperties.AddChild(pageSize, true);
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

        private static PageSize ToOpenXmlPageSize(WordPageSizeDefinition definition) =>
            new PageSize {
                Width = definition.WidthTwips,
                Height = definition.HeightTwips,
                Code = definition.PaperCode
            };
    }
}
