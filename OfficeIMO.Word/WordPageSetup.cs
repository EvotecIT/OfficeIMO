using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Page size, margins, orientation and running headers and footers of one section.
    /// </summary>
    /// <remarks>
    /// The setup is a view over the section's properties; it holds no state of its own.
    /// <see cref="Header"/>, <see cref="Footer"/> and their first-page and even-page variants
    /// create the requested story when it is missing. Use <see cref="GetHeader"/> and
    /// <see cref="GetFooter"/> to inspect a story without creating it.
    /// </remarks>
    public sealed class WordPageSetup {
        private readonly WordSection _section;

        internal WordPageSetup(WordSection section) {
            _section = section;
        }

        /// <summary>
        /// Gets the page size settings of the section.
        /// </summary>
        public WordPageSizes Size => _section._pageSizes;

        /// <summary>
        /// Gets the page margins of the section.
        /// </summary>
        public WordMargins Margins => _section._margins;

        /// <summary>
        /// Gets or sets the page orientation of the section.
        /// </summary>
        public OfficePageOrientation Orientation {
            get => WordPageSizes.GetOrientation(_section._sectionProperties).ToOfficeEnum();
            set => WordPageSizes.SetOrientation(_section._sectionProperties, value.ToOpenXml());
        }

        /// <summary>
        /// Gets or sets whether the first page of the section uses its own header and footer.
        /// Setting the value to <see langword="true"/> also creates the first-page header and footer
        /// when they are missing.
        /// </summary>
        public bool DifferentFirstPage {
            get {
                var titlePage = _section._sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
                return titlePage != null && (titlePage.Val?.Value ?? true);
            }
            set {
                var sectionProperties = _section._sectionProperties;
                var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
                if (!value) {
                    titlePage?.Remove();
                    return;
                }

                if (titlePage == null) sectionProperties.Append(new TitlePage());
                else titlePage.Val = true;
                WordHeadersAndFooters.AddHeaderReference(_section._document, _section, HeaderFooterValues.First);
                WordHeadersAndFooters.AddFooterReference(_section._document, _section, HeaderFooterValues.First);
            }
        }

        /// <summary>
        /// Gets whether this section has an explicit or inherited even-page header or footer while the
        /// document-wide odd/even setting is enabled. Setting the value to
        /// <see langword="true"/> creates even-page references for this section and enables that setting;
        /// setting it to <see langword="false"/> disables the document-wide setting.
        /// </summary>
        public bool DifferentOddAndEvenPages {
            get {
                if (!_section.DocumentOddEvenSettingEnabled) return false;
                return _section.ResolveEvenHeader() != null || _section.ResolveEvenFooter() != null;
            }
            set {
                var mainPart = _section._document._wordprocessingDocument.MainDocumentPart;
                if (!value) {
                    mainPart?.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>()?.Remove();
                    return;
                }

                _ = _section._document.Settings;
                WordHeadersAndFooters.AddHeaderReference(_section._document, _section, HeaderFooterValues.Even);
                WordHeadersAndFooters.AddFooterReference(_section._document, _section, HeaderFooterValues.Even);

                Settings settings = mainPart!.DocumentSettingsPart!.Settings!;
                EvenAndOddHeaders? setting = settings.GetFirstChild<EvenAndOddHeaders>();
                if (setting == null) settings.AddChild(new EvenAndOddHeaders(), true);
                else setting.Val = true;
            }
        }

        /// <summary>
        /// Gets the default header of the section, creating the default header and footer when missing.
        /// </summary>
        public WordHeader Header => GetOrCreateHeader(WordHeaderFooterType.Default);

        /// <summary>
        /// Gets the first-page header of the section. When it is missing, this enables
        /// <see cref="DifferentFirstPage"/>, which creates the first-page header and footer.
        /// </summary>
        public WordHeader FirstPageHeader => GetOrCreateHeader(WordHeaderFooterType.First);

        /// <summary>
        /// Gets the even-page header of the section. When it is missing, this enables
        /// <see cref="DifferentOddAndEvenPages"/>, which creates the even-page header and footer
        /// and turns on the document-wide odd/even setting.
        /// </summary>
        public WordHeader EvenPageHeader => GetOrCreateHeader(WordHeaderFooterType.Even);

        /// <summary>
        /// Gets the default footer of the section, creating the default header and footer when missing.
        /// </summary>
        public WordFooter Footer => GetOrCreateFooter(WordHeaderFooterType.Default);

        /// <summary>
        /// Gets the first-page footer of the section. When it is missing, this enables
        /// <see cref="DifferentFirstPage"/>, which creates the first-page header and footer.
        /// </summary>
        public WordFooter FirstPageFooter => GetOrCreateFooter(WordHeaderFooterType.First);

        /// <summary>
        /// Gets the even-page footer of the section. When it is missing, this enables
        /// <see cref="DifferentOddAndEvenPages"/>, which creates the even-page header and footer
        /// and turns on the document-wide odd/even setting.
        /// </summary>
        public WordFooter EvenPageFooter => GetOrCreateFooter(WordHeaderFooterType.Even);

        /// <summary>
        /// Returns the section's own header of the requested type, or <see langword="null"/> when the
        /// section has none. Never creates a header and does not resolve headers inherited from
        /// earlier sections.
        /// </summary>
        public WordHeader? GetHeader(WordHeaderFooterType type) {
            var headers = _section.Headers;
            if (type == WordHeaderFooterType.First) return headers.First;
            if (type == WordHeaderFooterType.Even) return headers.Even;
            return headers.Default;
        }

        /// <summary>
        /// Returns the section's own footer of the requested type, or <see langword="null"/> when the
        /// section has none. Never creates a footer and does not resolve footers inherited from
        /// earlier sections.
        /// </summary>
        public WordFooter? GetFooter(WordHeaderFooterType type) {
            var footers = _section.Footers;
            if (type == WordHeaderFooterType.First) return footers.First;
            if (type == WordHeaderFooterType.Even) return footers.Even;
            return footers.Default;
        }

        internal WordHeader GetOrCreateHeader(WordHeaderFooterType type) {
            var headers = _section.Headers;
            if (type == WordHeaderFooterType.Default) {
                if (headers.Default == null) AddDefaultHeaderAndFooter();
                return headers.Default ?? throw new InvalidOperationException("Failed to create default header.");
            }

            if (type == WordHeaderFooterType.First) {
                if (headers.First == null) DifferentFirstPage = true;
                return headers.First ?? throw new InvalidOperationException("Failed to create first-page header.");
            }

            if (type == WordHeaderFooterType.Even) {
                if (headers.Even == null) DifferentOddAndEvenPages = true;
                return headers.Even ?? throw new InvalidOperationException("Failed to create even-page header.");
            }

            throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
        }

        internal WordFooter GetOrCreateFooter(WordHeaderFooterType type) {
            var footers = _section.Footers;
            if (type == WordHeaderFooterType.Default) {
                if (footers.Default == null) AddDefaultHeaderAndFooter();
                return footers.Default ?? throw new InvalidOperationException("Failed to create default footer.");
            }

            if (type == WordHeaderFooterType.First) {
                if (footers.First == null) DifferentFirstPage = true;
                return footers.First ?? throw new InvalidOperationException("Failed to create first-page footer.");
            }

            if (type == WordHeaderFooterType.Even) {
                if (footers.Even == null) DifferentOddAndEvenPages = true;
                return footers.Even ?? throw new InvalidOperationException("Failed to create even-page footer.");
            }

            throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");
        }

        private void AddDefaultHeaderAndFooter() {
            WordHeadersAndFooters.AddHeaderReference(_section._document, _section, HeaderFooterValues.Default);
            WordHeadersAndFooters.AddFooterReference(_section._document, _section, HeaderFooterValues.Default);
        }
    }
}
