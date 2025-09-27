using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Diagnostics;

namespace OfficeIMO.Word {
    /// <summary>
    /// Manages document headers and footers.
    /// </summary>
    public partial class WordDocument {
        // internal header properties for easy usage
        internal Header? _headerFirst;
        internal Header? _headerEven;
        internal Header? _headerDefault;
        // internal footer properties for easy usage
        internal Footer? _footerFirst;
        internal Footer? _footerDefault;
        internal Footer? _footerEven;

        private void WarnIfMultipleSections(string componentName) {
            if (this.Sections.Count > 1) {
                Trace.TraceWarning($"This document contains more than 1 section. Consider using Sections[wantedSection].{componentName}.");
            }
        }

        private WordSection GetFirstSectionOrThrow(string propertyName, bool createIfMissing) {
            if (this.Sections.Count > 0) {
                return this.Sections[0];
            }

            if (!createIfMissing) {
                throw new InvalidOperationException($"Cannot access '{propertyName}' because the document does not contain any sections. Call AddSection() before using this property.");
            }

            return new WordSection(this, sectionProperties: null, paragraph: null);
        }

        /// <summary>
        /// Gets the headers of the first section.
        /// </summary>
        public WordHeaders? Header {
            get {
                WarnIfMultipleSections(nameof(Header));
                return GetFirstSectionOrThrow(nameof(Header), false).Header;
            }
        }
        /// <summary>
        /// Gets the footers of the first section.
        /// </summary>
        public WordFooters? Footer {
            get {
                WarnIfMultipleSections(nameof(Footer));
                return GetFirstSectionOrThrow(nameof(Footer), false).Footer;
            }
        }

        /// <summary>
        /// Returns the default header of the first section, creating it on demand.
        /// Prefer this over <see cref="Header"/> when you want a non-null result.
        /// </summary>
        public WordHeader HeaderDefaultOrCreate {
            get {
                WarnIfMultipleSections(nameof(HeaderDefaultOrCreate));
                return GetFirstSectionOrThrow(nameof(HeaderDefaultOrCreate), true).GetOrCreateHeader(HeaderFooterValues.Default);
            }
        }

        /// <summary>
        /// Returns the first-page header of the first section, creating it on demand
        /// by enabling <see cref="DifferentFirstPage"/> if necessary.
        /// </summary>
        public WordHeader HeaderFirstOrCreate {
            get {
                WarnIfMultipleSections(nameof(HeaderFirstOrCreate));
                return GetFirstSectionOrThrow(nameof(HeaderFirstOrCreate), true).GetOrCreateHeader(HeaderFooterValues.First);
            }
        }

        /// <summary>
        /// Returns the even-page header of the first section, creating it on demand
        /// by enabling <see cref="DifferentOddAndEvenPages"/> if necessary.
        /// </summary>
        public WordHeader HeaderEvenOrCreate {
            get {
                WarnIfMultipleSections(nameof(HeaderEvenOrCreate));
                return GetFirstSectionOrThrow(nameof(HeaderEvenOrCreate), true).GetOrCreateHeader(HeaderFooterValues.Even);
            }
        }

        /// <summary>
        /// Returns the default footer of the first section, creating it on demand.
        /// Prefer this over <see cref="Footer"/> when you want a non-null result.
        /// </summary>
        public WordFooter FooterDefaultOrCreate {
            get {
                WarnIfMultipleSections(nameof(FooterDefaultOrCreate));
                return GetFirstSectionOrThrow(nameof(FooterDefaultOrCreate), true).GetOrCreateFooter(HeaderFooterValues.Default);
            }
        }

        /// <summary>
        /// Returns the first-page footer of the first section, creating it on demand
        /// by enabling <see cref="DifferentFirstPage"/> if necessary.
        /// </summary>
        public WordFooter FooterFirstOrCreate {
            get {
                WarnIfMultipleSections(nameof(FooterFirstOrCreate));
                return GetFirstSectionOrThrow(nameof(FooterFirstOrCreate), true).GetOrCreateFooter(HeaderFooterValues.First);
            }
        }

        /// <summary>
        /// Returns the even-page footer of the first section, creating it on demand
        /// by enabling <see cref="DifferentOddAndEvenPages"/> if necessary.
        /// </summary>
        public WordFooter FooterEvenOrCreate {
            get {
                WarnIfMultipleSections(nameof(FooterEvenOrCreate));
                return GetFirstSectionOrThrow(nameof(FooterEvenOrCreate), true).GetOrCreateFooter(HeaderFooterValues.Even);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the first page has different headers and footers.
        /// </summary>
        public bool DifferentFirstPage {
            get {
                WarnIfMultipleSections(nameof(DifferentFirstPage));
                return GetFirstSectionOrThrow(nameof(DifferentFirstPage), true).DifferentFirstPage;
            }
            set {
                GetFirstSectionOrThrow(nameof(DifferentFirstPage), true).DifferentFirstPage = value;
            }

        }
        /// <summary>
        /// Gets or sets a value indicating whether odd and even pages use different headers and footers.
        /// </summary>
        public bool DifferentOddAndEvenPages {
            get {
                WarnIfMultipleSections(nameof(DifferentOddAndEvenPages));
                return GetFirstSectionOrThrow(nameof(DifferentOddAndEvenPages), true).DifferentOddAndEvenPages;
            }
            set {
                GetFirstSectionOrThrow(nameof(DifferentOddAndEvenPages), true).DifferentOddAndEvenPages = value;
            }
        }

    }
}
