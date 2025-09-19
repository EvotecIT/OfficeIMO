using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;

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

        /// <summary>
        /// Gets the headers of the first section.
        /// </summary>
        public WordHeaders? Header {
            get {
                if (this.Sections.Count == 0) {
                    return null;
                }
                WarnIfMultipleSections(nameof(Header));
                return this.Sections[0].Header;
            }
        }
        /// <summary>
        /// Gets the footers of the first section.
        /// </summary>
        public WordFooters? Footer {
            get {
                if (this.Sections.Count == 0) {
                    return null;
                }
                WarnIfMultipleSections(nameof(Footer));
                return this.Sections[0].Footer;
            }
        }

        /// <summary>
        /// Returns the default header of the first section, creating it on demand.
        /// Prefer this over <see cref="Header"/> when you want a non-null result.
        /// </summary>
        public WordHeader HeaderDefaultOrCreate {
            get {
                WarnIfMultipleSections(nameof(HeaderDefaultOrCreate));
                return this.Sections[0].GetOrCreateHeader(HeaderFooterValues.Default);
            }
        }

        /// <summary>
        /// Returns the first-page header of the first section, creating it on demand
        /// by enabling <see cref="DifferentFirstPage"/> if necessary.
        /// </summary>
        public WordHeader HeaderFirstOrCreate {
            get {
                WarnIfMultipleSections(nameof(HeaderFirstOrCreate));
                return this.Sections[0].GetOrCreateHeader(HeaderFooterValues.First);
            }
        }

        /// <summary>
        /// Returns the even-page header of the first section, creating it on demand
        /// by enabling <see cref="DifferentOddAndEvenPages"/> if necessary.
        /// </summary>
        public WordHeader HeaderEvenOrCreate {
            get {
                WarnIfMultipleSections(nameof(HeaderEvenOrCreate));
                return this.Sections[0].GetOrCreateHeader(HeaderFooterValues.Even);
            }
        }

        /// <summary>
        /// Returns the default footer of the first section, creating it on demand.
        /// Prefer this over <see cref="Footer"/> when you want a non-null result.
        /// </summary>
        public WordFooter FooterDefaultOrCreate {
            get {
                WarnIfMultipleSections(nameof(FooterDefaultOrCreate));
                return this.Sections[0].GetOrCreateFooter(HeaderFooterValues.Default);
            }
        }

        /// <summary>
        /// Returns the first-page footer of the first section, creating it on demand
        /// by enabling <see cref="DifferentFirstPage"/> if necessary.
        /// </summary>
        public WordFooter FooterFirstOrCreate {
            get {
                WarnIfMultipleSections(nameof(FooterFirstOrCreate));
                return this.Sections[0].GetOrCreateFooter(HeaderFooterValues.First);
            }
        }

        /// <summary>
        /// Returns the even-page footer of the first section, creating it on demand
        /// by enabling <see cref="DifferentOddAndEvenPages"/> if necessary.
        /// </summary>
        public WordFooter FooterEvenOrCreate {
            get {
                WarnIfMultipleSections(nameof(FooterEvenOrCreate));
                return this.Sections[0].GetOrCreateFooter(HeaderFooterValues.Even);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the first page has different headers and footers.
        /// </summary>
        public bool DifferentFirstPage {
            get {
                WarnIfMultipleSections(nameof(DifferentFirstPage));
                return this.Sections[0].DifferentFirstPage;
            }
            set {
                this.Sections[0].DifferentFirstPage = value;
            }

        }
        /// <summary>
        /// Gets or sets a value indicating whether odd and even pages use different headers and footers.
        /// </summary>
        public bool DifferentOddAndEvenPages {
            get {
                WarnIfMultipleSections(nameof(DifferentOddAndEvenPages));
                return this.Sections[0].DifferentOddAndEvenPages;
            }
            set {
                this.Sections[0].DifferentOddAndEvenPages = value;
            }
        }

    }
}
