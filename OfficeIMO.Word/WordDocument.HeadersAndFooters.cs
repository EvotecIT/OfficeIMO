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
        public WordHeaders Header {
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
        public WordFooters Footer {
            get {
                if (this.Sections.Count == 0) {
                    return null;
                }
                WarnIfMultipleSections(nameof(Footer));
                return this.Sections[0].Footer;
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
