using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Helper methods for adding and retrieving headers and footers in Word documents.
    /// </summary>
    public static partial class WordHeadersAndFooters {
        /// <summary>
        /// Add default header and footers to section. You can control odd/even/first with DifferentOddAndEventPages/DifferentFirstPage properties.
        /// </summary>
        /// <param name="section"></param>
        public static void AddHeadersAndFooters(this WordSection section) {
            AddHeaderReference(section._document, section, HeaderFooterValues.Default);
            AddFooterReference(section._document, section, HeaderFooterValues.Default);
        }

        /// <summary>
        /// Add default header and footers to document (section 0). You can control odd/even/first with DifferentOddAndEventPages/DifferentFirstPage properties.
        /// </summary>
        /// <param name="document"></param>
        public static void AddHeadersAndFooters(this WordDocument document) {
            AddHeaderReference(document, document.Sections[0], HeaderFooterValues.Default);
            AddFooterReference(document, document.Sections[0], HeaderFooterValues.Default);
        }

        /// <summary>
        /// Checks for existing header reference. Allow checking if different odd and even pages are set
        /// </summary>
        /// <param name="document"></param>
        /// <param name="section"></param>
        /// <param name="headerFooterValue"></param>
        /// <returns></returns>
        internal static bool GetHeaderReference(WordDocument document, WordSection section, HeaderFooterValues headerFooterValue) {
            var sectionProperties = section._sectionProperties;

            foreach (var element in sectionProperties.ChildElements.OfType<HeaderReference>()) {
                if (element.Type?.Value == headerFooterValue) {
                    // we found the header reference already exists; we do nothing;
                    return true;
                }
            }
            return false;
        }


        /// <summary>
        /// Checks for existing footer reference. Allow checking if different odd and even pages are set
        /// </summary>
        /// <param name="document"></param>
        /// <param name="section"></param>
        /// <param name="headerFooterValue"></param>
        /// <returns></returns>
        internal static bool GetFooterReference(WordDocument document, WordSection section, HeaderFooterValues headerFooterValue) {
            var sectionProperties = section._sectionProperties;
            foreach (var element in sectionProperties.ChildElements.OfType<FooterReference>()) {
                if (element.Type?.Value == headerFooterValue) {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Creates Header reference in the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="section"></param>
        /// <param name="headerFooterValue"></param>
        internal static void AddHeaderReference(WordDocument document, WordSection section, HeaderFooterValues headerFooterValue) {
            var sectionProperties = section._sectionProperties;

            foreach (var element in sectionProperties.ChildElements.OfType<HeaderReference>()) {
                if (element.Type?.Value == headerFooterValue) {
                    // we found the header reference already exists; we do nothing;
                    return;
                }
            }

            var headerPart = document._wordprocessingDocument.MainDocumentPart!.AddNewPart<HeaderPart>();

            var header = new Header();
            header.Save(headerPart);
            var headerElement = headerPart.Header ?? throw new InvalidOperationException("Header element is missing.");

            if (headerPart != null) {
                var id = document._wordprocessingDocument.MainDocumentPart!.GetIdOfPart(headerPart);
                //var id1 = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(part.HeaderParts.FirstOrDefault());

                if (id != null) {
                    var headerReference = new HeaderReference() {
                        Type = headerFooterValue,
                        Id = id
                    };

                    // Header/footer references must appear before other section
                    // properties (such as pgSz/pgMar) to satisfy the Open XML
                    // schema content model.
                    var lastHdrFtrRef = sectionProperties
                        .ChildElements
                        .Where(e => e is HeaderReference || e is FooterReference)
                        .LastOrDefault();

                    if (lastHdrFtrRef != null) {
                        sectionProperties.InsertAfter(headerReference, lastHdrFtrRef);
                    } else {
                        sectionProperties.InsertAt(headerReference, 0);
                    }
                }
            }

            var headers = section.Header ?? throw new InvalidOperationException("Headers collection is missing.");
            if (headerFooterValue == HeaderFooterValues.Default) {
                headers.Default = new WordHeader(document, HeaderFooterValues.Default, headerElement, section);
            } else if (headerFooterValue == HeaderFooterValues.First) {
                headers.First = new WordHeader(document, HeaderFooterValues.First, headerElement, section);
            } else {
                headers.Even = new WordHeader(document, HeaderFooterValues.Even, headerElement, section);
            }
        }

        /// <summary>
        /// Creates Footer reference in the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="section"></param>
        /// <param name="headerFooterValue"></param>
        internal static void AddFooterReference(WordDocument document, WordSection section, HeaderFooterValues headerFooterValue) {
            var sectionProperties = section._sectionProperties;
            foreach (var element in sectionProperties.ChildElements.OfType<FooterReference>()) {
                if (element.Type?.Value == headerFooterValue) {
                    // we found the footer reference already exists; we do nothing;
                    return;
                }
            }

            var footerPart = document._wordprocessingDocument.MainDocumentPart!.AddNewPart<FooterPart>();

            var footer = new Footer();
            footer.Save(footerPart);
            var footerElement = footerPart.Footer ?? throw new InvalidOperationException("Footer element is missing.");

            if (footerPart != null) {
                var id = document._wordprocessingDocument.MainDocumentPart!.GetIdOfPart(footerPart);
                if (id != null) {
                    var footerReference = new FooterReference() {
                        Type = headerFooterValue,
                        Id = id
                    };

                    // Footer references must live in the same leading group as
                    // header references inside sectPr.
                    var lastHdrFtrRef = sectionProperties
                        .ChildElements
                        .Where(e => e is HeaderReference || e is FooterReference)
                        .LastOrDefault();

                    if (lastHdrFtrRef != null) {
                        sectionProperties.InsertAfter(footerReference, lastHdrFtrRef);
                    } else {
                        sectionProperties.InsertAt(footerReference, 0);
                    }
                }
            }

            var footers = section.Footer ?? throw new InvalidOperationException("Footers collection is missing.");
            if (headerFooterValue == HeaderFooterValues.Default) {
                footers.Default = new WordFooter(document, HeaderFooterValues.Default, footerElement, section);
            } else if (headerFooterValue == HeaderFooterValues.First) {
                footers.First = new WordFooter(document, HeaderFooterValues.First, footerElement, section);
            } else {
                footers.Even = new WordFooter(document, HeaderFooterValues.Even, footerElement, section);
            }
        }

        /// <summary>
        /// Add section property to the document
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <returns></returns>
        internal static SectionProperties AddSectionProperties(this WordprocessingDocument wordDocument) {
            // When attaching section properties directly to the document body
            // we want explicit page size and margins so that the document
            // validates cleanly against the Open XML schema.
            var sectionProperties = CreateSectionProperties(includeDefaultPageSettings: true);
            wordDocument.MainDocumentPart!.Document!.Body!.Append(sectionProperties);
            return sectionProperties;
        }

        /// <summary>
        /// Some documents might not have section properties with proper rsidR. This method will fix that.
        /// Due to the nature of the rsidR, it is important to have unique rsidR for each section.
        /// Otherwise, comparison of sections will not work properly.
        /// </summary>
        /// <param name="sectionProperties"></param>
        /// <returns></returns>
        internal static SectionProperties MakeSureSectionIsValid(this SectionProperties sectionProperties) {
            if (sectionProperties.RsidR == null) {
                sectionProperties.RsidR = GenerateRsid();
            }

            return sectionProperties;
        }

        /// <summary>
        /// Generate a unique rsid
        /// </summary>
        /// <returns></returns>
        internal static string GenerateRsid() {
            // Generate a unique rsid using a GUID
            return Guid.NewGuid().ToString("N").Substring(0, 8).ToUpper();
        }

        private static uint _revisionIdCounter = 1;

        /// <summary>
        /// Generate a unique revision id used by <see cref="InsertedRun"/> and
        /// <see cref="DeletedRun"/> elements.
        /// </summary>
        /// <returns>Revision identifier as decimal string.</returns>
        internal static string GenerateRevisionId() {
            return (_revisionIdCounter++).ToString();
        }

        /// <summary>
        /// Create a new section properties container.
        /// </summary>
        /// <remarks>
        /// For new sections inserted into an existing document we only want a bare
        /// <see cref="SectionProperties"/> element so that <see cref="WordSection"/>
        /// can copy settings such as headers, footers, margins and numbering from
        /// the previous section using <c>CopySectionProperties</c>.
        /// </remarks>
        /// <returns>Empty section properties element with a unique rsid.</returns>
        internal static SectionProperties CreateSectionProperties() {
            // New sections should start with an empty sectPr so that existing
            // section properties can be copied correctly.
            return CreateSectionProperties(includeDefaultPageSettings: false);
        }

        /// <summary>
        /// Create a new section properties container.
        /// </summary>
        /// <param name="includeDefaultPageSettings">
        /// When <c>true</c>, includes default Letter page size and Normal margins.
        /// </param>
        /// <returns>Section properties element.</returns>
        internal static SectionProperties CreateSectionProperties(bool includeDefaultPageSettings) {
            SectionProperties sectionProperties = new SectionProperties() { RsidR = GenerateRsid() };

            if (includeDefaultPageSettings) {
                // Align new documents with Word defaults:
                // Letter page size with 1" Normal margins.
                var pageSize = WordPageSizes.Letter;
                // Clone to avoid sharing instances between sections.
                PageSize pageSizeClone = new PageSize() {
                    Width = pageSize.Width,
                    Height = pageSize.Height,
                    Code = pageSize.Code,
                    Orient = pageSize.Orient
                };

                // Match WordMargins.Normal so that margin presets are detected correctly.
                PageMargin pageMargin = new PageMargin() {
                    Top = 1440,    // 1 inch
                    Right = 1440,  // 1 inch
                    Bottom = 1440, // 1 inch
                    Left = 1440,   // 1 inch
                    Header = 720,
                    Footer = 720,
                    Gutter = 0
                };

                sectionProperties.Append(pageSizeClone);
                sectionProperties.Append(pageMargin);
            }

            return sectionProperties;
        }
    }
}
