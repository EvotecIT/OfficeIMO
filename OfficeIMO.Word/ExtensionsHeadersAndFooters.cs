using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
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
                if (element.Type == headerFooterValue) {
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
                if (element.Type == headerFooterValue) {
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
                if (element.Type == headerFooterValue) {
                    // we found the header reference already exists; we do nothing;
                    return;
                }
            }

            var headerPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();

            var header = new Header();
            header.Save(headerPart);

            if (headerPart != null) {
                var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(headerPart);
                //var id1 = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(part.HeaderParts.FirstOrDefault());

                if (id != null) {
                    var headerReference = new HeaderReference() {
                        Type = headerFooterValue,
                        Id = id
                    };
                    sectionProperties.Append(headerReference);
                }
            }

            if (headerFooterValue == HeaderFooterValues.Default) {
                //  section._headerDefault = headerPart.Header;
                section.Header.Default = new WordHeader(document, HeaderFooterValues.Default, headerPart.Header, section);
            } else if (headerFooterValue == HeaderFooterValues.First) {
                //  section._headerFirst = headerPart.Header;
                section.Header.First = new WordHeader(document, HeaderFooterValues.First, headerPart.Header, section);
            } else {
                // section._headerEven = headerPart.Header;
                section.Header.Even = new WordHeader(document, HeaderFooterValues.Even, headerPart.Header, section);
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
                if (element.Type == headerFooterValue) {
                    // we found the footer reference already exists; we do nothing;
                    return;
                }
            }

            var footerPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>();

            MainDocumentPart part = document._wordprocessingDocument.MainDocumentPart;
            Body body = document._wordprocessingDocument.MainDocumentPart.Document.Body;

            var footer = new Footer();
            footer.Save(footerPart);

            if (footerPart != null) {
                var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(footerPart);
                if (id != null) {
                    var footerReference = new FooterReference() {
                        Type = headerFooterValue,
                        Id = id
                    };
                    sectionProperties.Append(footerReference);
                }
            }

            if (headerFooterValue == HeaderFooterValues.Default) {
                //section._footerDefault = footerPart.Footer;
                section.Footer.Default = new WordFooter(document, HeaderFooterValues.Default, footerPart.Footer, section);
            } else if (headerFooterValue == HeaderFooterValues.First) {
                //section._footerFirst = footerPart.Footer;
                section.Footer.First = new WordFooter(document, HeaderFooterValues.First, footerPart.Footer, section);
            } else {
                //section._footerEven = footerPart.Footer;
                section.Footer.Even = new WordFooter(document, HeaderFooterValues.Even, footerPart.Footer, section);
            }
        }

        /// <summary>
        /// Add section property to the document
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <returns></returns>
        internal static SectionProperties AddSectionProperties(this WordprocessingDocument wordDocument) {
            var sectionProperties = CreateSectionProperties();
            wordDocument.MainDocumentPart.Document.Body.Append(sectionProperties);
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

        /// <summary>
        /// Create a new section properties
        /// </summary>
        /// <returns></returns>
        internal static SectionProperties CreateSectionProperties() {
            SectionProperties sectionProperties = new SectionProperties() { RsidR = GenerateRsid() };

            // Set the page size and margins
            //PageSize pageSize = new PageSize() { Width = 12240, Height = 15840 }; // A4 size
            //PageMargin pageMargin = new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }; // 1 inch margins

            //sectionProperties.Append(pageSize);
            //sectionProperties.Append(pageMargin);

            return sectionProperties;
        }
    }
}
