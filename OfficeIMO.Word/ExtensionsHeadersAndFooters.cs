using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordHeadersAndFooters {
        public static void AddHeadersAndFooters(this WordSection section) {
            AddHeaderReference1(section._document, section, HeaderFooterValues.Default);
            AddFooterReference1(section._document, section, HeaderFooterValues.Default);
        }

        public static void AddHeadersAndFooters(this WordDocument document) {
            AddHeaderReference1(document, document.Sections[0], HeaderFooterValues.Default);
            AddFooterReference1(document, document.Sections[0], HeaderFooterValues.Default);
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

        internal static void AddHeaderReference1(WordDocument document, WordSection section, HeaderFooterValues headerFooterValue) {
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

            if (section == null) {
                if (headerFooterValue == HeaderFooterValues.Default) {
                    // document._headerDefault = headerPart.Header;
                    document.Header.Default = new WordHeader(document, HeaderFooterValues.Default, headerPart.Header);
                } else if (headerFooterValue == HeaderFooterValues.First) {
                    // document._headerFirst = headerPart.Header;
                    document.Header.First = new WordHeader(document, HeaderFooterValues.First, headerPart.Header);
                } else {
                    // document._headerEven = headerPart.Header;
                    document.Header.Even = new WordHeader(document, HeaderFooterValues.Even, headerPart.Header);
                }
            } else {
                if (headerFooterValue == HeaderFooterValues.Default) {
                    //  section._headerDefault = headerPart.Header;
                    section.Header.Default = new WordHeader(document, HeaderFooterValues.Default, headerPart.Header);
                } else if (headerFooterValue == HeaderFooterValues.First) {
                    //  section._headerFirst = headerPart.Header;
                    section.Header.First = new WordHeader(document, HeaderFooterValues.First, headerPart.Header);
                } else {
                    // section._headerEven = headerPart.Header;
                    section.Header.Even = new WordHeader(document, HeaderFooterValues.Even, headerPart.Header);
                }
            }
        }

        internal static void AddFooterReference1(WordDocument document, WordSection section, HeaderFooterValues headerFooterValue) {
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

            if (section == null) {
                if (headerFooterValue == HeaderFooterValues.Default) {
                    //document._footerDefault = footerPart.Footer;
                    document.Footer.Default = new WordFooter(document, HeaderFooterValues.Default, footerPart.Footer);
                } else if (headerFooterValue == HeaderFooterValues.First) {
                    //document._footerFirst = footerPart.Footer;
                    document.Footer.First = new WordFooter(document, HeaderFooterValues.First, footerPart.Footer);
                } else {
                    //document._footerEven = footerPart.Footer;
                    document.Footer.Even = new WordFooter(document, HeaderFooterValues.Even, footerPart.Footer);
                }
            } else {
                if (headerFooterValue == HeaderFooterValues.Default) {
                    //section._footerDefault = footerPart.Footer;
                    section.Footer.Default = new WordFooter(document, HeaderFooterValues.Default, footerPart.Footer);
                } else if (headerFooterValue == HeaderFooterValues.First) {
                    //section._footerFirst = footerPart.Footer;
                    section.Footer.First = new WordFooter(document, HeaderFooterValues.First, footerPart.Footer);
                } else {
                    //section._footerEven = footerPart.Footer;
                    section.Footer.Even = new WordFooter(document, HeaderFooterValues.Even, footerPart.Footer);
                }
            }
        }

        internal static SectionProperties AddSectionProperties(this WordprocessingDocument wordDocument) {
            var sectionProperties = new SectionProperties();
            wordDocument.MainDocumentPart.Document.Body.Append(sectionProperties);
            return sectionProperties;
        }
    }
}