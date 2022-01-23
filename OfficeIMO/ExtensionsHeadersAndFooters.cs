using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public static partial class WordHeadersAndFooters {

        public static void AddHeadersAndFooters(this WordSection section) {
            var document = section._document;

            //var index = document.Sections.IndexOf(section);
            //var body = document._wordprocessingDocument.MainDocumentPart.Document.Body;
            //var sectionProperties = document._wordprocessingDocument.MainDocumentPart.Document.Body.Elements<SectionProperties>().Last();
            //var sectionBefore = document.Sections[index - 1];
            //var Headers = sectionBefore._sectionProperties.ChildElements.OfType<HeaderReference>();
            //var Footers = sectionBefore._sectionProperties.ChildElements.OfType<FooterReference>();
            //var HEA = sectionBefore._sectionProperties.OfType<HeaderReference>(); //<HeaderReference>().Remove();


            //var properties = document.Sections[index - 1]._sectionProperties;

            //AddHeaderReference1(document, section._sectionProperties, HeaderFooterValues.Default, section);
            //AddFooterReference1(document, section._sectionProperties, HeaderFooterValues.Default, section);

            //var currentProperties = section._sectionProperties;

            //document.Sections[index]._sectionProperties = currentProperties;

            //section._sectionProperties = properties;


            AddHeaderReference1(document, section, HeaderFooterValues.Default);
            AddFooterReference1(document, section, HeaderFooterValues.Default);

            //AddHeaderReference1(document, section._sectionProperties, HeaderFooterValues.Even, section);
            //AddFooterReference1(document, section._sectionProperties, HeaderFooterValues.Even, section);

            //AddHeaderReference1(document, section._sectionProperties, HeaderFooterValues.First, section);
            //AddFooterReference1(document, section._sectionProperties, HeaderFooterValues.First, section);
        }
        public static void AddHeadersAndFooters(this WordDocument document) {

            AddHeaderReference1(document, document.Sections[0], HeaderFooterValues.Default);
            AddFooterReference1(document, document.Sections[0], HeaderFooterValues.Default);
            //AddHeaderReference1(document, document.Sections[0]._sectionProperties, HeaderFooterValues.Even);
            //AddFooterReference1(document, document.Sections[0]._sectionProperties, HeaderFooterValues.Even);
            //AddHeaderReference1(document, document.Sections[0]._sectionProperties, HeaderFooterValues.First);
            //AddFooterReference1(document, document.Sections[0]._sectionProperties, HeaderFooterValues.First);
       }

        public static void CreateHeadersAndFooters(this WordDocument document) {
            //var firstPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();
            //var firstPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>();
            //var evenPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();
            //var evenPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>();
            //var defaultPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();
            //var defaultPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>();

            //firstPageFooterPart.AddFooters();
            //evenPageFooterPart.AddFooters();
            //defaultPageFooterPart.AddFooters();

            //firstPageHeaderPart.AddHeaders();
            //evenPageHeaderPart.AddHeaders();
            //defaultPageHeaderPart.AddHeaders();

            //document._footerFirst = firstPageFooterPart.Footer;
            //document._footerDefault = defaultPageFooterPart.Footer;
            //document._footerEven = evenPageFooterPart.Footer;

            //document._headerFirst = firstPageHeaderPart.Header;
            //document._headerEven = evenPageHeaderPart.Header;
            //document._headerDefault = defaultPageHeaderPart.Header;

            //document.Sections[0]._sectionProperties.AddHeaderFooterToSectionProperties(document, firstPageHeaderPart, defaultPageHeaderPart, evenPageHeaderPart, firstPageFooterPart, defaultPageFooterPart, evenPageFooterPart);

            //// lets set proper 
            //document.Footer.Even = new WordFooter(document, HeaderFooterValues.Even);
            //document.Footer.Default = new WordFooter(document, HeaderFooterValues.Default);
            //document.Footer.First = new WordFooter(document, HeaderFooterValues.First);

            //document.Header.Even = new WordHeader(document, HeaderFooterValues.Even);
            //document.Header.Default = new WordHeader(document, HeaderFooterValues.Default);
            //document.Header.First = new WordHeader(document, HeaderFooterValues.First);

         
        }

        //public static SectionProperties GetSectionProperties(this WordDocument document) {


        //}

        private static void GetHeaderReference(this WordDocument document, WordSection section) {
            IEnumerable<HeaderPart> headerPart = document._wordprocessingDocument.MainDocumentPart.HeaderParts;
            foreach (HeaderPart header in headerPart) {

            }
        }

        private static void AddHeaderRef(WordDocument document, WordSection section, SectionProperties sectionProperties, HeaderFooterValues headerFooterValue) {
           
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
                    section.Header.Default = new WordHeader(document, HeaderFooterValues.Default, headerPart.Header);

                } else if (headerFooterValue == HeaderFooterValues.First) {
                    //  section._headerFirst = headerPart.Header;
                    section.Header.First = new WordHeader(document, HeaderFooterValues.First, headerPart.Header);
                } else {
                    // section._headerEven = headerPart.Header;
                    section.Header.Even = new WordHeader(document, HeaderFooterValues.Even, headerPart.Header);
                }
        }

        internal static void AddHeaderReference1(WordDocument document, SectionProperties sectionProperties, HeaderFooterValues headerFooterValue, WordSection section = null) {
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
        public static void CreateHeadersAndFooters(this WordSection section) {

        }


        private static Document GenerateMainDocumentPart() {
            var element = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text("Page 1 content"))
                        ),
                        new Paragraph(
                            new Run(
                                new Break() { Type = BreakValues.Page })
                        ),
                        new Paragraph(
                            new Run(
                                new LastRenderedPageBreak(),
                                new Text("Page 2 content"))
                        ),
                        new Paragraph(
                            new Run(
                                new Break() { Type = BreakValues.Page })
                        ),
                        new Paragraph(
                            new Run(
                                new LastRenderedPageBreak(),
                                new Text("Page 3 content"))
                        ),
                        new Paragraph(
                            new Run(
                                new Break() { Type = BreakValues.Page })
                        ),
                        new Paragraph(
                            new Run(
                                new LastRenderedPageBreak(),
                                new Text("Page 4 content"))
                        ),
                        new Paragraph(
                            new Run(
                                new Break() { Type = BreakValues.Page })
                        ),
                        new Paragraph(
                            new Run(
                                new LastRenderedPageBreak(),
                                new Text("Page 5 content"))
                        ),
                        new SectionProperties(
                            new HeaderReference() {
                                Type = HeaderFooterValues.First,
                                Id = "rId2"
                            },
                            new FooterReference() {
                                Type = HeaderFooterValues.First,
                                Id = "rId3"
                            },
                            new HeaderReference() {
                                Type = HeaderFooterValues.Even,
                                Id = "rId4"
                            },
                            new FooterReference() {
                                Type = HeaderFooterValues.Even,
                                Id = "rId5"
                            },
                            new HeaderReference() {
                                Type = HeaderFooterValues.Default,
                                Id = "rId6"
                            },
                            new FooterReference() {
                                Type = HeaderFooterValues.Default,
                                Id = "rId7"
                            },
                            new PageMargin() {
                                Top = 1440,
                                Right = (UInt32Value)1440UL,
                                Bottom = 1440,
                                Left = (UInt32Value)1440UL,
                                Header = (UInt32Value)720UL,
                                Footer = (UInt32Value)720UL,
                                Gutter = (UInt32Value)0UL
                            },
                            new TitlePage()
                        )));

            return element;
        }

        internal static SectionProperties AddSectionProperties(this WordprocessingDocument wordDocument) {
            var sectionProperties = new SectionProperties();
            wordDocument.MainDocumentPart.Document.Body.Append(sectionProperties);
            return sectionProperties;
        }

        internal static SectionProperties AddHeaderFooterToSectionProperties(this SectionProperties sectionProperties, WordDocument document, HeaderPart firstHeaderPart, HeaderPart defaultHeaderPart, HeaderPart evenHeaderPart, FooterPart firstFooterPart, FooterPart defaultFooterPart, FooterPart evenFooterPart) {
            AddHeaderReference(document, sectionProperties, firstHeaderPart, HeaderFooterValues.First);
            AddHeaderReference(document, sectionProperties, defaultHeaderPart, HeaderFooterValues.Default);
            AddHeaderReference(document, sectionProperties, evenHeaderPart, HeaderFooterValues.Even);

            AddFooterReference(document, sectionProperties, firstFooterPart, HeaderFooterValues.First);
            AddFooterReference(document, sectionProperties, defaultFooterPart, HeaderFooterValues.Default);
            AddFooterReference(document, sectionProperties, evenFooterPart, HeaderFooterValues.Even);

            return sectionProperties;
        }

        private static void AddFooterReference(WordDocument document, SectionProperties sectionProperties, FooterPart footerPart, HeaderFooterValues headerFooterValue) {
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
        }
        private static void AddHeaderReference(WordDocument document, SectionProperties sectionProperties, HeaderPart headerPart, HeaderFooterValues headerFooterValue) {
            if (headerPart != null) {
                var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(headerPart);
                if (id != null) {
                    var headerReference = new HeaderReference() {
                        Type = headerFooterValue,
                        Id = id
                    };
                    sectionProperties.Append(headerReference);
                }
            }
        }
    }
}
