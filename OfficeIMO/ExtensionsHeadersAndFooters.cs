using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public static partial class WordHeadersAndFooters {

        public static void AddHeadersAndFooters(this WordDocument document) {
            var sectionProperties = document._wordprocessingDocument.AddSectionProperties();
            //sectionProperties.AddHeaderFooterToSectionProperties();
           // var id = document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart;
            var documentSettingsPart = document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart;
            if (documentSettingsPart == null) {
                documentSettingsPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>("rId1");
            }
            //var documentSettingsPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>("rId1");

            var settings = new Settings();
            settings.Save(documentSettingsPart);

            //GenerateDocumentSettingsPart().Save(documentSettingsPart);
            //documentSettingsPart.Settings.HideSpellingErrors = new HideSpellingErrors(){ Val = false };

            var firstPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();
            var firstPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>();
            var evenPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();
            var evenPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>();
            var oddPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();
            var oddPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>();

            //var firstPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId2");
            //var firstPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>("rId3");
            //var evenPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId4");
            //var evenPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>("rId5");
            //var oddPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId6");
            //var oddPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>("rId7");
            var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(firstPageFooterPart);


            firstPageFooterPart.AddFooters();
            evenPageFooterPart.AddFooters();
            oddPageFooterPart.AddFooters();

            document._footerFirst = firstPageFooterPart.Footer;
            document._footerOdd = oddPageFooterPart.Footer;
            document._footerEven = evenPageFooterPart.Footer;

            firstPageHeaderPart.AddHeaders();
            evenPageHeaderPart.AddHeaders();
            oddPageHeaderPart.AddHeaders();

            document._headerFirst = firstPageHeaderPart.Header;
            document._headerEven = evenPageHeaderPart.Header;
            document._headerOdd = oddPageHeaderPart.Header;

            sectionProperties.AddHeaderFooterToSectionProperties(document, firstPageHeaderPart, oddPageHeaderPart, evenPageHeaderPart, firstPageFooterPart, oddPageFooterPart, evenPageFooterPart);

            // lets set proper 
            document.Footer.Even = new WordFooter(document, "even");
            document.Footer.Odd = new WordFooter(document, "odd");
            document.Footer.First = new WordFooter(document, "first");

            document.Header.Even = new WordHeader(document, "even");
            document.Header.Odd = new WordHeader(document, "odd");
            document.Header.First = new WordHeader(document, "first");

        }
        //public static void AddHeadersAndFooters(this WordprocessingDocument wordDocument, WordDocument document) {
        //    wordDocument.AddSection();

        //    var documentSettingsPart = wordDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>("rId1");

        //    var settings = new Settings();
        //    //var settings = new Settings(new EvenAndOddHeaders());

        //    settings.Save(documentSettingsPart);

        //    //GenerateDocumentSettingsPart().Save(documentSettingsPart);
        //    //documentSettingsPart.Settings.HideSpellingErrors = new HideSpellingErrors(){ Val = false };

        //    var firstPageHeaderPart = wordDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId2");
        //    var firstPageFooterPart = wordDocument.MainDocumentPart.AddNewPart<FooterPart>("rId3");
        //    var evenPageHeaderPart = wordDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId4");
        //    var evenPageFooterPart = wordDocument.MainDocumentPart.AddNewPart<FooterPart>("rId5");
        //    var oddPageHeaderPart = wordDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId6");
        //    var oddPageFooterPart = wordDocument.MainDocumentPart.AddNewPart<FooterPart>("rId7");

        //    firstPageFooterPart.AddFooters();
        //    evenPageFooterPart.AddFooters();
        //    oddPageFooterPart.AddFooters();

        //    document._footerFirst = firstPageFooterPart.Footer;
        //    document._footerOdd = oddPageFooterPart.Footer;
        //    document._footerEven = evenPageFooterPart.Footer;

        //    firstPageHeaderPart.AddHeaders();
        //    evenPageHeaderPart.AddHeaders();
        //    oddPageHeaderPart.AddHeaders();

        //    document._headerFirst = firstPageHeaderPart.Header;
        //    document._headerEven = evenPageHeaderPart.Header;
        //    document._headerOdd = oddPageHeaderPart.Header;

        //}

        private static void AddFooters(this FooterPart footerPart, Document document) {
            var element = new Footer();
            element.Save(footerPart);
        }

        private static void AddFooters(this FooterPart footerPart) {
            var element = new Footer();
            element.Save(footerPart);
        }

        private static void AddHeaders(this HeaderPart headerPart) {
            var element = new Header();
            element.Save(headerPart);
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

        private static void AddSection(this WordprocessingDocument wordDocument) {

            wordDocument.AddSectionProperties();
            wordDocument.MainDocumentPart.Document.Body.Append(
                AddSectionProperties()
            );
            wordDocument.AddSectionProperties();
        }

        internal static SectionProperties AddSectionProperties(this WordprocessingDocument wordDocument) {
            //var sections = wordDocument.MainDocumentPart.Document.Descendants<SectionProperties>();
            //if (sections.Count() > 0) {
            //    foreach (SectionProperties section in sections) {
            //        //var sectionProperties = wordDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
            //        if (section == null) {
            //            var sectionProperties = new SectionProperties();
            //            wordDocument.MainDocumentPart.Document.Body.Append(sectionProperties);
            //        }
            //        //section.AddHeaderFooterToSectionProperties();
            //    }
            //} else {
            //    var sectionProperties = new SectionProperties();
            //    wordDocument.MainDocumentPart.Document.Body.Append(sectionProperties);
            //}
                       var sectionProperties = new SectionProperties();
                        wordDocument.MainDocumentPart.Document.Body.Append(sectionProperties);
                        return sectionProperties;
        }

        internal static SectionProperties AddHeaderFooterToSectionProperties(this SectionProperties sectionProperties, WordDocument document, HeaderPart firstHeaderPart, HeaderPart oddHeaderPart, HeaderPart evenHeaderPart, FooterPart firstFooterPart, FooterPart oddFooterPart, FooterPart evenFooterPart) {
            //foreach (HeaderReference header in sectionProperties) {
                
            //}

            //if (firstFooterPart != null) {
            //    var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(firstFooterPart);
            //    if (id != null) {
            //        var footerReference = new FooterReference() {
            //            Type = HeaderFooterValues.First,
            //            Id = id
            //        };
            //        sectionProperties.Append(footerReference);
            //    }
            //}

            AddHeaderReference(document, sectionProperties, firstHeaderPart, HeaderFooterValues.First);
            AddHeaderReference(document, sectionProperties, oddHeaderPart, HeaderFooterValues.Default);
            AddHeaderReference(document, sectionProperties, evenHeaderPart, HeaderFooterValues.Even);

            AddFooterReference(document, sectionProperties, firstFooterPart, HeaderFooterValues.First);
            AddFooterReference(document, sectionProperties, oddFooterPart, HeaderFooterValues.Default);
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

        internal static SectionProperties AddHeaderFooterToSectionProperties1(this SectionProperties sectionProperties) {
            sectionProperties.Append(
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
                }
            );
            return sectionProperties;
        }

        internal static SectionProperties AddSectionProperties() {
            SectionProperties sectionProperties = new SectionProperties();
            sectionProperties.Append(
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
                }
                //new TitlePage()
            );
            return sectionProperties;
        }
    }
}
