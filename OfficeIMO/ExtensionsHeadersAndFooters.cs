using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public static partial class WordHeadersAndFooters {

        public static void AddHeadersAndFooters(this WordDocument document) {

            document._wordprocessingDocument.AddSection();

            var documentSettingsPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>("rId1");

            var settings = new Settings();
            //var settings = new Settings(new EvenAndOddHeaders());

            settings.Save(documentSettingsPart);

            //GenerateDocumentSettingsPart().Save(documentSettingsPart);
            //documentSettingsPart.Settings.HideSpellingErrors = new HideSpellingErrors(){ Val = false };

            var firstPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId2");
            var firstPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>("rId3");
            var evenPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId4");
            var evenPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>("rId5");
            var oddPageHeaderPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>("rId6");
            var oddPageFooterPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<FooterPart>("rId7");

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
            wordDocument.MainDocumentPart.Document.Body.Append(
                AddSectionProperties()
            );
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
