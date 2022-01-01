using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordSection {
        public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        public List<WordParagraph> PageBreaks = new List<WordParagraph>();
        public List<WordImage> Images = new List<WordImage>();

        public readonly WordFooters Footer = new WordFooters();
        public readonly WordHeaders Header = new WordHeaders();

        // internal header properties for easy usage
        internal Header _headerFirst;
        internal Header _headerEven;
        internal Header _headerOdd;
        // internal footer properties for easy usage
        internal Footer _footerFirst;
        internal Footer _footerOdd;
        internal Footer _footerEven;

   
        public WordDocument _document;
        public SectionProperties _sectionProperties;

        public WordSection() {

        }
        public WordSection(WordDocument wordDocument, Paragraph paragraph = null) {
            this._document = wordDocument;
            if (paragraph != null) {
                
                var sectionProperties = paragraph.ParagraphProperties.SectionProperties;
                if (sectionProperties == null) {
                    return;
                }
                this._sectionProperties = sectionProperties;
            } else {
                var sectionProperties = wordDocument._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                if (sectionProperties == null) {
                    wordDocument._wordprocessingDocument.AddSectionProperties();
                }
                this._sectionProperties = sectionProperties;
            }
            wordDocument.Sections.Add(this);

            // this is added for tracking current section of the document
            wordDocument._currentSection = this;
        }
        
        //public bool DifferentFirstPage {
        //    get {
        //        var sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
        //        if (sectionProperties != null) {
        //            var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
        //            if (titlePage != null) {
        //                return true;
        //            }
        //        }
        //        return false;
        //    }
        //    set {
        //        var sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
        //        if (sectionProperties == null) {
        //            if (value == false) {
        //                // section properties doesn't exists, so we don't do anything
        //                return;
        //            } else {
        //                _document.Body.Append(
        //                //WordHeadersAndFooters.AddSectionProperties()
        //                );
        //            }
        //        }

        //        sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().First();
        //        var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
        //        if (value == false) {
        //            if (titlePage == null) {
        //                return;
        //            } else {
        //                titlePage.Remove();
        //            }
        //        } else {
        //            sectionProperties.Append(new TitlePage());
        //        }

        //    }

        //}

        //public bool DifferentOddAndEvenPages {
        //    get {
        //        var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
        //        if (settings != null) {
        //            return true;
        //        } else {
        //            return false;
        //        }
        //    }
        //    set {
        //        var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
        //        if (value == false) {

        //        } else {
        //            if (settings == null) {
        //                _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.Append(new EvenAndOddHeaders());
        //            } else {
        //                // noting to do, already enabled
        //            }
        //        }
        //    }
        //}

        //public WordSection(WordDocument document) {
        //    WordParagraph paragraph = new WordParagraph();
        //    WordSection section = new WordSection(document, paragraph);
        //}
        //public WordSection(WordDocument document, WordParagraph paragraph) {
        //    ParagraphProperties paragraphProperties = new ParagraphProperties();

        //    SectionProperties sectionProperties = new SectionProperties();
        //    SectionType sectionType = new SectionType() {Val = SectionMarkValues.NextPage};


        //    sectionProperties.Append(sectionType);
        //    paragraphProperties.Append(sectionProperties);
        //    paragraph._paragraph.Append(paragraphProperties);
        //}
        //private static void AddSectionBreakToTheDocument(string fileName) {
        //    using (WordprocessingDocument mydoc = WordprocessingDocument.Open(fileName, true)) {
        //        MainDocumentPart myMainPart = mydoc.MainDocumentPart;
        //        Paragraph paragraphSectionBreak = new Paragraph();
        //        ParagraphProperties paragraphSectionBreakProperties = new ParagraphProperties();
        //        SectionProperties SectionBreakProperties = new SectionProperties();
        //        SectionType SectionBreakType = new SectionType() { Val = SectionMarkValues.NextPage };
        //        SectionBreakProperties.Append(SectionBreakType);
        //        paragraphSectionBreakProperties.Append(SectionBreakProperties);
        //        paragraphSectionBreak.Append(paragraphSectionBreakProperties);
        //        myMainPart.Document.Body.InsertAfter(paragraphSectionBreak, myMainPart.Document.Body.LastChild);
        //        myMainPart.Document.Save();
        //    }
        //}

        public WordParagraph InsertParagraph(string text) {
            if (this.Paragraphs.Count == 0) {

                var paragraph = this._document.InsertParagraph(text);
                paragraph._section = this;
                return paragraph;
            } else {
                var lastParagraphWithinSection = this.Paragraphs.Last();

                var paragraph = lastParagraphWithinSection.InsertParagraphAfterSelf();
                paragraph._document = this._document;
                paragraph._section = this;
                this.Paragraphs.Add(paragraph);
                paragraph.Text = text;
                return paragraph;
                //paragraph._paragraph.InsertAfterSelf(new )
                //paragraph.InsertText(text);
                //_sectionProperties.Parent.Ap
                //return paragraph.InsertParagraphAfterSelf();
            }

            //this._document.InsertParagraph(text);
            //_document._currentSection.InsertParagraph("test");
        }
    }
}
