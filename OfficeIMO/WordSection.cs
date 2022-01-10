using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public partial class WordSection {
        public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        public List<WordParagraph> PageBreaks = new List<WordParagraph>();
        public List<WordImage> Images = new List<WordImage>();

        public WordFooters Footer = new WordFooters();
        public WordHeaders Header = new WordHeaders();

        public List<WordList> Lists = new List<WordList>();
        
        //public List<WordList> Lists {
        //    get {
        //        List<WordList> returnList = new List<WordList>();
        //        foreach (WordParagraph paragraph in this.Paragraphs) {
        //            if (paragraph.IsListItem) {
        //                if (!_document._listNumbersUsed.Contains(paragraph._listNumberId.Value)) {
        //                    WordList list = new WordList(paragraph._document, paragraph._section, paragraph._listNumberId.Value);
        //                    returnList.Add(list);
        //                    _document._listNumbersUsed.Add(paragraph._listNumberId.Value);
        //                }
        //            }
        //        }

        //        return returnList;
        //    }
        //}

        public List<WordTable> Tables = new List<WordTable>();
        
        internal WordDocument _document;
        internal SectionProperties _sectionProperties;
        internal WordprocessingDocument _wordprocessingDocument;

        public WordSection(WordDocument wordDocument, Paragraph paragraph = null) {
            this._document = wordDocument;
            this._wordprocessingDocument = wordDocument._wordprocessingDocument;

            WordSection lastSection;
            if (this._document.Sections.Count > 0) {
                lastSection = this._document.Sections[this._document.Sections.Count - 1];
                //lastSection._sectionProperties.
            }

            //PageMargin pageMargin1;
            if (paragraph != null) {
                var sectionProperties = paragraph.ParagraphProperties.SectionProperties;
                if (sectionProperties == null) {
                    return;
                }
                this._sectionProperties = sectionProperties;
                //pageMargin1 = new PageMargin() { Top = 40, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };

            } else {
                var sectionProperties = wordDocument._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                if (sectionProperties == null) {
                    sectionProperties = wordDocument._wordprocessingDocument.AddSectionProperties();
                }
                this._sectionProperties = sectionProperties;
                // pageMargin1 = new PageMargin() { Top = 2040, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            }

            // defaults 
            //PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            //Columns columns1 = new Columns() { Space = "720" };
            //DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };
            //this._sectionProperties.Append(pageSize1);
            //this._sectionProperties.Append(pageMargin1);
            //this._sectionProperties.Append(columns1);
            //this._sectionProperties.Append(docGrid1);
            
            wordDocument.Sections.Add(this);

            // this is added for tracking current section of the document
            wordDocument._currentSection = this;

            var listSectionEntries = this._sectionProperties.ChildElements.ToList();
            foreach (var element in listSectionEntries) {
                if (element is HeaderReference) {
                    WordHeader wordHeader = new WordHeader(wordDocument, (HeaderReference)element);
                } else if (element is FooterReference) {
                    WordFooter wordHeader = new WordFooter(wordDocument, (FooterReference)element);
                } else if (element is PageSize) {

                } else if (element is PageMargin) {

                } else if (element is Columns) {

                } else if (element is DocGrid) {

                } else if (element is SectionType) {

                } else if (element is TitlePage) {

                } else {
                    throw new NotImplementedException("This isn't implemented yet?");
                }
            }
        }

        public bool DifferentFirstPage {
            get {
                var sectionProperties = _sectionProperties;
                if (sectionProperties != null) {
                    var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
                    if (titlePage != null) {
                        return true;
                    }
                }
                return false;
            }
            set {
                var sectionProperties = _sectionProperties;
                if (sectionProperties == null) {
                    if (value == false) {
                        // section properties doesn't exists, so we don't do anything
                        return;
                    } else {
                        throw new InvalidOperationException("this is bad");
                    }
                }

                sectionProperties = _sectionProperties;
                var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
                if (value == false) {
                    if (titlePage == null) {
                        return;
                    } else {
                        titlePage.Remove();
                    }
                } else {
                    sectionProperties.Append(new TitlePage());
                }

            }

        }

        public static HeaderFooterValues GetType(string type) {
            if (type == "default") {
                return HeaderFooterValues.Default;
            } else if (type == "even") {
                return HeaderFooterValues.Even;
            } else {
                return HeaderFooterValues.First;
            }
        }

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
                
                //this.Paragraphs.Add(paragraph);
                paragraph.Text = text;
                return paragraph;
                //paragraph._paragraph.InsertAfterSelf(new )
                //paragraph.InsertText(text);
                //_sectionProperties.Parent.Ap
                //return paragraph.InsertParagraphAfterSelf();
            }

            //this._document.InsertParagraph(text);
           // _document._currentSection.InsertParagraph("test");
        }
    }
}
