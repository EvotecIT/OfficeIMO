using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word {
    public partial class WordSection {
        public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        public List<WordParagraph> PageBreaks = new List<WordParagraph>();
        public List<WordImage> Images = new List<WordImage>();

        public WordFooters Footer = new WordFooters();
        public WordHeaders Header = new WordHeaders();

        public List<WordList> Lists = new List<WordList>();

        public WordBorders Borders;
        public WordMargins Margins;

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
        private WordprocessingDocument _wordprocessingDocument;


        public WordSection SetMargins(PageMargin pageMargins) {
            var pageMargin = _sectionProperties.GetFirstChild<PageMargin>();
            if (pageMargin == null) {
                _sectionProperties.Append(pageMargins);
                // pageMargin = _sectionProperties.GetFirstChild<PageMargin>();
            } else {
                pageMargin.Remove();
                _sectionProperties.Append(pageMargins);
            }

            return this;
        }


        /// <summary>
        /// This method moves headers and footers and title page to section before it.
        /// It also copies copies all other parts of sections (PageSize,PageMargin and others) to section before it.
        /// This is because headers/footers when applied to section apply to the rest of the document
        /// unless there are headers/footers on next section.
        /// On the other hand page size doesn't apply to other sections
        /// and word uses default values. 
        /// </summary>
        /// <param name="sectionProperties"></param>
        /// <param name="newSectionProperties"></param>
        private static void CopySectionProperties(SectionProperties sectionProperties, SectionProperties newSectionProperties) {
            if (newSectionProperties.ChildElements.Count == 0) {
                var listSectionEntries = sectionProperties.ChildElements.ToList();
                foreach (var element in listSectionEntries) {
                    if (element is HeaderReference) {
                        newSectionProperties.Append(element.CloneNode(true));
                        sectionProperties.RemoveChild(element);
                    } else if (element is FooterReference) {
                        newSectionProperties.Append(element.CloneNode(true));
                        sectionProperties.RemoveChild(element);
                        //} else if (element is PageSize) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                    } else if (element is PageMargin) {
                        newSectionProperties.Append(element.CloneNode(true));
                        //sectionProperties.RemoveChild(element);
                        //} else if (element is Columns) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                        //} else if (element is DocGrid) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                        //} else if (element is SectionType) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                    } else if (element is TitlePage) {
                        newSectionProperties.Append(element.CloneNode(true));
                        sectionProperties.RemoveChild(element);
                    } else {
                        newSectionProperties.Append(element.CloneNode(true));
                        //throw new NotImplementedException("This isn't implemented yet?");
                    }
                }
                //sectionProperties.RemoveAllChildren();
                //newSectionProperties.Append(listSectionEntries);
            }
        }

        /// <summary>
        /// Used to load WordSection withing word document
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <param name="sectionProperties"></param>
        /// <param name="paragraph"></param>
        /// <exception cref="NotImplementedException"></exception>
        internal WordSection(WordDocument wordDocument, SectionProperties sectionProperties = null, Paragraph paragraph = null) {
            this._document = wordDocument;
            this._wordprocessingDocument = wordDocument._wordprocessingDocument;
            if (sectionProperties != null) {
                this._sectionProperties = sectionProperties;
            } else {
                sectionProperties = wordDocument._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                if (sectionProperties == null) {
                    // most likely not necessary during load - but lets see
                    // would require a broken document created by some app
                    sectionProperties = wordDocument._wordprocessingDocument.AddSectionProperties();
                }

                this._sectionProperties = sectionProperties;
            }

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
                } else if (element is PageBorders) {
                } else if (element is Columns) {
                } else if (element is DocGrid) {
                } else if (element is SectionType) {
                } else if (element is TitlePage) {
                } else {
                    throw new NotImplementedException("This isn't implemented yet?");
                }
            }

            this.Margins = new WordMargins(wordDocument, this);
            this.Borders = new WordBorders(wordDocument, this);
        }


        /// <summary>
        /// Used for creating WordSection in new documents
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <param name="paragraph"></param>
        internal WordSection(WordDocument wordDocument, Paragraph paragraph = null) {
            this._document = wordDocument;
            this._wordprocessingDocument = wordDocument._wordprocessingDocument;

            if (paragraph != null) {
                var sectionProperties = paragraph.ParagraphProperties.SectionProperties;
                if (sectionProperties == null) {
                    return;
                }

                this._sectionProperties = sectionProperties;
            } else {
                var sectionProperties = wordDocument._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                if (sectionProperties == null) {
                    sectionProperties = wordDocument._wordprocessingDocument.AddSectionProperties();
                }
                this._sectionProperties = sectionProperties;
            }

            if (this._document.Sections.Count > 0) {
                WordSection lastSection = this._document.Sections[this._document.Sections.Count - 1];
                //lastSection._sectionProperties.

                var temporarySectionProperties = lastSection._sectionProperties;
                if (temporarySectionProperties != null) {

                    CopySectionProperties(lastSection._sectionProperties, this._sectionProperties);
                    var old = this._sectionProperties;

                    this._sectionProperties = lastSection._sectionProperties;
                    lastSection._sectionProperties = old;


                }
            }


            this.Margins = new WordMargins(wordDocument, this);
            this.Borders = new WordBorders(wordDocument, this);

            wordDocument.Sections.Add(this);

            // this is added for tracking current section of the document
            wordDocument._currentSection = this;

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
                    WordHeadersAndFooters.AddHeaderReference1(this._document, this, HeaderFooterValues.First);
                    WordHeadersAndFooters.AddFooterReference1(this._document, this, HeaderFooterValues.First);
                }
            }
        }

        public bool DifferentOddAndEvenPages {
            get {
                if (this == this._document.Sections[0]) {
                    var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
                    if (settings != null) {
                        return true;
                    } else {
                        return false;
                    }
                } else {
                    throw new NotImplementedException("Not implemented for other sections");
                    return false;
                }
            }
            set {
                var sectionProperties = _sectionProperties;
                WordHeadersAndFooters.AddHeaderReference1(this._document, this, HeaderFooterValues.Even);
                WordHeadersAndFooters.AddFooterReference1(this._document, this, HeaderFooterValues.Even);

                if (this == this._document.Sections[0]) {
                    var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
                    if (value == false) {
                    } else {
                        if (settings == null) {
                            _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.Append(new EvenAndOddHeaders());
                        } else {
                            // noting to do, already enabled
                        }
                    }
                } else {
                }
            }
        }

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
        internal static HeaderFooterValues GetType(string type) {
            if (type == "default") {
                return HeaderFooterValues.Default;
            } else if (type == "even") {
                return HeaderFooterValues.Even;
            } else {
                return HeaderFooterValues.First;
            }
        }

        public WordParagraph AddParagraph(string text) {
            if (this.Paragraphs.Count == 0) {
                WordParagraph paragraph = this._document.AddParagraph(text);
                paragraph._section = this;
                return paragraph;
            } else {
                WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();

                WordParagraph paragraph = lastParagraphWithinSection.AddParagraphAfterSelf(this);
                paragraph._document = this._document;
                paragraph._section = this;
                //this.Paragraphs.Add(paragraph);
                paragraph.Text = text;
                return paragraph;
            }
        }

        public WordWatermark AddWatermark(WordWatermarkStyle watermarkStyle, string text) {
            return new WordWatermark(this._document, this, this.Header.Default, watermarkStyle, text);
        }

        public WordSection SetBorders(WordBorder wordBorder) {
            this.Borders.SetBorder(wordBorder);

            return this;
        }
    }
}