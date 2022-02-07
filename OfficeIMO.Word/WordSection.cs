using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {
        public List<WordParagraph> Paragraphs => GetParagraphsList();

        public List<WordParagraph> ParagraphsPageBreaks {
            get { return Paragraphs.Where(p => p.IsPageBreak).ToList(); }
        }

        public List<WordParagraph> ParagraphsHyperLinks {
            get { return Paragraphs.Where(p => p.IsHyperLink).ToList(); }
        }

        public List<WordParagraph> ParagraphsFields {
            get { return Paragraphs.Where(p => p.IsField).ToList(); }
        }

        public List<WordParagraph> ParagraphsBookmarks {
            get { return Paragraphs.Where(p => p.IsBookmark).ToList(); }
        }

        public List<WordParagraph> ParagraphsEquations {
            get { return Paragraphs.Where(p => p.IsEquation).ToList(); }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain Structured Document Tags
        /// </summary>
        public List<WordParagraph> ParagraphsStructuredDocumentTags {
            get { return Paragraphs.Where(p => p.IsStructuredDocumentTag).ToList(); }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain Image
        /// </summary>
        public List<WordParagraph> ParagraphsImages {
            get { return Paragraphs.Where(p => p.IsImage).ToList(); }
        }

        public List<WordPageBreak> PageBreaks {
            get {
                List<WordPageBreak> list = new List<WordPageBreak>();
                var paragraphs = Paragraphs.Where(p => p.IsPageBreak).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.PageBreak);
                }
                return list;
            }
        }

        /// <summary>
        /// Exposes Images in their Image form for easy access (saving, modifying)
        /// </summary>
        public List<WordImage> Images {
            get {
                List<WordImage> list = new List<WordImage>();
                var paragraphs = Paragraphs.Where(p => p.IsImage).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Image);
                }
                return list;
            }
        }
        public List<WordBookmark> Bookmarks {
            get {
                List<WordBookmark> list = new List<WordBookmark>();
                var paragraphs = Paragraphs.Where(p => p.IsBookmark).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Bookmark);
                }
                return list;
            }
        }

        public List<WordField> Fields {
            get {
                List<WordField> list = new List<WordField>();
                var paragraphs = Paragraphs.Where(p => p.IsField).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Field);
                }
                return list;
            }
        }

        public List<WordHyperLink> HyperLinks {
            get {
                List<WordHyperLink> list = new List<WordHyperLink>();
                var paragraphs = Paragraphs.Where(p => p.IsHyperLink).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Hyperlink);
                }
                return list;
            }
        }

        public List<WordEquation> Equations {
            get {
                List<WordEquation> list = new List<WordEquation>();
                var paragraphs = Paragraphs.Where(p => p.IsEquation).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Equation);
                }
                return list;
            }
        }

        public List<WordStructuredDocumentTag> StructuredDocumentTags {
            get {
                List<WordStructuredDocumentTag> list = new List<WordStructuredDocumentTag>();
                var paragraphs = Paragraphs.Where(p => p.IsStructuredDocumentTag).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.StructuredDocumentTag);
                }
                return list;
            }
        }

        public WordFooters Footer = new WordFooters();
        public WordHeaders Header = new WordHeaders();

        public WordBorders Borders;
        public WordMargins Margins;

        public List<WordList> Lists {
            get {
                Dictionary<int, List<WordList>> dataLists = new Dictionary<int, List<WordList>>();

                List<WordList> returnList = new List<WordList>();
                if (_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart != null) {
                    var numbering = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    var ids = new List<int>();
                    foreach (var element in numbering.ChildElements.OfType<NumberingInstance>()) {
                        WordList list = new WordList(_document, this, element.NumberID);
                        returnList.Add(list);
                    }
                }

                return returnList;
            }
        }

        public List<WordTable> Tables {
            get { return GetTablesList(); }
        }

        internal WordDocument _document;
        internal SectionProperties _sectionProperties;
        private WordprocessingDocument _wordprocessingDocument;
        private readonly Paragraph _paragraph;


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
            this._paragraph = paragraph;
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
            this._paragraph = paragraph;

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

        internal static HeaderFooterValues GetType(string type) {
            if (type == "default") {
                return HeaderFooterValues.Default;
            } else if (type == "even") {
                return HeaderFooterValues.Even;
            } else {
                return HeaderFooterValues.First;
            }
        }
    }
}