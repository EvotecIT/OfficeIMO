using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {

        /// <summary>
        /// Provides a list of all paragraphs within the section
        /// </summary>
        public List<WordParagraph> Paragraphs => GetParagraphsList();

        /// <summary>
        /// Provides a list of all paragraphs with page breaks within the section
        /// </summary>
        public List<WordParagraph> ParagraphsPageBreaks {
            get { return Paragraphs.Where(p => p.IsPageBreak).ToList(); }
        }

        /// <summary>
        /// Provides a list of all paragraphs with breaks within the section
        /// </summary>
        public List<WordParagraph> ParagraphsBreaks {
            get { return Paragraphs.Where(p => p.IsBreak).ToList(); }
        }

        internal List<WordParagraph> ParagraphsIsListItem {
            get { return Paragraphs.Where(p => p.IsListItem).ToList(); }
        }

        internal List<int> ParagraphListItemsNumbers {
            get {
                var listNumbers = new List<int>();
                var listItems = Paragraphs.Where(p => p.IsListItem).ToList();
                foreach (var item in listItems) {
                    listNumbers.Add(item._listNumberId.Value);
                }

                return listNumbers.Distinct().ToList();
            }
        }

        public List<WordParagraph> ParagraphsHyperLinks {
            get { return Paragraphs.Where(p => p.IsHyperLink).ToList(); }
        }

        public List<WordParagraph> ParagraphsFields {
            get { return Paragraphs.Where(p => p.IsField).ToList(); }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain Bookmarks
        /// </summary>
        public List<WordParagraph> ParagraphsBookmarks {
            get { return Paragraphs.Where(p => p.IsBookmark).ToList(); }
        }

        /// <summary>
        /// Provies a list of paragraphs that contain Equations
        /// </summary>
        public List<WordParagraph> ParagraphsEquations {
            get { return Paragraphs.Where(p => p.IsEquation).ToList(); }
        }

        /// <summary>
        /// Provies a list of paragraphs that contain Tabs
        /// </summary>
        public List<WordParagraph> ParagraphsTabs {
            get { return Paragraphs.Where(p => p.IsTab).ToList(); }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain TabStops
        /// </summary>
        public List<WordParagraph> ParagraphsTabStops {
            get { return Paragraphs.Where(p => p.TabStops.Count > 0).ToList(); }
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

        public List<WordParagraph> ParagraphsCharts {
            get { return Paragraphs.Where(p => p.IsChart).ToList(); }
        }

        public List<WordParagraph> ParagraphsEndNotes {
            get { return Paragraphs.Where(p => p.IsEndNote).ToList(); }
        }

        public List<WordParagraph> ParagraphsFootNotes {
            get { return Paragraphs.Where(p => p.IsFootNote).ToList(); }
        }

        public List<WordParagraph> ParagraphsTextBoxes {
            get { return Paragraphs.Where(p => p.IsTextBox).ToList(); }
        }

        public List<WordBreak> PageBreaks {
            get {
                List<WordBreak> list = new List<WordBreak>();
                var paragraphs = Paragraphs.Where(p => p.IsPageBreak).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.PageBreak);
                }
                return list;
            }
        }

        public List<WordChart> Charts {
            get {
                List<WordChart> list = new List<WordChart>();
                var paragraphs = Paragraphs.Where(p => p.IsChart).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Chart);
                }
                return list;
            }
        }

        public List<WordBreak> Breaks {
            get {
                List<WordBreak> list = new List<WordBreak>();
                var paragraphs = Paragraphs.Where(p => p.IsBreak).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Break);
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

        public List<WordEndNote> EndNotes {
            get {
                List<WordEndNote> list = new List<WordEndNote>();
                var paragraphs = Paragraphs.Where(p => p.IsEndNote).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.EndNote);
                }
                return list;
            }
        }

        public List<WordFootNote> FootNotes {
            get {
                List<WordFootNote> list = new List<WordFootNote>();
                var paragraphs = Paragraphs.Where(p => p.IsFootNote).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.FootNote);
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

                foreach (var table in this.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsHyperLink) {
                            list.Add(paragraph.Hyperlink);
                        }
                    }
                }
                return list;
            }
        }

        public List<WordTabChar> Tabs {
            get {
                List<WordTabChar> list = new List<WordTabChar>();
                var paragraphs = Paragraphs.Where(p => p.IsTab).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.Tab);
                }
                return list;
            }
        }


        public List<WordTextBox> TextBoxes {
            get {
                List<WordTextBox> list = new List<WordTextBox>();
                var paragraphs = Paragraphs.Where(p => p.IsTextBox).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.TextBox);
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
        public WordPageSizes PageSettings;

        /// <summary>
        /// Provides a list of all lists within the section
        /// </summary>
        public List<WordList> Lists => GetLists();

        /// <summary>
        /// Provides a list of all tables within the section, excluding nested tables
        /// </summary>
        public List<WordTable> Tables => GetTablesList();

        /// <summary>
        /// Provides a list of all embedded documents within the section
        /// </summary>
        public List<WordEmbeddedDocument> EmbeddedDocuments => GetEmbeddedDocumentsList();

        /// <summary>
        /// Provides a list of all watermarks within the section
        /// </summary>
        public List<WordWatermark> Watermarks {
            get {
                var sdtBlockList = GetSdtBlockList();
                return WordSection.ConvertStdBlockToWatermark(_document, sdtBlockList);
            }
        }

        /// <summary>
        /// Provides a list of all tables within the section, including nested tables
        /// </summary>
        public List<WordTable> TablesIncludingNestedTables {
            get {
                List<WordTable> list = new List<WordTable>();
                foreach (var table in Tables) {
                    list.Add(table);
                    // if (table.NestedTables.Count > 0) {
                    list.AddRange(table.NestedTables);
                    //}
                }
                return list;
            }
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
                    WordHeader wordHeader = new WordHeader(wordDocument, (HeaderReference)element, this);
                } else if (element is FooterReference) {
                    WordFooter wordHeader = new WordFooter(wordDocument, (FooterReference)element, this);
                } else if (element is PageSize) {
                } else if (element is PageMargin) {
                } else if (element is PageBorders) {
                } else if (element is Columns) {
                } else if (element is DocGrid) {
                } else if (element is SectionType) {
                } else if (element is TitlePage) {
                } else {
                    Debug.WriteLine($"The section '{element.GetType().Name}' is currently not supported. "
                        + "To request support, open an issue at https://github.com/EvotecIT/OfficeIMO/issues");
                }
            }

            this.Margins = new WordMargins(wordDocument, this);
            this.Borders = new WordBorders(wordDocument, this);
            this.PageSettings = new WordPageSizes(wordDocument, this);
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
            this.PageSettings = new WordPageSizes(wordDocument, this);
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
                        throw new InvalidOperationException("Section doesn't exits. Weird :-)");
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
                    WordHeadersAndFooters.AddHeaderReference(this._document, this, HeaderFooterValues.First);
                    WordHeadersAndFooters.AddFooterReference(this._document, this, HeaderFooterValues.First);
                }
            }
        }

        public bool DifferentOddAndEvenPages {
            get {
                var headerReference = WordHeadersAndFooters.GetHeaderReference(this._document, this, HeaderFooterValues.Even);
                var footerReference = WordHeadersAndFooters.GetFooterReference(this._document, this, HeaderFooterValues.Even);

                var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
                if (headerReference == true && footerReference == true && settings != null) {
                    return true;
                }

                return false;

            }
            set {
                var sectionProperties = _sectionProperties;
                WordHeadersAndFooters.AddHeaderReference(this._document, this, HeaderFooterValues.Even);
                WordHeadersAndFooters.AddFooterReference(this._document, this, HeaderFooterValues.Even);

                var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
                if (value != false) {
                    if (settings == null) {
                        _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.Append(new EvenAndOddHeaders());
                    }
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
