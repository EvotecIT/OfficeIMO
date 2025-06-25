using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the shared functionality for headers and footers and
    /// exposes their content through various collections.
    /// </summary>
    public partial class WordHeaderFooter {
        private protected HeaderFooterValues _type;
        private protected HeaderPart _headerPart;
        protected internal Header _header;
        protected internal Footer _footer;
        protected private FooterPart _footerPart;
        private protected string _id;
        protected WordDocument _document;

        /// <summary>
        /// Gets all paragraphs contained in the header or footer.
        /// </summary>
        public List<WordParagraph> Paragraphs {
            get {
                if (_header != null) {
                    return WordSection.ConvertParagraphsToWordParagraphs(_document, _header.ChildElements.OfType<Paragraph>());
                } else if (_footer != null) {
                    return WordSection.ConvertParagraphsToWordParagraphs(_document, _footer.ChildElements.OfType<Paragraph>());
                }

                return new List<WordParagraph>();
            }
        }

        /// <summary>
        /// Gets all tables contained in the header or footer.
        /// </summary>
        public List<WordTable> Tables {
            get {
                if (_header != null) {
                    return WordSection.ConvertTableToWordTable(_document, _header.ChildElements.OfType<Table>());
                } else if (_footer != null) {
                    return WordSection.ConvertTableToWordTable(_document, _footer.ChildElements.OfType<Table>());
                }

                return new List<WordTable>();
            }
        }

        /// <summary>
        /// Gets paragraphs that contain a page break.
        /// </summary>
        public List<WordParagraph> ParagraphsPageBreaks {
            get { return Paragraphs.Where(p => p.IsPageBreak).ToList(); }
        }

        /// <summary>
        /// Gets paragraphs that contain a hyperlink.
        /// </summary>
        public List<WordParagraph> ParagraphsHyperLinks {
            get { return Paragraphs.Where(p => p.IsHyperLink).ToList(); }
        }

        /// <summary>
        /// Gets paragraphs that contain a field.
        /// </summary>
        public List<WordParagraph> ParagraphsFields {
            get { return Paragraphs.Where(p => p.IsField).ToList(); }
        }

        /// <summary>
        /// Gets paragraphs that contain a bookmark.
        /// </summary>
        public List<WordParagraph> ParagraphsBookmarks {
            get { return Paragraphs.Where(p => p.IsBookmark).ToList(); }
        }

        /// <summary>
        /// Gets paragraphs that contain an equation.
        /// </summary>
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
        /// Gets paragraphs that contain a check box.
        /// </summary>
        public List<WordParagraph> ParagraphsCheckBoxes {
            get { return Paragraphs.Where(p => p.IsCheckBox).ToList(); }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain Image
        /// </summary>
        public List<WordParagraph> ParagraphsImages {
            get { return Paragraphs.Where(p => p.IsImage).ToList(); }
        }

        /// <summary>
        /// Gets the page breaks contained in the header or footer.
        /// </summary>
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

        /// <summary>
        /// Gets the bookmarks contained in the header or footer.
        /// </summary>
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

        /// <summary>
        /// Gets the fields contained in the header or footer.
        /// </summary>
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

        /// <summary>
        /// Gets the hyperlinks contained in the header or footer.
        /// </summary>
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

        /// <summary>
        /// Gets the equations contained in the header or footer.
        /// </summary>
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

        /// <summary>
        /// Gets the structured document tags contained in the header or footer.
        /// </summary>
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

        /// <summary>
        /// Gets the check boxes contained in the header or footer.
        /// </summary>
        public List<WordCheckBox> CheckBoxes {
            get {
                List<WordCheckBox> list = new List<WordCheckBox>();
                var paragraphs = Paragraphs.Where(p => p.IsCheckBox).ToList();
                foreach (var paragraph in paragraphs) {
                    list.Add(paragraph.CheckBox);
                }

                return list;
            }
        }

        /// <summary>
        /// Gets the watermarks contained in the header or footer.
        /// </summary>
        public List<WordWatermark> Watermarks {
            get {
                if (_header != null) {
                    return WordSection.ConvertStdBlockToWatermark(_document, _header.ChildElements.OfType<SdtBlock>());
                } else if (_footer != null) {
                    return WordSection.ConvertStdBlockToWatermark(_document, _footer.ChildElements.OfType<SdtBlock>());
                }

                return new List<WordWatermark>();
            }
        }
    }
}
