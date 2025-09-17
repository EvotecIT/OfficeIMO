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
        private protected HeaderPart? _headerPart;

        /// <summary>
        /// Underlying OpenXML header element associated with this header or footer.
        /// </summary>
        protected internal Header? _header;

        /// <summary>
        /// Underlying OpenXML footer element associated with this header or footer.
        /// </summary>
        protected internal Footer? _footer;

        protected private FooterPart? _footerPart;
        private protected string? _id;

        /// <summary>
        /// Parent <see cref="WordDocument"/> instance this header or footer belongs to.
        /// </summary>
        protected WordDocument _document = null!;

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
        /// Gets paragraphs that contain a date picker control.
        /// </summary>
        public List<WordParagraph> ParagraphsDatePickers {
            get { return Paragraphs.Where(p => p.IsDatePicker).ToList(); }
        }

        /// <summary>
        /// Gets paragraphs that contain a dropdown list control.
        /// </summary>
        public List<WordParagraph> ParagraphsDropDownLists {
            get { return Paragraphs.Where(p => p.IsDropDownList).ToList(); }
        }

        /// <summary>
        /// Gets paragraphs that contain a repeating section control.
        /// </summary>
        public List<WordParagraph> ParagraphsRepeatingSections {
            get { return Paragraphs.Where(p => p.IsRepeatingSection).ToList(); }
        }
        /// <summary>
        /// Provides a list of paragraphs that contain Image
        /// </summary>
        public List<WordParagraph> ParagraphsImages {
            get { return Paragraphs.Where(p => p.IsImage).ToList(); }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain SmartArt diagrams.
        /// </summary>
        public List<WordParagraph> ParagraphsSmartArts {
            get { return Paragraphs.Where(p => p.IsSmartArt).ToList(); }
        }

        /// <summary>
        /// Gets the page breaks contained in the header or footer.
        /// </summary>
        public List<WordBreak> PageBreaks {
            get {
                List<WordBreak> list = new List<WordBreak>();
                var paragraphs = Paragraphs.Where(p => p.IsPageBreak).ToList();
                foreach (var paragraph in paragraphs) {
                    var pb = paragraph.PageBreak;
                    if (pb != null) list.Add(pb);
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
                    var img = paragraph.Image;
                    if (img != null) list.Add(img);
                }

                return list;
            }
        }

        /// <summary>
        /// Gets the SmartArt diagrams contained in the header or footer.
        /// </summary>
        public List<WordSmartArt> SmartArts {
            get {
                List<WordSmartArt> list = new List<WordSmartArt>();
                var paragraphs = Paragraphs.Where(p => p.IsSmartArt).ToList();
                foreach (var paragraph in paragraphs) {
                    var sa = paragraph.SmartArt;
                    if (sa != null) list.Add(sa);
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
                    var bm = paragraph.Bookmark;
                    if (bm != null) list.Add(bm);
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
                    var f = paragraph.Field;
                    if (f != null) list.Add(f);
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
                    var hl = paragraph.Hyperlink;
                    if (hl != null) list.Add(hl);
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
                    var eq = paragraph.Equation;
                    if (eq != null) list.Add(eq);
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
                    var sdt = paragraph.StructuredDocumentTag;
                    if (sdt != null) list.Add(sdt);
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
                    var cb = paragraph.CheckBox;
                    if (cb != null) list.Add(cb);
                }

                return list;
            }
        }
        /// <summary>
        /// Gets the date pickers contained in the header or footer.
        /// </summary>
        public List<WordDatePicker> DatePickers {
            get {
                List<WordDatePicker> list = new List<WordDatePicker>();
                var paragraphs = Paragraphs.Where(p => p.IsDatePicker).ToList();
                foreach (var paragraph in paragraphs) {
                    var dp = paragraph.DatePicker;
                    if (dp != null) list.Add(dp);
                }

                return list;
            }
        }

        /// <summary>
        /// Gets the dropdown lists contained in the header or footer.
        /// </summary>
        public List<WordDropDownList> DropDownLists {
            get {
                List<WordDropDownList> list = new List<WordDropDownList>();
                var paragraphs = Paragraphs.Where(p => p.IsDropDownList).ToList();
                foreach (var paragraph in paragraphs) {
                    var ddl = paragraph.DropDownList;
                    if (ddl != null) list.Add(ddl);
                }

                return list;
            }
        }

        /// <summary>
        /// Gets the repeating sections contained in the header or footer.
        /// </summary>
        public List<WordRepeatingSection> RepeatingSections {
            get {
                List<WordRepeatingSection> list = new List<WordRepeatingSection>();
                var paragraphs = Paragraphs.Where(p => p.IsRepeatingSection).ToList();
                foreach (var paragraph in paragraphs) {
                    var rs = paragraph.RepeatingSection;
                    if (rs != null) list.Add(rs);
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
