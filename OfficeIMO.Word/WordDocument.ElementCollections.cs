using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
using OfficeIMO.Word.Fluent;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// List of all elements in the document from all the sections
        /// </summary>
        public List<WordElement> Elements {
            get {
                List<WordElement> list = new List<WordElement>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Elements);
                }
                return list;
            }
        }

        /// <summary>
        /// List of all elements in the document from all the sections by their subtype
        /// </summary>
        public List<WordElement> ElementsByType {
            get {
                List<WordElement> list = new List<WordElement>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ElementsByType);
                }
                return list;
            }
        }

        /// <summary>
        /// List of all PageBreaks in the document from all the sections
        /// </summary>
        public List<WordBreak> PageBreaks {
            get {
                List<WordBreak> list = new List<WordBreak>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.PageBreaks);
                }

                return list;
            }
        }

        /// <summary>
        /// Collection of all break elements (page, column, or text wrapping) found across the document.
        /// </summary>
        /// <returns>List of <see cref="WordBreak"/> items representing every break instance.</returns>
        public List<WordBreak> Breaks {
            get {
                List<WordBreak> list = new List<WordBreak>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Breaks);
                }

                return list;
            }
        }

        /// <summary>
        /// Collection of all endnotes referenced throughout the document.
        /// </summary>
        /// <returns>List of <see cref="WordEndNote"/> items representing endnote references.</returns>
        public List<WordEndNote> EndNotes {
            get {
                List<WordEndNote> list = new List<WordEndNote>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.EndNotes);
                }
                return list;
            }
        }

        /// <summary>
        /// Collection of all footnotes referenced throughout the document.
        /// </summary>
        /// <returns>List of <see cref="WordFootNote"/> items representing footnote references.</returns>
        public List<WordFootNote> FootNotes {
            get {
                List<WordFootNote> list = new List<WordFootNote>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.FootNotes);
                }
                if (list.Count == 0) {
                    // Fallback: enumerate footnotes part when no body references were materialized yet.
                    var footnotes = _wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes;
                    if (footnotes != null) {
                        foreach (var fn in footnotes.ChildElements.OfType<DocumentFormat.OpenXml.Wordprocessing.Footnote>()) {
                            if (fn.Type != null) continue; // skip separators
                            // create a lightweight run containing a reference to this id so WordFootNote can resolve paragraphs
                            var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                            var r = new DocumentFormat.OpenXml.Wordprocessing.Run();
                            r.Append(new DocumentFormat.OpenXml.Wordprocessing.RunProperties());
                            r.Append(new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference() { Id = fn.Id });
                            p.Append(r);
                            list.Add(new WordFootNote(this, p, r));
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets the lists in the document
        /// </summary>
        /// <value>
        /// The lists.
        /// </value>
        public List<WordList> Lists => WordSection.GetAllDocumentsLists(this);

        /// <summary>
        /// Provides a list of Bookmarks in the document from all the sections
        /// </summary>
        public List<WordBookmark> Bookmarks {
            get {
                List<WordBookmark> list = new List<WordBookmark>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Bookmarks);
                }

                return list;
            }
        }

        /// <summary>
        /// Provides a list of all tables within the document from all the sections, excluding nested tables
        /// </summary>
        public List<WordTable> Tables {
            get {
                List<WordTable> list = new List<WordTable>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Tables);
                }

                return list;
            }
        }

        /// <summary>
        /// Provides a list of all watermarks within the document from all the
        /// sections, including watermarks defined in headers.
        /// </summary>
        public List<WordWatermark> Watermarks {
            get {
                List<WordWatermark> list = new List<WordWatermark>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Watermarks);
                }

                return list;
            }
        }

        /// <summary>
        /// Provides a list of all embedded documents within the document.
        /// </summary>
        public List<WordEmbeddedDocument> EmbeddedDocuments {
            get {
                List<WordEmbeddedDocument> list = new List<WordEmbeddedDocument>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.EmbeddedDocuments);
                }
                return list;
            }
        }

        /// <summary>
        /// Provides a list of all tables within the document from all the sections, including nested tables
        /// </summary>
        public List<WordTable> TablesIncludingNestedTables {
            get {
                List<WordTable> list = new List<WordTable>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.TablesIncludingNestedTables);
                }
                return list;
            }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain Image
        /// </summary>
        public List<WordParagraph> ParagraphsImages {
            get {
                return EnumerateBodyParagraphs().Where(p => p.IsImage).ToList();
            }
        }

        /// <summary>
        /// Exposes Images in their Image form for easy access (saving, modifying)
        /// </summary>
        public List<WordImage> Images {
            get {
                return ParagraphsImages
                    .Select(p => p.Image)
                    .Where(image => image != null)
                    .Cast<WordImage>()
                    .ToList();
            }
        }

        /// <summary>
        /// Provides a list of all embedded objects within the document.
        /// </summary>
        public List<WordEmbeddedObject> EmbeddedObjects {
            get {
                List<WordEmbeddedObject> list = new List<WordEmbeddedObject>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.EmbeddedObjects);
                }
                return list;
            }
        }

        /// <summary>
        /// Provides a list of all fields within the document.
        /// </summary>
        public List<WordField> Fields {
            get {
                List<WordField> list = new List<WordField>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Fields);
                }

                return list;
            }
        }

        /// <summary>
        /// Provides a list of all charts within the document.
        /// </summary>
        public List<WordChart> Charts {
            get {
                return ParagraphsCharts
                    .Select(p => p.Chart)
                    .Where(chart => chart != null)
                    .Cast<WordChart>()
                    .ToList();
            }
        }


        /// <summary>
        /// Collection of all hyperlinks in the document.
        /// </summary>
        public List<WordHyperLink> HyperLinks {
            get {
                List<WordHyperLink> list = new List<WordHyperLink>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.HyperLinks);
                }

                return list;
            }
        }

        /// <summary>
        /// Collection of all text boxes in the document.
        /// </summary>
        public List<WordTextBox> TextBoxes {
            get {
                List<WordTextBox> list = new List<WordTextBox>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.TextBoxes);
                }
                return list;
            }

        }

        /// <summary>
        /// Collection of all shapes in the document.
        /// </summary>
        public List<WordShape> Shapes {
            get {
                List<WordShape> list = new List<WordShape>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Shapes);
                }
                return list;
            }

        }

        /// <summary>
        /// Collection of all SmartArt diagrams in the document.
        /// </summary>
        public List<WordSmartArt> SmartArts {
            get {
                List<WordSmartArt> list = new List<WordSmartArt>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.SmartArts);
                }
                return list;
            }

        }

        /// <summary>
        /// Collection of tab character elements in the document.
        /// </summary>
        public List<WordTabChar> TabChars {
            get {
                List<WordTabChar> list = new List<WordTabChar>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Tabs);
                }

                return list;
            }
        }

        /// <summary>
        /// Collection of structured document tags in the document.
        /// </summary>
        public List<WordStructuredDocumentTag> StructuredDocumentTags {
            get {
                List<WordStructuredDocumentTag> list = new List<WordStructuredDocumentTag>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.StructuredDocumentTags);
                }

                return list;
            }
        }

        /// <summary>
        /// Collection of all check boxes in the document.
        /// </summary>
        public List<WordCheckBox> CheckBoxes {
            get {
                List<WordCheckBox> list = new List<WordCheckBox>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.CheckBoxes);
                }

                return list;
            }
        }

        /// <summary>
        /// Collection of all date picker controls in the document.
        /// </summary>
        public List<WordDatePicker> DatePickers {
            get {
                List<WordDatePicker> list = new List<WordDatePicker>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.DatePickers);
                }
                return list;
            }
        }

        /// <summary>
        /// Collection of all dropdown list controls in the document.
        /// </summary>
        public List<WordDropDownList> DropDownLists {
            get {
                List<WordDropDownList> list = new List<WordDropDownList>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.DropDownLists);
                }
                return list;
            }
        }

        /// <summary>
        /// Collection of all combo box controls in the document.
        /// </summary>
        public List<WordComboBox> ComboBoxes {
            get {
                List<WordComboBox> list = new List<WordComboBox>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ComboBoxes);
                }
                return list;
            }
        }

        /// <summary>
        /// Collection of all picture controls in the document.
        /// </summary>
        public List<WordPictureControl> PictureControls {
            get {
                List<WordPictureControl> list = new List<WordPictureControl>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.PictureControls);
                }
                return list;
            }
        }

        /// <summary>
        /// Collection of all repeating section controls in the document.
        /// </summary>
        public List<WordRepeatingSection> RepeatingSections {
            get {
                List<WordRepeatingSection> list = new List<WordRepeatingSection>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.RepeatingSections);
                }
                return list;
            }
        }
        /// <summary>
        /// Collection of all equations in the document.
        /// </summary>
        public List<WordEquation> Equations {
            get {
                List<WordEquation> list = new List<WordEquation>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Equations);
                }

                return list;
            }
        }
    }
}
