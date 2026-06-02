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
        /// Returns all paragraphs from every section of the document.
        /// </summary>
        public List<WordParagraph> Paragraphs {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Paragraphs);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain page breaks.
        /// </summary>
        public List<WordParagraph> ParagraphsPageBreaks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsPageBreaks);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain any break elements.
        /// </summary>
        public List<WordParagraph> ParagraphsBreaks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsBreaks);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that include hyperlinks.
        /// </summary>
        public List<WordParagraph> ParagraphsHyperLinks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsHyperLinks);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain tab characters.
        /// </summary>
        public List<WordParagraph> ParagraphsTabs {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsTabs);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that define tab stops.
        /// </summary>
        public List<WordParagraph> ParagraphsTabStops {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsTabStops);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that include fields.
        /// </summary>
        public List<WordParagraph> ParagraphsFields {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsFields);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain bookmarks.
        /// </summary>
        public List<WordParagraph> ParagraphsBookmarks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsBookmarks);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs containing equations.
        /// </summary>
        public List<WordParagraph> ParagraphsEquations {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsEquations);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that host structured document tags.
        /// </summary>
        public List<WordParagraph> ParagraphsStructuredDocumentTags {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsStructuredDocumentTags);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain check boxes.
        /// </summary>
        public List<WordParagraph> ParagraphsCheckBoxes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsCheckBoxes);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain date picker controls.
        /// </summary>
        public List<WordParagraph> ParagraphsDatePickers {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsDatePickers);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain dropdown list controls.
        /// </summary>
        public List<WordParagraph> ParagraphsDropDownLists {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsDropDownLists);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain combo box controls.
        /// </summary>
        public List<WordParagraph> ParagraphsComboBoxes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsComboBoxes);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain picture controls.
        /// </summary>
        public List<WordParagraph> ParagraphsPictureControls {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsPictureControls);
                }

                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain repeating section controls.
        /// </summary>
        public List<WordParagraph> ParagraphsRepeatingSections {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsRepeatingSections);
                }

                return list;
            }
        }
        /// <summary>
        /// Returns paragraphs with embedded charts.
        /// </summary>
        public List<WordParagraph> ParagraphsCharts {
            get {
                return EnumerateBodyParagraphs().Where(p => p.IsChart).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs referencing endnotes.
        /// </summary>
        public List<WordParagraph> ParagraphsEndNotes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsEndNotes);
                }
                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain text boxes.
        /// </summary>
        public List<WordParagraph> ParagraphsTextBoxes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsTextBoxes);
                }
                return list;
            }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain shapes.
        /// </summary>
        public List<WordParagraph> ParagraphsShapes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsShapes);
                }
                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs that contain SmartArt diagrams.
        /// </summary>
        public List<WordParagraph> ParagraphsSmartArts {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsSmartArts);
                }
                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs containing embedded objects.
        /// </summary>
        public List<WordParagraph> ParagraphsEmbeddedObjects {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsEmbeddedObjects);
                }
                return list;
            }
        }

        /// <summary>
        /// Returns paragraphs referencing footnotes.
        /// </summary>
        public List<WordParagraph> ParagraphsFootNotes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsFootNotes);
                }

                return list;
            }
        }
    }
}
