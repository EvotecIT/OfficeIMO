using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Returns all paragraphs from every section of the document.
        /// </summary>
        public List<WordParagraph> Paragraphs {
            get {
                return WordSection.GetParagraphsOfAllSections(this);
            }
        }

        /// <summary>
        /// Returns paragraphs that contain page breaks.
        /// </summary>
        public List<WordParagraph> ParagraphsPageBreaks {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsPageBreak).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain any break elements.
        /// </summary>
        public List<WordParagraph> ParagraphsBreaks {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsBreak).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that include hyperlinks.
        /// </summary>
        public List<WordParagraph> ParagraphsHyperLinks {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsHyperLink).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain tab characters.
        /// </summary>
        public List<WordParagraph> ParagraphsTabs {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsTab).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that define tab stops.
        /// </summary>
        public List<WordParagraph> ParagraphsTabStops {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.TabStops.Count > 0).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that include fields.
        /// </summary>
        public List<WordParagraph> ParagraphsFields {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsField).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain bookmarks.
        /// </summary>
        public List<WordParagraph> ParagraphsBookmarks {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsBookmark).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs containing equations.
        /// </summary>
        public List<WordParagraph> ParagraphsEquations {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsEquation).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that host structured document tags.
        /// </summary>
        public List<WordParagraph> ParagraphsStructuredDocumentTags {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsStructuredDocumentTag).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain check boxes.
        /// </summary>
        public List<WordParagraph> ParagraphsCheckBoxes {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsCheckBox).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain date picker controls.
        /// </summary>
        public List<WordParagraph> ParagraphsDatePickers {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsDatePicker).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain dropdown list controls.
        /// </summary>
        public List<WordParagraph> ParagraphsDropDownLists {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsDropDownList).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain combo box controls.
        /// </summary>
        public List<WordParagraph> ParagraphsComboBoxes {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsComboBox).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain picture controls.
        /// </summary>
        public List<WordParagraph> ParagraphsPictureControls {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsPictureControl).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain repeating section controls.
        /// </summary>
        public List<WordParagraph> ParagraphsRepeatingSections {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsRepeatingSection).ToList();
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
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsEndNote).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain text boxes.
        /// </summary>
        public List<WordParagraph> ParagraphsTextBoxes {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsTextBox).ToList();
            }
        }

        /// <summary>
        /// Provides a list of paragraphs that contain shapes.
        /// </summary>
        public List<WordParagraph> ParagraphsShapes {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsShape).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs that contain SmartArt diagrams.
        /// </summary>
        public List<WordParagraph> ParagraphsSmartArts {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsSmartArt).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs containing embedded objects.
        /// </summary>
        public List<WordParagraph> ParagraphsEmbeddedObjects {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsEmbeddedObject).ToList();
            }
        }

        /// <summary>
        /// Returns paragraphs referencing footnotes.
        /// </summary>
        public List<WordParagraph> ParagraphsFootNotes {
            get {
                return WordSection.EnumerateParagraphsOfAllSections(this).Where(p => p.IsFootNote).ToList();
            }
        }
    }
}
