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
    /// <summary>
    /// Provides functionality for creating, loading and manipulating Word documents.
    /// </summary>
    public partial class WordDocument : IDisposable {

        internal int BookmarkId {
            get {
                List<int> bookmarksList = new List<int>() { 0 };
                foreach (var paragraph in this.ParagraphsBookmarks) {
                    if (paragraph.Bookmark != null) {
                        bookmarksList.Add(paragraph.Bookmark.Id);
                    }
                }

                return bookmarksList.Max() + 1;
            }
        }

        /// <summary>
        /// Gets the table of contents defined in the document.
        /// </summary>
        public WordTableOfContent? TableOfContent {
            get {
                var sdtBlocks = _document.Body?.ChildElements.OfType<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();
                return WordSection.ConvertStdBlockToTableOfContent(this, sdtBlocks);
            }
        }

        /// <summary>
        /// Gets the cover page if one is defined in the document.
        /// </summary>
        public WordCoverPage? CoverPage {
            get {
                var sdtBlocks = _document.Body?.ChildElements.OfType<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();
                return WordSection.ConvertStdBlockToCoverPage(this, sdtBlocks);
            }
        }

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
        /// Collection of all comments inserted into the document.
        /// </summary>
        /// <returns>List of <see cref="WordComment"/> objects for each comment.</returns>
        public List<WordComment> Comments {
            get { return WordComment.GetAllComments(this); }
        }

        /// <summary>
        /// Removes comment with the specified id.
        /// </summary>
        /// <param name="commentId">Id of the comment to remove.</param>
        public void RemoveComment(string commentId) {
            var comment = this.Comments.FirstOrDefault(c => c.Id == commentId);
            comment?.Delete();
        }

        /// <summary>
        /// Removes the specified comment from the document.
        /// </summary>
        /// <param name="comment">Comment instance to remove.</param>
        public void RemoveComment(WordComment comment) {
            comment?.Delete();
        }

        /// <summary>
        /// Removes all comments from the document.
        /// </summary>
        public void RemoveAllComments() {
            foreach (var comment in this.Comments.ToList()) {
                comment.Delete();
            }
        }

        /// <summary>
        /// Gets the value of a document variable or <c>null</c> if the variable does not exist.
        /// </summary>
        /// <param name="name">Variable name.</param>
        public string? GetDocumentVariable(string name) {
            return DocumentVariables.TryGetValue(name, out var value) ? value : null;
        }

        /// <summary>
        /// Sets the value of a document variable. Creates it if it does not exist.
        /// </summary>
        /// <param name="name">Variable name.</param>
        /// <param name="value">Variable value.</param>
        public void SetDocumentVariable(string name, string value) {
            DocumentVariables[name] = value;
        }

        /// <summary>
        /// Removes the document variable with the specified name if present.
        /// </summary>
        /// <param name="name">Variable name.</param>
        public void RemoveDocumentVariable(string name) {
            DocumentVariables.Remove(name);
        }

        /// <summary>
        /// Removes the document variable at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the variable to remove.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when index is out of range.</exception>
        public void RemoveDocumentVariableAt(int index) {
            if (index < 0 || index >= DocumentVariables.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            string key = DocumentVariables.Keys.ElementAt(index);
            DocumentVariables.Remove(key);
        }

        /// <summary>
        /// Determines whether the document contains any document variables.
        /// </summary>
        public bool HasDocumentVariables => DocumentVariables.Count > 0;

        /// <summary>
        /// Returns a read-only view of all document variables.
        /// </summary>
        public IReadOnlyDictionary<string, string> GetDocumentVariables() {
            return new Dictionary<string, string>(DocumentVariables);
        }

        /// <summary>
        /// Enable or disable tracking of comment changes.
        /// </summary>
        public bool TrackComments {
            get => this.Settings.TrackComments;
            set => this.Settings.TrackComments = value;
        }

        /// <summary>
        /// Enable or disable tracking of all revisions, moves and formatting changes.
        /// </summary>
        public bool TrackChanges {
            get => this.Settings.TrackRevisions;
            set {
                this.Settings.TrackRevisions = value;
                this.Settings.TrackFormatting = value;
                this.Settings.TrackMoves = value;
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

        /// <summary>
        /// Collection of sections contained in the document.
        /// </summary>
        public List<WordSection> Sections = new List<WordSection>();

        /// <summary>
        /// Path to the file backing this document.
        /// </summary>
        public string FilePath { get; set; } = null!;

        /// <summary>
        /// Original stream where this document was created / loaded from.
        /// </summary>
        internal Stream OriginalStream { get; set; } = null!;

        /// <summary>
        /// Provides access to document settings.
        /// </summary>
        public WordSettings Settings = null!;

        /// <summary>
        /// Manages application related properties.
        /// </summary>
        public ApplicationProperties ApplicationProperties = null!;

        /// <summary>
        /// Provides access to built-in document properties.
        /// </summary>
        public BuiltinDocumentProperties BuiltinDocumentProperties = null!;

        private WordCoverPageProperties? _coverPageProperties;

        /// <summary>
        /// Provides access to the cover page properties custom XML part used by built-in templates.
        /// </summary>
        public WordCoverPageProperties CoverPageProperties => _coverPageProperties ??= new WordCoverPageProperties(this);

        /// <summary>
        /// Collection of custom document properties.
        /// </summary>
        public readonly Dictionary<string, WordCustomProperty> CustomDocumentProperties = new Dictionary<string, WordCustomProperty>();
        /// <summary>
        /// Collection of document variables accessible via <see cref="WordFieldType.DocVariable"/> fields.
        /// </summary>
        public Dictionary<string, string> DocumentVariables { get; } = new Dictionary<string, string>();

        /// <summary>
        /// Collection of bibliographic sources used in the document.
        /// </summary>
        public Dictionary<string, WordBibliographySource> BibliographySources { get; } = new Dictionary<string, WordBibliographySource>();

        /// <summary>
        /// Provides basic statistics for the document.
        /// </summary>
        public WordDocumentStatistics Statistics { get; internal set; } = null!;

        /// <summary>
        /// Indicates whether the document is saved automatically.
        /// </summary>
        public bool AutoSave => _wordprocessingDocument.AutoSave;

        /// <summary>
        /// When <c>true</c> the table of contents is flagged to update before saving.
        /// </summary>
        public bool AutoUpdateToc { get; set; }


        // we expose them to help with integration
        /// <summary>
        /// Underlying Open XML word processing document.
        /// </summary>
        public WordprocessingDocument _wordprocessingDocument = null!;

        /// <summary>
        /// Root document element.
        /// </summary>
        public Document _document = null!;
        //public WordCustomProperties _customDocumentProperties;


        /// <summary>
        /// FileOpenAccess of the document
        /// </summary>
        public FileAccess FileOpenAccess => _wordprocessingDocument.FileOpenAccess;
    }
}
