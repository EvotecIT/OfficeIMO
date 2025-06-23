using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides functionality for creating, loading and manipulating Word documents.
    /// </summary>
    public partial class WordDocument : IDisposable {
        internal List<int> _listNumbersUsed = new List<int>();
        internal int? _tableOfContentIndex;
        internal TableOfContentStyle? _tableOfContentStyle;

        internal int BookmarkId {
            get {
                List<int> bookmarksList = new List<int>() { 0 };
                ;
                foreach (var paragraph in this.ParagraphsBookmarks) {
                    bookmarksList.Add(paragraph.Bookmark.Id);
                }

                return bookmarksList.Max() + 1;
            }
        }

        public WordTableOfContent TableOfContent {
            get {
                var sdtBlocks = _document.Body?.ChildElements.OfType<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();
                return WordSection.ConvertStdBlockToTableOfContent(this, sdtBlocks);
            }
        }

        public WordCoverPage CoverPage {
            get {
                var sdtBlocks = _document.Body?.ChildElements.OfType<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();
                return WordSection.ConvertStdBlockToCoverPage(this, sdtBlocks);
            }
        }

        public List<WordParagraph> Paragraphs {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Paragraphs);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsPageBreaks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsPageBreaks);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsBreaks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsBreaks);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsHyperLinks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsHyperLinks);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsTabs {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsTabs);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsTabStops {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsTabStops);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsFields {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsFields);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsBookmarks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsBookmarks);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsEquations {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsEquations);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsStructuredDocumentTags {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsStructuredDocumentTags);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsCheckBoxes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsCheckBoxes);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsCharts {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsCharts);
                }

                return list;
            }
        }

        public List<WordParagraph> ParagraphsEndNotes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsEndNotes);
                }
                return list;
            }
        }

        public List<WordParagraph> ParagraphsTextBoxes {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsTextBoxes);
                }
                return list;
            }
        }

        public List<WordParagraph> ParagraphsEmbeddedObjects {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsEmbeddedObjects);
                }
                return list;
            }
        }

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

        public List<WordBreak> Breaks {
            get {
                List<WordBreak> list = new List<WordBreak>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Breaks);
                }

                return list;
            }
        }

        public List<WordEndNote> EndNotes {
            get {
                List<WordEndNote> list = new List<WordEndNote>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.EndNotes);
                }
                return list;
            }
        }

        public List<WordFootNote> FootNotes {
            get {
                List<WordFootNote> list = new List<WordFootNote>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.FootNotes);
                }
                return list;
            }
        }

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
        public string GetDocumentVariable(string name) {
            return DocumentVariables.ContainsKey(name) ? DocumentVariables[name] : null;
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
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.ParagraphsImages);
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
                foreach (var section in this.Sections) {
                    list.AddRange(section.Images);
                }

                return list;
            }
        }

        public List<WordEmbeddedObject> EmbeddedObjects {
            get {
                List<WordEmbeddedObject> list = new List<WordEmbeddedObject>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.EmbeddedObjects);
                }
                return list;
            }
        }

        public List<WordField> Fields {
            get {
                List<WordField> list = new List<WordField>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Fields);
                }

                return list;
            }
        }

        public List<WordChart> Charts {
            get {
                List<WordChart> list = new List<WordChart>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Charts);
                }
                return list;
            }
        }


        public List<WordHyperLink> HyperLinks {
            get {
                List<WordHyperLink> list = new List<WordHyperLink>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.HyperLinks);
                }

                return list;
            }
        }

        public List<WordTextBox> TextBoxes {
            get {
                List<WordTextBox> list = new List<WordTextBox>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.TextBoxes);
                }
                return list;
            }

        }

        public List<WordTabChar> TabChars {
            get {
                List<WordTabChar> list = new List<WordTabChar>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Tabs);
                }

                return list;
            }
        }

        public List<WordStructuredDocumentTag> StructuredDocumentTags {
            get {
                List<WordStructuredDocumentTag> list = new List<WordStructuredDocumentTag>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.StructuredDocumentTags);
                }

                return list;
            }
        }

        public List<WordCheckBox> CheckBoxes {
            get {
                List<WordCheckBox> list = new List<WordCheckBox>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.CheckBoxes);
                }

                return list;
            }
        }

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
        public string FilePath { get; set; }

        /// <summary>
        /// Provides access to document settings.
        /// </summary>
        public WordSettings Settings;

        /// <summary>
        /// Manages application related properties.
        /// </summary>
        public ApplicationProperties ApplicationProperties;

        /// <summary>
        /// Provides access to built-in document properties.
        /// </summary>
        public BuiltinDocumentProperties BuiltinDocumentProperties;

        /// <summary>
        /// Collection of custom document properties.
        /// </summary>
        public readonly Dictionary<string, WordCustomProperty> CustomDocumentProperties = new Dictionary<string, WordCustomProperty>();
        /// <summary>
        /// Collection of document variables accessible via <see cref="WordField.DocVariable"/> fields.
        /// </summary>
        public Dictionary<string, string> DocumentVariables { get; } = new Dictionary<string, string>();

        /// <summary>
        /// Indicates whether the document is saved automatically.
        /// </summary>
        public bool AutoSave => _wordprocessingDocument.AutoSave;


        // we expose them to help with integration
        /// <summary>
        /// Underlying Open XML word processing document.
        /// </summary>
        public WordprocessingDocument _wordprocessingDocument;

        /// <summary>
        /// Root document element.
        /// </summary>
        public Document _document;
        //public WordCustomProperties _customDocumentProperties;

        private FileStream _fileStream;

        /// <summary>
        /// FileOpenAccess of the document
        /// </summary>
        public FileAccess FileOpenAccess => _wordprocessingDocument.FileOpenAccess;

        private static string GetUniqueFilePath(string filePath) {
            if (File.Exists(filePath)) {
                string folderPath = Path.GetDirectoryName(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string fileExtension = Path.GetExtension(filePath);
                int number = 1;

                Match regex = Regex.Match(fileName, @"^(.+) \((\d+)\)$");

                if (regex.Success) {
                    fileName = regex.Groups[1].Value;
                    number = int.Parse(regex.Groups[2].Value);
                }

                do {
                    number++;
                    string newFileName = $"{fileName} ({number}){fileExtension}";
                    filePath = Path.Combine(folderPath, newFileName);
                } while (File.Exists(filePath));
            }

            return filePath;
        }

        /// <summary>
        /// Create a new WordDocument
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="autoSave"></param>
        /// <returns></returns>
        public static WordDocument Create(string filePath = "", bool autoSave = false) {
            WordDocument word = new WordDocument();

            WordprocessingDocumentType documentType = WordprocessingDocumentType.Document;
            if (!string.IsNullOrEmpty(filePath) && Path.GetExtension(filePath).Equals(".docm", StringComparison.OrdinalIgnoreCase)) {
                documentType = WordprocessingDocumentType.MacroEnabledDocument;
            }
            WordprocessingDocument wordDocument;

            if (filePath != "") {
                //Open the file for writing so as to get lock
                word._fileStream = new FileStream(filePath, FileMode.Create);
            }

            //Always create package in memory.
            wordDocument = WordprocessingDocument.Create(new MemoryStream(), documentType, autoSave);

            wordDocument.AddMainDocumentPart();
            wordDocument.MainDocumentPart.Document = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            wordDocument.MainDocumentPart.Document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            wordDocument.MainDocumentPart.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body();

            word.FilePath = filePath;
            word._wordprocessingDocument = wordDocument;
            word._document = wordDocument.MainDocumentPart.Document;

            StyleDefinitionsPart styleDefinitionsPart1 = wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            WebSettingsPart webSettingsPart1 = wordDocument.MainDocumentPart.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DocumentSettingsPart documentSettingsPart1 = wordDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            EndnotesPart endnotesPart1 = wordDocument.MainDocumentPart.AddNewPart<EndnotesPart>("rId4");
            GenerateEndNotesPart1Content(endnotesPart1);

            FootnotesPart footnotesPart1 = wordDocument.MainDocumentPart.AddNewPart<FootnotesPart>("rId5");
            GenerateFootNotesPart1Content(footnotesPart1);

            FontTablePart fontTablePart1 = wordDocument.MainDocumentPart.AddNewPart<FontTablePart>("rId6");
            GenerateFontTablePart1Content(fontTablePart1);

            ThemePart themePart1 = wordDocument.MainDocumentPart.AddNewPart<ThemePart>("rId7");
            GenerateThemePart1Content(themePart1);

            WordSettings wordSettings = new WordSettings(word);
            WordCompatibilitySettings compatibilitySettings = new WordCompatibilitySettings(word);
            ApplicationProperties applicationProperties = new ApplicationProperties(word);
            BuiltinDocumentProperties builtinDocumentProperties = new BuiltinDocumentProperties(word);
            //CustomDocumentProperties customDocumentProperties = new CustomDocumentProperties(word);
            WordSection wordSection = new WordSection(word, null);
            WordBackground wordBackground = new WordBackground(word);

            // initialize abstract number id for lists to make sure those are unique
            WordListStyles.InitializeAbstractNumberId(word._wordprocessingDocument);

            // initialize abstract number id for lists to make sure those are unique
            WordListStyles.InitializeAbstractNumberId(word._wordprocessingDocument);

            //word.Save();
            return word;
        }

        /// <summary>
        /// PreSaving function to be called before saving the document
        /// </summary>
        private void LoadDocument() {
            Sections.Clear();
            // add settings if not existing
            var wordSettings = new WordSettings(this);
            var applicationProperties = new ApplicationProperties(this);
            var builtinDocumentProperties = new BuiltinDocumentProperties(this);
            var wordCustomProperties = new WordCustomProperties(this);
            var wordDocumentVariables = new WordDocumentVariables(this);
            var wordBackground = new WordBackground(this);
            var compatibilitySettings = new WordCompatibilitySettings(this);
            //CustomDocumentProperties customDocumentProperties = new CustomDocumentProperties(this);
            // add a section that's assigned to top of the document
            var wordSection = new WordSection(this, null, null);

            var list = this._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.ToList(); //.OfType<Paragraph>().ToList();
            foreach (var element in list) {
                if (element is Paragraph) {
                    Paragraph paragraph = (Paragraph)element;
                    if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                        wordSection = new WordSection(this, paragraph.ParagraphProperties.SectionProperties, paragraph);
                    }
                } else if (element is Table) {
                    // WordTable wordTable = new WordTable(this, wordSection, (Table)element);
                } else if (element is SectionProperties sectionProperties) {
                    // we don't do anything as we already created it above - i think
                } else if (element is SdtBlock sdtBlock) {
                    // we don't do anything as we load stuff with get on demand
                } else if (element is OpenXmlUnknownElement) {
                    // this happens when adding dirty element - mainly during TOC Update() function
                } else if (element is BookmarkEnd) {

                } else {
                    //throw new NotImplementedException("This isn't implemented yet");
                }
            }

            RearrangeSectionsAfterLoad();
        }

        /// <summary>
        /// Rearrange sections after loading the document
        /// </summary>
        private void RearrangeSectionsAfterLoad() {
            if (Sections.Count > 0) {
                //var firstElement = Sections[0];
                var firstElementHeader = Sections[0].Header;
                var firstElementFooter = Sections[0].Footer;
                var firstElementSection = Sections[0]._sectionProperties;

                for (int i = 0; i < Sections.Count; i++) {
                    var element = Sections[i];
                    //var tempFooter = element.Footer;
                    //var tempHeader = element.Header;
                    //var tempSectionProp = element._sectionProperties;

                    if (i + 1 < Sections.Count) {
                        Sections[i].Footer = Sections[i + 1].Footer;
                        Sections[i].Header = Sections[i + 1].Header;
                        Sections[i]._sectionProperties = Sections[i + 1]._sectionProperties;

                        Sections[i + 1].Footer = element.Footer;
                        Sections[i + 1].Header = element.Header;
                        Sections[i + 1]._sectionProperties = element._sectionProperties;
                    } else {
                        Sections[i].Footer = firstElementFooter;
                        Sections[i].Header = firstElementHeader;
                        Sections[i]._sectionProperties = firstElementSection;
                    }
                }
            }
        }

        /// <summary>
        /// Load WordDocument from filePath
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="readOnly"></param>
        /// <param name="autoSave"></param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        public static WordDocument Load(string filePath, bool readOnly = false, bool autoSave = false) {
            if (filePath != null) {
                if (!File.Exists(filePath)) {
                    throw new FileNotFoundException("File doesn't exists", filePath);
                }
            }

            var word = new WordDocument();

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            word._fileStream = new FileStream(filePath, FileMode.Open, readOnly ? FileAccess.Read : FileAccess.ReadWrite);
            var memoryStream = new MemoryStream();
            word._fileStream.CopyTo(memoryStream);
            memoryStream.Seek(0, SeekOrigin.Begin);

            var wordDocument = WordprocessingDocument.Open(memoryStream, !readOnly, openSettings);

            InitialiseStyleDefinitions(wordDocument, readOnly);

            word.FilePath = filePath;
            word._wordprocessingDocument = wordDocument;
            word._document = wordDocument.MainDocumentPart.Document;
            word.LoadDocument();
            WordChart.InitializeAxisIdSeed(wordDocument);
            WordChart.InitializeDocPrIdSeed(wordDocument);

            // initialize abstract number id for lists to make sure those are unique
            WordListStyles.InitializeAbstractNumberId(word._wordprocessingDocument);
            return word;
        }

        /// <summary>
        /// Asynchronously loads a <see cref="WordDocument"/> from the given file.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <param name="readOnly">Open the document in read-only mode.</param>
        /// <param name="autoSave">Enable auto-save on dispose.</param>
        /// <returns>Loaded <see cref="WordDocument"/> instance.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public static async Task<WordDocument> LoadAsync(string filePath, bool readOnly = false, bool autoSave = false) {
            if (filePath != null) {
                if (!File.Exists(filePath)) {
                    throw new FileNotFoundException("File doesn't exists", filePath);
                }
            }

            using var fileStream = new FileStream(filePath, FileMode.Open, readOnly ? FileAccess.Read : FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.Asynchronous);
            var memoryStream = new MemoryStream();
            await fileStream.CopyToAsync(memoryStream);
            memoryStream.Seek(0, SeekOrigin.Begin);

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            var wordDocument = WordprocessingDocument.Open(memoryStream, !readOnly, openSettings);

            var word = new WordDocument {
                FilePath = filePath,
                _wordprocessingDocument = wordDocument,
                _document = wordDocument.MainDocumentPart.Document
            };

            InitialiseStyleDefinitions(wordDocument, readOnly);
            word.LoadDocument();
            WordChart.InitializeAxisIdSeed(wordDocument);
            WordChart.InitializeDocPrIdSeed(wordDocument);
            WordListStyles.InitializeAbstractNumberId(word._wordprocessingDocument);
            return word;
        }

        /// <summary>
        /// Load WordDocument from stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="readOnly"></param>
        /// <param name="autoSave"></param>
        /// <returns></returns>
        public static WordDocument Load(Stream stream, bool readOnly = false, bool autoSave = false) {
            var document = new WordDocument();

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            var wordDocument = WordprocessingDocument.Open(stream, !readOnly, openSettings);
            InitialiseStyleDefinitions(wordDocument, readOnly);

            document._wordprocessingDocument = wordDocument;
            document._document = wordDocument.MainDocumentPart.Document;
            document.LoadDocument();
            WordChart.InitializeAxisIdSeed(wordDocument);
            WordChart.InitializeDocPrIdSeed(wordDocument);

            // initialize abstract number id for lists to make sure those are unique
            WordListStyles.InitializeAbstractNumberId(document._wordprocessingDocument);
            return document;
        }

        /// <summary>
        /// Open WordDocument in Microsoft Word (if Word is present)
        /// </summary>
        /// <param name="openWord"></param>
        public void Open(bool openWord = true) {
            this.Open("", openWord);
        }

        /// <summary>
        /// Open WordDocument in Microsoft Word (if Word is present)
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="openWord"></param>
        public void Open(string filePath = "", bool openWord = true) {
            if (filePath == "") {
                filePath = this.FilePath;
            }

            Helpers.Open(filePath, openWord);
        }

        /// <summary>
        /// Copies package properties. Clone and SaveAs don't actually clone document properties for some reason, so they must be copied manually
        /// </summary>
        /// <param name="src"></param>
        /// <param name="dest"></param>
        // IPackageProperties is currently marked as experimental (OOXML0001).
        // There is no non-experimental alternative available yet.
        #pragma warning disable 0618
        private static void CopyPackageProperties(IPackageProperties src, IPackageProperties dest) {
            dest.Category = src.Category;
            dest.ContentStatus = src.ContentStatus;
            dest.ContentType = src.ContentType;
            dest.Created = src.Created;
            dest.Creator = src.Creator;
            dest.Description = src.Description;
            dest.Identifier = src.Identifier;
            dest.Keywords = src.Keywords;
            dest.Language = src.Language;
            dest.LastModifiedBy = src.LastModifiedBy;
            dest.LastPrinted = src.LastPrinted;
            dest.Modified = src.Modified;
            dest.Revision = src.Revision;
            dest.Subject = src.Subject;
            dest.Title = src.Title;
            dest.Version = src.Version;
        }
        #pragma warning restore 0618

        /// <summary>
        /// Save WordDocument to filePath (SaveAs), and open the file in Microsoft Word
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="openWord"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Save(string filePath, bool openWord) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            PreSaving();

            if (this._wordprocessingDocument != null) {
                try {
                    //Save to the memory stream
                    this._wordprocessingDocument.Save();

                    //Open the specified file and copy the bytes
                    if (filePath != "") {
                        //Close existing fileStream
                        if (_fileStream != null) {
                            _fileStream.Dispose();
                        }

                        _fileStream = new FileStream(filePath, FileMode.Create);
                        //Clone and SaveAs don't actually clone document properties for some reason, so they must be copied manually
                        using (var clone = this._wordprocessingDocument.Clone(_fileStream)) {
                            CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                        }
                        Helpers.MakeOpenOfficeCompatible(_fileStream);
                        _fileStream.Flush();
                        FilePath = filePath;
                    } else {
                        if (_fileStream != null) {
                            _fileStream.Seek(0, SeekOrigin.Begin);
                            _fileStream.SetLength(0);
                            //Clone and SaveAs don't actually clone document properties for some reason, so they must be copied manually
                            using (var clone = this._wordprocessingDocument.Clone(_fileStream)) {
                                CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                            }
                            Helpers.MakeOpenOfficeCompatible(_fileStream);
                            _fileStream.Flush();
                        }
                    }
                } catch {
                    throw;
                } finally {
                    if (_fileStream != null) {
                        _fileStream.Dispose();
                        _fileStream = null;
                    }
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (openWord) {
                this.Open(filePath, true);
            }
        }

        /// <summary>
        /// Save WordDocument to where it was open from
        /// </summary>
        public void Save() {
            this.Save("", false);
        }

        /// <summary>
        /// Save WordDocument to given filePath
        /// </summary>
        /// <param name="filePath"></param>
        public void Save(string filePath) {
            this.Save(filePath, false);
        }

        /// <summary>
        /// Save WordDocument and open it in Microsoft Word (if Word is present)
        /// </summary>
        /// <param name="openWord"></param>
        public void Save(bool openWord) {
            this.Save("", openWord);
        }

        /// <summary>
        /// Asynchronously saves the document.
        /// </summary>
        /// <param name="filePath">Optional path to save to.</param>
        /// <param name="openWord">Whether to open Word after saving.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task SaveAsync(string filePath, bool openWord, CancellationToken cancellationToken = default) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            PreSaving();

            if (this._wordprocessingDocument != null) {
                try {
                    this._wordprocessingDocument.Save();

                    if (filePath != "") {
                        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.Asynchronous);
                        using (var clone = this._wordprocessingDocument.Clone(fs)) {
                            CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                        }
                        Helpers.MakeOpenOfficeCompatible(fs);
                        await fs.FlushAsync(cancellationToken);
                        FilePath = filePath;
                    } else if (_fileStream != null) {
                        _fileStream.Seek(0, SeekOrigin.Begin);
                        _fileStream.SetLength(0);
                        using (var clone = this._wordprocessingDocument.Clone(_fileStream)) {
                            CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
                        }
                        Helpers.MakeOpenOfficeCompatible(_fileStream);
                        await _fileStream.FlushAsync(cancellationToken);
                    }
                } finally {
                    if (_fileStream != null) {
                        _fileStream.Dispose();
                        _fileStream = null;
                    }
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (openWord) {
                this.Open(filePath, true);
            }
        }

        public Task SaveAsync(CancellationToken cancellationToken = default) {
            return SaveAsync("", false, cancellationToken);
        }

        public Task SaveAsync(string filePath, CancellationToken cancellationToken = default) {
            return SaveAsync(filePath, false, cancellationToken);
        }

        public Task SaveAsync(bool openWord, CancellationToken cancellationToken = default) {
            return SaveAsync("", openWord, cancellationToken);
        }

        /// <summary>
        /// Save the WordDocument to Stream
        /// </summary>
        /// <param name="outputStream"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Save(Stream outputStream) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("Document is read only, and cannot be saved.");
            }
            PreSaving();

            // Clone document once and copy package properties in the same operation
            using (var clone = this._wordprocessingDocument.Clone(outputStream)) {
                CopyPackageProperties(_wordprocessingDocument.PackageProperties, clone.PackageProperties);
            }

            if (outputStream.CanSeek) {
                outputStream.Seek(0, SeekOrigin.Begin);
            }
        }

        /// <summary>
        /// This moves section within body from top to bottom to allow footers/headers to move
        /// Needs more work, but this is what Word does all the time
        /// </summary>
        private void MoveSectionProperties() {
            var body = this._wordprocessingDocument.MainDocumentPart.Document.Body;
            var sectionProperties = this._wordprocessingDocument.MainDocumentPart.Document.Body.Elements<SectionProperties>().Last();
            body.RemoveChild(sectionProperties);
            body.Append(sectionProperties);
        }

        public void Dispose() {
            if (this._wordprocessingDocument.AutoSave) {
                Save();
            }

            if (this._wordprocessingDocument != null) {
                try {
                    this._wordprocessingDocument.Dispose();
                } catch {
                    // ignored
                }
            }

            if (_fileStream != null) {
                _fileStream.Dispose();
            }
        }

        private static void InitialiseStyleDefinitions(WordprocessingDocument wordDocument, bool readOnly) {
            // if document is read only we shouldn't be doing any new styles, hopefully it doesn't break anything
            if (readOnly == false) {
                var styleDefinitionsPart = wordDocument.MainDocumentPart.GetPartsOfType<StyleDefinitionsPart>()
                    .FirstOrDefault();
                if (styleDefinitionsPart != null) {
                    AddStyleDefinitions(styleDefinitionsPart);
                } else {

                    var styleDefinitionsPart1 = wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
                    GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

                }
            }
        }

        internal WordSection _currentSection => this.Sections.Last();


        /// <summary>
        /// Provides access to the document background settings.
        /// </summary>
        public WordBackground Background { get; set; }

        /// <summary>
        /// Indicates whether the document passes Open XML validation.
        /// </summary>
        public bool DocumentIsValid {
            get {
                if (DocumentValidationErrors.Count > 0) {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        /// Gets the list of validation errors for the document.
        /// </summary>
        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return ValidateDocument();
            }
        }

        /// <summary>
        /// Validates the document using the specified file format version.
        /// </summary>
        /// <param name="fileFormatVersions">File format version to validate against.</param>
        /// <returns>List of validation errors.</returns>
        public List<ValidationErrorInfo> ValidateDocument(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            List<ValidationErrorInfo> listErrors = new List<ValidationErrorInfo>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersions);
            foreach (ValidationErrorInfo error in validator.Validate(this._wordprocessingDocument)) {
                listErrors.Add(error);
            }
            return listErrors;
        }

        /// <summary>
        /// Gets or sets compatibility settings for the document.
        /// </summary>
        public WordCompatibilitySettings CompatibilitySettings { get; set; }

        internal void HeadingModified() {
            if (TableOfContent != null) {
                Settings.UpdateFieldsOnOpen = true;
            }
        }

        private void PreSaving() {
            MoveSectionProperties();
            SaveNumbering();
            _ = new WordCustomProperties(this, true);
            var settingsPart = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart;
            bool hasVariables = settingsPart?.Settings?.GetFirstChild<DocumentVariables>() != null;
            if (hasVariables || DocumentVariables.Count > 0) {
                _ = new WordDocumentVariables(this, true);
            }
        }
    }
}
