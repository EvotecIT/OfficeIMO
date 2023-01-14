using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument : IDisposable {
        internal List<int> _listNumbersUsed = new List<int>();

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

                foreach (var sdtBlock in sdtBlocks) {
                    var sdtProperties = sdtBlock?.ChildElements.OfType<SdtProperties>().FirstOrDefault();
                    var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
                    var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

                    if (docPartGallery != null && docPartGallery.Val == "Table of Contents") {
                        return new WordTableOfContent(this, sdtBlock);
                    }
                }

                return null;
            }
        }

        public WordCoverPage CoverPage {
            get {
                var sdtBlocks = _document.Body?.ChildElements.OfType<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();

                foreach (var sdtBlock in sdtBlocks) {
                    var sdtProperties = sdtBlock?.ChildElements.OfType<SdtProperties>().FirstOrDefault();
                    var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
                    var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

                    if (docPartGallery != null && docPartGallery.Val == "Cover Pages") {
                        return new WordCoverPage(this, sdtBlock);
                    }
                }

                return null;
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

        public List<WordComment> Comments {
            get { return WordComment.GetAllComments(this); }
        }

        public List<WordList> Lists {
            get {
                return WordSection.GetAllDocumentsLists(this);
                //List<WordList> list = new List<WordList>();
                //foreach (var section in this.Sections) {
                //    list.AddRange(section.Lists);
                //}

                //return list;
            }
        }

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

        public List<WordField> Fields {
            get {
                List<WordField> list = new List<WordField>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Fields);
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

        public List<WordStructuredDocumentTag> StructuredDocumentTags {
            get {
                List<WordStructuredDocumentTag> list = new List<WordStructuredDocumentTag>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.StructuredDocumentTags);
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

        public List<WordSection> Sections = new List<WordSection>();

        public string FilePath { get; set; }

        public WordSettings Settings;

        public ApplicationProperties ApplicationProperties;
        public BuiltinDocumentProperties BuiltinDocumentProperties;

        public readonly Dictionary<string, WordCustomProperty> CustomDocumentProperties = new Dictionary<string, WordCustomProperty>();

        public bool AutoSave => _wordprocessingDocument.AutoSave;


        // we expose them to help with integration
        public WordprocessingDocument _wordprocessingDocument;
        public Document _document;
        //public WordCustomProperties _customDocumentProperties;

        private FileStream _fileStream;

        public FileAccess FileOpenAccess {
            get { return _wordprocessingDocument.MainDocumentPart.OpenXmlPackage.Package.FileOpenAccess; }
        }

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

        public static WordDocument Create(string filePath = "", bool autoSave = false) {
            WordDocument word = new WordDocument();

            WordprocessingDocumentType documentType = WordprocessingDocumentType.Document;
            WordprocessingDocument wordDocument;

            if (filePath != "") {
                //Open the file for writing so as to get lock
                word._fileStream = new FileStream(filePath, FileMode.Create);
            }

            //Always create package in memory.
            wordDocument = WordprocessingDocument.Create(new MemoryStream(), documentType, autoSave);

            wordDocument.AddMainDocumentPart();
            wordDocument.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
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

            //FontTablePart fontTablePart1 = wordDocument.MainDocumentPart.AddNewPart<FontTablePart>("rId4");
            //GenerateFontTablePart1Content(fontTablePart1);

            //ThemePart themePart1 = wordDocument.MainDocumentPart.AddNewPart<ThemePart>("rId5");
            //GenerateThemePart2Content(themePart1);

            WordSettings wordSettings = new WordSettings(word);
            WordCompatibilitySettings compatibilitySettings = new WordCompatibilitySettings(word);
            ApplicationProperties applicationProperties = new ApplicationProperties(word);
            BuiltinDocumentProperties builtinDocumentProperties = new BuiltinDocumentProperties(word);
            //CustomDocumentProperties customDocumentProperties = new CustomDocumentProperties(word);
            WordSection wordSection = new WordSection(word, null);
            WordBackground wordBackground = new WordBackground(word);

            word.Save();
            return word;
        }

        private void LoadDocument() {
            Sections.Clear();
            // add settings if not existing
            var wordSettings = new WordSettings(this);
            var applicationProperties = new ApplicationProperties(this);
            var builtinDocumentProperties = new BuiltinDocumentProperties(this);
            var wordCustomProperties = new WordCustomProperties(this);
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
            return word;
        }

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
            return document;
        }

        public void Open(bool openWord = true) {
            this.Open("", openWord);
        }

        public void Open(string filePath = "", bool openWord = true) {
            if (filePath == "") {
                filePath = this.FilePath;
            }

            Helpers.Open(filePath, openWord);
        }


        //private void LoadNumbering() {
        //    if (_wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart != null) {
        //        Numbering numbering = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;
        //        if (numbering == null) {
        //        } else {
        //            var tempAbstractNumList = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<AbstractNum>();
        //            foreach (AbstractNum abstractNum in tempAbstractNumList) {
        //               // _ListAbstractNum.Add(abstractNum);
        //            }

        //            var tempNumberingInstance = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<NumberingInstance>();
        //            foreach (NumberingInstance numberingInstance in tempNumberingInstance) {
        //                //_listNumberingInstances.Add(numberingInstance);
        //            }
        //        }
        //    }
        //}
        //private void SaveSections() {
        //    WordSection temporarySection = null;
        //    if (this.Sections.Count > 0) {
        //        for (int i = 0; i < Sections.Count; i++) {
        //            if (temporarySection != null) {

        //            } else {
        //                temporarySection = Sections[i];
        //                Sections[i]._sectionProperties.Remove();
        //            }
        //        }
        //    }
        //}

        private static void CopyPackageProperties(PackageProperties src, PackageProperties dest) {
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

        public void Save(string filePath, bool openWord) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new Exception("Document is read only, and cannot be saved.");
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
                            _fileStream.Flush();
                        }
                    }
                } catch {
                    throw;
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (openWord) {
                _fileStream.Dispose();
                _fileStream = null;
                this.Open(filePath, openWord);
            }
        }

        public void Save() {
            this.Save("", false);
        }

        public void Save(string filePath) {
            this.Save(filePath, false);
        }

        public void Save(bool openWord) {
            this.Save("", openWord);
        }

        public void Save(Stream outputStream) {
            if (FileOpenAccess == FileAccess.Read) {
                throw new Exception("Document is read only, and cannot be saved.");
            }
            PreSaving();

            this._wordprocessingDocument.Clone(outputStream);

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
                    this._wordprocessingDocument.Close();
                } catch {
                    // ignored
                }
                this._wordprocessingDocument.Dispose();
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


        public WordBackground Background { get; set; }

        public bool DocumentIsValid {
            get {
                if (DocumentValidationErrors.Count > 0) {
                    return false;
                }

                return true;
            }
        }

        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return ValidateDocument();
            }
        }

        public List<ValidationErrorInfo> ValidateDocument(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            List<ValidationErrorInfo> listErrors = new List<ValidationErrorInfo>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersions);
            foreach (ValidationErrorInfo error in validator.Validate(this._wordprocessingDocument)) {
                listErrors.Add(error);
            }
            return listErrors;
        }

        public WordCompatibilitySettings CompatibilitySettings { get; set; }

        private void PreSaving() {
            MoveSectionProperties();
            SaveNumbering();
            _ = new WordCustomProperties(this, true);
        }
    }
}
