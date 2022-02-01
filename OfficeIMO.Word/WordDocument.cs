using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word {
    public partial class WordDocument : IDisposable {
        internal List<int> _listNumbersUsed = new List<int>();

        //internal int _listNumbers;

        //internal List<int> _listNumbersUsed = new List<int>();
        //internal List<NumberingInstance> _listNumberingInstances = new List<NumberingInstance>();
        //internal List<AbstractNum> _ListAbstractNum = new List<AbstractNum>();

        //public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        public WordTableOfContent TableOfContent {
            get {
                SdtBlock sdtBlock = _document.Body.ChildElements.OfType<SdtBlock>().FirstOrDefault();
                if (sdtBlock != null) {
                    return new WordTableOfContent(this, sdtBlock);
                }
                return null;
            }
        }

        public List<WordParagraph> Paragraphs {
            get {
                //if (this.Sections.Count > 1) {
                //    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Headers.");
                //}
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Paragraphs);
                }

                return list;
            }
        }

        public List<WordParagraph> PageBreaks {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.PageBreaks);
                }

                return list;
            }
        }

        public List<WordList> Lists {
            get {
                List<WordList> list = new List<WordList>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Lists);
                }

                return list;
            }
        }

        public List<WordTable> Tables {
            get {
                List<WordTable> list = new List<WordTable>();
                foreach (var section in this.Sections) {
                    list.AddRange(section.Tables);
                }

                return list;
            }
        }

        //public List<WordParagraph> PageBreaks = new List<WordParagraph>();
        public List<WordImage> Images = new List<WordImage>();
        public readonly List<WordSection> Sections = new List<WordSection>();

        public string FilePath { get; set; }

        public WordSettings Settings;

        public ApplicationProperties ApplicationProperties;
        public BuiltinDocumentProperties BuiltinDocumentProperties;

        public Dictionary<string, WordCustomProperty> CustomDocumentProperties = new Dictionary<string, WordCustomProperty>();
        public WordCustomProperties _customDocumentProperties;
        //internal WordLists WordLists;

        public bool AutoSave {
            get { return _wordprocessingDocument.AutoSave; }
        }

        internal WordprocessingDocument _wordprocessingDocument = null;
        public Document _document = null;


        public FileAccess FileOpenAccess {
            get { return _wordprocessingDocument.MainDocumentPart.OpenXmlPackage.Package.FileOpenAccess; }
        }
        public static string GetUniqueFilePath(string filePath) {
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
                }
                while (File.Exists(filePath));
            }

            return filePath;
        }

        public static WordDocument Create(string filePath = "", bool autoSave = false) {
            WordDocument word = new WordDocument();

            WordprocessingDocumentType documentType = WordprocessingDocumentType.Document;
            WordprocessingDocument wordDocument;

            if (filePath != "") {
                //System.IO.IOException: 'The process cannot access the file 'C:\Support\GitHub\OfficeIMO\OfficeIMO.Examples\bin\Debug\net5.0\Documents\Basic Document with some sections 1.docx' because it is being used by another process.'
                try {
                    wordDocument = WordprocessingDocument.Create(filePath, documentType, autoSave);
                } catch {
                    filePath = GetUniqueFilePath(filePath);
                    wordDocument = WordprocessingDocument.Create(filePath, documentType, autoSave);
                }
            } else {
                MemoryStream mem = new MemoryStream();
                //word._memory = mem;
                wordDocument = WordprocessingDocument.Create(mem, documentType, autoSave);
            }


            //ExtendedFilePropertiesPart extendedFilePropertiesPart1 = wordDocument.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            //GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            //MainDocumentPart mainDocumentPart1 = wordDocument.AddMainDocumentPart();
            //GenerateMainDocumentPart1Content(mainDocumentPart1);

            //WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            //GenerateWebSettingsPart1Content(webSettingsPart1);

            //DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            //GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            //StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            //GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            //ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId5");
            //GenerateThemePart1Content(themePart1);

            //FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId4");
            //GenerateFontTablePart1Content(fontTablePart1);

            wordDocument.AddMainDocumentPart();
            wordDocument.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            wordDocument.MainDocumentPart.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body();

            //wordDocument.AddHeadersAndFooters(word);

            word.FilePath = filePath;
            word._wordprocessingDocument = wordDocument;
            word._document = wordDocument.MainDocumentPart.Document;

            StyleDefinitionsPart styleDefinitionsPart1 = wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            WordSettings wordSettings = new WordSettings(word);
            ApplicationProperties applicationProperties = new ApplicationProperties(word);
            BuiltinDocumentProperties builtinDocumentProperties = new BuiltinDocumentProperties(word);
            //CustomDocumentProperties customDocumentProperties = new CustomDocumentProperties(word);
            WordSection wordSection = new WordSection(word, null);
            WordBackground wordBackground = new WordBackground(word);
            return word;
        }

        private void LoadDocument() {
            // add settings if not existing
            var wordSettings = new WordSettings(this);
            var applicationProperties = new ApplicationProperties(this);
            var builtinDocumentProperties = new BuiltinDocumentProperties(this);
            var wordCustomProperties = new WordCustomProperties(this);
            var wordBackground = new WordBackground(this);
            //CustomDocumentProperties customDocumentProperties = new CustomDocumentProperties(this);
            // add a section thats assigned to top of the document
            WordSection wordSection = new WordSection(this, null, null);

            var list = this._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.ToList(); //.OfType<Paragraph>().ToList();
            foreach (var element in list) {
                if (element is Paragraph) {
                    Paragraph paragraph = (Paragraph)element;
                    WordParagraph wordParagraph = new WordParagraph(this, paragraph, wordSection);
                    if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                        // find sections added via section page breaks
                        //var sectionType = paragraph.ParagraphProperties.SectionProperties.ChildElements.OfType<SectionType>().FirstOrDefault();
                        //if (sectionType != null) {
                        //    if (sectionType.Val == SectionMarkValues.NextPage) {
                        //        wordSection = new WordSection(this, paragraph);
                        //    }
                        //} else {
                        //    wordSection.Paragraphs.Add(wordParagraph);
                        //}
                        wordSection = new WordSection(this, paragraph.ParagraphProperties.SectionProperties, paragraph);
                    }
                } else if (element is Table) {
                    WordTable wordTable = new WordTable(this, wordSection, (Table)element);
                } else if (element is SectionProperties sectionProperties) {
                    // we don't do anything as we already created it above - i think
                } else if (element is SdtBlock sdtBlock) {
                    // we don't do anything as we load stuff with get on demand
                } else if (element is OpenXmlUnknownElement) {
                    // this happens when adding dirty element - mainly during TOC Update() function
                } else {
                    throw new NotImplementedException("This isn't implemented yet");
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

            WordDocument word = new WordDocument();

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            // this seems to solve an issue where custom properties wouldn't want to save when opening file
            // no problem with creating empty
            FileMode fileMode = readOnly ? FileMode.Open : FileMode.OpenOrCreate;
            var package = Package.Open(filePath, fileMode);
            //WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, readOnly, openSettings);
            WordprocessingDocument wordDocument = WordprocessingDocument.Open(package, openSettings);
            word.FilePath = filePath;
            word._wordprocessingDocument = wordDocument;
            word._document = wordDocument.MainDocumentPart.Document;
            word.LoadDocument();
            return word;
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
        private void SaveSections() {
            WordSection temporarySection = null;
            if (this.Sections.Count > 0) {
                for (int i = 0; i < Sections.Count; i++) {
                    if (temporarySection != null) {

                    } else {
                        temporarySection = Sections[i];
                        Sections[i]._sectionProperties.Remove();
                    }
                }
            }
        }

        private void SaveNumbering() {
            // it seems the order of numbering instance/abstractnums in numbering matters...
            List<AbstractNum> listAbstractNum = new List<AbstractNum>();
            List<NumberingInstance> listNumberingInstance = new List<NumberingInstance>();

            if (_wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart != null) {
                var tempAbstractNumList = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<AbstractNum>();
                foreach (AbstractNum abstractNum in tempAbstractNumList) {
                    listAbstractNum.Add(abstractNum);
                }

                var tempNumberingInstance = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<NumberingInstance>();
                foreach (NumberingInstance numberingInstance in tempNumberingInstance) {
                    listNumberingInstance.Add(numberingInstance);
                }

                if (_wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart != null) {
                    Numbering numbering = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    if (numbering != null) {
                        numbering.RemoveAllChildren();
                    }
                }

                //var tempAbstractNumList = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<AbstractNum>();
                //foreach (AbstractNum abstractNum in tempAbstractNumList) {
                //    abstractNum.Remove();
                //}
                //var tempNumberingInstance = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<NumberingInstance>();
                //foreach (NumberingInstance numberingInstance in tempNumberingInstance) {
                //    numberingInstance.Remove();
                //}
                //tempNumberingInstance = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<NumberingInstance>();

                foreach (AbstractNum abstractNum in listAbstractNum) {
                    _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Append(abstractNum);
                }

                foreach (NumberingInstance numberingInstance in listNumberingInstance) {
                    _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Append(numberingInstance);
                }
            }
        }

        public void Save(string filePath, bool openWord) {
            MoveSectionProperties();
            SaveNumbering();
            WordCustomProperties wordCustomProperties = new WordCustomProperties(this, true);

            if (this._wordprocessingDocument != null) {
                try {
                    if (filePath != "") {
                        // doesn't work correctly with packages
                        this._wordprocessingDocument.SaveAs(filePath);
                    } else {
                        this._wordprocessingDocument.Save();
                    }
                } catch {
                    throw;
                }
                finally {
                    this._wordprocessingDocument.Close();
                    this._wordprocessingDocument.Dispose();
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            this.Open(filePath, openWord);
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

        public WordParagraph AddParagraph(WordParagraph wordParagraph = null) {
            if (wordParagraph == null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this);
            } else {
                // since we created paragraph without adding it to document, we now need to add it to document
                //this.Paragraphs.Add(wordParagraph);
            }

            //this._currentSection.Paragraphs.Add(wordParagraph);
            // wordParagraph._section = this._currentSection;
            this._wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(wordParagraph._paragraph);
            return wordParagraph;
        }

        public WordParagraph AddParagraph(string text) {
            return AddParagraph().SetText(text);
        }

        public WordParagraph AddPageBreak() {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = BreakValues.Page }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            this._document.Body.Append(newWordParagraph._paragraph);

            this._currentSection.PageBreaks.Add(newWordParagraph);
            this._currentSection.Paragraphs.Add(newWordParagraph);
            //this.PageBreaks.Add(newWordParagraph); 
            //this.Paragraphs.Add(newWordParagraph);
            return newWordParagraph;
        }

        public WordParagraph AddBreak(BreakValues breakType = BreakValues.Page) {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = breakType }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            this._document.Body.Append(newWordParagraph._paragraph);
            this.Paragraphs.Add(newWordParagraph);
            return newWordParagraph;
        }

        public void Dispose() {
            if (this._wordprocessingDocument != null) {
                //this._wordprocessingDocument.Close();
                this._wordprocessingDocument.Dispose();
            }
        }

        public WordSection AddSection(SectionMarkValues? sectionMark = null) {
            //Paragraph paragraph = new Paragraph() { RsidParagraphAddition = "fff0", RsidRunAdditionDefault = "fff0"};
            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();

            SectionProperties sectionProperties = new SectionProperties();
            // SectionProperties sectionProperties = new SectionProperties() { RsidR = "fff0"  };

            if (sectionMark != null) {
                SectionType sectionType = new SectionType() { Val = sectionMark };
                sectionProperties.Append(sectionType);
            }

            paragraphProperties.Append(sectionProperties);
            paragraph.Append(paragraphProperties);


            this._document.MainDocumentPart.Document.Body.Append(paragraph);


            WordSection wordSection = new WordSection(this, paragraph);

            return wordSection;
        }

        public WordSection _currentSection { get; set; }
        public WordBackground Background { get; set; }

        public bool ValidateDocument() {
            bool foundIssue = false;
            try {
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in validator.Validate(this._wordprocessingDocument)) {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("ErrorType: " + error.ErrorType);
                    Console.WriteLine("Node: " + error.Node);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                    foundIssue = true;
                }

                Console.WriteLine("count={0}", count);
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }

            return foundIssue;
        }

        //public WordList AddList(CustomListStyles style) {
        //    WordList wordList = new WordList(this, this._currentSection);
        //    wordList.AddList(style, "o", 0);
        //    return wordList;
        //}

        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this, this._currentSection);
            wordList.AddList(style);
            return wordList;
        }

        public WordList AddTableOfContentList(WordListStyle style) {
            WordList wordList = new WordList(this, this._currentSection, true);
            wordList.AddList(style);
            return wordList;
        }

        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this, this._currentSection, rows, columns, tableStyle);
            return wordTable;
        }

        public WordTableOfContent AddTableOfContent(TableOfContentStyle tableOfContentStyle = TableOfContentStyle.Template1) {
            WordTableOfContent wordTableContent = new WordTableOfContent(this, tableOfContentStyle);
            return wordTableContent;
        }
    }
}