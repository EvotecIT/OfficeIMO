using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public partial class WordDocument : IDisposable {
        public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        public List<WordParagraph> PageBreaks = new List<WordParagraph>();
        public List<WordImage> Images = new List<WordImage>();
        public readonly List<WordSection> Sections = new List<WordSection>();

        public string FilePath { get; set; }

        public WordSettings Settings;

        public bool AutoSave {
            get { return _wordprocessingDocument.AutoSave; }
        }

        internal WordprocessingDocument _wordprocessingDocument = null;
        public Document _document = null;

        public static WordDocument Create(string filePath = "", bool autoSave = false) {
            WordDocument word = new WordDocument();

            WordprocessingDocumentType documentType = WordprocessingDocumentType.Document;
            try {
                WordprocessingDocument wordDocument;
                if (filePath != "") {
                    wordDocument = WordprocessingDocument.Create(filePath, documentType, autoSave);
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

                WordSettings wordSettings = new WordSettings(word);
                WordSection wordSection = new WordSection(word);

                return word;
            } catch {
                return word;
            }
        }

        private List<WordParagraph> LoadDocument() {
            // add settings if not existing
            new WordSettings(this);
            // add a section thats assigned to top of the document
            WordSection wordSection = new WordSection(this);


            var list = this._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>().ToList();
            foreach (Paragraph paragraph in list) {
                WordParagraph wordParagraph = new WordParagraph(this, paragraph);
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
                    wordSection = new WordSection(this, paragraph);
                    Debug.WriteLine(wordSection.Paragraphs.Count);
                } else {
                    wordSection.Paragraphs.Add(wordParagraph);
                }

                Debug.WriteLine(wordSection.Paragraphs.Count);
            }

            return this.Paragraphs;
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

            WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, readOnly, openSettings);
            word.FilePath = filePath;
            word._wordprocessingDocument = wordDocument;
            word._document = wordDocument.MainDocumentPart.Document;

            word.LoadDocument();
            //word.GetImages();
            return word;
        }

        public void Save(string filePath, bool openWord) {
            MoveSectionProperties();
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
                } finally {
                    this._wordprocessingDocument.Close();
                    this._wordprocessingDocument.Dispose();
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }

            if (openWord) {
                if (filePath == "") {
                    filePath = this.FilePath;
                }

                ProcessStartInfo startInfo = new ProcessStartInfo(filePath) {
                    UseShellExecute = true
                };
                Process.Start(startInfo);
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

        /// <summary>
        /// This moves section within body from top to bottom to allow footers/headers to move
        /// Needs more work
        /// </summary>
        public void MoveSectionProperties() {
            var body = this._wordprocessingDocument.MainDocumentPart.Document.Body;
            var sectionProperties = this._wordprocessingDocument.MainDocumentPart.Document.Body.Elements<SectionProperties>().Last();
            body.RemoveChild(sectionProperties);
            body.Append(sectionProperties);
        }

        public WordParagraph InsertParagraph(WordParagraph wordParagraph = null) {
            if (wordParagraph == null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this);
            } else {
                // since we created paragraph without adding it to document, we now need to add it to document
                this.Paragraphs.Add(wordParagraph);
            }

            this._currentSection.Paragraphs.Add(wordParagraph);
            wordParagraph._section = this._currentSection;
            this._wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(wordParagraph._paragraph);
            return wordParagraph;
        }

        public WordParagraph InsertParagraph(string text) {
            return InsertParagraph().InsertText(text);
        }

        public WordParagraph InsertPageBreak() {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() {Type = BreakValues.Page}),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            this._document.Body.Append(newWordParagraph._paragraph);
            this.PageBreaks.Add(newWordParagraph);
            this.Paragraphs.Add(newWordParagraph);
            return newWordParagraph;
        }

        public WordParagraph InsertBreak(BreakValues breakType = BreakValues.Page) {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() {Type = breakType}),
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

        public WordSection InsertSection(SectionMarkValues? sectionMark = null) {
            //Paragraph paragraph = new Paragraph() { RsidParagraphAddition = "fff0", RsidRunAdditionDefault = "fff0"};
            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();

            SectionProperties sectionProperties = new SectionProperties();
            // SectionProperties sectionProperties = new SectionProperties() { RsidR = "fff0"  };

            if (sectionMark != null) {
                SectionType sectionType = new SectionType() {Val = sectionMark};
                sectionProperties.Append(sectionType);
            }

            paragraphProperties.Append(sectionProperties);
            paragraph.Append(paragraphProperties);


            this._document.MainDocumentPart.Document.Body.Append(paragraph);


            WordSection wordSection = new WordSection(this, paragraph);

            return wordSection;
        }

        public WordSection _currentSection { get; set; }
    }
}