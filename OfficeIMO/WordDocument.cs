using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordDocument : IDisposable {
        public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        public List<WordImage> Images = new List<WordImage>();

        public string filePath = null;
        public bool AutoSave
        {
            get {
                return _wordprocessingDocument.AutoSave;
            }
            set
            {
               //_wordprocessingDocument = value;
                
            }
        }

        internal WordprocessingDocument _wordprocessingDocument = null;
        internal Document _document = null;

        //private MemoryStream _memory = null;

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

                wordDocument.AddMainDocumentPart();
                wordDocument.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                wordDocument.MainDocumentPart.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body();
                OfficeIMO.Word.WordDocument.AddDefaultStyleDefinitions(wordDocument, null);

                word.filePath = filePath;
                word._wordprocessingDocument = wordDocument;
                word._document = wordDocument.MainDocumentPart.Document;

                return word;
            } catch {
                return word;
            }
        }

        internal List<WordParagraph> GetParagraphs() {
            //var listWord = new List<WordParagraph>();
            var list = this._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>().ToList();
            foreach (Paragraph paragraph in list)
            {
                WordParagraph wordParagraph = new WordParagraph(this, paragraph);
            }

            return this.Paragraphs;
            //return listWord;
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
            word.filePath = filePath;
            word._wordprocessingDocument = wordDocument;
            word._document = wordDocument.MainDocumentPart.Document;

            word.GetParagraphs();
            
            return word;
        }

        public void Save(string filePath = "", bool openWord = false) {
            if (this._wordprocessingDocument != null) {
                try {
                    if (filePath != "") {
                        this._wordprocessingDocument.SaveAs(filePath);
                    } else {
                        this._wordprocessingDocument.Save();
                    }
                } catch {
                    throw;
                } finally {
                    //this._memory.Close();
                    this._wordprocessingDocument.Close();
                }
            } else {
                throw new InvalidOperationException("Document couldn't be saved as WordDocument wasn't provided.");
            }


            //this._document = null;
            this._wordprocessingDocument.Dispose();
            //this._wordprocessingDocument = null;


            // TODO this needs fixing because Examples are showing that if Example2 runs too long it won't open up on 1st example
            if (openWord) {
                if (filePath == "") {
                    filePath = this.filePath;
                }
                ProcessStartInfo startInfo = new ProcessStartInfo(filePath) {
                    UseShellExecute = true
                };
                Process.Start(startInfo);
            }
        }

        public WordParagraph InsertParagraph(WordParagraph wordParagraph = null) {
            if (wordParagraph == null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this);
            } else {
                // since we created paragraph without adding it to document, we now need to add it to document
                this.Paragraphs.Add(wordParagraph);
            }
            this._wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(wordParagraph._paragraph);
            return wordParagraph;
        }

        public WordParagraph InsertParagraph(string text) {
            return InsertParagraph().InsertText(text);
        }

        public void Dispose() {
            if (this._wordprocessingDocument != null) {
                this._wordprocessingDocument.Close();
                this._wordprocessingDocument.Dispose();
            }
        }
    }
}