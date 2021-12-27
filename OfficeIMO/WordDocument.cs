using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordDocument {
        public List<WordParagraph> Paragraphs = new List<WordParagraph>();

        public string filePath = null;


        private WordprocessingDocument _wordprocessingDocument = null;
        private Document _document = null;

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

            this._wordprocessingDocument.Dispose();

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
                wordParagraph = new WordParagraph();
            }
            //var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            //paragraph.AppendChild(wordParagraph._run);
            this._wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(wordParagraph._paragraph);
            return wordParagraph;
        }

        public WordParagraph InsertParagraph(string text) {
            return InsertParagraph().InsertText(text);
        }
    }
}