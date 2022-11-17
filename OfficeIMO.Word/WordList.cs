using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordList {
        private readonly WordprocessingDocument _wordprocessingDocument;
        private readonly WordDocument _document;
        // private readonly WordSection _section;
        private int _abstractId;
        internal int _numberId;

        /// <summary>
        /// This provides a way to set items to be treated with heading style
        /// </summary>
        private readonly bool _isToc;

        /// <summary>
        /// This provides a way to set it teams to be treated with heading style during load 
        /// </summary>
        public bool IsToc {
            get {
                foreach (var paragraph in ListItems) {
                    var style = paragraph.Style.ToString();
                    if (style.Remove(style.Length - 1) == "Heading") {
                        return true;
                    }
                }
                return false;
            }
        }

        public List<WordParagraph> ListItems {
            get {
                var listItems = new List<WordParagraph>();
                foreach (WordParagraph paragraph in this._document.Paragraphs) {
                    if (paragraph.IsListItem) {
                        if (_numberId == paragraph._listNumberId) {
                            listItems.Add(paragraph);
                        }
                    }
                }
                return listItems;
            }
        }

        public WordList(WordDocument wordDocument, WordSection section, bool isToc = false) {
            _document = wordDocument;
            _wordprocessingDocument = wordDocument._wordprocessingDocument;
            //_section = section;
            _isToc = isToc;
            // section.Lists.Add(this);
        }

        public WordList(WordDocument wordDocument, WordSection section, int numberId) {
            _document = wordDocument;
            _wordprocessingDocument = wordDocument._wordprocessingDocument;
            //  _section = section;
            _numberId = numberId;
        }

        private void CreateNumberingDefinition(WordDocument document) {
            NumberingDefinitionsPart numberingDefinitionsPart = document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingDefinitionsPart == null) {
                numberingDefinitionsPart = _wordprocessingDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            }

            Numbering numbering = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;
            if (numbering == null) {
                numbering = new Numbering();
                numbering.Save(_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);
            }
        }

        internal static int GetNextAbstractNum(Numbering numbering) {
            var ids = new List<int>();
            foreach (var element in numbering.ChildElements.OfType<AbstractNum>()) {
                ids.Add(element.AbstractNumberId);
            }
            if (ids.Count > 0) {
                return ids.Max() + 1;
            } else {
                return 1;
            }
        }

        internal static int GetNextNumberingInstance(Numbering numbering) {
            var ids = new List<int>();
            foreach (var element in numbering.ChildElements.OfType<NumberingInstance>()) {
                ids.Add(element.NumberID);
            }

            if (ids.Count > 0) {
                return ids.Max() + 1;
            } else {
                return 1;
            }
        }

        internal void AddList(WordListStyle style) {
            CreateNumberingDefinition(_document);
            var numbering = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;

            _abstractId = GetNextAbstractNum(numbering);
            _numberId = GetNextNumberingInstance(numbering);

            AbstractNum abstractNum = WordListStyles.GetStyle(style);
            abstractNum.AbstractNumberId = _abstractId;
            AbstractNumId abstractNumId = new AbstractNumId();
            abstractNumId.Val = _abstractId;
            NumberingInstance numberingInstance = new NumberingInstance(abstractNumId);
            numberingInstance.NumberID = _numberId;
            numbering.Append(numberingInstance, abstractNum);
        }

        internal void AddList(CustomListStyles style = CustomListStyles.Bullet, string levelText = "·", int levelIndex = 0) {
            // TODO this isn't working yet, needs implementation
            CreateNumberingDefinition(_document);
            if (_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering == null) {
                Numbering numbering = new Numbering();
                numbering.Save(_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);
            }

            // we take current list number from the document
            //_numberId = _document._listNumbers;

            var numberingFormatValues = CustomListStyle.GetStyle(style);

            Level level = new Level(
                new NumberingFormat() { Val = numberingFormatValues },
                new LevelText() { Val = levelText }
            );
            level.LevelIndex = 1;

            Level level1 = new Level(
                new NumberingFormat() { Val = numberingFormatValues },
                new LevelText() { Val = levelText }
            );
            level1.LevelIndex = 2;

            AbstractNum abstractNum = new AbstractNum(level, level1);
            abstractNum.AbstractNumberId = 0;
            //abstractNum.Nsid = new Nsid();

            AbstractNumId abstractNumId = new AbstractNumId();
            abstractNumId.Val = 0;

            NumberingInstance numberingInstance = new NumberingInstance(abstractNumId);
            numberingInstance.NumberID = _numberId;


            //LevelOverride levelOverride = new LevelOverride();
            //levelOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue();
            //levelOverride.StartOverrideNumberingValue.Val = 1;
            //numberingInstance.Append(levelOverride);


            _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Append(abstractNum);
            _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Append(numberingInstance);
        }

        public WordParagraph AddItem(string text, int level = 0) {
            Text textProperty = new Text() { Space = SpaceProcessingModeValues.Preserve };
            RunProperties runProperties = new RunProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };
            NumberingProperties numberingProperties = new NumberingProperties(
                new NumberingLevelReference() { Val = level },
                new NumberingId() { Val = this._numberId }
            );
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(numberingProperties);

            Run run = new Run();
            run.Append(runProperties);
            run.Append(textProperty);

            Paragraph paragraph = new Paragraph();
            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            this._wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(paragraph);

            WordParagraph wordParagraph = new WordParagraph(_document, paragraph, run);
            wordParagraph.Text = text;
            // this simplifies TOC for user usage
            if (_isToc || IsToc) {
                wordParagraph.Style = WordParagraphStyle.GetStyle(level);
            }

            return wordParagraph;
        }
    }
}