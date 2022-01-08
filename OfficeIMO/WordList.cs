using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordList {
        private WordprocessingDocument _wordprocessingDocument;
        private WordDocument _document;
        private WordSection _section;
        internal int _numberId;
        private bool _continueNumbering;

        public List<WordParagraph> ListItems = new List<WordParagraph>();



        public WordList() {
            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference() {Val = 0},
                        new NumberingId() {Val = 1})),
                new Run(
                    new RunProperties(),
                    new Text("Hello, ") {Space = SpaceProcessingModeValues.Preserve}
                )
            );
            //paragraph = new Paragraph(
            //    new ParagraphProperties(
            //        new NumberingProperties(
            //            new NumberingLevelReference() {Val = 0},
            //            new NumberingId() {Val = 1})),
            //    new Run(
            //        new RunProperties(),
            //        new Text("world!") {Space = SpaceProcessingModeValues.Preserve}));
        }

        public WordList(WordDocument wordDocument, WordSection section) {
            _document = wordDocument;
            _wordprocessingDocument = wordDocument._wordprocessingDocument;
            _section = section;
            //_continueNumbering = continueNumbering;
            section.Lists.Add(this);

            _document._listNumbers++;
        }

        public WordList(WordDocument wordDocument, WordSection section, int numberId) {
            _document = wordDocument;
            _wordprocessingDocument = wordDocument._wordprocessingDocument;
            _section = section;
            _numberId = numberId;
            //_continueNumbering = continueNumbering;
            section.Lists.Add(this);

            _document._listNumbers++;
            _document._listNumbersUsed.Add(numberId);
        }

        public void BuiltinStyle(ListStyles style, ref AbstractNum abstractNum, ref NumberingInstance numberingInstance) {
            abstractNum = ListStyle.GetStyle(style);
            AbstractNumId abstractNumId = new AbstractNumId();
            abstractNumId.Val = abstractNum.AbstractNumberId;
            numberingInstance = new NumberingInstance(abstractNumId);
            numberingInstance.NumberID = _numberId - 1;
            
            //Numbering numbering = new Numbering(abstractNum, numberingInstance);
            //return numbering;
        }

        public void CreateNumberingDefinition(WordDocument document) {
            NumberingDefinitionsPart numberingDefinitionsPart = document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingDefinitionsPart == null) {
                numberingDefinitionsPart = _wordprocessingDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            }
            
            //_numberingDefinitionsPart = numberingDefinitionsPart;
        }

        public void AddList(ListStyles style) {
            CreateNumberingDefinition(_document);
            
            _numberId = _document._listNumbers;
            
            Numbering numbering = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;
            if (numbering == null) {
                numbering = new Numbering();
                numbering.Save(_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);
            }

            AbstractNum abstractNum = ListStyle.GetStyle(style);
            AbstractNumId abstractNumId = new AbstractNumId();
            abstractNumId.Val = abstractNum.AbstractNumberId;
            NumberingInstance numberingInstance = new NumberingInstance(abstractNumId);
            numberingInstance.NumberID = _numberId;

            Debug.WriteLine("NumberID " + numberingInstance.NumberID);
            //BuiltinStyle(style, ref abstractNum, ref numberingInstance);



            _document._ListAbstractNum.Add(abstractNum);
            _document._listNumberingInstances.Add(numberingInstance);
            //numbering.Append(numberingInstance, abstractNum);
        }

        public void AddList(CustomListStyles style = CustomListStyles.Bullet, string levelText = "·", int levelIndex = 0) {
            CreateNumberingDefinition(_document);
            if (_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering == null) {
                Numbering numbering = new Numbering();
                numbering.Save(_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);
            }

            // we take current list number from the document
            _numberId = _document._listNumbers;

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

            AbstractNum abstractNum = new AbstractNum(level,level1);
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
            Text textProperty = new Text() {Space = SpaceProcessingModeValues.Preserve};
            RunProperties runProperties = new RunProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };
            ParagraphProperties paragraphProperties = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference() {Val = level},
                    new NumberingId() {Val =  this._numberId }
                    ));
            paragraphProperties.Append(paragraphStyleId);

            Run run = new Run();

            WordParagraph wordParagraph = new WordParagraph(_document, true, paragraphProperties, runProperties, run);
            wordParagraph.Text = text;

            _document.InsertParagraph(wordParagraph);

            ListItems.Add(wordParagraph);

            // we add internal tracking in the paragraph for a list
            wordParagraph._list = this;

            return wordParagraph;
        }
    }
}