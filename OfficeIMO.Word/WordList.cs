using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

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
            return ListItems
                .Select(paragraph => paragraph.Style.ToString())
                .Any(style => style.StartsWith("Heading", StringComparison.Ordinal));
        }
    }

    public List<WordParagraph> ListItems {
        get {
            return _document.Paragraphs
                .Where(paragraph => paragraph.IsListItem && paragraph._listNumberId == _numberId)
                .ToList();
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

    public WordParagraph AddItem(string text, int level = 0) {
        var paragraph = new Paragraph();

        var run = new Run();
        run.Append(new RunProperties());
        run.Append(new Text { Space = SpaceProcessingModeValues.Preserve });
        paragraph.Append(run);

        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
        paragraphProperties.Append(
            new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = _numberId }
            ));
        paragraph.Append(paragraphProperties);

        _wordprocessingDocument.MainDocumentPart!.Document.Body!.AppendChild(paragraph);

        var wordParagraph = new WordParagraph(_document, paragraph, run) {
            Text = text
        };

        // this simplifies TOC for user usage
        if (_isToc || IsToc) {
            wordParagraph.Style = WordParagraphStyle.GetStyle(level);
        }

        return wordParagraph;
    }

    internal static int GetNextAbstractNum(Numbering numbering) {
        var ids = numbering.ChildElements
            .OfType<AbstractNum>()
            .Select(element => (int) element.AbstractNumberId)
            .ToList();
        return ids.Count > 0 ? ids.Max() + 1 : 1;
    }

    internal static int GetNextNumberingInstance(Numbering numbering) {
        var ids = numbering.ChildElements
            .OfType<NumberingInstance>()
            .Select(element => (int) element.NumberID)
            .ToList();
        return ids.Count > 0 ? ids.Max() + 1 : 1;
    }

    internal void AddList(WordListStyle style) {
        CreateNumberingDefinition(_document);
        var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;

        _abstractId = GetNextAbstractNum(numbering);
        _numberId = GetNextNumberingInstance(numbering);

        var abstractNum = WordListStyles.GetStyle(style);
        abstractNum.AbstractNumberId = _abstractId;
        var abstractNumId = new AbstractNumId {
            Val = _abstractId
        };
        var numberingInstance = new NumberingInstance(abstractNumId) {
            NumberID = _numberId
        };
        numbering.Append(numberingInstance, abstractNum);
    }

    // TODO this isn't working yet, needs implementation
    internal void AddList(CustomListStyles style = CustomListStyles.Bullet, string levelText = "·", int levelIndex = 0) {
        CreateNumberingDefinition(_document);

        // we take current list number from the document
        //_numberId = _document._listNumbers;

        var numberingFormatValues = CustomListStyle.GetStyle(style);

        var level = new Level(
            new NumberingFormat { Val = numberingFormatValues },
            new LevelText { Val = levelText }
        ) {
            LevelIndex = 1
        };
        var level1 = new Level(
            new NumberingFormat { Val = numberingFormatValues },
            new LevelText { Val = levelText }
        ) {
            LevelIndex = 2
        };
        var abstractNum = new AbstractNum(level, level1) {
            AbstractNumberId = 0
        };
        //abstractNum.Nsid = new Nsid();

        var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
        numbering.Append(abstractNum);

        var abstractNumId = new AbstractNumId {
            Val = 0
        };
        var numberingInstance = new NumberingInstance(abstractNumId) {
            NumberID = _numberId
        };

        //LevelOverride levelOverride = new LevelOverride();
        //levelOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue();
        //levelOverride.StartOverrideNumberingValue.Val = 1;
        //numberingInstance.Append(levelOverride);

        numbering.Append(numberingInstance);
    }

    private void CreateNumberingDefinition(WordDocument document) {
        var numberingDefinitionsPart =
            document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart
            ?? _wordprocessingDocument.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();

        if (numberingDefinitionsPart.Numbering == null) {
            numberingDefinitionsPart.Numbering = new Numbering();
            numberingDefinitionsPart.Numbering.Save(_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);
        }
    }
}
