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
    /// This provides a way to set items to be treated with heading style during load
    /// </summary>
    public bool IsToc {
        get {
            return ListItems
                .Select(paragraph => paragraph.Style.ToString())
                .Any(style => style.StartsWith("Heading", StringComparison.Ordinal));
        }
    }

    private string NsidId {
        get {
            if (AbstractNum == null) {
                return null;
            }

            return AbstractNum.Nsid.Val;

        }
        set {
            if (AbstractNum != null) {
                AbstractNum.Nsid.Val = value;
            }
        }
    }

    private string GenerateNsidId() {
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.nsid?view=openxml-2.8.1
        // Specifies a number value specified as a four digit hexadecimal number),
        // whose contents of this decimal number are interpreted based on the context of the parent XML element.
        // for example FFFFFF89 or D9842532
        return Guid.NewGuid().ToString().ToUpper().Substring(0, 8);

    }

    private AbstractNum AbstractNum {
        get {
            var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            var abstractNumList = numbering.ChildElements.OfType<AbstractNum>();
            foreach (AbstractNum abstractNum in abstractNumList) {
                if (abstractNum.AbstractNumberId == _abstractId) {
                    return abstractNum;
                }
            }

            return null;
        }
    }

    public List<WordParagraph> ListItems {
        get {
            return _document.Paragraphs
                .Where(paragraph => paragraph.IsListItem && paragraph._listNumberId == _numberId)
                .ToList();
        }
    }

    public bool RestartNumbering { get; set; }

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

        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
        paragraphProperties.Append(
            new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = _numberId }
            ));
        paragraph.Append(paragraphProperties);
        paragraph.Append(run);
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

    private static int GetNextAbstractNum(Numbering numbering) {
        var ids = numbering.ChildElements
            .OfType<AbstractNum>()
            .Select(element => (int)element.AbstractNumberId)
            .ToList();
        return ids.Count > 0 ? ids.Max() + 1 : 0;
    }

    private static int GetNextNumberingInstance(Numbering numbering) {
        var ids = numbering.ChildElements
            .OfType<NumberingInstance>()
            .Select(element => (int)element.NumberID)
            .ToList();
        return ids.Count > 0 ? ids.Max() + 1 : 1;
    }

    internal void AddList(WordListStyle style, bool continueNumbering) {
        CreateNumberingDefinition(_document);
        var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;

        _abstractId = GetNextAbstractNum(numbering);
        _numberId = GetNextNumberingInstance(numbering);

        var abstractNum = WordListStyles.GetStyle(style);
        abstractNum.AbstractNumberId = _abstractId;
        var abstractNumId = new AbstractNumId {
            Val = _abstractId
        };
        NumberingInstance numberingInstance = new NumberingInstance();
        if (continueNumbering) {
            numberingInstance = DefaultNumberingInstance(abstractNumId, _numberId);
        } else {
            numberingInstance = RestartNumberingInstance(abstractNumId, _numberId);
        }

        numbering.Append(numberingInstance, abstractNum);
    }

    // TODO this isn't working yet, needs implementation
    //internal void AddList(CustomListStyles style = CustomListStyles.Bullet, string levelText = "·", int levelIndex = 0) {
    //    CreateNumberingDefinition(_document);

    //    // we take current list number from the document
    //    //_numberId = _document._listNumbers;

    //    var numberingFormatValues = CustomListStyle.GetStyle(style);

    //    var level = new Level(
    //        new NumberingFormat { Val = numberingFormatValues },
    //        new LevelText { Val = levelText }
    //    ) {
    //        LevelIndex = 1
    //    };
    //    var level1 = new Level(
    //        new NumberingFormat { Val = numberingFormatValues },
    //        new LevelText { Val = levelText }
    //    ) {
    //        LevelIndex = 2
    //    };
    //    var abstractNum = new AbstractNum(level, level1) {
    //        AbstractNumberId = 0
    //    };
    //    //abstractNum.Nsid = new Nsid();

    //    var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
    //    numbering.Append(abstractNum);

    //    var abstractNumId = new AbstractNumId {
    //        Val = 0
    //    };
    //    var numberingInstance = new NumberingInstance(abstractNumId) {
    //        NumberID = _numberId
    //    };

    //    //LevelOverride levelOverride = new LevelOverride();
    //    //levelOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue();
    //    //levelOverride.StartOverrideNumberingValue.Val = 1;
    //    //numberingInstance.Append(levelOverride);

    //    numbering.Append(numberingInstance);
    //}

    private void CreateNumberingDefinition(WordDocument document) {
        var numberingDefinitionsPart =
            document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart
            ?? _wordprocessingDocument.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();

        if (numberingDefinitionsPart.Numbering == null) {
            numberingDefinitionsPart.Numbering = new Numbering();
            numberingDefinitionsPart.Numbering.Save(_document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart);
        }
    }

    private NumberingInstance DefaultNumberingInstance(AbstractNumId abstractNumId, int numberId) {
        var numberingInstance = new NumberingInstance(abstractNumId) { NumberID = numberId };
        return numberingInstance;
    }

    private NumberingInstance RestartNumberingInstance(AbstractNumId abstractNumId, int numberId) {
        NumberingInstance numberingInstance1 = new NumberingInstance(abstractNumId) { NumberID = numberId };

        LevelOverride levelOverride1 = new LevelOverride() { LevelIndex = 0 };
        StartOverrideNumberingValue startOverrideNumberingValue1 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride1.Append(startOverrideNumberingValue1);

        LevelOverride levelOverride2 = new LevelOverride() { LevelIndex = 1 };
        StartOverrideNumberingValue startOverrideNumberingValue2 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride2.Append(startOverrideNumberingValue2);

        LevelOverride levelOverride3 = new LevelOverride() { LevelIndex = 2 };
        StartOverrideNumberingValue startOverrideNumberingValue3 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride3.Append(startOverrideNumberingValue3);

        LevelOverride levelOverride4 = new LevelOverride() { LevelIndex = 3 };
        StartOverrideNumberingValue startOverrideNumberingValue4 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride4.Append(startOverrideNumberingValue4);

        LevelOverride levelOverride5 = new LevelOverride() { LevelIndex = 4 };
        StartOverrideNumberingValue startOverrideNumberingValue5 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride5.Append(startOverrideNumberingValue5);

        LevelOverride levelOverride6 = new LevelOverride() { LevelIndex = 5 };
        StartOverrideNumberingValue startOverrideNumberingValue6 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride6.Append(startOverrideNumberingValue6);

        LevelOverride levelOverride7 = new LevelOverride() { LevelIndex = 6 };
        StartOverrideNumberingValue startOverrideNumberingValue7 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride7.Append(startOverrideNumberingValue7);

        LevelOverride levelOverride8 = new LevelOverride() { LevelIndex = 7 };
        StartOverrideNumberingValue startOverrideNumberingValue8 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride8.Append(startOverrideNumberingValue8);

        LevelOverride levelOverride9 = new LevelOverride() { LevelIndex = 8 };
        StartOverrideNumberingValue startOverrideNumberingValue9 = new StartOverrideNumberingValue() { Val = 1 };

        levelOverride9.Append(startOverrideNumberingValue9);

        numberingInstance1.Append(levelOverride1);
        numberingInstance1.Append(levelOverride2);
        numberingInstance1.Append(levelOverride3);
        numberingInstance1.Append(levelOverride4);
        numberingInstance1.Append(levelOverride5);
        numberingInstance1.Append(levelOverride6);
        numberingInstance1.Append(levelOverride7);
        numberingInstance1.Append(levelOverride8);
        numberingInstance1.Append(levelOverride9);
        return numberingInstance1;
    }

}
