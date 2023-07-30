using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp;

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

    private WordParagraph _wordParagraph;
    private readonly WordHeaderFooter _headerFooter;

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

    //private string NsidId {
    //    get {
    //        if (AbstractNum == null) {
    //            return null;
    //        }

    //        return AbstractNum.Nsid.Val;

    //    }
    //    set {
    //        if (AbstractNum != null) {
    //            AbstractNum.Nsid.Val = value;
    //        }
    //    }
    //}

    //private string GenerateNsidId() {
    //    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.nsid?view=openxml-2.8.1
    //    // Specifies a number value specified as a four digit hexadecimal number),
    //    // whose contents of this decimal number are interpreted based on the context of the parent XML element.
    //    // for example FFFFFF89 or D9842532
    //    return Guid.NewGuid().ToString().ToUpper().Substring(0, 8);

    //}

    //private AbstractNum AbstractNum {
    //    get {
    //        var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
    //        var abstractNumList = numbering.ChildElements.OfType<AbstractNum>();
    //        foreach (AbstractNum abstractNum in abstractNumList) {
    //            if (abstractNum.AbstractNumberId == _abstractId) {
    //                return abstractNum;
    //            }
    //        }

    //        return null;
    //    }
    //}

    public List<WordParagraph> ListItems {
        get {
            List<WordParagraph> list = new List<WordParagraph>();
            foreach (var paragraph in _document.Paragraphs) {
                if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                    list.Add(paragraph);
                }
            }

            foreach (var table in _document.Tables) {
                foreach (var paragraph in table.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
            }

            if (_document.Header.Default != null) {
                foreach (var paragraph in _document.Header.Default.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Header.Default.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Header.Even != null) {
                foreach (var paragraph in _document.Header.Even.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Header.Even.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Header.First != null) {
                foreach (var paragraph in _document.Header.First.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Header.First.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }


            if (_document.Footer.Default != null) {
                foreach (var paragraph in _document.Footer.Default.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Footer.Default.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Footer.Even != null) {
                foreach (var paragraph in _document.Footer.Even.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Footer.Even.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Footer.First != null) {
                foreach (var paragraph in _document.Footer.First.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Footer.First.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }
            return list;


            //if (_wordParagraph != null) {
            //    var list = new List<Paragraph>();
            //    var parent = _wordParagraph._paragraph.Parent;
            //    var elementsAfter = parent.ChildElements.OfType<Paragraph>();
            //    foreach (var element in elementsAfter) {
            //        if (element.ParagraphProperties != null && element.ParagraphProperties.NumberingProperties != null) {
            //            if (element.ParagraphProperties.NumberingProperties.NumberingId.Val == _numberId) {
            //                list.Add(element);
            //            }
            //        }
            //    }
            //    var listWord = WordSection.ConvertParagraphsToWordParagraphs(_document, list);
            //    return listWord;
            //} else {
            //    return new List<WordParagraph>();
            //}
            //elementsAfter.Where(paragraph => paragraph.IsListItem && paragraph._listNumberId == _numberId).ToList();
            //return _document.Paragraphs
            //    .Where(paragraph => paragraph.IsListItem && paragraph._listNumberId == _numberId)
            //    .ToList();
        }
    }

    public bool RestartNumbering {
        get {
            var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            var listNumbering = numbering.ChildElements.OfType<NumberingInstance>();
            foreach (var numberingInstance in listNumbering) {
                if (numberingInstance.NumberID == _numberId) {
                    var level = numberingInstance.ChildElements.OfType<LevelOverride>().FirstOrDefault();
                    if (level != null) {
                        return true;
                    }
                }
            }
            return false;
        }
        set {
            var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            var listNumbering = numbering.ChildElements.OfType<NumberingInstance>();
            foreach (var numberingInstance in listNumbering) {
                if (numberingInstance.NumberID == _numberId) {
                    var abstractNumId = new AbstractNumId {
                        Val = _abstractId
                    };
                    NumberingInstance foundNumberingInstance;
                    if (value == false) {
                        // continue numbering as it was by default
                        foundNumberingInstance = DefaultNumberingInstance(abstractNumId, _numberId);
                    } else {
                        // restart numbering from 1
                        foundNumberingInstance = RestartNumberingInstance(abstractNumId, _numberId);
                    }
                    numberingInstance.InsertBeforeSelf(foundNumberingInstance);
                    numberingInstance.Remove();
                }
            }
        }
    }

    /// <summary>
    /// Restarts numbering of a list after a break. Requires a list to be set to RestartNumbering overall.
    /// </summary>
    public bool RestartNumberingAfterBreak {
        get {
            var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            var listAbstracts = numbering.ChildElements.OfType<AbstractNum>();
            foreach (var abstractInstance in listAbstracts) {
                if (abstractInstance.AbstractNumberId == _abstractId) {
                    var currentValue = abstractInstance.GetAttribute("restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml");
                    return currentValue.Value != "0";
                }
            }
            return false;
        }
        set {
            var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            var listAbstracts = numbering.ChildElements.OfType<AbstractNum>();
            foreach (var abstractInstance in listAbstracts) {
                if (abstractInstance.AbstractNumberId == _abstractId) {
                    var setValue = value ? "1" : "0";
                    abstractInstance.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", setValue));
                }
            }
        }
    }

    public WordList(WordDocument wordDocument, bool isToc = false) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //_section = section;
        _isToc = isToc;
        // section.Lists.Add(this);
    }

    public WordList(WordDocument wordDocument, WordParagraph paragraph, bool isToc = false) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //_section = section;
        _isToc = isToc;
        _wordParagraph = paragraph;
        // section.Lists.Add(this);
    }


    public WordList(WordDocument wordDocument, int numberId) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //  _section = section;
        _numberId = numberId;
    }

    public WordList(WordDocument wordDocument, WordHeaderFooter headerFooter) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        _headerFooter = headerFooter;
    }

    public WordParagraph AddItem(WordParagraph wordParagraph, int level = 0) {
        return AddItem(null, level, wordParagraph);
    }

    public WordParagraph AddItem(string text, int level = 0, WordParagraph wordParagraph = null) {
        if (wordParagraph != null) {
            wordParagraph._paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
            wordParagraph._paragraphProperties.Append(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = _numberId }
                ));
            if (text != null) {
                wordParagraph.Text = text;
            }
        } else {
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

            if (_wordParagraph != null) {

                if (this.ListItems.Count > 0) {
                    var lastItem = this.ListItems.Last();
                    var allElements = lastItem._paragraph.Parent.ChildElements.OfType<Paragraph>();
                    if (allElements.Count() > 0) {
                        var lastParagraph = allElements.Last();
                        lastParagraph.Parent.Append(paragraph);
                    }
                } else {
                    var allElements = _wordParagraph._paragraph.Parent.ChildElements.OfType<Paragraph>();
                    var lastElement = allElements.Last();
                    lastElement.Parent.Append(paragraph);
                }

                // _wordParagraph._paragraph.Append(paragraph);
            } else {
                if (this.ListItems.Count > 0) {
                    var lastItem = this.ListItems.Last();
                    var allElementsAfter = lastItem._paragraph.ElementsAfter();
                    if (allElementsAfter.Count() > 0) {
                        var lastParagraph = allElementsAfter.Last();
                        lastParagraph.InsertAfterSelf(paragraph);
                    } else {
                        lastItem._paragraph.InsertAfterSelf(paragraph);
                    }
                } else {
                    if (_headerFooter != null && _headerFooter._header != null) {
                        _headerFooter._header.Append(paragraph);
                    } else if (_headerFooter != null && _headerFooter._footer != null) {
                        _headerFooter._footer.Append(paragraph);
                    } else {
                        _wordprocessingDocument.MainDocumentPart!.Document.Body!.AppendChild(paragraph);
                    }
                }
            }
            wordParagraph = new WordParagraph(_document, paragraph, run) {
                Text = text
            };
        }

        // this simplifies TOC for user usage
        if (_isToc || IsToc) {
            wordParagraph.Style = WordParagraphStyle.GetStyle(level);
        }

        if (_wordParagraph == null) {
            _wordParagraph = wordParagraph;
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
        var numberingDefinitionsPart = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart ?? _wordprocessingDocument.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
        if (numberingDefinitionsPart.Numbering == null) {
            // the check for null is required even tho Resharper claims it's not
            Numbering numbering1 = new Numbering() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            numbering1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            numbering1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            numbering1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            numbering1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            numbering1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            numbering1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            numbering1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            numbering1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            numbering1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            numbering1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            numbering1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            numbering1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
            numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            numbering1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            numbering1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            numbering1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            numbering1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            numbering1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            numbering1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            numbering1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            numbering1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            numbering1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            numberingDefinitionsPart.Numbering = numbering1;
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
