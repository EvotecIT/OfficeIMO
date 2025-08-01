using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace OfficeIMO.Word;

/// <summary>
/// Contains private methods for list handling.
/// </summary>
public partial class WordList : WordElement {
    /// <summary>
    /// Retrieves the <see cref="AbstractNum"/> associated with this list.
    /// </summary>
    private AbstractNum GetAbstractNum() {
        var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
        if (_abstractId == 0) {
            var instance = numbering.Elements<NumberingInstance>()
                .FirstOrDefault(n => n.NumberID.Value == _numberId);
            if (instance?.AbstractNumId?.Val != null) {
                _abstractId = (int)instance.AbstractNumId.Val.Value;
            }
        }
        return numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId.Value == _abstractId);
    }

    /// <summary>
    /// Gets the next available abstract numbering ID.
    /// </summary>
    /// <param name="numbering">The numbering definitions.</param>
    private static int GetNextAbstractNum(Numbering numbering) {
        var ids = numbering.ChildElements
            .OfType<AbstractNum>()
            .Select(element => (int)element.AbstractNumberId)
            .ToList();
        return ids.Count > 0 ? ids.Max() + 1 : 0;
    }

    /// <summary>
    /// Gets the next available numbering instance ID.
    /// </summary>
    /// <param name="numbering">The numbering definitions.</param>
    private static int GetNextNumberingInstance(Numbering numbering) {
        var ids = numbering.ChildElements
            .OfType<NumberingInstance>()
            .Select(element => (int)element.NumberID)
            .ToList();
        return ids.Count > 0 ? ids.Max() + 1 : 1;
    }

    /// <summary>
    /// Gets a numbering property using a selector function.
    /// </summary>
    /// <typeparam name="T">The type of the property.</typeparam>
    /// <param name="propertySelector">The selector function to extract the property.</param>
    /// <param name="defaultValue">The default value if the property is not found.</param>
    private T GetNumberingProperty<T>(Func<NumberingSymbolRunProperties, T> propertySelector, T defaultValue = default) {
        var abstractNum = GetAbstractNum();
        var level = abstractNum?.Elements<Level>().FirstOrDefault();
        if (level != null) {
            var props = level.NumberingSymbolRunProperties;
            if (props != null) {
                return propertySelector(props);
            }
        }
        return defaultValue;
    }

    /// <summary>
    /// Sets a numbering property using the specified action.
    /// </summary>
    /// <param name="setProperty">The action to set the property.</param>
    /// <param name="applyWhenNull">Indicates whether to apply the property when it is null.</param>
    private void SetNumberingProperty(Action<NumberingSymbolRunProperties> setProperty, bool applyWhenNull = false) {
        var abstractNum = GetAbstractNum();
        if (abstractNum != null) {
            foreach (var level in abstractNum.Elements<Level>()) {
                var props = level.GetFirstChild<NumberingSymbolRunProperties>();
                if (props == null && applyWhenNull) {
                    props = new NumberingSymbolRunProperties();
                    level.Append(props);
                }
                if (props != null) {
                    setProperty(props);
                }
            }
        }
    }

    /// <summary>
    /// Creates the numbering definitions part if it doesn't exist.
    /// </summary>
    /// <param name="document">The Word document.</param>
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

    /// <summary>
    /// Creates a default numbering instance.
    /// </summary>
    /// <param name="abstractNumId">The abstract numbering ID.</param>
    /// <param name="numberId">The numbering instance ID.</param>
    private NumberingInstance DefaultNumberingInstance(AbstractNumId abstractNumId, int numberId) {
        var numberingInstance = new NumberingInstance(abstractNumId) { NumberID = numberId };
        return numberingInstance;
    }

    /// <summary>
    /// Creates a numbering instance that restarts numbering after a break.
    /// </summary>
    /// <param name="abstractNumId">The abstract numbering ID.</param>
    /// <param name="numberId">The numbering instance ID.</param>
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

    /// <summary>
    /// Adds a list to the document with the specified style.
    /// </summary>
    /// <param name="style">The list style to apply.</param>
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
        NumberingInstance numberingInstance = new NumberingInstance();
        numberingInstance = RestartNumberingInstance(abstractNumId, _numberId);
        numbering.Append(numberingInstance, abstractNum);
    }

    private WordList Clone(OpenXmlElement referenceParagraph, bool after) {
        var numberingPart = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
        if (numberingPart == null) {
            numberingPart = _document._wordprocessingDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }
        var numbering = numberingPart.Numbering;

        var originalAbstract = numbering.Elements<AbstractNum>().First(a => a.AbstractNumberId.Value == _abstractId);
        var originalInstance = numbering.Elements<NumberingInstance>().First(n => n.NumberID.Value == _numberId);

        int newAbstractId = GetNextAbstractNum(numbering);
        int newNumberId = GetNextNumberingInstance(numbering);

        var newAbstract = (AbstractNum)originalAbstract.CloneNode(true);
        newAbstract.AbstractNumberId = newAbstractId;

        var restartAttr = originalAbstract.GetAttribute("restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml");
        if (!string.IsNullOrEmpty(restartAttr.Value)) {
            EnsureW15Namespace(numbering);
            newAbstract.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", restartAttr.Value));
        }

        numbering.Append(newAbstract);

        var newInstance = new NumberingInstance { NumberID = newNumberId };
        newInstance.Append(new AbstractNumId { Val = newAbstractId });
        foreach (var levelOverride in originalInstance.Elements<LevelOverride>()) {
            newInstance.Append((LevelOverride)levelOverride.CloneNode(true));
        }
        numbering.Append(newInstance);

        WordList clonedList;
        if (_headerFooter != null) {
            clonedList = new WordList(_document, _headerFooter);
        } else if (_wordParagraph != null) {
            clonedList = new WordList(_document, _wordParagraph, _isToc);
        } else {
            clonedList = new WordList(_document, _isToc);
        }

        clonedList._abstractId = newAbstractId;
        clonedList._numberId = newNumberId;

        OpenXmlElement currentRef = referenceParagraph;
        WordParagraph firstInserted = null;
        foreach (var item in ListItems) {
            var clonedParagraph = (Paragraph)item._paragraph.CloneNode(true);
            var numberingProps = clonedParagraph.GetFirstChild<ParagraphProperties>()?.NumberingProperties;
            if (numberingProps != null) {
                numberingProps.NumberingId.Val = newNumberId;
            }
            currentRef = after ? currentRef.InsertAfterSelf(clonedParagraph) : currentRef.InsertBeforeSelf(clonedParagraph);
            var run = ((Paragraph)currentRef).GetFirstChild<Run>() ?? new Run();
            if (run.Parent == null) ((Paragraph)currentRef).AppendChild(run);
            var wp = new WordParagraph(_document, (Paragraph)currentRef, run);
            wp.Text = item.Text;
            wp._list = clonedList;
            clonedList._listItems.Add(wp);
            if (firstInserted == null) {
                firstInserted = wp;
            }
        }

        if (firstInserted != null) {
            clonedList._wordParagraph = firstInserted;
        }

        return clonedList;
    }

    private static Level CreateBulletLevel(char symbol, string fontName, string colorHex, int? fontSize) {
        var level = new Level();
        level.Append(new StartNumberingValue() { Val = 1 });
        level.Append(new NumberingFormat() { Val = NumberFormatValues.Bullet });
        level.Append(new LevelText() { Val = symbol.ToString() });
        level.Append(new LevelJustification() { Val = LevelJustificationValues.Left });

        var prevProps = new PreviousParagraphProperties();
        prevProps.Append(new Indentation() { Left = "720", Hanging = "360" });
        level.Append(prevProps);

        var symbolProps = new NumberingSymbolRunProperties();
        if (!string.IsNullOrEmpty(fontName)) {
            symbolProps.Append(new RunFonts { Ascii = fontName, HighAnsi = fontName });
        }
        if (!string.IsNullOrEmpty(colorHex)) {
            symbolProps.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = colorHex.Replace("#", "").ToLowerInvariant() });
        }
        if (fontSize.HasValue) {
            var size = (fontSize.Value * 2).ToString();
            symbolProps.Append(new FontSize { Val = size });
            symbolProps.Append(new FontSizeComplexScript { Val = size });
        }
        level.Append(symbolProps);
        return level;
    }

    /// <summary>
    /// Replaces the underlying abstract numbering definition while keeping the current numbering instance.
    /// </summary>
    /// <param name="newAbstract">The new abstract numbering definition.</param>
    private void ReplaceAbstractNum(AbstractNum newAbstract) {
        var numberingPart = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!;
        var numbering = numberingPart.Numbering;

        var oldAbstract = numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId.Value == _abstractId);
        if (oldAbstract == null) {
            return;
        }

        // preserve indentation from existing levels
        var oldLevels = oldAbstract.Elements<Level>().ToList();
        var newLevels = newAbstract.Elements<Level>().ToList();
        for (int i = 0; i < Math.Min(oldLevels.Count, newLevels.Count); i++) {
            var oldIndent = oldLevels[i].GetFirstChild<PreviousParagraphProperties>()?.GetFirstChild<Indentation>();
            if (oldIndent != null) {
                var prev = newLevels[i].GetFirstChild<PreviousParagraphProperties>();
                if (prev == null) {
                    prev = new PreviousParagraphProperties();
                    newLevels[i].Append(prev);
                }
                var indent = prev.GetFirstChild<Indentation>();
                if (indent == null) {
                    prev.Append((Indentation)oldIndent.CloneNode(true));
                } else {
                    indent.Left = oldIndent.Left;
                    indent.Hanging = oldIndent.Hanging;
                }
            }
        }

        newAbstract.AbstractNumberId = _abstractId;
        numbering.InsertAfter(newAbstract, oldAbstract);
        oldAbstract.Remove();
        numbering.Save();
    }

    private static char GetBulletSymbol(WordListLevelKind kind) {
        return kind switch {
            WordListLevelKind.Bullet => '\u2022',
            WordListLevelKind.BulletSquareSymbol => '\u25A0',
            WordListLevelKind.BulletBlackCircle => '\u25CF',
            WordListLevelKind.BulletDiamondSymbol => '\u25C6',
            WordListLevelKind.BulletArrowSymbol => '\u25BA',
            WordListLevelKind.BulletSolidRound => '·',
            WordListLevelKind.BulletOpenCircle => 'o',
            WordListLevelKind.BulletSquare2 => '■',
            WordListLevelKind.BulletSquare => '§',
            WordListLevelKind.BulletClubs => 'v',
            WordListLevelKind.BulletArrow => 'Ø',
            WordListLevelKind.BulletDiamond => '¨',
            WordListLevelKind.BulletCheckmark => 'ü',
            _ => throw new ArgumentOutOfRangeException(nameof(kind), "Only bullet kinds are supported")
        };
    }

    private static void EnsureW15Namespace(Numbering numbering) {
        const string prefix = "w15";
        const string ns = "http://schemas.microsoft.com/office/word/2012/wordml";
        if (numbering.LookupNamespace(prefix) == null) {
            numbering.AddNamespaceDeclaration(prefix, ns);
        }
        if (numbering.MCAttributes == null) {
            numbering.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = prefix };
        } else {
            var ignorable = numbering.MCAttributes.Ignorable?.Value;
            if (string.IsNullOrEmpty(ignorable)) {
                numbering.MCAttributes.Ignorable = prefix;
            } else if (!ignorable.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Contains(prefix)) {
                numbering.MCAttributes.Ignorable = ignorable + " " + prefix;
            }
        }
    }
}
