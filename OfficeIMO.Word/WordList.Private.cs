using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordList : WordElement {
    /// <summary>
    /// Retrieves the <see cref="AbstractNum"/> associated with this list.
    /// </summary>
    private AbstractNum GetAbstractNum() {
        return _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering
            .Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId.Value == _abstractId);
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
}
