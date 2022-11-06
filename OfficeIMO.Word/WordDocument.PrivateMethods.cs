using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private void SaveNumbering() {
        var numbering = GetNumbering();
        if (numbering == null) {
            return;
        }

        // it seems the order of numbering instance/abstractnums in numbering matters...

        var listAbstractNum = numbering.ChildElements.OfType<AbstractNum>().ToArray();
        var listNumberingInstance = numbering.ChildElements.OfType<NumberingInstance>().ToArray();
        var listNumberPictures = numbering.ChildElements.OfType<NumberingPictureBullet>().ToArray();

        numbering.RemoveAllChildren();

        foreach (var pictureBullet in listNumberPictures) {
            numbering.Append(pictureBullet);
        }

        foreach (var abstractNum in listAbstractNum) {
            numbering.Append(abstractNum);
        }

        foreach (var numberingInstance in listNumberingInstance) {
            numbering.Append(numberingInstance);
        }
    }

    private Numbering GetNumbering() {
        return _wordprocessingDocument?.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
    }
}
