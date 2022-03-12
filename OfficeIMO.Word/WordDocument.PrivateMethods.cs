using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private void SaveNumbering() {
            // it seems the order of numbering instance/abstractnums in numbering matters...
            List<AbstractNum> listAbstractNum = new List<AbstractNum>();
            List<NumberingInstance> listNumberingInstance = new List<NumberingInstance>();
            List<NumberingPictureBullet> listNumberPictures = new List<NumberingPictureBullet>();

            if (_wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart != null) {
                var tempAbstractNumList = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<AbstractNum>();
                foreach (AbstractNum abstractNum in tempAbstractNumList) {
                    listAbstractNum.Add(abstractNum);
                }

                var tempNumberingInstance = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<NumberingInstance>();
                foreach (NumberingInstance numberingInstance in tempNumberingInstance) {
                    listNumberingInstance.Add(numberingInstance);
                }

                var tempNumberingPictureBullet = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements.OfType<NumberingPictureBullet>();
                foreach (NumberingPictureBullet pictureBullet in tempNumberingPictureBullet) {
                    listNumberPictures.Add(pictureBullet);
                }

                if (_wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart != null) {
                    Numbering numbering = _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    if (numbering != null) {
                        numbering.RemoveAllChildren();
                    }
                }

                foreach (NumberingPictureBullet pictureBullet in listNumberPictures) {
                    _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Append(pictureBullet);
                }

                foreach (AbstractNum abstractNum in listAbstractNum) {
                    _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Append(abstractNum);
                }

                foreach (NumberingInstance numberingInstance in listNumberingInstance) {
                    _wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Append(numberingInstance);
                }
            }
        }

    }
}
