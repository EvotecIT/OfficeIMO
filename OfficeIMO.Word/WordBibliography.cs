using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Bibliography;

namespace OfficeIMO.Word {
    /// <summary>
    /// Loads and writes bibliographic sources to the document.
    /// </summary>
    internal class WordBibliography {
        private readonly WordDocument _document;
        private readonly WordprocessingDocument _wordDocument;
        private CustomXmlPart _part;

        public WordBibliography(WordDocument document, bool? create = null) {
            _document = document;
            _wordDocument = document._wordprocessingDocument;

            if (create == true) {
                SaveBibliography();
            } else {
                LoadBibliography();
            }
        }

        private void LoadBibliography() {
            // Loading existing bibliography sources is not currently implemented.
        }

        private void SaveBibliography() {
            if (_wordDocument.MainDocumentPart == null) return;

            if (_part == null) {
                if (_document.FileOpenAccess == FileAccess.Read) {
                    throw new System.ArgumentException("Document is read only!");
                }
                _part = _wordDocument.MainDocumentPart.AddCustomXmlPart(CustomXmlPartType.Bibliography);
            }

            var sources = new Sources();
            foreach (var pair in _document.BibliographySources) {
                sources.Append(pair.Value.ToOpenXml());
            }

            if (sources.ChildElements.Count > 0) {
                sources.Save(_part);
            } else {
                if (_part != null) {
                    _wordDocument.MainDocumentPart.DeletePart(_part);
                    _part = null;
                }
            }
        }
    }
}
