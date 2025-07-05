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
            if (_wordDocument.MainDocumentPart == null) return;

            _part = _wordDocument.MainDocumentPart.CustomXmlParts
                .FirstOrDefault(p => string.Equals(p.ContentType, "application/bibliography+xml", System.StringComparison.OrdinalIgnoreCase));

            if (_part == null) return;

            var sources = new Sources();
            sources.Load(_part);

            foreach (var source in sources.Elements<Source>()) {
                var wrapper = new WordBibliographySource(source);
                if (!string.IsNullOrEmpty(wrapper.Tag)) {
                    _document.BibliographySources[wrapper.Tag] = wrapper;
                }
            }
        }

        private void SaveBibliography() {
            if (_wordDocument.MainDocumentPart == null) return;

            if (_part == null) {
                if (_document.BibliographySources.Count == 0) return;

                if (_document.FileOpenAccess == FileAccess.Read) {
                    throw new System.ArgumentException("Document is read only!");
                }

                _part = _wordDocument.MainDocumentPart.AddCustomXmlPart(CustomXmlPartType.Bibliography);
            }

            if (_document.BibliographySources.Count == 0) {
                _wordDocument.MainDocumentPart.DeletePart(_part);
                _part = null;
                return;
            }

            var sources = new Sources();
            foreach (var pair in _document.BibliographySources) {
                sources.Append(pair.Value.ToOpenXml());
            }

            sources.Save(_part);
        }
    }
}
