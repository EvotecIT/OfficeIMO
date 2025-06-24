using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents an embedded document such as RTF or HTML content inserted into a Word document.
    /// </summary>
    public class WordEmbeddedDocument : WordElement {
        private string _id;
        private AltChunk _altChunk;
        private readonly AlternativeFormatImportPart _altContent;
        private readonly WordDocument _document;

        /// <summary>
        /// Gets the content type of the embedded document.
        /// </summary>
        public string ContentType => _altContent.ContentType;


        /// <summary>
        /// Saves the embedded document to the specified <paramref name="fileName"/>.
        /// </summary>
        /// <param name="fileName">Path of the file to create.</param>
        public void Save(string fileName) {
            using (FileStream stream = new FileStream(fileName, FileMode.Create)) {
                var altStream = _altContent.GetStream();
                altStream.CopyTo(stream);
                altStream.Close();
            }
        }

        /// <summary>
        /// Removes the embedded document from the parent Word document.
        /// </summary>
        public void Remove() {
            _altChunk.Remove();

            var list = _document._document.MainDocumentPart.AlternativeFormatImportParts;
            AlternativeFormatImportPart itemToDelete = null;
            foreach (var item in list) {
                var relationshipId = _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(item);
                if (relationshipId == _id) {
                    itemToDelete = item;
                }
            }

            if (itemToDelete != null) {
                _document._wordprocessingDocument.MainDocumentPart.DeletePart(itemToDelete);
            }
        }

        public WordEmbeddedDocument(WordDocument wordDocument, AltChunk altChunk) {
            _id = altChunk.Id;
            _altChunk = altChunk;
            _document = wordDocument;

            var list = wordDocument._document.MainDocumentPart.AlternativeFormatImportParts;
            foreach (var item in list) {
                var relationshipId = wordDocument._wordprocessingDocument.MainDocumentPart.GetIdOfPart(item);
                if (relationshipId == _id) {
                    _altContent = item;
                }
            }
        }

        public WordEmbeddedDocument(WordDocument wordDocument, string fileNameOrContent, WordAlternativeFormatImportPartType? alternativeFormatImportPartType, bool htmlFragment) {
            WordAlternativeFormatImportPartType partType;
            if (alternativeFormatImportPartType == null) {
                FileInfo fileInfo = new FileInfo(fileNameOrContent);
                if (fileInfo.Extension == ".rtf") {
                    partType = WordAlternativeFormatImportPartType.Rtf;
                } else if (fileInfo.Extension == ".html") {
                    partType = WordAlternativeFormatImportPartType.Html;
                } else if (fileInfo.Extension == ".log" || fileInfo.Extension == ".txt") {
                    partType = WordAlternativeFormatImportPartType.TextPlain;
                } else {
                    throw new InvalidOperationException("Only RTF and HTML files are supported for now :-)");
                }
            } else {
                partType = alternativeFormatImportPartType.Value;
            }

            MainDocumentPart mainDocPart = wordDocument._document.MainDocumentPart;

            PartTypeInfo partTypeInfo = partType switch {
                WordAlternativeFormatImportPartType.Rtf => AlternativeFormatImportPartType.Rtf,
                WordAlternativeFormatImportPartType.Html => AlternativeFormatImportPartType.Html,
                WordAlternativeFormatImportPartType.TextPlain => AlternativeFormatImportPartType.TextPlain,
                _ => throw new InvalidOperationException("Unsupported format type")
            };

            AlternativeFormatImportPart chunk = mainDocPart.AddAlternativeFormatImportPart(partTypeInfo);
            string altChunkId = mainDocPart.GetIdOfPart(chunk);
            AltChunk altChunk = new AltChunk { Id = altChunkId };

            // if it's a fragment, we don't need to read the file
            var documentContent = htmlFragment ? fileNameOrContent : File.ReadAllText(fileNameOrContent, Encoding.ASCII);

            using (MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(documentContent))) {
                chunk.FeedData(ms);
            }

            _id = altChunkId;
            _altChunk = altChunk;
            _altContent = chunk;
            _document = wordDocument;

            mainDocPart.Document.Body.Append(altChunk);

            mainDocPart.Document.Save();
        }
    }
}
