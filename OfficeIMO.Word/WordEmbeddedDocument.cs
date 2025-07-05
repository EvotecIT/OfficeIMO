using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents an embedded document within a <see cref="WordDocument"/>.
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
        /// Saves the embedded document to the specified file.
        /// </summary>
        /// <param name="fileName">Target file path.</param>
        public void Save(string fileName) {
            using (FileStream stream = new FileStream(fileName, FileMode.Create)) {
                using var altStream = _altContent.GetStream();
                altStream.CopyTo(stream);
            }
        }

        /// <summary>
        /// Removes the embedded document from the parent <see cref="WordDocument"/>.
        /// </summary>
        public void Remove() {
            _altChunk.Remove();

            var list = _document._document.MainDocumentPart.AlternativeFormatImportParts;
            foreach (var item in list) {
                var relationshipId = _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(item);
                if (relationshipId == _id) {
                    _document._wordprocessingDocument.MainDocumentPart.DeletePart(item);
                    break;
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordEmbeddedDocument"/> class
        /// based on an existing <see cref="AltChunk"/> element.
        /// </summary>
        /// <param name="wordDocument">Parent <see cref="WordDocument"/>.</param>
        /// <param name="altChunk">AltChunk that defines the embedded content.</param>
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

        /// <summary>
        /// Initializes a new instance of the <see cref="WordEmbeddedDocument"/> class
        /// using the specified file or HTML fragment.
        /// </summary>
        /// <param name="wordDocument">Parent <see cref="WordDocument"/>.</param>
        /// <param name="fileNameOrContent">File path or HTML content to embed.</param>
        /// <param name="alternativeFormatImportPartType">Explicit part type or <c>null</c> to infer from the file extension.</param>
        /// <param name="htmlFragment">When <c>true</c>, <paramref name="fileNameOrContent"/> is treated as HTML markup rather than a file path.</param>
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

            try {
                // if it's a fragment, we don't need to read the file
                var documentContent = htmlFragment
                    ? fileNameOrContent
                    : File.ReadAllText(fileNameOrContent, Encoding.UTF8);

                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(documentContent))) {
                    chunk.FeedData(ms);
                }

                _id = altChunkId;
                _altChunk = altChunk;
                _altContent = chunk;
                _document = wordDocument;

                mainDocPart.Document.Body.Append(altChunk);

                mainDocPart.Document.Save();
            } catch {
                mainDocPart.DeletePart(chunk);
                throw;
            }
        }
    }
}
