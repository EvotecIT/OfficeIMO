using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents an embedded document within a <see cref="WordDocument"/>.
    /// </summary>
    public class WordEmbeddedDocument : WordElement {
        private readonly string _id;
        private readonly AltChunk _altChunk;
        private readonly AlternativeFormatImportPart _altContent;
        private readonly WordDocument _document;

        /// <summary>
        /// Gets the content type of the embedded document.
        /// </summary>
        public string ContentType => _altContent.ContentType;


        /// <summary>
        /// Retrieves the HTML markup of the embedded document when available.
        /// </summary>
        /// <returns>HTML content or <c>null</c> if the embedded document is not HTML.</returns>
        public string? GetHtml() {
            if (!string.Equals(ContentType, "text/html", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            using (var stream = _altContent.GetStream())
            using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
                return reader.ReadToEnd();
            }
        }

        /// <summary>Asynchronously retrieves the HTML markup of the embedded document when available.</summary>
        public async Task<string?> GetHtmlAsync(CancellationToken cancellationToken = default) {
            if (!string.Equals(ContentType, "text/html", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            cancellationToken.ThrowIfCancellationRequested();
            using var stream = OpenRead();
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            string html = await reader.ReadToEndAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return html;
        }

        /// <summary>Opens a readable stream for the embedded document payload.</summary>
        public Stream OpenRead() => _altContent.GetStream(FileMode.Open, FileAccess.Read);

        /// <summary>Returns a copy of the embedded document payload.</summary>
        public byte[] ToBytes() {
            using Stream stream = OpenRead();
            return OfficeStreamReader.ReadAllBytes(stream);
        }

        /// <summary>Returns the embedded document payload in a new stream positioned at the beginning.</summary>
        public MemoryStream ToStream() => new MemoryStream(ToBytes());


        /// <summary>
        /// Saves the embedded document to the specified file.
        /// </summary>
        /// <param name="fileName">Target file path.</param>
        public void Save(string fileName) {
            OfficeFileCommit.WriteAllBytes(fileName, ToBytes());
        }

        /// <summary>Saves the embedded document to a caller-owned stream.</summary>
        public void Save(Stream stream) => OfficeStreamWriter.WriteAllBytes(stream, ToBytes());

        /// <summary>Saves the embedded document to the specified file asynchronously.</summary>
        public async Task SaveAsync(string fileName, CancellationToken cancellationToken = default) {
            using Stream source = OpenRead();
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(source, cancellationToken).ConfigureAwait(false);
            await OfficeFileCommit.WriteAllBytesAsync(fileName, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Saves the embedded document to a caller-owned stream asynchronously.</summary>
        public async Task SaveAsync(Stream stream, CancellationToken cancellationToken = default) {
            using Stream source = OpenRead();
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(source, cancellationToken).ConfigureAwait(false);
            await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Removes the embedded document from the parent <see cref="WordDocument"/>.
        /// </summary>
        public void Remove() {
            MainDocumentPart mainPart = GetMainDocumentPart(_document);
            _altChunk.Remove();
            mainPart.DeletePart(_altContent);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordEmbeddedDocument"/> class
        /// based on an existing <see cref="AltChunk"/> element.
        /// </summary>
        /// <param name="wordDocument">Parent <see cref="WordDocument"/>.</param>
        /// <param name="altChunk">AltChunk that defines the embedded content.</param>
        public WordEmbeddedDocument(WordDocument wordDocument, AltChunk altChunk) {
            if (wordDocument == null) throw new ArgumentNullException(nameof(wordDocument));
            if (altChunk == null) throw new ArgumentNullException(nameof(altChunk));

            _document = wordDocument;
            _altChunk = altChunk;

            string? chunkId = altChunk.Id?.Value ?? altChunk.Id;
            if (string.IsNullOrWhiteSpace(chunkId)) {
                throw new InvalidOperationException("The supplied AltChunk does not declare a relationship id.");
            }

            _id = chunkId!;

            MainDocumentPart mainPart = GetMainDocumentPart(wordDocument);
            AlternativeFormatImportPart? matchingPart = mainPart.AlternativeFormatImportParts
                .FirstOrDefault(part => string.Equals(mainPart.GetIdOfPart(part), _id, StringComparison.Ordinal));

            _altContent = matchingPart ?? throw new InvalidOperationException($"Could not find an alternative format part with id '{_id}'.");
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
            if (wordDocument == null) throw new ArgumentNullException(nameof(wordDocument));
            if (string.IsNullOrWhiteSpace(fileNameOrContent)) throw new ArgumentException("Value cannot be null or whitespace.", nameof(fileNameOrContent));

            WordAlternativeFormatImportPartType partType;
            if (alternativeFormatImportPartType == null) {
                FileInfo fileInfo = new FileInfo(fileNameOrContent);
                string extension = fileInfo.Extension;
                if (extension.Equals(".rtf", StringComparison.OrdinalIgnoreCase)) {
                    partType = WordAlternativeFormatImportPartType.Rtf;
                } else if (extension.Equals(".html", StringComparison.OrdinalIgnoreCase) || extension.Equals(".htm", StringComparison.OrdinalIgnoreCase)) {
                    partType = WordAlternativeFormatImportPartType.Html;
                } else if (extension.Equals(".log", StringComparison.OrdinalIgnoreCase) || extension.Equals(".txt", StringComparison.OrdinalIgnoreCase)) {
                    partType = WordAlternativeFormatImportPartType.TextPlain;
                } else {
                    throw new InvalidOperationException("Only RTF and HTML files are supported for now :-)");
                }
            } else {
                partType = alternativeFormatImportPartType.Value;
            }

            MainDocumentPart mainDocPart = GetMainDocumentPart(wordDocument);

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

                var documentRoot = mainDocPart.Document ?? throw new InvalidOperationException("The document is missing its root element.");
                if (documentRoot.Body == null) {
                    throw new InvalidOperationException("The document does not contain a body element.");
                }

                wordDocument.AppendBlockToBody(altChunk);

                documentRoot.Save();
            } catch {
                mainDocPart.DeletePart(chunk);
                throw;
            }
        }

        private static MainDocumentPart GetMainDocumentPart(WordDocument document) {
            MainDocumentPart? mainPart = document._wordprocessingDocument?.MainDocumentPart;
            return mainPart ?? throw new InvalidOperationException("The Word document is not associated with a main document part.");
        }
    }
}
