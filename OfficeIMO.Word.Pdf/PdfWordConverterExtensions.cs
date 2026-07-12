using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Provides semantic PDF to Word import helpers over the first-party OfficeIMO.Pdf logical reader.
    /// </summary>
    public static class PdfWordConverterExtensions {
        /// <summary>Converts a first-party PDF read model into an editable Word document using semantic extraction.</summary>
        public static WordDocument ToWordDocument(this PdfCore.PdfReadDocument document, PdfWordReadOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            PdfWordReadOptions readOptions = options ?? new PdfWordReadOptions();
            readOptions.ResetImportState();
            return PdfWordConverter.Convert(LoadLogical(document, readOptions), readOptions);
        }

        /// <summary>Converts a first-party logical PDF read model into an editable Word document using semantic extraction.</summary>
        public static WordDocument ToWordDocument(this PdfCore.PdfLogicalDocument document, PdfWordReadOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            return PdfWordConverter.Convert(document, options);
        }

        /// <summary>Reads source PDF bytes and converts parser-supported semantic content to a Word document.</summary>
        public static WordDocument ToWordDocumentFromPdf(this byte[] pdfBytes, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            return LoadPdf(pdfBytes, importOptions, readOptions).ToWordDocument(importOptions);
        }

        /// <summary>Reads a PDF stream from the current position and converts parser-supported semantic content to a Word document.</summary>
        public static WordDocument ToWordDocumentFromPdf(this Stream pdfStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            return LoadPdf(pdfStream, importOptions, readOptions).ToWordDocument(importOptions);
        }

        /// <summary>Reads a PDF file and converts parser-supported semantic content to a Word document.</summary>
        public static WordDocument ToWordDocumentFromPdfFile(this string pdfPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            return LoadPdf(pdfPath, importOptions, readOptions).ToWordDocument(importOptions);
        }

        /// <summary>Reads source PDF bytes and returns a serialized Word document package.</summary>
        public static byte[] ToWordBytesFromPdf(this byte[] pdfBytes, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            using var stream = new MemoryStream();
            pdfBytes.SavePdfAsWord(stream, options, readOptions);
            return stream.ToArray();
        }

        /// <summary>Reads a PDF stream and returns a serialized Word document package.</summary>
        public static byte[] ToWordBytesFromPdf(this Stream pdfStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            using var stream = new MemoryStream();
            pdfStream.SavePdfAsWord(stream, options, readOptions);
            return stream.ToArray();
        }

        /// <summary>Reads a PDF file and returns a serialized Word document package.</summary>
        public static byte[] ToWordBytesFromPdfFile(this string pdfPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            using var stream = new MemoryStream();
            SavePdfAsWord(pdfPath, stream, options, readOptions);
            return stream.ToArray();
        }

        /// <summary>Reads source PDF bytes and saves semantic Word output to a file.</summary>
        public static void SavePdfAsWord(this byte[] pdfBytes, string documentPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            if (string.IsNullOrWhiteSpace(documentPath)) {
                throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            PdfCore.PdfLogicalDocument logical = LoadPdf(pdfBytes, importOptions, readOptions);
            SaveWordDocument(logical, documentPath, importOptions);
        }

        /// <summary>Reads a PDF stream and saves semantic Word output to a file.</summary>
        public static void SavePdfAsWord(this Stream pdfStream, string documentPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            if (string.IsNullOrWhiteSpace(documentPath)) {
                throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            PdfCore.PdfLogicalDocument logical = LoadPdf(pdfStream, importOptions, readOptions);
            SaveWordDocument(logical, documentPath, importOptions);
        }

        /// <summary>Reads a PDF file and saves semantic Word output to a file.</summary>
        public static void SavePdfAsWord(string pdfPath, string documentPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            if (string.IsNullOrWhiteSpace(documentPath)) {
                throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            PdfCore.PdfLogicalDocument logical = LoadPdf(pdfPath, importOptions, readOptions);
            SaveWordDocument(logical, documentPath, importOptions);
        }

        /// <summary>Reads source PDF bytes and saves semantic Word output to a writable stream.</summary>
        public static void SavePdfAsWord(this byte[] pdfBytes, Stream documentStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            if (documentStream == null) {
                throw new ArgumentNullException(nameof(documentStream));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            PdfCore.PdfLogicalDocument logical = LoadPdf(pdfBytes, importOptions, readOptions);
            SaveWordDocument(logical, documentStream, importOptions);
        }

        /// <summary>Reads a PDF stream and saves semantic Word output to a writable stream.</summary>
        public static void SavePdfAsWord(this Stream pdfStream, Stream documentStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            if (documentStream == null) {
                throw new ArgumentNullException(nameof(documentStream));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            PdfCore.PdfLogicalDocument logical = LoadPdf(pdfStream, importOptions, readOptions);
            SaveWordDocument(logical, documentStream, importOptions);
        }

        /// <summary>Reads a PDF file and saves semantic Word output to a writable stream.</summary>
        public static void SavePdfAsWord(string pdfPath, Stream documentStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            if (documentStream == null) {
                throw new ArgumentNullException(nameof(documentStream));
            }

            PdfWordReadOptions importOptions = options ?? new PdfWordReadOptions();
            PdfCore.PdfLogicalDocument logical = LoadPdf(pdfPath, importOptions, readOptions);
            SaveWordDocument(logical, documentStream, importOptions);
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(byte[] pdfBytes, PdfWordReadOptions options, PdfCore.PdfReadOptions? readOptions) {
            PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Load(pdfBytes, readOptions);
            return LoadLogical(document, options);
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(Stream pdfStream, PdfWordReadOptions options, PdfCore.PdfReadOptions? readOptions) {
            PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Load(pdfStream, readOptions);
            return LoadLogical(document, options);
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(string pdfPath, PdfWordReadOptions options, PdfCore.PdfReadOptions? readOptions) {
            PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Load(pdfPath, readOptions);
            return LoadLogical(document, options);
        }

        private static PdfCore.PdfLogicalDocument LoadLogical(PdfCore.PdfReadDocument document, PdfWordReadOptions options) {
            PdfCore.PdfPageRange[] ranges = GetPageRanges(options);
            return ranges.Length == 0
                ? PdfCore.PdfLogicalDocument.From(document, options.LayoutOptions)
                : PdfCore.PdfLogicalDocument.FromPageRanges(document, options.LayoutOptions, ranges);
        }

        private static PdfCore.PdfPageRange[] GetPageRanges(PdfWordReadOptions options) {
            return options.PageRanges == null || options.PageRanges.Count == 0
                ? Array.Empty<PdfCore.PdfPageRange>()
                : options.PageRanges.ToArray();
        }

        private static void SaveWordDocument(PdfCore.PdfLogicalDocument logical, string documentPath, PdfWordReadOptions options) {
            using WordDocument word = WordDocument.Create(documentPath);
            options.ResetImportState();
            PdfWordConverter.ImportInto(logical, word, options);
            word.Save();
        }

        private static void SaveWordDocument(PdfCore.PdfLogicalDocument logical, Stream documentStream, PdfWordReadOptions options) {
            using WordDocument word = WordDocument.Create(documentStream);
            options.ResetImportState();
            PdfWordConverter.ImportInto(logical, word, options);
            word.Save();
            if (documentStream.CanSeek) {
                documentStream.Position = 0;
            }
        }
    }
}
