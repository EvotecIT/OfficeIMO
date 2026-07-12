using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Provides semantic PDF to Word import helpers over the first-party OfficeIMO.Pdf logical reader.
    /// </summary>
    public static class PdfWordConverterExtensions {
        /// <summary>Converts a first-party PDF read model into an editable Word document using semantic extraction.</summary>
        public static WordDocument ToWordDocument(this PdfCore.PdfReadDocument document, PdfWordReadOptions? options = null) {
            return document.ToWordDocumentResult(options).Value;
        }

        /// <summary>Converts a PDF read model into an editable Word document with import diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentResult(this PdfCore.PdfReadDocument document, PdfWordReadOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            PdfWordReadOptions operation = (options ?? new PdfWordReadOptions()).CloneForConversion();
            WordDocument word = PdfWordConverter.Convert(LoadLogical(document, operation), operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Converts a first-party logical PDF read model into an editable Word document using semantic extraction.</summary>
        public static WordDocument ToWordDocument(this PdfCore.PdfLogicalDocument document, PdfWordReadOptions? options = null) {
            return document.ToWordDocumentResult(options).Value;
        }

        /// <summary>Converts a logical PDF model into an editable Word document with import diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentResult(this PdfCore.PdfLogicalDocument document, PdfWordReadOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            PdfWordReadOptions operation = (options ?? new PdfWordReadOptions()).CloneForConversion();
            WordDocument word = PdfWordConverter.Convert(document, operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Reads source PDF bytes and converts parser-supported semantic content to a Word document.</summary>
        public static WordDocument ToWordDocumentFromPdf(this byte[] pdfBytes, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            return pdfBytes.ToWordDocumentFromPdfResult(options, readOptions).Value;
        }

        /// <summary>Reads PDF bytes and returns editable Word output with import diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentFromPdfResult(this byte[] pdfBytes, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            PdfWordReadOptions operation = (options ?? new PdfWordReadOptions()).CloneForConversion();
            WordDocument word = PdfWordConverter.Convert(LoadPdf(pdfBytes, operation, readOptions), operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Reads a PDF stream from the current position and converts parser-supported semantic content to a Word document.</summary>
        public static WordDocument ToWordDocumentFromPdf(this Stream pdfStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            return pdfStream.ToWordDocumentFromPdfResult(options, readOptions).Value;
        }

        /// <summary>Reads a PDF stream and returns editable Word output with import diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentFromPdfResult(this Stream pdfStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            PdfWordReadOptions operation = (options ?? new PdfWordReadOptions()).CloneForConversion();
            WordDocument word = PdfWordConverter.Convert(LoadPdf(pdfStream, operation, readOptions), operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Reads a PDF file and converts parser-supported semantic content to a Word document.</summary>
        public static WordDocument ToWordDocumentFromPdfFile(this string pdfPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            return pdfPath.ToWordDocumentFromPdfFileResult(options, readOptions).Value;
        }

        /// <summary>Reads a PDF file and returns editable Word output with import diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentFromPdfFileResult(this string pdfPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            PdfWordReadOptions operation = (options ?? new PdfWordReadOptions()).CloneForConversion();
            WordDocument word = PdfWordConverter.Convert(LoadPdf(pdfPath, operation, readOptions), operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Reads source PDF bytes and returns a serialized Word document package.</summary>
        public static byte[] ToWordBytesFromPdf(this byte[] pdfBytes, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            using WordDocument document = pdfBytes.ToWordDocumentFromPdf(options, readOptions);
            return document.ToBytes();
        }

        /// <summary>Reads a PDF stream and returns a serialized Word document package.</summary>
        public static byte[] ToWordBytesFromPdf(this Stream pdfStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            using WordDocument document = pdfStream.ToWordDocumentFromPdf(options, readOptions);
            return document.ToBytes();
        }

        /// <summary>Reads a PDF file and returns a serialized Word document package.</summary>
        public static byte[] ToWordBytesFromPdfFile(this string pdfPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            using WordDocument document = pdfPath.ToWordDocumentFromPdfFile(options, readOptions);
            return document.ToBytes();
        }

        /// <summary>Reads source PDF bytes and saves semantic Word output to a file.</summary>
        public static void SavePdfAsWord(this byte[] pdfBytes, string documentPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            if (string.IsNullOrWhiteSpace(documentPath)) {
                throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));
            }

            using WordDocument document = pdfBytes.ToWordDocumentFromPdf(options, readOptions);
            document.Save(documentPath);
        }

        /// <summary>Reads a PDF stream and saves semantic Word output to a file.</summary>
        public static void SavePdfAsWord(this Stream pdfStream, string documentPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            if (string.IsNullOrWhiteSpace(documentPath)) {
                throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));
            }

            using WordDocument document = pdfStream.ToWordDocumentFromPdf(options, readOptions);
            document.Save(documentPath);
        }

        /// <summary>Reads a PDF file and saves semantic Word output to a file.</summary>
        public static void SavePdfAsWord(string pdfPath, string documentPath, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            if (string.IsNullOrWhiteSpace(documentPath)) {
                throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));
            }

            using WordDocument document = pdfPath.ToWordDocumentFromPdfFile(options, readOptions);
            document.Save(documentPath);
        }

        /// <summary>Reads source PDF bytes and saves semantic Word output to a writable stream.</summary>
        public static void SavePdfAsWord(this byte[] pdfBytes, Stream documentStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfBytes == null) {
                throw new ArgumentNullException(nameof(pdfBytes));
            }

            if (documentStream == null) {
                throw new ArgumentNullException(nameof(documentStream));
            }

            using WordDocument document = pdfBytes.ToWordDocumentFromPdf(options, readOptions);
            document.Save(documentStream);
        }

        /// <summary>Reads a PDF stream and saves semantic Word output to a writable stream.</summary>
        public static void SavePdfAsWord(this Stream pdfStream, Stream documentStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfStream == null) {
                throw new ArgumentNullException(nameof(pdfStream));
            }

            if (documentStream == null) {
                throw new ArgumentNullException(nameof(documentStream));
            }

            using WordDocument document = pdfStream.ToWordDocumentFromPdf(options, readOptions);
            document.Save(documentStream);
        }

        /// <summary>Reads a PDF file and saves semantic Word output to a writable stream.</summary>
        public static void SavePdfAsWord(string pdfPath, Stream documentStream, PdfWordReadOptions? options = null, PdfCore.PdfReadOptions? readOptions = null) {
            if (pdfPath == null) {
                throw new ArgumentNullException(nameof(pdfPath));
            }

            if (documentStream == null) {
                throw new ArgumentNullException(nameof(documentStream));
            }

            using WordDocument document = pdfPath.ToWordDocumentFromPdfFile(options, readOptions);
            document.Save(documentStream);
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

    }
}
