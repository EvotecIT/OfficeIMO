using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Converts structured logical PDF tables into Word tables.
    /// </summary>
    public static class PdfWordTableConverterExtensions {
        /// <summary>
        /// Extracts logical PDF tables into a new Word document written to <paramref name="documentPath"/>.
        /// </summary>
        /// <param name="document">Logical PDF document to import.</param>
        /// <param name="documentPath">Destination Word document path.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfWordTableImportResult> SavePdfTablesAsWord(
            this PdfCore.PdfLogicalDocument document,
            string documentPath,
            PdfWordTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(documentPath)) throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));

            using WordDocument word = WordDocument.Create(documentPath);
            IReadOnlyList<PdfWordTableImportResult> results = ImportTables(document, word, options ?? new PdfWordTableImportOptions());
            word.Save();
            return results;
        }

        /// <summary>
        /// Extracts logical PDF tables into a new Word document written to <paramref name="documentStream"/>.
        /// </summary>
        /// <param name="document">Logical PDF document to import.</param>
        /// <param name="documentStream">Writable destination stream for the document package.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfWordTableImportResult> SavePdfTablesAsWord(
            this PdfCore.PdfLogicalDocument document,
            Stream documentStream,
            PdfWordTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (documentStream == null) throw new ArgumentNullException(nameof(documentStream));

            using WordDocument word = WordDocument.Create(documentStream);
            IReadOnlyList<PdfWordTableImportResult> results = ImportTables(document, word, options ?? new PdfWordTableImportOptions());
            word.Save();
            return results;
        }

        /// <summary>
        /// Extracts logical PDF tables into Word document bytes.
        /// </summary>
        /// <param name="document">Logical PDF document to import.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Word document package bytes.</returns>
        public static byte[] ToWordTableDocumentBytes(
            this PdfCore.PdfLogicalDocument document,
            PdfWordTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            using var stream = new MemoryStream();
            document.SavePdfTablesAsWord(stream, options);
            return stream.ToArray();
        }

        /// <summary>
        /// Loads a PDF file, extracts logical tables, and writes them to a new Word document.
        /// </summary>
        /// <param name="pdfPath">Source PDF path.</param>
        /// <param name="documentPath">Destination Word document path.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfWordTableImportResult> SavePdfTablesAsWord(
            string pdfPath,
            string documentPath,
            PdfWordTableImportOptions? options = null) {
            if (string.IsNullOrWhiteSpace(pdfPath)) throw new ArgumentException("PDF path cannot be empty.", nameof(pdfPath));
            if (string.IsNullOrWhiteSpace(documentPath)) throw new ArgumentException("Document path cannot be empty.", nameof(documentPath));

            options ??= new PdfWordTableImportOptions();
            PdfCore.PdfLogicalDocument document = LoadPdf(pdfPath, options);
            return document.SavePdfTablesAsWord(documentPath, options);
        }

        /// <summary>
        /// Loads PDF bytes, extracts logical tables, and writes them to a new Word document stream.
        /// </summary>
        /// <param name="pdfBytes">Source PDF bytes.</param>
        /// <param name="documentStream">Writable destination stream for the document package.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfWordTableImportResult> SavePdfTablesAsWord(
            byte[] pdfBytes,
            Stream documentStream,
            PdfWordTableImportOptions? options = null) {
            if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));
            if (documentStream == null) throw new ArgumentNullException(nameof(documentStream));

            options ??= new PdfWordTableImportOptions();
            PdfCore.PdfLogicalDocument document = LoadPdf(pdfBytes, options);
            return document.SavePdfTablesAsWord(documentStream, options);
        }

        /// <summary>
        /// Loads a PDF stream, extracts logical tables, and writes them to a new Word document stream.
        /// </summary>
        /// <param name="pdfStream">Readable source PDF stream.</param>
        /// <param name="documentStream">Writable destination stream for the document package.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfWordTableImportResult> SavePdfTablesAsWord(
            Stream pdfStream,
            Stream documentStream,
            PdfWordTableImportOptions? options = null) {
            if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
            if (documentStream == null) throw new ArgumentNullException(nameof(documentStream));

            options ??= new PdfWordTableImportOptions();
            PdfCore.PdfLogicalDocument document = LoadPdf(pdfStream, options);
            return document.SavePdfTablesAsWord(documentStream, options);
        }

        private static IReadOnlyList<PdfWordTableImportResult> ImportTables(
            PdfCore.PdfLogicalDocument document,
            WordDocument word,
            PdfWordTableImportOptions options) {
            IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = PdfCore.PdfLogicalTableAnalysis.ExtractTables(document, options.MaxRows);
            if (tables.Count == 0) {
                AddEmptyDocumentParagraph(word, options);
                return Array.Empty<PdfWordTableImportResult>();
            }

            var results = new List<PdfWordTableImportResult>(tables.Count);
            for (int i = 0; i < tables.Count; i++) {
                PdfCore.PdfLogicalTableExtraction extraction = tables[i];
                PdfCore.PdfLogicalTableData data = extraction.Data;
                bool headerRowIncluded = HasHeaderRow(data);

                if (i > 0 && options.PageBreakBetweenTables) {
                    word.AddPageBreak();
                }

                if (options.IncludeSourceCaptions) {
                    word.AddParagraph(BuildCaption(extraction));
                }

                WordTable table = word.AddTable(data.Rows.Count + (headerRowIncluded ? 1 : 0), data.Columns.Count, options.TableStyle);
                PopulateTable(table, data, headerRowIncluded, options);

                results.Add(new PdfWordTableImportResult(
                    extraction.PageIndex,
                    extraction.PageNumber,
                    extraction.TableIndex,
                    extraction.DetectionKind,
                    data.Columns.Count,
                    data.Rows.Count,
                    data.TotalRowCount,
                    data.Truncated,
                    headerRowIncluded));
            }

            return results.AsReadOnly();
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(string path, PdfWordTableImportOptions options) {
            PdfCore.PdfPageRange[] ranges = GetPageRanges(options);
            return ranges.Length == 0
                ? PdfCore.PdfLogicalDocument.Load(path, options.LayoutOptions)
                : PdfCore.PdfLogicalDocument.LoadPageRanges(path, options.LayoutOptions, ranges);
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(byte[] pdfBytes, PdfWordTableImportOptions options) {
            PdfCore.PdfPageRange[] ranges = GetPageRanges(options);
            return ranges.Length == 0
                ? PdfCore.PdfLogicalDocument.Load(pdfBytes, options.LayoutOptions)
                : PdfCore.PdfLogicalDocument.LoadPageRanges(pdfBytes, options.LayoutOptions, ranges);
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(Stream stream, PdfWordTableImportOptions options) {
            PdfCore.PdfPageRange[] ranges = GetPageRanges(options);
            return ranges.Length == 0
                ? PdfCore.PdfLogicalDocument.Load(stream, options.LayoutOptions)
                : PdfCore.PdfLogicalDocument.LoadPageRanges(stream, options.LayoutOptions, ranges);
        }

        private static PdfCore.PdfPageRange[] GetPageRanges(PdfWordTableImportOptions options) {
            return options.PageRanges == null || options.PageRanges.Count == 0
                ? Array.Empty<PdfCore.PdfPageRange>()
                : options.PageRanges.ToArray();
        }

        private static void AddEmptyDocumentParagraph(WordDocument word, PdfWordTableImportOptions options) {
            string message = string.IsNullOrWhiteSpace(options.EmptyDocumentMessage)
                ? "No PDF tables detected."
                : options.EmptyDocumentMessage;
            word.AddParagraph(message);
        }

        private static bool HasHeaderRow(PdfCore.PdfLogicalTableData data) {
            return data.Columns.Count > 0
                && (data.Structure.HasHeaderRow || data.Structure.IsKeyValueTable)
                && data.Columns.Any(column => !string.IsNullOrWhiteSpace(column));
        }

        private static string BuildCaption(PdfCore.PdfLogicalTableExtraction extraction) {
            return "PDF page "
                + extraction.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + ", table "
                + (extraction.TableIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void PopulateTable(
            WordTable table,
            PdfCore.PdfLogicalTableData data,
            bool headerRowIncluded,
            PdfWordTableImportOptions options) {
            List<WordTableRow> rows = table.Rows;
            int rowOffset = headerRowIncluded ? 1 : 0;

            if (headerRowIncluded) {
                WriteRow(rows[0], data.Columns, data, alignNumericColumns: false);
                if (options.RepeatHeaderRows) {
                    rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
                }
            }

            for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
                WriteRow(rows[rowIndex + rowOffset], data.Rows[rowIndex], data, options.AlignNumericColumns);
            }

            if (options.FitTablesToPageWidth) {
                table.WidthType = TableWidthUnitValues.Pct;
                table.Width = 5000;
                table.DistributeColumnsEvenly();
            }
        }

        private static void WriteRow(
            WordTableRow row,
            IReadOnlyList<string> values,
            PdfCore.PdfLogicalTableData data,
            bool alignNumericColumns) {
            List<WordTableCell> cells = row.Cells;
            for (int columnIndex = 0; columnIndex < cells.Count; columnIndex++) {
                string value = columnIndex < values.Count ? values[columnIndex] : string.Empty;
                WordParagraph paragraph = cells[columnIndex].AddParagraph(value ?? string.Empty, removeExistingParagraphs: true);
                if (alignNumericColumns && data.IsNumericColumn(columnIndex)) {
                    paragraph.ParagraphAlignment = JustificationValues.Right;
                }
            }
        }
    }
}
