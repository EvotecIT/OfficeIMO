using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// First-party Excel workbook to PDF conversion helpers.
    /// </summary>
    public static partial class ExcelPdfConverterExtensions {
        /// <summary>
        /// Converts an Excel workbook to a first-party OfficeIMO PDF document model.
        /// </summary>
        public static PdfCore.PdfDocument ToPdfDocument(this ExcelDocument document, ExcelPdfSaveOptions? options = null) {
            return document.ToPdfDocumentResult(options).Value;
        }

        private static PdfCore.PdfDocument ConvertToPdfDocument(ExcelDocument document, ExcelPdfSaveOptions options) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            PdfCore.PdfOptions pdfOptions = CreatePdfOptions(options, out bool preserveConfiguredFontSlots);
            PdfCore.PdfStandardFont defaultFontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(pdfOptions.DefaultFont);
            using ExcelDocumentReader reader = document.CreateReader();
            IReadOnlyList<string> sheetNames = GetSheetNames(reader, options);
            bool hasExplicitSheetSelection = HasExplicitSheetSelection(options);
            IReadOnlyList<WorksheetPdfExportPlan> exportPlans = BuildWorksheetExportPlans(document, reader, sheetNames, options, hasExplicitSheetSelection, defaultFontFamily);
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = RegisterWorksheetFonts(pdfOptions, exportPlans, options, preserveConfiguredFontSlots);
            ApplyTextFallbacks(pdfOptions, options, preserveConfiguredFontSlots, registeredFontSlots);
            var pdf = PdfCore.PdfDocument.Create(pdfOptions);
            IReadOnlyDictionary<string, string> sheetDestinations = BuildSheetDestinationMap(exportPlans);
            IReadOnlyDictionary<string, string> cellDestinations = BuildCellDestinationMap(exportPlans);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                if (options.WorksheetLayout == ExcelPdfWorksheetLayoutMode.WorksheetCanvas) {
                    AddWorksheetCanvasPages(pdf, document, plan, options, sheetDestinations, cellDestinations, defaultFontFamily);
                    continue;
                }

                object?[,] values = plan.ExportData.Values;
                int columns = values.GetLength(1);

                pdf.Section(page => {
                    ApplyWorksheetPageSetup(page, plan.PageSetup, options);
                    ApplyWorksheetHeaderFooter(page, plan.HeaderFooter, plan.SheetName, document.FilePath, options);
                    page.Content(content => content.Item(item => {
                        item.Bookmark(plan.BookmarkName);
                        if (options.IncludeSheetHeadings) {
                            item.H1(plan.SheetName);
                        }

                        IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>> imagesByCellReference = CreateWorksheetImageMap(plan);
                        foreach (WorksheetImageExportData image in plan.Images) {
                            if (!imagesByCellReference.ContainsKey(NormalizeCellReference(image.CellReference))) {
                                item.Image(image.Bytes, image.WidthPoints, image.HeightPoints, PdfCore.PdfAlign.Left, spacingBefore: 4, spacingAfter: 6, style: CreateConverterImageStyle(image));
                            }
                        }

                        foreach (WorksheetChartExportData chart in plan.Charts) {
                            AddWorksheetChart(item, chart, plan.SheetName, options);
                        }

                        if (plan.HasTable) {
                            IReadOnlyList<TableChunk> chunks = CreateTableChunks(plan, options, columns);
                            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                                TableChunk chunk = chunks[chunkIndex];
                                if (chunkIndex > 0) {
                                    item.PageBreak();
                                }

                                item.Table(
                                    CreatePdfRows(values, plan.ExportData.Styles, plan.ExportData.Hyperlinks, plan.ExportData.CellReferences, plan.ExportData.MergedCells, imagesByCellReference, chunk.RowIndexes, chunk.StartColumn, chunk.ColumnCount, options.EmptyCellText, sheetDestinations, cellDestinations, plan.SheetName, defaultFontFamily),
                                    style: CreateTableStyle(options, plan.PageSetup, chunk.RowIndexes, chunk.HeaderRowCount, plan.ExportData.Styles, plan.ExportData.ConditionalFills, plan.ExportData.ColumnWidths, plan.ExportData.RowHeights, chunk.StartColumn, chunk.ColumnCount));
                            }
                        }
                    }));
                });
            }

            if (exportPlans.Count == 0) {
                pdf.H1("Workbook");
                pdf.Table(new[] { new[] { "No worksheet data found." } }, style: CreateEmptyWorkbookTableStyle(options));
            }

            return pdf;
        }

        /// <summary>
        /// Converts an Excel workbook to a PDF document and returns conversion diagnostics with it.
        /// </summary>
        public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this ExcelDocument document, ExcelPdfSaveOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            ExcelPdfSaveOptions operation = (options ?? new ExcelPdfSaveOptions()).CloneForConversion();
            PdfCore.PdfDocument pdf = ConvertToPdfDocument(document, operation);
            return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
        }

        private static PdfCore.PdfImageStyle CreateConverterImageStyle() => new() {
            ScaleDownToFit = true
        };

        private static PdfCore.PdfImageStyle CreateConverterImageStyle(WorksheetImageExportData image) {
            PdfCore.PdfImageStyle style = CreateConverterImageStyle();
            style.RotationAngle = -image.RotationDegrees;
            return style;
        }

        /// <summary>
        /// Converts an Excel workbook to PDF bytes.
        /// </summary>
        /// <example><code>byte[] pdf = workbook.ToPdf();</code></example>
        public static byte[] ToPdf(this ExcelDocument document, ExcelPdfSaveOptions? options = null) {
            return document.ToPdfDocument(options).ToBytes();
        }

        /// <summary>
        /// Saves an Excel workbook as a PDF file.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdf(this ExcelDocument document, string path, ExcelPdfSaveOptions? options = null) =>
            document.ToPdfDocumentResult(options).Save(path);

        /// <summary>
        /// Attempts to save an Excel workbook as a PDF file and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult TrySaveAsPdf(this ExcelDocument document, string path, ExcelPdfSaveOptions? options = null) {
            try {
                return document.ToPdfDocumentResult(options).TrySave(path);
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(path, ex);
            }
        }

        /// <summary>
        /// Writes an Excel workbook as PDF to a stream.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdf(this ExcelDocument document, Stream stream, ExcelPdfSaveOptions? options = null) =>
            document.ToPdfDocumentResult(options).Save(stream);

        /// <summary>
        /// Attempts to write an Excel workbook as PDF to a stream and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult TrySaveAsPdf(this ExcelDocument document, Stream stream, ExcelPdfSaveOptions? options = null) {
            try {
                return document.ToPdfDocumentResult(options).TrySave(stream);
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
        }

        /// <summary>Converts synchronously, then asynchronously saves an Excel workbook PDF at the specified path.</summary>
        public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
            this ExcelDocument document,
            string path,
            ExcelPdfSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return document.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
        }

        /// <summary>Converts synchronously, then asynchronously saves an Excel workbook PDF to a caller-owned stream.</summary>
        public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
            this ExcelDocument document,
            Stream stream,
            ExcelPdfSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return document.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
        }

        /// <summary>Attempts to asynchronously save an Excel workbook as PDF at the specified path.</summary>
        public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
            this ExcelDocument document,
            string path,
            ExcelPdfSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                return await document.ToPdfDocumentResult(options)
                    .TrySaveAsync(path, cancellationToken)
                    .ConfigureAwait(false);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(path, ex);
            }
        }

        /// <summary>Attempts to asynchronously save an Excel workbook as PDF to a caller-owned stream.</summary>
        public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
            this ExcelDocument document,
            Stream stream,
            ExcelPdfSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                return await document.ToPdfDocumentResult(options)
                    .TrySaveAsync(stream, cancellationToken)
                    .ConfigureAwait(false);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
        }

    }
}
