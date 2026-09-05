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
        public static PdfCore.PdfDocument ToPdfDocument(this ExcelDocument document, ExcelToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            return document.ToPdfDocumentResult(options, cancellationToken).Value;
        }

        private static PdfCore.PdfDocument ConvertToPdfDocument(ExcelDocument document, ExcelToPdfOptions options) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            CancellationToken cancellationToken = options.CancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfOptions pdfOptions = CreatePdfOptions(options, out bool preserveConfiguredFontSlots);
            PdfCore.PdfStandardFont defaultFontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(pdfOptions.DefaultFont);
            using ExcelDocumentReader reader = document.CreateReader();
            IReadOnlyList<string> sheetNames = GetSheetNames(reader, options);
            bool hasExplicitSheetSelection = HasExplicitSheetSelection(options);
            IReadOnlyList<WorksheetPdfExportPlan> exportPlans = BuildWorksheetExportPlans(document, reader, sheetNames, options, hasExplicitSheetSelection, defaultFontFamily);
            cancellationToken.ThrowIfCancellationRequested();
            ReportAccountingUnderlineApproximations(exportPlans, options);
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = RegisterWorksheetFonts(pdfOptions, exportPlans, options, preserveConfiguredFontSlots);
            ApplyTextFallbacks(pdfOptions, options, preserveConfiguredFontSlots, registeredFontSlots, exportPlans);
            var pdf = PdfCore.PdfDocument.Create(pdfOptions);
            IReadOnlyDictionary<string, string> sheetDestinations = BuildSheetDestinationMap(exportPlans);
            IReadOnlyDictionary<string, string> cellDestinations = BuildCellDestinationMap(exportPlans);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                cancellationToken.ThrowIfCancellationRequested();
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
                            cancellationToken.ThrowIfCancellationRequested();
                            if (!imagesByCellReference.ContainsKey(NormalizeCellReference(image.CellReference))) {
                                item.Image(image.Bytes, image.WidthPoints, image.HeightPoints, PdfCore.PdfAlign.Left, spacingBefore: 4, spacingAfter: 6, style: CreateConverterImageStyle(image));
                            }
                        }

                        foreach (WorksheetChartExportData chart in plan.Charts) {
                            cancellationToken.ThrowIfCancellationRequested();
                            AddWorksheetChart(item, chart, plan.SheetName, options);
                        }

                        if (plan.HasTable) {
                            IReadOnlyList<TableChunk> chunks = CreateTableChunks(plan, options, columns);
                            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                                cancellationToken.ThrowIfCancellationRequested();
                                TableChunk chunk = chunks[chunkIndex];
                                if (chunkIndex > 0) {
                                    item.PageBreak();
                                }

                                item.Table(
                                    CreatePdfRows(values, plan.ExportData.Styles, plan.ExportData.Hyperlinks, plan.ExportData.CellReferences, plan.ExportData.StructuredTables, plan.ExportData.MergedCells, imagesByCellReference, chunk.RowIndexes, chunk.StartColumn, chunk.ColumnCount, options.EmptyCellText, sheetDestinations, cellDestinations, plan.SheetName, defaultFontFamily),
                                    style: CreateTableStyle(options, plan.PageSetup, chunk.RowIndexes, chunk.HeaderRowCount, plan.ExportData.Styles, plan.ExportData.ConditionalFills, plan.ExportData.CellReferences, plan.ExportData.StructuredTables, plan.ExportData.ColumnWidths, plan.ExportData.RowHeights, chunk.StartColumn, chunk.ColumnCount));
                            }
                        }
                    }));
                });
            }

            if (exportPlans.Count == 0) {
                pdf.H1("Workbook");
                pdf.Table(new[] { new[] { "No worksheet data found." } }, style: CreateEmptyWorkbookTableStyle(options));
            }

            cancellationToken.ThrowIfCancellationRequested();
            return pdf;
        }

        /// <summary>
        /// Converts an Excel workbook to a PDF document and returns conversion diagnostics with it.
        /// </summary>
        public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this ExcelDocument document, ExcelToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            ExcelToPdfOptions operation = (options ?? new ExcelToPdfOptions()).CloneForConversion();
            operation.CancellationToken = cancellationToken;
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
        /// <example><code>byte[] pdf = workbook.ToPdfBytes();</code></example>
        public static byte[] ToPdfBytes(this ExcelDocument document, ExcelToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            return document.ToPdfDocument(options, cancellationToken).ToBytes(cancellationToken);
        }

        /// <summary>
        /// Saves an Excel workbook as a PDF file.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdf(this ExcelDocument document, string path, ExcelToPdfOptions? options = null, CancellationToken cancellationToken = default) =>
            document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

        /// <summary>
        /// Attempts to save an Excel workbook as a PDF file and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdfResult(this ExcelDocument document, string path, ExcelToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            try {
                return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken);
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(path, ex);
            }
        }

        /// <summary>
        /// Writes an Excel workbook as PDF to a stream.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdf(this ExcelDocument document, Stream stream, ExcelToPdfOptions? options = null, CancellationToken cancellationToken = default) =>
            document.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

        /// <summary>
        /// Attempts to write an Excel workbook as PDF to a stream and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdfResult(this ExcelDocument document, Stream stream, ExcelToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            try {
                return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken);
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
        }

        /// <summary>Converts synchronously, then asynchronously saves an Excel workbook PDF at the specified path.</summary>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
            this ExcelDocument document,
            string path,
            ExcelToPdfOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return await document.ToPdfDocumentResult(options, cancellationToken).SaveAsync(path, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Converts synchronously, then asynchronously saves an Excel workbook PDF to a caller-owned stream.</summary>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
            this ExcelDocument document,
            Stream stream,
            ExcelToPdfOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return await document.ToPdfDocumentResult(options, cancellationToken).SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Attempts to asynchronously save an Excel workbook as PDF at the specified path.</summary>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
            this ExcelDocument document,
            string path,
            ExcelToPdfOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                return await document.ToPdfDocumentResult(options, cancellationToken)
                    .SaveResultAsync(path, cancellationToken)
                    .ConfigureAwait(false);
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(path, ex);
            }
        }

        /// <summary>Attempts to asynchronously save an Excel workbook as PDF to a caller-owned stream.</summary>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
            this ExcelDocument document,
            Stream stream,
            ExcelToPdfOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                return await document.ToPdfDocumentResult(options, cancellationToken)
                    .SaveResultAsync(stream, cancellationToken)
                    .ConfigureAwait(false);
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
        }



    }
}
