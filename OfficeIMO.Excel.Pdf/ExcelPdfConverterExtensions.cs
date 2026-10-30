using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// First-party Excel workbook to PDF conversion helpers.
    /// </summary>
    public static partial class ExcelPdfConverterExtensions {


        private static PdfCore.PdfDocument ConvertToPdfDocument(ExcelDocument document, ExcelPdfExportOptions options) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            CancellationToken cancellationToken = options.CancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfCreateOptions pdfOptions = CreatePdfOptions(options, out bool preserveConfiguredFontSlots);
            PdfCore.PdfStandardFont defaultFontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(pdfOptions.DefaultFont);
            // Styled output needs the original serial, including negative sub-day durations
            // and precision that a DateTime carrier cannot preserve.
            using ExcelDocumentReader reader = document.CreateReader(new ExcelReaderOptions {
                TreatDatesUsingNumberFormat = !options.UseWorksheetCellStyles
            });
            IReadOnlyList<string> sheetNames = GetSheetNames(reader, options);
            bool hasExplicitSheetSelection = HasExplicitSheetSelection(options);
            IReadOnlyList<WorksheetPdfExportPlan> exportPlans = BuildWorksheetExportPlans(document, reader, sheetNames, options, hasExplicitSheetSelection, defaultFontFamily);
            cancellationToken.ThrowIfCancellationRequested();
            ReportAccountingUnderlineApproximations(exportPlans, options);
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = RegisterWorksheetFonts(pdfOptions, exportPlans, options, preserveConfiguredFontSlots, document.GetWorkbookDefaultFontInfo());
            ApplyTextFallbacks(pdfOptions, options, preserveConfiguredFontSlots, registeredFontSlots, exportPlans);
            ApplyWorksheetColumnFontMetrics(document, pdfOptions, options, exportPlans);
            var pdf = PdfCore.PdfDocument.CreateForImportedLayout(pdfOptions);
            IReadOnlyDictionary<string, string> sheetDestinations = BuildSheetDestinationMap(exportPlans);
            IReadOnlyDictionary<string, string> cellDestinations = BuildCellDestinationMap(exportPlans);
            foreach (var worksheet in exportPlans.GroupBy(plan => plan.SheetName, StringComparer.OrdinalIgnoreCase)) {
                cancellationToken.ThrowIfCancellationRequested();
                WorksheetPdfExportPlan[] worksheetPlans = worksheet.ToArray();
                WorksheetPdfExportPlan firstPlan = worksheetPlans[0];
                if (options.WorksheetLayout == ExcelPdfWorksheetLayoutMode.WorksheetCanvas) {
                    AddWorksheetCanvasPages(pdf, document, worksheetPlans, options, sheetDestinations, cellDestinations, defaultFontFamily);
                    continue;
                }

                pdf.Section(page => {
                    ApplyWorksheetPageSetup(page, firstPlan.PageSetup, options);
                    ApplyWorksheetHeaderFooter(page, firstPlan.HeaderFooter, firstPlan.SheetName, document.FilePath, options);
                    page.Content(content => content.Item(item => {
                        for (int areaIndex = 0; areaIndex < worksheetPlans.Length; areaIndex++) {
                            WorksheetPdfExportPlan plan = worksheetPlans[areaIndex];
                            object?[,] values = plan.ExportData.Values;
                            int columns = values.GetLength(1);
                            if (areaIndex > 0) item.PageBreak();
                            item.Bookmark(plan.BookmarkName);
                            if (options.IncludeSheetHeadings && areaIndex == 0) {
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
                                    SheetExportData data = SelectPageCells(plan, chunk, options);
                                    int[] pageRows = Enumerable.Range(0, chunk.RowIndexes.Count).ToArray();
                                    PdfCore.PdfTableStyle tableStyle = CreateTableStyle(options, plan.PageSetup, pageRows, chunk.HeaderRowCount, data.Styles, data.ConditionalFills, data.CellReferences, data.StructuredTables, data.ColumnWidths, data.RowHeights, 0, chunk.ColumnIndexes.Count);
                                    ApplyGeneralCellAlignments(tableStyle, data.Values, pageRows);
                                    double originalFontSize = CreateBaseTableStyle(options).FontSize ?? GetDefaultTableFontSize(options);
                                    double cellFontScale = (tableStyle.FontSize ?? originalFontSize) / originalFontSize;
                                    if (chunkIndex > 0) {
                                        item.PageBreak();
                                    }

                                    item.Table(
                                        CreatePdfRows(data.Values, data.Styles, data.Hyperlinks, data.CellReferences, data.StructuredTables, data.MergedCells, imagesByCellReference, pageRows, 0, chunk.ColumnIndexes.Count, options.EmptyCellText, sheetDestinations, cellDestinations, plan.SheetName, defaultFontFamily, fontScale: cellFontScale, dateSystem: document.DateSystem),
                                        style: tableStyle);
                                }
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
        internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(this ExcelDocument document, ExcelPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            ExcelPdfExportOptions operation = (options ?? new ExcelPdfExportOptions()).CloneForConversion();
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





















    }
}
