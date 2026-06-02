using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// First-party Excel workbook to PDF conversion helpers.
    /// </summary>
    public static partial class ExcelPdfConverterExtensions {
        /// <summary>
        /// Converts an Excel workbook to a first-party OfficeIMO PDF document model.
        /// </summary>
        public static PdfCore.PdfDoc ToPdfDocument(this ExcelDocument document, ExcelPdfSaveOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            options ??= new ExcelPdfSaveOptions();
            options.Warnings.Clear();
            var pdf = PdfCore.PdfDoc.Create(CreatePdfOptions(options));
            using ExcelDocumentReader reader = document.CreateReader();
            IReadOnlyList<string> sheetNames = GetSheetNames(reader, options);
            bool hasExplicitSheetSelection = HasExplicitSheetSelection(options);
            IReadOnlyList<WorksheetPdfExportPlan> exportPlans = BuildWorksheetExportPlans(document, reader, sheetNames, options, hasExplicitSheetSelection);
            IReadOnlyDictionary<string, string> sheetDestinations = BuildSheetDestinationMap(exportPlans);
            IReadOnlyDictionary<string, string> cellDestinations = BuildCellDestinationMap(exportPlans);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
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
                                item.Image(image.Bytes, image.WidthPoints, image.HeightPoints, PdfCore.PdfAlign.Left, spacingBefore: 4, spacingAfter: 6);
                            }
                        }

                        foreach (WorksheetChartExportData chart in plan.Charts) {
                            AddWorksheetChart(item, chart);
                        }

                        if (plan.HasTable) {
                            IReadOnlyList<TableChunk> chunks = CreateTableChunks(plan, options, columns);
                            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                                TableChunk chunk = chunks[chunkIndex];
                                if (chunkIndex > 0) {
                                    item.PageBreak();
                                }

                                item.Table(
                                    CreatePdfRows(values, plan.ExportData.Styles, plan.ExportData.Hyperlinks, plan.ExportData.CellReferences, plan.ExportData.MergedCells, imagesByCellReference, chunk.RowIndexes, chunk.StartColumn, chunk.ColumnCount, options.EmptyCellText, sheetDestinations, cellDestinations, plan.SheetName),
                                    style: CreateTableStyle(options, plan.PageSetup, chunk.RowIndexes, chunk.HeaderRowCount, plan.ExportData.Styles, plan.ExportData.ConditionalFills, plan.ExportData.ColumnWidths, plan.ExportData.RowHeights, chunk.StartColumn, chunk.ColumnCount));
                            }
                        }
                    }));
                });
            }

            if (exportPlans.Count == 0) {
                pdf.H1("Workbook");
                pdf.Table(new[] { new[] { "No worksheet data found." } }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 0 });
            }

            return pdf;
        }

        /// <summary>
        /// Converts an Excel workbook to PDF bytes.
        /// </summary>
        public static byte[] SaveAsPdf(this ExcelDocument document, ExcelPdfSaveOptions? options = null) {
            return document.ToPdfDocument(options).ToBytes();
        }

        /// <summary>
        /// Saves an Excel workbook as a PDF file.
        /// </summary>
        public static void SaveAsPdf(this ExcelDocument document, string path, ExcelPdfSaveOptions? options = null) {
            document.ToPdfDocument(options).Save(path);
        }

        /// <summary>
        /// Writes an Excel workbook as PDF to a stream.
        /// </summary>
        public static void SaveAsPdf(this ExcelDocument document, Stream stream, ExcelPdfSaveOptions? options = null) {
            document.ToPdfDocument(options).Save(stream);
        }

    }
}
