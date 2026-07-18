using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Exports workbook sheets as supported raster formats or SVG images.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, ExcelWorkbookImageExportOptions? options = null) {
            var results = new List<OfficeImageExportResult>();
            ExportImages(format, results.Add, options);
            return results.AsReadOnly();
        }

        /// <summary>Streams workbook images to a consumer without retaining earlier payloads.</summary>
        public void ExportImages(
            OfficeImageExportFormat format,
            OfficeImageExportConsumer consumer,
            ExcelWorkbookImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (consumer == null) throw new ArgumentNullException(nameof(consumer));
            ExcelWorkbookImageExportOptions resolved = NormalizeWorkbookOptions(options);
            HashSet<string>? selected = resolved.SheetNames == null
                ? null
                : new HashSet<string>(resolved.SheetNames, StringComparer.OrdinalIgnoreCase);
            if (resolved.SheetNames != null) {
                ValidateRequestedSheetNames(resolved.SheetNames);
            }
            OfficeImageExportConsumer accept =
                OfficeImageExportBatchProcessor.CreateGuardedConsumer(
                    resolved,
                    consumer,
                    cancellationToken);

            foreach (ExcelSheet sheet in Sheets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (selected != null && !selected.Contains(sheet.Name)) {
                    continue;
                }

                if (selected == null && sheet.Hidden && !resolved.IncludeHiddenSheets) {
                    continue;
                }

                ExcelWorksheetImageExportOptions sheetOptions =
                    resolved.CopyExcelOptionsTo(new ExcelWorksheetImageExportOptions());
                sheetOptions.HeaderFooterDateTime = resolved.HeaderFooterDateTime;
                sheetOptions.UsePrintArea = resolved.UseWorksheetPrintAreas;
                sheetOptions.SplitByManualPageBreaks = resolved.SplitWorksheetsByManualPageBreaks;
                sheet.ExportImages(format, accept, sheetOptions, cancellationToken);
            }
        }

        /// <summary>
        /// Saves workbook sheets as PNG files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, ExcelWorkbookImageExportOptions? options = null) =>
            new ExcelWorkbookImageExportBuilder(this, options).AsPng().Save(folderPath);

        /// <summary>
        /// Saves workbook sheets as image files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, OfficeImageExportFormat format, ExcelWorkbookImageExportOptions? options = null) {
            ExcelWorkbookImageExportBuilder builder = new ExcelWorkbookImageExportBuilder(this, options);
            builder.As(format);
            return builder.Save(folderPath);
        }

        /// <summary>
        /// Asynchronously saves workbook sheets as PNG files in a folder.
        /// </summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
            string folderPath,
            ExcelWorkbookImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelWorkbookImageExportBuilder(this, options).AsPng().SaveAsync(folderPath, cancellationToken);

        /// <summary>
        /// Asynchronously saves workbook sheets as image files in a folder.
        /// </summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
            string folderPath,
            OfficeImageExportFormat format,
            ExcelWorkbookImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            ExcelWorkbookImageExportBuilder builder = new ExcelWorkbookImageExportBuilder(this, options);
            builder.As(format);
            return builder.SaveAsync(folderPath, cancellationToken);
        }

        private static ExcelWorkbookImageExportOptions NormalizeWorkbookOptions(ExcelWorkbookImageExportOptions? options) {
            ExcelWorkbookImageExportOptions resolved = options?.CloneWorkbook() ?? new ExcelWorkbookImageExportOptions();
            resolved.ConditionalFormattingDate ??= DateTime.Today;
            resolved.HeaderFooterDateTime ??= DateTime.Now;
            resolved.Validate();

            return resolved;
        }

        private void ValidateRequestedSheetNames(IReadOnlyList<string> sheetNames) {
            if (sheetNames.Count == 0) {
                return;
            }

            var available = new HashSet<string>(Sheets.Select(sheet => sheet.Name), StringComparer.OrdinalIgnoreCase);
            var missing = sheetNames
                .Where(sheetName => !available.Contains(sheetName))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (missing.Length > 0) {
                throw new ArgumentException(
                    "Workbook image export requested worksheet names that do not exist: " + string.Join(", ", missing) + ".",
                    nameof(sheetNames));
            }
        }
    }
}
