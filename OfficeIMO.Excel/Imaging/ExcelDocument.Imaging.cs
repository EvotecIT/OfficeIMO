using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Exports workbook sheets as supported raster formats or SVG images.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, ExcelImageExportOptions? options = null) {
            var results = new List<OfficeImageExportResult>();
            ExportImages(format, results.Add, options);
            return results.AsReadOnly();
        }

        /// <summary>Streams workbook images to a consumer without retaining earlier payloads.</summary>
        public void ExportImages(
            OfficeImageExportFormat format,
            OfficeImageExportConsumer consumer,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (consumer == null) throw new ArgumentNullException(nameof(consumer));
            ExcelImageExportOptions resolved = NormalizeWorkbookOptions(options);
            var plan = new List<ExcelWorksheetImageExportPlan>();
            var encodingBudget = new OfficeImageExportEncodingBudget(resolved.MaximumTotalEncodedBytes);
            OfficeImageExportBatchProcessor.RunWithPreflight(
                resolved,
                operationCancellationToken => {
                    operationCancellationToken.ThrowIfCancellationRequested();
                    HashSet<string>? selected = resolved.SheetNames == null
                        ? null
                        : new HashSet<string>(resolved.SheetNames, StringComparer.OrdinalIgnoreCase);
                    if (resolved.SheetNames != null) {
                        ValidateRequestedSheetNames(resolved.SheetNames);
                    }

                    int expectedOutputCount = 0;
                    foreach (ExcelSheet sheet in Sheets) {
                        operationCancellationToken.ThrowIfCancellationRequested();
                        if (selected != null && !selected.Contains(sheet.Name)) continue;
                        if (selected == null && sheet.Hidden && !resolved.IncludeHiddenSheets) continue;

                        ExcelImageExportOptions sheetOptions = resolved.Clone();
                        int outputCount = sheet.GetImageExportResultCount(sheetOptions, operationCancellationToken);
                        expectedOutputCount = checked(expectedOutputCount + outputCount);
                        plan.Add(new ExcelWorksheetImageExportPlan(sheet, sheetOptions));
                    }

                    return expectedOutputCount;
                },
                (accept, operationCancellationToken) => {
                    foreach (ExcelWorksheetImageExportPlan item in plan) {
                        operationCancellationToken.ThrowIfCancellationRequested();
                        item.Sheet.ExportImagesCore(format, accept, item.Options, operationCancellationToken, encodingBudget);
                    }
                },
                consumer,
                cancellationToken);
        }

        private static ExcelImageExportOptions NormalizeWorkbookOptions(ExcelImageExportOptions? options) {
            ExcelImageExportOptions resolved = options?.Clone() ?? new ExcelImageExportOptions();
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

        private sealed class ExcelWorksheetImageExportPlan {
            internal ExcelWorksheetImageExportPlan(
                ExcelSheet sheet,
                ExcelImageExportOptions options) {
                Sheet = sheet;
                Options = options;
            }

            internal ExcelSheet Sheet { get; }

            internal ExcelImageExportOptions Options { get; }
        }
    }
}
