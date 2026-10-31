using System.IO;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        /// <summary>
        /// A pivot table written without its output cells prints nothing. Render such pivots
        /// from their cache in an in-memory copy, so the caller's workbook is never changed,
        /// and report every pivot that cannot be rendered instead of dropping it silently.
        /// </summary>
        private static ExcelDocument? CreatePivotRenderedCopy(ExcelDocument document, ExcelPdfExportOptions options) {
            if (options.UseBoundedWorksheetRead && options.MaxRowsPerSheet.HasValue) return null;
            var pending = new List<(string Sheet, string Pivot)>();
            foreach (ExcelSheet sheet in document.Sheets) {
                foreach (string pivot in sheet.GetPivotTablesWithoutRenderedCells()) pending.Add((sheet.Name, pivot));
            }

            if (pending.Count == 0) return null;
            options.CancellationToken.ThrowIfCancellationRequested();
            ExcelDocument copy;
            try {
                var stream = new MemoryStream();
                document.Save(stream);
                stream.Position = 0;
                copy = ExcelDocument.Load(stream);
            } catch (Exception ex) when (ex is IOException or InvalidOperationException or NotSupportedException) {
                foreach ((string sheet, string pivot) in pending) ReportPivotNotRendered(options, sheet, pivot, ex.Message);
                return null;
            }

            int rendered = 0;
            foreach ((string sheetName, string pivot) in pending) {
                options.CancellationToken.ThrowIfCancellationRequested();
                ExcelSheet? sheet = GetWorkbookSheet(copy, sheetName);
                try {
                    if (sheet == null) throw new InvalidOperationException("The worksheet was not found in the rendering copy.");
                    sheet.MaterializePivotTable(pivot, cancellationToken: options.CancellationToken);
                    rendered++;
                } catch (Exception ex) when (ex is NotSupportedException or InvalidOperationException or ArgumentException) {
                    ReportPivotNotRendered(options, sheetName, pivot, ex.Message);
                }
            }

            if (rendered > 0) return copy;
            copy.Dispose();
            return null;
        }

        private static void ReportPivotNotRendered(ExcelPdfExportOptions options, string sheetName, string pivotName, string reason) =>
            AddWarning(
                options,
                sheetName,
                "pivot-table",
                $"Pivot table '{pivotName}' stores no rendered cells and could not be calculated from its cache, so it was not exported: {reason}",
                OfficeConversionLossKind.Omission);
    }
}
