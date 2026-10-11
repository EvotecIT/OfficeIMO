using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class PdfExcelTableConverterExtensions {
        /// <summary>
        /// Applies canonical detected proportions to the existing destination width budget.
        /// A single scale respects Excel's maximum width without independently clipping column ratios.
        /// </summary>
        private static bool ApplyColumnProportions(
            ExcelSheet sheet,
            IReadOnlyList<PdfCore.PdfLogicalTableColumn> columns,
            int columnCount,
            PdfTablesToExcelOptions options) {
            if (!options.PreserveColumnProportions || columnCount == 0 || columns.Count != columnCount) return false;

            bool horizontal = PdfCore.PdfTableColumnGeometry.HasHorizontalProgression(
                columns, static column => column.VisualBounds, true, options.CancellationToken);
            var lengths = new double[columnCount];
            double maximumLength = 0D;
            for (int index = 0; index < columnCount; index++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                double length = PdfCore.PdfTableColumnGeometry.GetProgressionLength(columns[index], horizontal);
                if (!IsPositiveFiniteWidth(length)) return false;
                lengths[index] = length;
                maximumLength = Math.Max(maximumLength, length);
            }

            double totalWeight = 0D;
            for (int index = 0; index < lengths.Length; index++) {
                lengths[index] /= maximumLength;
                if (!IsPositiveFiniteWidth(lengths[index])) return false;
                totalWeight += lengths[index];
            }

            // New import sheets use Excel's application default (8.43 character units)
            // unless the worksheet owner supplies an explicit default width.
            double defaultWidth = sheet.DefaultColumnWidth ?? 8.43D;
            if (!IsPositiveFiniteWidth(defaultWidth)) return false;
            var destinationWidths = Enumerable.Repeat(defaultWidth, columnCount).ToArray();
            foreach (ExcelColumnSnapshot definition in sheet.GetColumnDefinitions(options.CancellationToken)) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (!definition.CustomWidth || !definition.Width.HasValue) continue;
                double width = definition.Width.Value;
                if (!IsPositiveFiniteWidth(width)) return false;
                int first = Math.Max(1, definition.StartIndex);
                int last = Math.Min(columnCount, definition.EndIndex);
                for (int index = first; index <= last; index++) destinationWidths[index - 1] = width;
            }

            double budget = destinationWidths.Sum();
            if (!IsPositiveFiniteWidth(budget)) return false;
            double scale = Math.Min(budget / totalWeight, 255D);
            var widths = new Dictionary<int, double>(columnCount);
            for (int index = 0; index < columnCount; index++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                double width = scale * lengths[index];
                if (!IsPositiveFiniteWidth(width)) return false;
                widths.Add(index + 1, width);
            }

            options.CancellationToken.ThrowIfCancellationRequested();
            sheet.SetColumnWidths(widths);
            return true;
        }

        private static bool IsPositiveFiniteWidth(double width) =>
            width > 0D && !double.IsNaN(width) && !double.IsInfinity(width);
    }
}
