using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets explicit worksheet column definitions such as custom widths and hidden ranges.
        /// </summary>
        public IReadOnlyList<ExcelColumnSnapshot> GetColumnDefinitions() =>
            GetColumnDefinitions(CancellationToken.None);

        /// <summary>
        /// Gets explicit worksheet column definitions, observing cancellation while deferred data is materialized.
        /// </summary>
        /// <param name="cancellationToken">Cancels deferred materialization and column inspection.</param>
        /// <exception cref="OperationCanceledException">The operation was canceled.</exception>
        public IReadOnlyList<ExcelColumnSnapshot> GetColumnDefinitions(CancellationToken cancellationToken) =>
            _excelDocument.ExecuteReadAfterMaterializing(() => GetColumnDefinitionsCore(cancellationToken), cancellationToken);

        private IReadOnlyList<ExcelColumnSnapshot> GetColumnDefinitionsCore(CancellationToken cancellationToken) {
            Columns? columns = WorksheetRoot.GetFirstChild<Columns>();
            IReadOnlyList<double>? pendingWidths = _excelDocument.GetDirectColumnWidthsForSnapshot(this, cancellationToken);
            if (pendingWidths != null) {
                // A nondeferred direct-save candidate can hold newer auto-fit widths
                // than the worksheet XML. Project them without changing either state.
                columns = (Columns?)columns?.CloneNode(true) ?? new Columns();
                var indexes = new List<int>();
                var widths = new List<double>();
                for (int i = 0; i < pendingWidths.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    double width = NormalizeColumnWidth(pendingWidths[i]);
                    if (width <= 0D) continue;
                    indexes.Add(i + 1);
                    widths.Add(width);
                }
                if (indexes.Count > 0) SetColumnWidthsCore(columns, indexes, widths.ToArray());
            }
            if (columns == null) {
                return Array.Empty<ExcelColumnSnapshot>();
            }

            var result = new List<ExcelColumnSnapshot>();
            foreach (Column column in columns.Elements<Column>()) {
                cancellationToken.ThrowIfCancellationRequested();
                int start = checked((int)(column.Min?.Value ?? 0U));
                int end = checked((int)(column.Max?.Value ?? 0U));
                if (start <= 0 || end <= 0 || end < start) {
                    continue;
                }

                result.Add(new ExcelColumnSnapshot {
                    StartIndex = start,
                    EndIndex = end,
                    Width = column.Width?.Value,
                    Hidden = column.Hidden?.Value == true,
                    CustomWidth = column.CustomWidth?.Value == true,
                    StyleIndex = column.Style?.Value,
                    OutlineLevel = column.OutlineLevel?.Value,
                    Collapsed = column.Collapsed?.Value == true
                });
            }

            cancellationToken.ThrowIfCancellationRequested();
            return result.AsReadOnly();
        }
    }
}
