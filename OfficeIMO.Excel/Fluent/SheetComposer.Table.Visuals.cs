namespace OfficeIMO.Excel.Fluent {
    public sealed partial class SheetComposer {
        private string CompleteTable(
            string range,
            IReadOnlyList<string> paths,
            List<string> headers,
            int headerRow,
            int lastRow,
            ExcelTableStyle style,
            bool freezeHeaderRow,
            Action<TableVisualOptions>? visuals) {
            _row = lastRow + 1;
            var visualOptions = new TableVisualOptions {
                FreezeHeaderRow = freezeHeaderRow
            };
            visuals?.Invoke(visualOptions);
            Sheet.SetTableStyle(
                range,
                style,
                visualOptions.ShowFirstColumn,
                visualOptions.ShowLastColumn,
                visualOptions.ShowRowStripes,
                visualOptions.ShowColumnStripes);
            if (visualOptions.FreezeHeaderRow) Sheet.Freeze(topRows: headerRow, leftCols: 0);

            const int startColumn = 1;
            for (int index = 0; index < headers.Count; index++) {
                string header = headers[index];
                string columnRange = $"{ColumnLetter(startColumn + index)}{headerRow + 1}:{ColumnLetter(startColumn + index)}{lastRow}";
                var columnStyle = new ExcelColumnStyleByHeaderBuilder(
                    Sheet,
                    startColumn + index,
                    headerRow + 1,
                    lastRow);

                if (visualOptions.NumericColumnFormats.TryGetValue(header, out string? format)) {
                    columnStyle.NumberFormat(format);
                } else if (visualOptions.NumericColumnDecimals.TryGetValue(header, out int decimals)) {
                    columnStyle.Number(decimals);
                }

                if (visualOptions.DataBars.TryGetValue(header, out var color))
                    Sheet.AddConditionalDataBar(columnRange, color);

                if (visualOptions.IconSets.TryGetValue(header, out var iconOptions))
                    Sheet.AddConditionalIconSet(columnRange, iconOptions.IconSet, iconOptions.ShowValue, iconOptions.ReverseOrder, iconOptions.PercentThresholds, iconOptions.NumberThresholds);
                else if (visualOptions.IconSetColumns.Contains(header))
                    Sheet.AddConditionalIconSet(columnRange);

                if (visualOptions.TextBackgrounds.TryGetValue(header, out var backgroundMap)) {
                    columnStyle.BackgroundByTextMap(backgroundMap);
                }

                if (visualOptions.BoldByText.TryGetValue(header, out var boldValues)) {
                    columnStyle.BoldByTextSet(boldValues);
                }
            }

            if (visualOptions.AutoFormatDynamicCollections) {
                for (int index = 0; index < paths.Count; index++) {
                    if (!paths[index].Contains('.')) continue;
                    var columnStyle = new ExcelColumnStyleByHeaderBuilder(
                        Sheet,
                        startColumn + index,
                        headerRow + 1,
                        lastRow);
                    columnStyle.Number(visualOptions.AutoFormatDecimals);
                    string columnRange = $"{ColumnLetter(startColumn + index)}{headerRow + 1}:{ColumnLetter(startColumn + index)}{lastRow}";
                    Sheet.AddConditionalDataBar(columnRange, visualOptions.AutoFormatDataBarColor);
                }
            }

            Spacer();
            return range;
        }

        private static void EnsureUniqueTableHeaders(List<string> headers) {
            var usedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < headers.Count; index++) {
                string baseName = string.IsNullOrWhiteSpace(headers[index]) ? $"Column{index + 1}" : headers[index];
                string candidate = baseName;
                int suffix = 2;
                while (!usedHeaders.Add(candidate)) {
                    candidate = $"{baseName} ({suffix++})";
                }
                headers[index] = candidate;
            }
        }
    }
}
