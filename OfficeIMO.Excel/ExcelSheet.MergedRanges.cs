using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets worksheet merged ranges as reusable one-based A1 metadata.
        /// </summary>
        public IReadOnlyList<ExcelMergedRangeSnapshot> GetMergedRanges() {
            MergeCells? merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null) {
                return Array.Empty<ExcelMergedRangeSnapshot>();
            }

            var ranges = new List<ExcelMergedRangeSnapshot>();
            foreach (MergeCell merge in merges.Elements<MergeCell>()) {
                string? reference = merge.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                string normalized = NormalizeMergeRangeReference(reference!);
                if (!A1.TryParseRange(normalized, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    continue;
                }

                ranges.Add(new ExcelMergedRangeSnapshot {
                    A1Range = normalized,
                    StartRow = firstRow,
                    StartColumn = firstColumn,
                    EndRow = lastRow,
                    EndColumn = lastColumn
                });
            }

            return ranges.AsReadOnly();
        }

        private static string NormalizeMergeRangeReference(string reference) {
            int sheetSeparator = reference.LastIndexOf('!');
            string range = sheetSeparator >= 0 ? reference.Substring(sheetSeparator + 1) : reference;
            return range.Replace("$", string.Empty);
        }
    }
}
