using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets worksheet merged ranges as reusable one-based A1 metadata.
        /// </summary>
        public IReadOnlyList<ExcelMergedRangeSnapshot> GetMergedRanges() => GetMergedRanges(int.MaxValue);

        /// <summary>
        /// Gets worksheet merged ranges while rejecting collections beyond the configured inspection limit.
        /// </summary>
        /// <param name="maximumRanges">Maximum merge records inspected.</param>
        public IReadOnlyList<ExcelMergedRangeSnapshot> GetMergedRanges(int maximumRanges) {
            if (maximumRanges <= 0) throw new ArgumentOutOfRangeException(nameof(maximumRanges));
            MergeCells? merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null) {
                return Array.Empty<ExcelMergedRangeSnapshot>();
            }

            var ranges = new List<ExcelMergedRangeSnapshot>();
            int inspected = 0;
            foreach (MergeCell merge in merges.Elements<MergeCell>()) {
                if (++inspected > maximumRanges) {
                    throw new InvalidOperationException("The merged-range count exceeds the configured inspection limit.");
                }
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
