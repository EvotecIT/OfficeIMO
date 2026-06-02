using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Merges the specified A1 range.
        /// </summary>
        public void MergeRange(string a1Range) {
            A1.ParseRange(a1Range);
            WriteLock(() => {
                var ws = WorksheetRoot;
                var merges = ws.GetFirstChild<MergeCells>();
                uint mergeCount = 0;

                if (merges == null) {
                    var customSheetViews = ws.GetFirstChild<CustomSheetViews>();
                    merges = new MergeCells();
                    if (customSheetViews != null) {
                        ws.InsertBefore(merges, customSheetViews);
                    } else {
                        ws.Append(merges);
                    }
                } else if (MergeCellsContainReference(merges, a1Range, out mergeCount)) {
                    return;
                }

                merges.Append(new MergeCell { Reference = a1Range });
                merges.Count = mergeCount + 1U;
                ws.Save();
            });
        }

        private static bool MergeCellsContainReference(MergeCells merges, string reference, out uint count) {
            count = 0;
            foreach (var merge in merges.Elements<MergeCell>()) {
                count++;
                if (string.Equals(merge.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Removes merge definitions that overlap the supplied A1 range.
        /// </summary>
        public void UnmergeRange(string a1Range) {
            var bounds = A1.ParseRange(a1Range);
            WriteLock(() => UnmergeRangeCore(bounds));
        }

        private void UnmergeRangeCore((int r1, int c1, int r2, int c2) bounds) {
            var merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null) return;
            if (!MergeCellsOverlap(merges, bounds)) return;

            bool changed = false;
            uint remainingCount = 0;
            foreach (var merge in merges.Elements<MergeCell>().ToList()) {
                if (merge.Reference?.Value is string reference
                    && TryParseReference(reference, out var mergeBounds)
                    && RangesOverlapInclusive(bounds, mergeBounds)) {
                    merge.Remove();
                    changed = true;
                } else {
                    remainingCount++;
                }
            }

            if (changed) {
                merges.Count = remainingCount;
                WorksheetRoot.Save();
            }
        }

        private static bool MergeCellsOverlap(MergeCells merges, (int r1, int c1, int r2, int c2) bounds) {
            foreach (var merge in merges.Elements<MergeCell>()) {
                if (merge.Reference?.Value is string reference
                    && TryParseReference(reference, out var mergeBounds)
                    && RangesOverlapInclusive(bounds, mergeBounds)) {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Writes rich inline text runs into a cell.
        /// </summary>
        public void SetRichText(int row, int column, IEnumerable<ExcelRichTextRun> runs) {
            if (runs == null) throw new ArgumentNullException(nameof(runs));
            WriteLock(() => {
                var cell = GetCell(row, column);
                var inline = new InlineString();
                foreach (var run in runs) {
                    var text = new Text(run.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve };
                    var properties = new RunProperties();
                    if (run.Bold) properties.Append(new Bold());
                    if (run.Italic) properties.Append(new Italic());
                    if (run.Underline) properties.Append(new Underline());
                    if (!string.IsNullOrWhiteSpace(run.FontColor)) properties.Append(new Color { Rgb = NormalizeHexColor(run.FontColor!) });
                    if (!string.IsNullOrWhiteSpace(run.FontName)) properties.Append(new RunFont { Val = run.FontName });
                    if (run.FontSize.HasValue) properties.Append(new FontSize { Val = run.FontSize.Value });

                    var openXmlRun = new Run();
                    if (properties.HasChildren) {
                        openXmlRun.Append(properties);
                    }
                    openXmlRun.Append(text);
                    inline.Append(openXmlRun);
                }

                cell.CellFormula = null;
                cell.CellValue = null;
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
                cell.InlineString = inline;
                ClearHeaderCache();
            });
        }

        /// <summary>
        /// Reads rich inline text runs from a cell.
        /// </summary>
        public IReadOnlyList<ExcelRichTextRun> GetRichText(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            if (cell?.InlineString == null) {
                return Array.Empty<ExcelRichTextRun>();
            }

            var runs = new List<ExcelRichTextRun>();
            foreach (var run in cell.InlineString.Elements<Run>()) {
                var properties = run.RunProperties;
                var text = run.Text?.Text ?? string.Empty;
                runs.Add(new ExcelRichTextRun(text) {
                    Bold = properties?.GetFirstChild<Bold>() != null,
                    Italic = properties?.GetFirstChild<Italic>() != null,
                    Underline = properties?.GetFirstChild<Underline>() != null,
                    FontColor = properties?.GetFirstChild<Color>()?.Rgb?.Value,
                    FontName = properties?.GetFirstChild<RunFont>()?.Val?.Value,
                    FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value
                });
            }

            return runs;
        }
    }
}
