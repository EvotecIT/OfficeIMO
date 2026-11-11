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
        internal void MergeRange(string a1Range) {
            ExcelA1.ParseRange(a1Range);
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
                }

                MergeReferenceIndex index = MergeIndexFor(merges);
                bool? contains = index.Contains(merges, a1Range, out mergeCount);
                if (contains == null) {
                    // A remembered merge is no longer what the worksheet holds, so answer from the worksheet itself.
                    index = new MergeReferenceIndex(merges);
                    _mergeReferenceIndex = index;
                    contains = index.Contains(merges, a1Range, out mergeCount);
                }
                if (contains == true) {
                    return;
                }

                var merge = new MergeCell { Reference = a1Range };
                merges.Append(merge);
                merges.Count = mergeCount + 1U;
                index.Appended(merge, a1Range);
            });
        }

        private MergeReferenceIndex MergeIndexFor(MergeCells merges) {
            if (_mergeReferenceIndex is not { } index || !index.Serves(merges)) {
                index = new MergeReferenceIndex(merges);
                _mergeReferenceIndex = index;
            }
            return index;
        }

        /// <summary>Drops the merge index after merge cells are removed or their references rewritten in place.</summary>
        private void InvalidateMergeReferenceIndex() => _mergeReferenceIndex = null;

        /// <summary>
        /// Remembers the merge references of one <see cref="MergeCells"/> element so merging many ranges does not
        /// rescan every earlier merge. It serves only the element it indexed while that element's last child is the
        /// last merge it saw, so a reloaded worksheet or a merge appended elsewhere rebuilds it; removals and in-place
        /// reference rewrites invalidate it explicitly. A remembered match is confirmed against the live element.
        /// </summary>
        private sealed class MergeReferenceIndex {
            private readonly MergeCells _merges;
            private readonly Dictionary<string, MergeCell> _references = new(StringComparer.OrdinalIgnoreCase);
            private OpenXmlElement? _lastChild;
            private uint _count;

            internal MergeReferenceIndex(MergeCells merges) {
                _merges = merges;
                foreach (var merge in merges.Elements<MergeCell>()) {
                    _count++;
                    if (merge.Reference?.Value is string reference && !_references.ContainsKey(reference)) _references.Add(reference, merge);
                }
                _lastChild = merges.LastChild;
            }

            internal bool Serves(MergeCells merges) => ReferenceEquals(merges, _merges) && ReferenceEquals(merges.LastChild, _lastChild);

            /// <summary>True when the range is merged, false when it is not, or null when a remembered merge no longer
            /// holds that reference and the index must be rebuilt.</summary>
            internal bool? Contains(MergeCells merges, string reference, out uint count) {
                count = _count;
                if (!_references.TryGetValue(reference, out MergeCell? merge)) return false;
                return ReferenceEquals(merge.Parent, merges)
                    && string.Equals(merge.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase)
                    ? true
                    : null;
            }

            internal void Appended(MergeCell merge, string reference) {
                _count++;
                if (!_references.ContainsKey(reference)) _references.Add(reference, merge);
                _lastChild = merge;
            }
        }

        /// <summary>
        /// Removes merge definitions that overlap the supplied A1 range.
        /// </summary>
        internal void UnmergeRange(string a1Range) {
            var bounds = ExcelA1.ParseRange(a1Range);
            WriteLock(() => UnmergeRangeCore(bounds));
        }

        private void UnmergeRangeCore((int r1, int c1, int r2, int c2) bounds) {
            var merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null) return;
            if (!MergeCellsOverlap(merges, bounds)) return;
            InvalidateMergeReferenceIndex();

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
        internal void SetRichText(int row, int column, IEnumerable<ExcelRichTextRun> runs) {
            ArgumentNullException.ThrowIfNull(runs);
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }
            WriteLock(() => {
                EnsureDynamicArrayCellWritable(row, column);
                var inline = new InlineString();
                bool hasLineBreak = false;
                foreach (var run in runs) {
                    var text = new Text(Utilities.ExcelSanitizer.EncodeXString(run.Text ?? string.Empty)) { Space = SpaceProcessingModeValues.Preserve };
                    hasLineBreak |= text.Text.IndexOf('\n') >= 0 || text.Text.IndexOf('\r') >= 0;
                    var properties = new RunProperties();
                    if (run.BoldSpecified) properties.Append(new Bold { Val = run.Bold });
                    if (run.ItalicSpecified) properties.Append(new Italic { Val = run.Italic });
                    if (run.UnderlineStyle.HasValue) {
                        properties.Append(new Underline { Val = run.UnderlineStyle.Value.ToOpenXml() });
                    } else if (run.UnderlineSpecified) {
                        properties.Append(new Underline { Val = run.Underline ? UnderlineValues.Single : UnderlineValues.None });
                    }
                    if (run.StrikethroughSpecified) properties.Append(new Strike { Val = run.Strikethrough });
                    if (run.Color.HasValue) properties.Append(ExcelColor.ToSpreadsheetColor(run.Color.Value, _excelDocument.ColorTheme));
                    if (!string.IsNullOrWhiteSpace(run.FontFamily)) properties.Append(new RunFont { Val = run.FontFamily });
                    if (run.FontSize.HasValue) properties.Append(new FontSize { Val = run.FontSize.Value.ToPoints() });
                    ExcelRichTextRun.AppendFontMetadata(properties, run);

                    var openXmlRun = new Run();
                    if (properties.HasChildren) {
                        openXmlRun.Append(properties);
                    }
                    openXmlRun.Append(text);
                    inline.Append(openXmlRun);
                }

                PrepareTableHeaderRichText(row, column, inline);
                var cell = GetCell(row, column);
                ClearCellValueMetadata(cell);
                cell.CellFormula = null;
                cell.CellValue = null;
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
                cell.InlineString = inline;
                if (hasLineBreak) ApplyWrapText(cell);
                _excelDocument.MarkFormulaInputMutation();
                ClearHeaderCache();
            });
        }

        /// <summary>
        /// Reads rich text runs from an inline-string or shared-string cell.
        /// </summary>
        internal IReadOnlyList<ExcelRichTextRun> GetRichText(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            IEnumerable<Run>? openXmlRuns = cell?.InlineString?.Elements<Run>();
            if (openXmlRuns == null
                && cell?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                && int.TryParse(cell.CellValue?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out int sharedStringIndex)
                && sharedStringIndex >= 0) {
                SharedStringItem? sharedStringItem = _spreadSheetDocument.WorkbookPart?
                    .SharedStringTablePart?
                    .SharedStringTable?
                    .Elements<SharedStringItem>()
                    .ElementAtOrDefault(sharedStringIndex);
                openXmlRuns = sharedStringItem?.Elements<Run>();
            }

            if (openXmlRuns == null) {
                return Array.Empty<ExcelRichTextRun>();
            }

            var runs = new List<ExcelRichTextRun>();
            foreach (var run in openXmlRuns) {
                var properties = run.RunProperties;
                var text = Utilities.ExcelSanitizer.DecodeXString(run.Text?.Text ?? string.Empty);
                runs.Add(ExcelRichTextRun.FromOpenXml(text, properties));
            }

            return runs;
        }
    }
}
