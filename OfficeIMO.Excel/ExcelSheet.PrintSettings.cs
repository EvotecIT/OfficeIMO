using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Readable worksheet page setup values relevant to print/export pipelines.
    /// </summary>
    public sealed class ExcelSheetPageSetup {
        /// <summary>
        /// Worksheet page orientation when present.
        /// </summary>
        public ExcelPageOrientation? Orientation { get; internal set; }

        /// <summary>
        /// Worksheet print margins in inches when present.
        /// </summary>
        public ExcelSheetPageMargins? Margins { get; internal set; }

        /// <summary>
        /// Number of pages to fit horizontally, when configured.
        /// </summary>
        public uint? FitToWidth { get; internal set; }

        /// <summary>
        /// Number of pages to fit vertically, when configured.
        /// </summary>
        public uint? FitToHeight { get; internal set; }

        /// <summary>
        /// Manual worksheet print scale percentage, when configured.
        /// </summary>
        public uint? Scale { get; internal set; }

        /// <summary>
        /// OpenXML worksheet paper-size code, when configured.
        /// </summary>
        public uint? PaperSizeCode { get; internal set; }

        /// <summary>
        /// Known worksheet paper size, when the configured code maps to a supported OfficeIMO value.
        /// </summary>
        public ExcelPaperSize? PaperSize { get; internal set; }

        /// <summary>
        /// Worksheet page order for multi-page print/export output, when configured.
        /// </summary>
        public ExcelPageOrder? PageOrder { get; internal set; }
    }

    /// <summary>
    /// Worksheet page margins in inches.
    /// </summary>
    public sealed class ExcelSheetPageMargins {
        /// <summary>Left margin in inches.</summary>
        public double Left { get; internal set; }
        /// <summary>Right margin in inches.</summary>
        public double Right { get; internal set; }
        /// <summary>Top margin in inches.</summary>
        public double Top { get; internal set; }
        /// <summary>Bottom margin in inches.</summary>
        public double Bottom { get; internal set; }
        /// <summary>Header margin in inches.</summary>
        public double Header { get; internal set; }
        /// <summary>Footer margin in inches.</summary>
        public double Footer { get; internal set; }
    }

    /// <summary>
    /// Worksheet print options that are stored in the Open XML printOptions element.
    /// </summary>
    public sealed class ExcelSheetPrintOptions {
        /// <summary>Gets whether worksheet gridlines are printed, when configured.</summary>
        public bool? PrintGridLines { get; internal set; }

        /// <summary>Gets whether row and column headings are printed, when configured.</summary>
        public bool? PrintHeadings { get; internal set; }

        /// <summary>Gets whether the sheet is centered horizontally when printed, when configured.</summary>
        public bool? HorizontalCentered { get; internal set; }

        /// <summary>Gets whether the sheet is centered vertically when printed, when configured.</summary>
        public bool? VerticalCentered { get; internal set; }
    }

    /// <summary>
    /// Worksheet print title rows and columns.
    /// </summary>
    public sealed class ExcelPrintTitles {
        /// <summary>First repeated row, one-based.</summary>
        public int? FirstRow { get; internal set; }
        /// <summary>Last repeated row, one-based.</summary>
        public int? LastRow { get; internal set; }
        /// <summary>First repeated column, one-based.</summary>
        public int? FirstColumn { get; internal set; }
        /// <summary>Last repeated column, one-based.</summary>
        public int? LastColumn { get; internal set; }
        /// <summary>Whether repeated rows are configured.</summary>
        public bool HasRows => FirstRow.HasValue && LastRow.HasValue;
        /// <summary>Whether repeated columns are configured.</summary>
        public bool HasColumns => FirstColumn.HasValue && LastColumn.HasValue;
    }

    public partial class ExcelSheet {
        /// <summary>
        /// Reads worksheet page setup values used by print/export pipelines.
        /// </summary>
        public ExcelSheetPageSetup GetPageSetup() {
            var result = new ExcelSheetPageSetup();
            PageSetup? pageSetup = WorksheetRoot.GetFirstChild<PageSetup>();
            if (pageSetup?.Orientation?.Value == OrientationValues.Landscape) {
                result.Orientation = ExcelPageOrientation.Landscape;
            } else if (pageSetup?.Orientation?.Value == OrientationValues.Portrait) {
                result.Orientation = ExcelPageOrientation.Portrait;
            }

            if (pageSetup != null) {
                result.FitToWidth = pageSetup.FitToWidth?.Value;
                result.FitToHeight = pageSetup.FitToHeight?.Value;
                result.Scale = pageSetup.Scale?.Value;
                result.PaperSizeCode = pageSetup.PaperSize?.Value;
                if (pageSetup.PaperSize?.Value is uint paperSizeCode && Enum.IsDefined(typeof(ExcelPaperSize), paperSizeCode)) {
                    result.PaperSize = (ExcelPaperSize)paperSizeCode;
                }

                if (pageSetup.PageOrder?.Value == PageOrderValues.OverThenDown) {
                    result.PageOrder = ExcelPageOrder.OverThenDown;
                } else if (pageSetup.PageOrder?.Value == PageOrderValues.DownThenOver) {
                    result.PageOrder = ExcelPageOrder.DownThenOver;
                }
            }

            PageMargins? margins = WorksheetRoot.GetFirstChild<PageMargins>();
            if (margins != null) {
                result.Margins = new ExcelSheetPageMargins {
                    Left = margins.Left?.Value ?? 0D,
                    Right = margins.Right?.Value ?? 0D,
                    Top = margins.Top?.Value ?? 0D,
                    Bottom = margins.Bottom?.Value ?? 0D,
                    Header = margins.Header?.Value ?? 0D,
                    Footer = margins.Footer?.Value ?? 0D
                };
            }

            return result;
        }

        /// <summary>
        /// Reads worksheet print options.
        /// </summary>
        public ExcelSheetPrintOptions GetPrintOptions() {
            PrintOptions? printOptions = WorksheetRoot.GetFirstChild<PrintOptions>();
            return new ExcelSheetPrintOptions {
                PrintGridLines = printOptions?.GridLines?.Value,
                PrintHeadings = printOptions?.Headings?.Value,
                HorizontalCentered = printOptions?.HorizontalCentered?.Value,
                VerticalCentered = printOptions?.VerticalCentered?.Value
            };
        }

        /// <summary>
        /// Sets worksheet print options. Null values leave the current option unchanged.
        /// </summary>
        public void SetPrintOptions(
            bool? printGridLines = null,
            bool? printHeadings = null,
            bool? horizontalCentered = null,
            bool? verticalCentered = null,
            bool save = true) {
            if (!printGridLines.HasValue
                && !printHeadings.HasValue
                && !horizontalCentered.HasValue
                && !verticalCentered.HasValue) {
                return;
            }

            WriteLock(() => {
                PrintOptions printOptions = GetOrCreatePrintOptions();
                if (printGridLines.HasValue) {
                    printOptions.GridLines = printGridLines.Value;
                    printOptions.GridLinesSet = true;
                }

                if (printHeadings.HasValue) {
                    printOptions.Headings = printHeadings.Value;
                }

                if (horizontalCentered.HasValue) {
                    printOptions.HorizontalCentered = horizontalCentered.Value;
                }

                if (verticalCentered.HasValue) {
                    printOptions.VerticalCentered = verticalCentered.Value;
                }

                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Gets the worksheet print area range, or null when no print area is configured.
        /// </summary>
        public string? GetPrintArea() {
            return _excelDocument.GetNamedRange("_xlnm.Print_Area", this);
        }

        /// <summary>
        /// Gets worksheet print title rows and columns.
        /// </summary>
        public ExcelPrintTitles GetPrintTitles() {
            var titles = new ExcelPrintTitles();
            string? definedName = _excelDocument.GetNamedRange("_xlnm.Print_Titles", this);
            if (string.IsNullOrWhiteSpace(definedName)) {
                return titles;
            }

            foreach (string part in SplitDefinedNameParts(definedName!)) {
                string reference = StripSheetPrefix(part).Replace("$", string.Empty);
                int separator = reference.IndexOf(':');
                if (separator <= 0 || separator >= reference.Length - 1) {
                    continue;
                }

                string start = reference.Substring(0, separator);
                string end = reference.Substring(separator + 1);
                if (int.TryParse(start, out int firstRow) &&
                    int.TryParse(end, out int lastRow) &&
                    firstRow > 0 &&
                    lastRow >= firstRow) {
                    titles.FirstRow = firstRow;
                    titles.LastRow = lastRow;
                    continue;
                }

                int firstColumn = A1.ColumnLettersToIndex(start);
                int lastColumn = A1.ColumnLettersToIndex(end);
                if (firstColumn > 0 && lastColumn >= firstColumn) {
                    titles.FirstColumn = firstColumn;
                    titles.LastColumn = lastColumn;
                }
            }

            return titles;
        }

        /// <summary>
        /// Adds a manual worksheet row page break after the specified one-based row.
        /// </summary>
        public void AddManualRowPageBreak(int row, bool save = true) {
            if (row <= 0 || row > A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(row), "Row page break must be between 1 and the Excel row limit.");
            }

            WriteLock(() => {
                RowBreaks breaks = GetOrCreateRowBreaks();
                AddManualPageBreak(breaks, (uint)row, 16383U);
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Adds a manual worksheet column page break after the specified one-based column.
        /// </summary>
        public void AddManualColumnPageBreak(int column, bool save = true) {
            if (column <= 0 || column > A1.MaxColumns) {
                throw new ArgumentOutOfRangeException(nameof(column), "Column page break must be between 1 and the Excel column limit.");
            }

            WriteLock(() => {
                ColumnBreaks breaks = GetOrCreateColumnBreaks();
                AddManualPageBreak(breaks, (uint)column, 1048575U);
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Gets one-based worksheet rows that have a manual page break after them.
        /// </summary>
        public IReadOnlyList<int> GetManualRowPageBreaks() {
            return GetManualPageBreaks(WorksheetRoot.GetFirstChild<RowBreaks>());
        }

        /// <summary>
        /// Gets one-based worksheet columns that have a manual page break after them.
        /// </summary>
        public IReadOnlyList<int> GetManualColumnPageBreaks() {
            return GetManualPageBreaks(WorksheetRoot.GetFirstChild<ColumnBreaks>());
        }

        /// <summary>
        /// Removes a manual worksheet row page break.
        /// </summary>
        public bool RemoveManualRowPageBreak(int row, bool save = true) {
            if (row <= 0 || row > A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(row), "Row page break must be between 1 and the Excel row limit.");
            }

            bool changed = false;
            WriteLock(() => {
                changed = RemoveManualPageBreak(WorksheetRoot.GetFirstChild<RowBreaks>(), (uint)row);
                if (changed && save) {
                    WorksheetRoot.Save();
                }
            });
            return changed;
        }

        /// <summary>
        /// Removes a manual worksheet column page break.
        /// </summary>
        public bool RemoveManualColumnPageBreak(int column, bool save = true) {
            if (column <= 0 || column > A1.MaxColumns) {
                throw new ArgumentOutOfRangeException(nameof(column), "Column page break must be between 1 and the Excel column limit.");
            }

            bool changed = false;
            WriteLock(() => {
                changed = RemoveManualPageBreak(WorksheetRoot.GetFirstChild<ColumnBreaks>(), (uint)column);
                if (changed && save) {
                    WorksheetRoot.Save();
                }
            });
            return changed;
        }

        /// <summary>
        /// Clears all manual worksheet row page breaks.
        /// </summary>
        public bool ClearManualRowPageBreaks(bool save = true) {
            bool changed = false;
            WriteLock(() => {
                changed = ClearManualPageBreaks(WorksheetRoot.GetFirstChild<RowBreaks>());
                if (changed && save) {
                    WorksheetRoot.Save();
                }
            });
            return changed;
        }

        /// <summary>
        /// Clears all manual worksheet column page breaks.
        /// </summary>
        public bool ClearManualColumnPageBreaks(bool save = true) {
            bool changed = false;
            WriteLock(() => {
                changed = ClearManualPageBreaks(WorksheetRoot.GetFirstChild<ColumnBreaks>());
                if (changed && save) {
                    WorksheetRoot.Save();
                }
            });
            return changed;
        }

        /// <summary>
        /// Clears all manual worksheet row and column page breaks.
        /// </summary>
        public bool ClearManualPageBreaks(bool save = true) {
            bool changed = false;
            WriteLock(() => {
                changed = ClearManualPageBreaks(WorksheetRoot.GetFirstChild<RowBreaks>());
                changed |= ClearManualPageBreaks(WorksheetRoot.GetFirstChild<ColumnBreaks>());
                if (changed && save) {
                    WorksheetRoot.Save();
                }
            });
            return changed;
        }

        private static IEnumerable<string> SplitDefinedNameParts(string text) {
            var parts = new List<string>();
            int start = 0;
            bool inQuote = false;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '\'') {
                    if (inQuote && i + 1 < text.Length && text[i + 1] == '\'') {
                        i++;
                    } else {
                        inQuote = !inQuote;
                    }
                } else if (ch == ',' && !inQuote) {
                    AddPart(text, start, i - start, parts);
                    start = i + 1;
                }
            }

            AddPart(text, start, text.Length - start, parts);
            return parts;
        }

        private static List<string> SplitDefinedNameParts(string text, int maximumParts, out bool exceeded) {
            if (maximumParts <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumParts));
            }

            var parts = new List<string>(Math.Min(maximumParts, 16));
            int start = 0;
            bool inQuote = false;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '\'') {
                    if (inQuote && i + 1 < text.Length && text[i + 1] == '\'') {
                        i++;
                    } else {
                        inQuote = !inQuote;
                    }
                } else if (ch == ',' && !inQuote) {
                    if (!TryAddBoundedDefinedNamePart(text, start, i - start, maximumParts, parts)) {
                        exceeded = true;
                        return parts;
                    }

                    start = i + 1;
                }
            }

            exceeded = !TryAddBoundedDefinedNamePart(text, start, text.Length - start, maximumParts, parts);
            return parts;
        }

        private static bool TryAddBoundedDefinedNamePart(string text, int start, int length, int maximumParts, List<string> parts) {
            int end = start + length;
            while (start < end && char.IsWhiteSpace(text[start])) {
                start++;
            }

            while (end > start && char.IsWhiteSpace(text[end - 1])) {
                end--;
            }

            if (start >= end) {
                return true;
            }

            if (parts.Count >= maximumParts) {
                return false;
            }

            parts.Add(text.Substring(start, end - start));
            return true;
        }

        private static void AddPart(string text, int start, int length, List<string> parts) {
            string part = text.Substring(start, length).Trim();
            if (part.Length > 0) {
                parts.Add(part);
            }
        }

        private static string StripSheetPrefix(string reference) {
            int separator = reference.LastIndexOf('!');
            return separator >= 0 ? reference.Substring(separator + 1) : reference;
        }

        private RowBreaks GetOrCreateRowBreaks() {
            RowBreaks? breaks = WorksheetRoot.GetFirstChild<RowBreaks>();
            if (breaks != null) {
                return breaks;
            }

            breaks = new RowBreaks();
            ColumnBreaks? columnBreaks = WorksheetRoot.GetFirstChild<ColumnBreaks>();
            if (columnBreaks != null) {
                WorksheetRoot.InsertBefore(breaks, columnBreaks);
            } else {
                InsertAfterPrintSetupElement(breaks);
            }

            return breaks;
        }

        private ColumnBreaks GetOrCreateColumnBreaks() {
            ColumnBreaks? breaks = WorksheetRoot.GetFirstChild<ColumnBreaks>();
            if (breaks != null) {
                return breaks;
            }

            breaks = new ColumnBreaks();
            InsertAfterPrintSetupElement(breaks);
            return breaks;
        }

        private PrintOptions GetOrCreatePrintOptions() {
            PrintOptions? printOptions = WorksheetRoot.GetFirstChild<PrintOptions>();
            if (printOptions != null) {
                return printOptions;
            }

            printOptions = new PrintOptions();
            PageMargins? pageMargins = WorksheetRoot.GetFirstChild<PageMargins>();
            if (pageMargins != null) {
                WorksheetRoot.InsertBefore(printOptions, pageMargins);
                return printOptions;
            }

            Hyperlinks? hyperlinks = WorksheetRoot.GetFirstChild<Hyperlinks>();
            if (hyperlinks != null) {
                WorksheetRoot.InsertAfter(printOptions, hyperlinks);
                return printOptions;
            }

            WorksheetRoot.Append(printOptions);
            return printOptions;
        }

        private void InsertAfterPrintSetupElement(OpenXmlElement element) {
            OpenXmlElement? previous = WorksheetRoot.GetFirstChild<RowBreaks>();
            previous ??= WorksheetRoot.GetFirstChild<HeaderFooter>();
            previous ??= WorksheetRoot.GetFirstChild<PageSetup>();
            previous ??= WorksheetRoot.GetFirstChild<PageMargins>();
            previous ??= WorksheetRoot.GetFirstChild<PrintOptions>();

            if (previous != null) {
                WorksheetRoot.InsertAfter(element, previous);
            } else {
                WorksheetRoot.Append(element);
            }
        }

        private static IReadOnlyList<int> GetManualPageBreaks(OpenXmlCompositeElement? breaks) {
            if (breaks == null) {
                return Array.Empty<int>();
            }

            return breaks.Elements<Break>()
                .Where(item => item.ManualPageBreak?.Value == true && item.Id?.Value > 0U)
                .Select(item => (int)item.Id!.Value)
                .Distinct()
                .OrderBy(item => item)
                .ToList();
        }

        private static void AddManualPageBreak(OpenXmlCompositeElement breaks, uint id, uint max) {
            Break? existing = breaks.Elements<Break>().FirstOrDefault(item => item.Id?.Value == id);
            if (existing == null) {
                breaks.Append(new Break {
                    Id = id,
                    Min = 0U,
                    Max = max,
                    ManualPageBreak = true
                });
            } else {
                existing.Min = 0U;
                existing.Max = max;
                existing.ManualPageBreak = true;
            }

            UpdatePageBreakCounts(breaks);
        }

        private static bool RemoveManualPageBreak(OpenXmlCompositeElement? breaks, uint id) {
            if (breaks == null) {
                return false;
            }

            Break? existing = breaks.Elements<Break>().FirstOrDefault(item => item.Id?.Value == id && item.ManualPageBreak?.Value == true);
            if (existing == null) {
                return false;
            }

            existing.Remove();
            uint count = (uint)breaks.Elements<Break>().Count();
            if (count == 0U) {
                breaks.Remove();
            } else {
                UpdatePageBreakCounts(breaks);
            }

            return true;
        }

        private static bool ClearManualPageBreaks(OpenXmlCompositeElement? breaks) {
            if (breaks == null) {
                return false;
            }

            var manualBreaks = breaks.Elements<Break>()
                .Where(item => item.ManualPageBreak?.Value == true)
                .ToList();
            if (manualBreaks.Count == 0) {
                return false;
            }

            foreach (Break manualBreak in manualBreaks) {
                manualBreak.Remove();
            }

            uint count = (uint)breaks.Elements<Break>().Count();
            if (count == 0U) {
                breaks.Remove();
            } else {
                UpdatePageBreakCounts(breaks);
            }

            return true;
        }

        private static void UpdatePageBreakCounts(OpenXmlCompositeElement breaks) {
            uint count = (uint)breaks.Elements<Break>().Count();
            uint manualCount = (uint)breaks.Elements<Break>().Count(item => item.ManualPageBreak?.Value == true);
            switch (breaks) {
                case RowBreaks rowBreaks:
                    rowBreaks.Count = count;
                    rowBreaks.ManualBreakCount = manualCount;
                    break;
                case ColumnBreaks columnBreaks:
                    columnBreaks.Count = count;
                    columnBreaks.ManualBreakCount = manualCount;
                    break;
            }
        }
    }
}
