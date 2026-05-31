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
            if (row <= 0) {
                throw new ArgumentOutOfRangeException(nameof(row), "Row page break must be one-based and positive.");
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
            if (column <= 0) {
                throw new ArgumentOutOfRangeException(nameof(column), "Column page break must be one-based and positive.");
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

            uint count = (uint)breaks.Elements<Break>().Count();
            switch (breaks) {
                case RowBreaks rowBreaks:
                    rowBreaks.Count = count;
                    rowBreaks.ManualBreakCount = count;
                    break;
                case ColumnBreaks columnBreaks:
                    columnBreaks.Count = count;
                    columnBreaks.ManualBreakCount = count;
                    break;
            }
        }
    }
}
