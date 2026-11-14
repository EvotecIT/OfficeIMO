using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly Regex TableSummaryTokenPattern = new(@"\{(page|broughtForward|running):([^{}:\s]+)\}", RegexOptions.CultureInvariant);

    /// <summary>Totals of one aggregate over a set of body rows.</summary>
    private readonly struct TableAggregateTotal {
        private TableAggregateTotal(int count, decimal sum, decimal minimum, decimal maximum) {
            Count = count; Sum = sum; Minimum = minimum; Maximum = maximum;
        }
        public int Count { get; }
        public decimal Sum { get; }
        public decimal Minimum { get; }
        public decimal Maximum { get; }

        public TableAggregateTotal Add(decimal value) => Count == 0
            ? new TableAggregateTotal(1, value, value, value)
            : new TableAggregateTotal(Count + 1, Sum + value, Math.Min(Minimum, value), Math.Max(Maximum, value));

        public static TableAggregateTotal Combine(TableAggregateTotal first, TableAggregateTotal second) {
            if (first.Count == 0) return second;
            if (second.Count == 0) return first;
            return new TableAggregateTotal(first.Count + second.Count, first.Sum + second.Sum,
                Math.Min(first.Minimum, second.Minimum), Math.Max(first.Maximum, second.Maximum));
        }

        public string Format(PdfTableAggregate aggregate) {
            IFormatProvider culture = aggregate.Culture;
            switch (aggregate.Function) {
                case PdfTableAggregateFunction.Count:
                    return ((decimal)Count).ToString(aggregate.Format ?? "N0", culture);
                case PdfTableAggregateFunction.Minimum:
                    return Count == 0 ? string.Empty : Minimum.ToString(aggregate.Format ?? "N2", culture);
                case PdfTableAggregateFunction.Maximum:
                    return Count == 0 ? string.Empty : Maximum.ToString(aggregate.Format ?? "N2", culture);
                default:
                    return Sum.ToString(aggregate.Format ?? "N2", culture);
            }
        }
    }

    /// <summary>Summary rows prepared with resolved values against the table's current column widths.</summary>
    private sealed class PreparedTableSummaryRows {
        public PreparedTableSummaryRows(TableBlock table, PreparedFlowTableRows rows, double height) {
            Table = table; Rows = rows; Height = height;
        }
        public TableBlock Table { get; }
        public PreparedFlowTableRows Rows { get; }
        public double Height { get; }
    }

    private enum TableSummaryRowKind { PageFooter, Continuation, FinalFooter }

    /// <summary>
    /// Per-table state for page aggregates and repeated footer rows. Body rows are counted when their first
    /// segment is drawn, so values always describe the rows that landed on each page fragment.
    /// </summary>
    private sealed class TablePageSummaryState {
        private readonly decimal?[][] _values;
        private readonly TableAggregateTotal[][] _suffix;
        private readonly bool[] _counted;
        private readonly int _headerRowCount;

        public TablePageSummaryState(PdfTableAggregate[] aggregates, decimal?[][] values, int headerRowCount, int footerStartRowIndex,
            PdfTableCell[][] pageFooterRows, int[] pageFooterSourceRows, PdfTableCell[][] continuationRows,
            PdfTableCell[][] finalFooterRows, int[] finalFooterSourceRows) {
            Aggregates = aggregates;
            _values = values;
            _headerRowCount = headerRowCount;
            FooterStartRowIndex = footerStartRowIndex;
            PageFooterRows = pageFooterRows;
            PageFooterSourceRows = pageFooterSourceRows;
            ContinuationRows = continuationRows;
            FinalFooterRows = finalFooterRows;
            FinalFooterSourceRows = finalFooterSourceRows;
            _counted = new bool[values.Length];
            Page = new TableAggregateTotal[aggregates.Length];
            Before = new TableAggregateTotal[aggregates.Length];
            // Suffix totals let the final-fragment check see the values the closing footer would show.
            _suffix = new TableAggregateTotal[Math.Max(footerStartRowIndex, headerRowCount) + 1][];
            for (int row = _suffix.Length - 1; row >= 0; row--) {
                var totals = new TableAggregateTotal[aggregates.Length];
                if (row + 1 < _suffix.Length) Array.Copy(_suffix[row + 1], totals, totals.Length);
                if (row >= headerRowCount && row < footerStartRowIndex) Accumulate(totals, row);
                _suffix[row] = totals;
            }
        }

        public PdfTableAggregate[] Aggregates { get; }
        public int FooterStartRowIndex { get; }
        public PdfTableCell[][] PageFooterRows { get; }
        public int[] PageFooterSourceRows { get; }
        public PdfTableCell[][] ContinuationRows { get; }
        public PdfTableCell[][] FinalFooterRows { get; }
        public int[] FinalFooterSourceRows { get; }
        public TableAggregateTotal[] Page { get; private set; }
        public TableAggregateTotal[] Before { get; private set; }
        public bool FragmentHasBodyRows { get; private set; }
        public bool HasRowsBefore { get; private set; }
        public PreparedTableSummaryRows? PendingContinuation { get; set; }
        public Dictionary<string, PreparedTableSummaryRows> Prepared { get; } = new(StringComparer.Ordinal);

        public bool IsBodyRow(int row) => row >= _headerRowCount && row < FooterStartRowIndex;
        public bool IsCounted(int row) => _counted[row];

        public void MarkDrawn(int row) {
            if (!IsBodyRow(row)) return;
            FragmentHasBodyRows = true;
            if (_counted[row]) return;
            _counted[row] = true;
            Accumulate(Page, row);
        }

        public void ClosePage() {
            if (!FragmentHasBodyRows) return;
            for (int i = 0; i < Page.Length; i++) Before[i] = TableAggregateTotal.Combine(Before[i], Page[i]);
            HasRowsBefore = true;
            Page = new TableAggregateTotal[Aggregates.Length];
            FragmentHasBodyRows = false;
        }

        public TableAggregateTotal[] PageWith(int row) {
            var totals = (TableAggregateTotal[])Page.Clone();
            if (IsBodyRow(row) && !_counted[row]) Accumulate(totals, row);
            return totals;
        }

        public TableAggregateTotal[] PageWithRemainder(int row) {
            var totals = new TableAggregateTotal[Page.Length];
            for (int i = 0; i < totals.Length; i++)
                totals[i] = TableAggregateTotal.Combine(Page[i], row < _suffix.Length ? _suffix[row][i] : default);
            return totals;
        }

        public Func<string, string> CreateResolver(TableAggregateTotal[] page, out string key) {
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            var keyBuilder = new StringBuilder();
            for (int i = 0; i < Aggregates.Length; i++) {
                PdfTableAggregate aggregate = Aggregates[i];
                string pageValue = page[i].Format(aggregate);
                string beforeValue = Before[i].Format(aggregate);
                string runningValue = TableAggregateTotal.Combine(Before[i], page[i]).Format(aggregate);
                values["page:" + aggregate.Name] = pageValue;
                values["broughtForward:" + aggregate.Name] = beforeValue;
                values["running:" + aggregate.Name] = runningValue;
                keyBuilder.Append(pageValue).Append('\u001f').Append(beforeValue).Append('\u001f').Append(runningValue).Append('\u001e');
            }
            key = keyBuilder.ToString();
            return text => text.IndexOf('{') < 0 ? text : TableSummaryTokenPattern.Replace(text,
                match => values.TryGetValue(match.Groups[1].Value + ":" + match.Groups[2].Value, out string? value) ? value : match.Value);
        }

        private void Accumulate(TableAggregateTotal[] totals, int row) {
            decimal?[] rowValues = _values[row];
            if (rowValues == null) return;
            for (int i = 0; i < totals.Length; i++)
                if (rowValues[i] is decimal value) totals[i] = totals[i].Add(value);
        }
    }

    private static bool HasTablePageSummary(PdfTableStyle style) =>
        style.PageSummary != null || style.RepeatFooterRows && style.FooterRowCount > 0;

    private static void ThrowIfTablePageSummaryUnsupported(PdfTableStyle? style, string context) {
        if (style != null && HasTablePageSummary(style))
            throw new NotSupportedException("Table page summaries and repeated footer rows are not supported in " + context + "; use an eager flow table.");
    }

    private static TablePageSummaryState? CreateTablePageSummaryState(TableBlock table, PdfTableStyle style, int columns,
        int headerRowCount, int footerStartRowIndex) {
        if (!HasTablePageSummary(style)) return null;
        if (style.Position != null)
            throw new NotSupportedException("Table page summaries and repeated footer rows are not supported for positioned tables.");
        PdfTablePageSummary? summary = style.PageSummary;
        PdfTableAggregate[] aggregates = summary?.Aggregates.ToArray() ?? Array.Empty<PdfTableAggregate>();
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfTableAggregate aggregate in aggregates) {
            if (aggregate == null) throw new ArgumentException("Table page summary aggregates cannot contain null entries.");
            if (!names.Add(aggregate.Name)) throw new ArgumentException($"Table aggregate '{aggregate.Name}' is declared more than once.");
            if (aggregate.Column >= columns)
                throw new ArgumentException($"Table aggregate '{aggregate.Name}' refers to column {aggregate.Column}, but the table has {columns} columns.");
        }

        int footerRowCount = table.Rows.Count - footerStartRowIndex;
        for (int row = 0; row < footerStartRowIndex; row++) {
            foreach (PdfTableCell cell in table.Cells[row]) {
                if (footerRowCount > 0 && row + cell.RowSpan > footerStartRowIndex)
                    throw new ArgumentException("Merged cells cannot extend into footer rows of a table with page summaries or repeated footer rows.");
            }
        }

        PdfTableCell[][] footerRows = table.Cells.Skip(footerStartRowIndex).Select(row => row.ToArray()).ToArray();
        int[] footerSources = Enumerable.Range(footerStartRowIndex, footerRowCount).ToArray();
        PdfTableCell[][] pageSummaryRows = summary?.PageFooterRows.ToArray() ?? Array.Empty<PdfTableCell[]>();
        PdfTableCell[][] pageFooterRows = style.RepeatFooterRows ? pageSummaryRows.Concat(footerRows).ToArray() : pageSummaryRows;
        int[] pageFooterSources = Enumerable.Repeat(-1, pageSummaryRows.Length)
            .Concat(style.RepeatFooterRows ? footerSources : Array.Empty<int>()).ToArray();
        PdfTableCell[][] continuationRows = summary?.ContinuationHeaderRows.ToArray() ?? Array.Empty<PdfTableCell[]>();

        foreach (PdfTableCell[] row in pageSummaryRows.Concat(footerRows)) {
            ValidateTableSummaryRow(row, columns, names, continuation: false);
        }
        foreach (PdfTableCell[] row in continuationRows) {
            ValidateTableSummaryRow(row, columns, names, continuation: true);
        }

        var values = new decimal?[table.Rows.Count][];
        for (int row = headerRowCount; row < footerStartRowIndex; row++) {
            if (aggregates.Length == 0) break;
            var rowValues = new decimal?[aggregates.Length];
            List<TableCellLayout> cells = GetTableCellLayouts(table, row, columns);
            for (int i = 0; i < aggregates.Length; i++) {
                PdfTableAggregate aggregate = aggregates[i];
                if (aggregate.ValueSelector != null) {
                    rowValues[i] = aggregate.ValueSelector(row - headerRowCount);
                    continue;
                }
                string? text = null;
                foreach (TableCellLayout cell in cells) {
                    if (aggregate.Column >= cell.Column && aggregate.Column < cell.Column + cell.ColumnSpan) { text = cell.Text; break; }
                }
                rowValues[i] = ParseTableAggregateValue(text, aggregate);
            }
            values[row] = rowValues;
        }

        return new TablePageSummaryState(aggregates, values, headerRowCount, footerStartRowIndex,
            pageFooterRows, pageFooterSources, continuationRows, footerRows, footerSources);
    }

    private static decimal? ParseTableAggregateValue(string? text, PdfTableAggregate aggregate) {
        if (string.IsNullOrWhiteSpace(text)) return null;
        if (aggregate.Function == PdfTableAggregateFunction.Count) return 1M;
        string trimmed = text!.Trim();
        const NumberStyles styles = NumberStyles.Number | NumberStyles.AllowCurrencySymbol | NumberStyles.AllowParentheses;
        return decimal.TryParse(trimmed, styles, aggregate.Culture, out decimal value) ? value : null;
    }

    private static void ValidateTableSummaryRow(PdfTableCell[] row, int columns, HashSet<string> aggregateNames, bool continuation) {
        int span = 0;
        foreach (PdfTableCell cell in row) {
            if (cell.RowSpan > 1)
                throw new ArgumentException("Table summary and footer rows cannot contain merged rows.");
            if (cell.FormFields.Count > 0 || cell.CheckBoxes.Count > 0 || cell.Blocks.Count > 0 || cell.Viewport != null ||
                !string.IsNullOrEmpty(cell.NamedDestinationName))
                throw new ArgumentException("Table summary and footer rows support text, links and images only.");
            span += cell.ColumnSpan;
            foreach (string text in cell.Runs.Select(run => run.Text).Concat(cell.Paragraphs.SelectMany(p => p.Runs).Select(run => run.Text))) {
                foreach (Match match in TableSummaryTokenPattern.Matches(text)) {
                    if (!aggregateNames.Contains(match.Groups[2].Value))
                        throw new ArgumentException($"Table summary token '{match.Value}' refers to an aggregate that is not declared.");
                    // Continuation rows are drawn when their page starts, before its rows are known.
                    if (continuation && match.Groups[1].Value != "broughtForward")
                        throw new ArgumentException($"Continuation header rows can only use brought-forward tokens; '{match.Value}' depends on rows that are not placed yet.");
                }
            }
        }
        if (span != columns)
            throw new ArgumentException($"Table summary and footer rows must cover all {columns} columns.");
        if (PdfFootnoteTraversal.References(new TableBlock(new[] { row }, OfficeAlignment.Left, null)).Any())
            throw new ArgumentException("Table summary and footer rows cannot contain footnote references.");
    }

    /// <summary>Copies the table style for a standalone block of summary rows that render with footer styling.</summary>
    private static PdfTableStyle CreateTableSummaryRowStyle(PdfTableStyle tableStyle, int[] sourceRows) {
        PdfTableStyle style = tableStyle.Clone();
        style.HeaderRowCount = 0;
        style.RepeatHeaderRowCount = 0;
        style.FooterRowCount = sourceRows.Length;
        style.RepeatFooterRows = false;
        style.PageSummary = null;
        style.Caption = null;
        style.KeepTogether = false;
        style.KeepWithNext = false;
        style.MinimumBodyRowsOnFirstPage = 0;
        style.MinimumBodyRowsOnLastPage = 0;
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Zero;
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        style.CornerRadius = global::OfficeIMO.OfficeLength.Zero;
        style.BorderFrame = null;
        style.AutoFitCellPreferredWidthPoints = null;
        style.CellFills = Remap(tableStyle.CellFills, value => value);
        style.CellDataBars = Remap(tableStyle.CellDataBars, value => value.Clone());
        style.CellIcons = Remap(tableStyle.CellIcons, value => value.Clone());
        style.CellBorders = Remap(tableStyle.CellBorders, value => value.Clone());
        style.CellPaddings = Remap(tableStyle.CellPaddings, value => value.Clone());
        style.CellAlignments = Remap(tableStyle.CellAlignments, value => value);
        style.CellVerticalAlignments = Remap(tableStyle.CellVerticalAlignments, value => value);
        style.CellHeaderScopes = Remap(tableStyle.CellHeaderScopes, value => value);
        style.RowMinHeights = RemapRows(tableStyle.RowMinHeights);
        style.FixedRowHeights = RemapRows(tableStyle.FixedRowHeights);
        style.RowAllowBreakAcrossPages = RemapRows(tableStyle.RowAllowBreakAcrossPages);
        return style;

        Dictionary<(int Row, int Column), T>? Remap<T>(Dictionary<(int Row, int Column), T>? source, Func<T, T> copy) {
            if (source == null) return null;
            var result = new Dictionary<(int Row, int Column), T>();
            for (int row = 0; row < sourceRows.Length; row++) {
                if (sourceRows[row] < 0) continue;
                foreach (var entry in source)
                    if (entry.Key.Row == sourceRows[row]) result[(row, entry.Key.Column)] = copy(entry.Value);
            }
            return result.Count == 0 ? null : result;
        }

        List<T?>? RemapRows<T>(List<T?>? source) {
            if (source == null) return null;
            var result = new List<T?>();
            bool any = false;
            foreach (int sourceRow in sourceRows) {
                T? value = sourceRow >= 0 && sourceRow < source.Count ? source[sourceRow] : default;
                any |= value != null;
                result.Add(value);
            }
            return any ? result : null;
        }
    }

    private sealed partial class LayoutContext {
        private PreparedTableSummaryRows? PrepareTableSummaryRows(TablePageSummaryState summary, TableSummaryRowKind kind,
            TableAggregateTotal[] page, PdfTableStyle tableStyle, int columns, double[] columnWidths, double gap) {
            PdfTableCell[][] templates;
            int[] sources;
            switch (kind) {
                case TableSummaryRowKind.PageFooter: templates = summary.PageFooterRows; sources = summary.PageFooterSourceRows; break;
                case TableSummaryRowKind.Continuation: templates = summary.ContinuationRows; sources = Enumerable.Repeat(-1, templates.Length).ToArray(); break;
                default: templates = summary.FinalFooterRows; sources = summary.FinalFooterSourceRows; break;
            }
            if (templates.Length == 0) return null;
            Func<string, string> resolve = summary.CreateResolver(page, out string valueKey);
            string key = ((int)kind).ToString(CultureInfo.InvariantCulture) + "|" + columnWidths.Length.ToString(CultureInfo.InvariantCulture) + "|" +
                string.Join(",", columnWidths.Select(width => width.ToString("R", CultureInfo.InvariantCulture))) + "|" +
                System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(tableStyle).ToString(CultureInfo.InvariantCulture) + "|" + valueKey;
            if (summary.Prepared.TryGetValue(key, out PreparedTableSummaryRows? cached)) return cached;
            PdfTableStyle style = CreateTableSummaryRowStyle(tableStyle, sources);
            var table = new TableBlock(templates.Select(row => row.Select(cell => cell.WithReplacedText(resolve)).ToArray()), OfficeAlignment.Left, style);
            PreparedFlowTableRows rows = PrepareFlowTableRows(table, table.Style!, columns, columnWidths, gap, gap, 0, 0);
            var prepared = new PreparedTableSummaryRows(table, rows, GetTableRowsHeight(rows.Heights, 0, rows.Heights.Length, gap));
            // Values change on almost every row, so keep only a small working set.
            if (summary.Prepared.Count >= 64) summary.Prepared.Clear();
            summary.Prepared[key] = prepared;
            return prepared;
        }

        private void PaintTableSummaryRows(PreparedTableSummaryRows prepared, double[] columnWidths, double tableWidth, double gap,
            double xOrigin, int? tableStructureElementIndex, ref StringBuilder? pairedBorders, bool logicalBottomBoundary) {
            TableBlock table = prepared.Table;
            PdfTableStyle style = table.Style!;
            var geometry = new PreparedTableGeometry(table, style, prepared.Rows, columnWidths, tableWidth, 0D);
            var spanFlow = new TableSpanFlow(table, style, columnWidths.Length, 0, 0);
            for (int row = 0; row < prepared.Rows.Heights.Length; row++) {
                double height = prepared.Rows.Heights[row];
                var slice = new TableRowSlice(row, 0, prepared.Rows.LineCounts[row], height, isHeader: false);
                PaintPreparedTableRow(geometry, spanFlow, slice, xOrigin, y, tableStructureElementIndex, ref pairedBorders,
                    logicalTopBoundary: false, logicalBottomBoundary: logicalBottomBoundary && row == prepared.Rows.Heights.Length - 1);
                y -= height + GetTableRowGapAfter(row, prepared.Rows.Heights.Length, gap);
            }
        }
    }
}
