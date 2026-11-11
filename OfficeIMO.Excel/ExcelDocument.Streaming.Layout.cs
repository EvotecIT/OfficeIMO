using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private const string StreamingHeaderStyle = "_OfficeIMO_StreamingHeader";
    private const string StreamingTitleStyle = "_OfficeIMO_StreamingTitle";

    private static ExcelWriterOptions PrepareStreamingStyles(ExcelWriterOptions options, IReadOnlyList<string> headers) {
        var result = CreateRowWriteOptions(options);
        var definitions = options.Styles?.ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, ExcelStyleDefinition>(StringComparer.OrdinalIgnoreCase);
        var columns = options.ColumnStyles?.ToDictionary(item => item.Key, item => item.Value) ?? new Dictionary<int, string>();
        void Add(string name, ExcelStyleDefinition definition) {
            if (definitions.ContainsKey(name)) throw new ArgumentException($"Style name '{name}' is reserved by the streaming layout.", nameof(options));
            definitions.Add(name, definition);
        }
        if (options.BoldHeaders || options.ColumnFormat?.Count > 0) Add(StreamingHeaderStyle, new ExcelStyleDefinition { Bold = options.BoldHeaders });
        if (options.Title != null) Add(StreamingTitleStyle, new ExcelStyleDefinition { Bold = true });
        if (options.ColumnFormat != null) {
            for (int index = 0; index < options.ColumnFormat.Rules.Count; index++) {
                ExcelColumnFormatRule rule = options.ColumnFormat.Rules[index];
                string name = StreamingFormatName(index);
                string format = rule.ResolveNumberFormat();
                Add(name, new ExcelStyleDefinition { NumberFormat = format });
                Add(name + "_Header", new ExcelStyleDefinition { NumberFormat = format, Bold = options.BoldHeaders });
                int column = ResolveStreamingColumn(headers, rule.Header);
                if (column < 0) {
                    if (!options.IgnoreMissingColumnFormat) throw new KeyNotFoundException($"Header '{rule.Header}' was not found in the final worksheet schema.");
                } else columns[column + 1] = name;
            }
        }
        result.Styles = definitions;
        result.ColumnStyles = columns;
        return result;
    }

    private static string StreamingFormatName(int index) => "_OfficeIMO_StreamingFormat_" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static int ResolveStreamingColumn(IReadOnlyList<string> headers, string header) {
        string normalized = ExcelHeaderNameHelper.NormalizeHeader(header, true);
        for (int index = 0; index < headers.Count; index++)
            if (string.Equals(ExcelHeaderNameHelper.NormalizeHeader(headers[index], true), normalized, StringComparison.OrdinalIgnoreCase)) return index;
        return -1;
    }

    private sealed class StreamingWorksheetLayout {
        internal StreamingWorksheetLayout(ExcelWriterOptions options, IReadOnlyList<string> headers, ExcelTabularStylePlan? styles) {
            (FirstRow, FirstColumn) = ExcelA1.ParseCellRef(options.Address);
            if (FirstRow < 1 || FirstRow > ExcelA1.MaxRows || FirstColumn < 1 || FirstColumn + headers.Count - 1 > ExcelA1.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(options.Address), "The exported schema must fit within the worksheet bounds.");
            Title = options.Title;
            HeaderRow = FirstRow + (Title == null ? 0 : 1);
            if (HeaderRow > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(options.Address), "The title exceeds the worksheet bounds.");
            DataRow = HeaderRow + (options.IncludeHeaders ? 1 : 0);
            MaximumRows = ExcelA1.MaxRows - DataRow + 1;
            if (MaximumRows < 0) throw new ArgumentOutOfRangeException(nameof(options.Address), "The title and header exceed the worksheet bounds.");
            HeaderAttributes = new string?[headers.Count];
            string? header = options.BoldHeaders || options.ColumnFormat?.Count > 0 ? styles!.Resolve(StreamingHeaderStyle)!.Attribute : null;
            for (int index = 0; index < headers.Count; index++) HeaderAttributes[index] = header ?? styles?.Columns[index]?.Attribute;
            TitleAttribute = Title == null ? null : styles!.Resolve(StreamingTitleStyle)!.Attribute;
            AutoFitColumns = Enumerable.Repeat(options.AutoFit, headers.Count).ToArray();
            if (options.ColumnFormat != null) {
                for (int index = 0; index < options.ColumnFormat.Count; index++) {
                    ExcelColumnFormatRule rule = options.ColumnFormat.Rules[index];
                    int column = ResolveStreamingColumn(headers, rule.Header);
                    if (column < 0) continue;
                    if (rule.IncludeHeader) HeaderAttributes[column] = styles!.Resolve(StreamingFormatName(index) + "_Header")!.Attribute;
                    AutoFitColumns[column] |= rule.AutoFit;
                }
            }
            int frozenRows = options.FreezeHeader && options.IncludeHeaders ? HeaderRow : 0;
            int frozenColumns = options.FreezeFirstColumn ? FirstColumn : 0;
            if (frozenRows >= ExcelA1.MaxRows || frozenColumns >= ExcelA1.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(options.Address), "A freeze pane must leave a visible worksheet cell.");
            ViewsXml = frozenRows != 0 || frozenColumns != 0 ? FrozenViewsXml(frozenRows, frozenColumns) : ExcelSheet.DefaultSheetViewsXml;
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static string FrozenViewsXml(int frozenRows, int frozenColumns) {
            var view = new SheetView { WorkbookViewId = 0U };
            PaneValues pane = frozenRows != 0 ? frozenColumns != 0 ? PaneValues.BottomRight : PaneValues.BottomLeft : PaneValues.TopRight;
            view.Append(new Pane {
                State = PaneStateValues.Frozen, VerticalSplit = frozenRows, HorizontalSplit = frozenColumns,
                TopLeftCell = ExcelA1.CellReference(frozenRows + 1, frozenColumns + 1), ActivePane = pane
            });
            view.Append(new Selection { Pane = pane, ActiveCell = ExcelA1.CellReference(frozenRows + 1, frozenColumns + 1),
                SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = ExcelA1.CellReference(frozenRows + 1, frozenColumns + 1) } });
            return new SheetViews(view).OuterXml;
        }
        internal int FirstRow { get; }
        internal int FirstColumn { get; }
        internal int HeaderRow { get; }
        internal int DataRow { get; }
        internal int MaximumRows { get; }
        internal string? Title { get; }
        internal string? TitleAttribute { get; }
        internal string?[] HeaderAttributes { get; }
        internal bool[] AutoFitColumns { get; }
        internal string ViewsXml { get; }
        internal string Range(int columns, int rows, bool includeHeaders) =>
            columns <= 0 || (rows == 0 && !includeHeaders) ? string.Empty :
            ExcelA1.CellReference(HeaderRow, FirstColumn) + ":" +
            ExcelA1.CellReference(Math.Max(HeaderRow, DataRow + rows - 1), FirstColumn + columns - 1);
    }
}
