namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Width-bound table geometry shared by ordinary flow and retained note fragments.</summary>
        private sealed class PreparedTableGeometry {
            public PreparedTableGeometry(TableBlock table, PdfTableStyle style, PreparedFlowTableRows rows,
                double[] widths, double width, double cornerRadius) {
                Table = table; Style = style; Rows = rows; Widths = widths; Width = width; CornerRadius = cornerRadius;
            }
            public TableBlock Table { get; }
            public PdfTableStyle Style { get; }
            public PreparedFlowTableRows Rows { get; }
            public double[] Widths { get; }
            public int Columns => Widths.Length;
            public double Width { get; }
            public double CornerRadius { get; }
            public double Gap => GetTableCellSpacing(Style);
        }

        /// <summary>An accepted physical row slice, including its measured box rather than an additive line estimate.</summary>
        private sealed class TableRowSlice {
            public TableRowSlice(int row, int startLine, int lineCount, double height, bool isHeader, bool repeated = false) {
                Row = row; StartLine = startLine; LineCount = lineCount; Height = height; IsHeader = isHeader; IsRepeatedHeader = repeated;
            }
            public int Row { get; }
            public int StartLine { get; }
            public int LineCount { get; }
            public double Height { get; }
            public bool IsHeader { get; }
            public bool IsRepeatedHeader { get; }
            public double GapAfter { get; set; }
            public Dictionary<(int Row, int Column), int>? CompletedSpans { get; set; }
        }

        private PreparedTableGeometry PrepareTableGeometry(TableBlock table, PdfTableStyle style,
            double frameWidth, bool logicalTopBoundary = true, bool logicalBottomBoundary = true) {
            cancellationToken.ThrowIfCancellationRequested();
            int columns = GetTableColumnCount(table);
            ValidateFlowTable(table, style, columns);
            var columnLayout = ResolveTableColumnLayout(table, currentOpts, style, columns, frameWidth,
                GetTableBodyFontSize(style, currentOpts.DefaultFontSize), style.HeaderRowCount,
                table.Rows.Count - style.FooterRowCount);
            double gap = GetTableCellSpacing(style);
            var rows = PrepareFlowTableRows(table, style, columns, columnLayout.Widths, gap, gap,
                style.HeaderRowCount, table.Rows.Count - style.FooterRowCount);
            double radius = ResolveTableCornerRadius(style.CornerRadius, columnLayout.Width,
                logicalTopBoundary ? rows.Heights[0] : double.PositiveInfinity,
                logicalBottomBoundary ? rows.Heights[rows.Heights.Length - 1] : double.PositiveInfinity,
                columnLayout.Widths[0], columnLayout.Widths[columnLayout.Widths.Length - 1]);
            return new PreparedTableGeometry(table, style, rows, columnLayout.Widths, columnLayout.Width, radius);
        }
    }
}
