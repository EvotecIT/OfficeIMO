namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>A width-bound contained story; only the accepted cursor advances its source position.</summary>
        private sealed class PreparedCellFlow {
            public PreparedCellFlow(TableCellLayout cell, List<FootnoteUnit> units, double width) {
                Cell = cell; Units = units; Width = width;
            }
            public TableCellLayout Cell { get; }
            public List<FootnoteUnit> Units { get; }
            public double Width { get; }
        }

        private readonly record struct CellFlowCursor(int Unit, TableFlowCursor? Table);

        private static bool HasContainedCellBlocks(TableBlock table) => table.Cells.Any(row => row.Any(cell => cell.Blocks.Count > 0));

        private static PdfTableStyle ResolveContainedTableWidth(PdfTableStyle style, double frameWidth) {
            if (!style.PreferredWidthFraction.HasValue && !style.MaxWidthFraction.HasValue) return style;
            if ((double.IsNaN(style.FrameWidthAdjustment) || double.IsInfinity(style.FrameWidthAdjustment))) throw new ArgumentException("A table's frame adjustment must be finite.");
            if (double.IsNaN(style.ResolvedWidthAdjustment) || double.IsInfinity(style.ResolvedWidthAdjustment))
                throw new ArgumentException("A table's resolved width adjustment must be finite.");
            double referenceWidth = frameWidth + style.FrameWidthAdjustment;
            if (!(referenceWidth > 0D) || (double.IsNaN(referenceWidth) || double.IsInfinity(referenceWidth))) throw new ArgumentException("A table's relative width requires a positive finite containing frame.");
            var resolved = style.Clone();
            // The source adapter already captured physical fixed-column minima,
            // including distributed span minima. Retain the same grid floor after
            // converting a source perimeter width into the enclosed cell grid.
            double minimumWidth = (!style.AutoFitColumns || style.PreserveAuthoredPreferencesAtWidthLimit) &&
                style.ColumnMinWidthPoints is { } minimums
                ? minimums.Sum(width => width ?? 0D) + Math.Max(0, minimums.Count - 1) * style.CellSpacing
                : 0D;
            double Resolve(double fraction) {
                if (!(fraction > 0D) || (double.IsNaN(fraction) || double.IsInfinity(fraction))) throw new ArgumentException("A table's relative width must be positive and finite.");
                double width = referenceWidth * fraction + style.ResolvedWidthAdjustment;
                return Math.Max(minimumWidth, width);
            }
            if (style.PreferredWidthFraction.HasValue) resolved.PreferredWidth = Resolve(style.PreferredWidthFraction.Value);
            if (style.MaxWidthFraction.HasValue) resolved.MaxWidth = Resolve(style.MaxWidthFraction.Value);
            resolved.PreferredWidthFraction = null; resolved.MaxWidthFraction = null;
            return resolved;
        }

        private Dictionary<int, PreparedCellFlow[]> PrepareContainedTableCells(TableBlock table, PdfTableStyle style,
            PreparedFlowTableRows rows, double[] widths, double gap) {
            var result = new Dictionary<int, PreparedCellFlow[]>();
            if (!HasContainedCellBlocks(table)) return result;
            if (style.Position != null || !style.ConsumesVerticalFlow || style.Caption != null || style.ShrinkTextToFit)
                throw new NotSupportedException("Ordered table cell content requires inline table flow without captions or table text shrinking.");
            if (table.Cells.SelectMany(row => row).Any(cell => cell.RowSpan > 1))
                throw new NotSupportedException("Ordered table cell content does not yet support spanning physical rows.");
            if (PdfFootnoteTraversal.References(table).Any())
                throw new NotSupportedException("Page-associated references inside ordered table cell content require a supported reference-admission profile.");
            for (int row = 0; row < table.Rows.Count; row++) {
                var cells = GetTableCellLayouts(table, row, widths.Length);
                if (!cells.Any(cell => cell.Blocks.Count > 0)) continue;
                if (GetTableRowFixedHeight(style, row).HasValue)
                    throw new NotSupportedException("Ordered table cell content requires flexible row height.");
                var flows = new List<PreparedCellFlow>();
                double height = 0D;
                foreach (TableCellLayout cell in cells) {
                    if (cell.TextRotation != 0 || cell.Viewport != null || cell.ShrinkTextToFit || cell.NoWrap ||
                        cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0 || cell.Images.Count > 0)
                        throw new NotSupportedException("Ordered table cell content requires horizontal flowing cells without viewports, shrinking, or sidecar objects.");
                    double width = GetTableCellWidth(widths, cell.Column, cell.ColumnSpan, gap) -
                        GetTableCellPaddingLeft(style, row, cell.Column) - GetTableCellPaddingRight(style, row, cell.Column);
                    if (!(width > 0D)) throw new ArgumentException("Cell padding leaves no width for contained flow.");
                    IPdfBlock[] blocks = GetContainedCellBlocks(cell, style, rows, row);
                    var units = new List<FootnoteUnit>();
                    AppendFootnoteUnits(blocks, 0D, width, Array.Empty<FootnotePanel>(), units);
                    CompleteFootnoteUnitMeasurements(units);
                    double padding = GetTableCellPaddingTop(style, row, cell.Column) + GetTableCellPaddingBottom(style, row, cell.Column);
                    foreach (FootnoteUnit unit in units) unit.MaximumKeepHeight = Math.Max(0D, GetFullPageContentHeight() - padding);
                    flows.Add(new PreparedCellFlow(cell, units, width));
                    height = Math.Max(height, MeasureFootnoteUnits(units) + padding);
                }
                result.Add(row, flows.ToArray());
                rows.Heights[row] = rows.IntrinsicHeights[row] = ResolveTableRowHeight(style, row, height);
            }
            return result;
        }

        private static IPdfBlock[] GetContainedCellBlocks(TableCellLayout cell, PdfTableStyle tableStyle,
            PreparedFlowTableRows rows, int row) {
            IReadOnlyList<IPdfBlock> captured = cell.Blocks.Count > 0 ? cell.Blocks : cell.Paragraphs.Count > 0
                ? cell.Paragraphs.Select(paragraph => (IPdfBlock)paragraph.ToBlock()).ToArray()
                : new IPdfBlock[] { new RichParagraphBlock(cell.Runs,
                    MapTableCellAlignment(GetTableCellAlignment(tableStyle, row, cell.Column, cell.Text)), null) };
            return captured.Select(block => {
                if (block is not RichParagraphBlock paragraph) return block;
                var style = paragraph.Style?.Clone() ?? new PdfParagraphStyle();
                if (style.AnchoredCanvas != null || paragraph.Runs.Any(run => run.FormField != null))
                    throw new NotSupportedException("Ordered cell paragraphs cannot contain anchored canvases or interactive form fields.");
                style.FontSize ??= global::OfficeIMO.OfficeLength.Points(rows.Sizes[row]);
                style.LineHeight ??= rows.Leadings[row] / rows.Sizes[row];
                // SpacingBefore is a non-nullable captured value; its default
                // already is zero. Only a missing after-spacing inherits zero
                // from the ordinary cell contract rather than page defaults.
                style.SpacingAfter ??= global::OfficeIMO.OfficeLength.Points(0D);
                return (IPdfBlock)new RichParagraphBlock(paragraph.Runs.Select(run => run.WithInheritedFontStyle(rows.Bold[row], false)),
                    paragraph.Align, paragraph.DefaultColor ?? (row < tableStyle.HeaderRowCount ? tableStyle.HeaderTextColor :
                        row >= rows.Heights.Length - tableStyle.FooterRowCount ? tableStyle.FooterTextColor : tableStyle.TextColor), style);
            }).ToArray();
        }
    }
}
