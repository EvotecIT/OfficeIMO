namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly Dictionary<TableBlock, ContainerRenderScope> tableFrameScopes = new();

        private void RecordTableFrameContentBottom(TableBlock table, double bottom) {
            if (tableFrameScopes.TryGetValue(table, out ContainerRenderScope? scope))
                scope.TableContentBottom = bottom;
        }

        /// <summary>Measures an independent table perimeter using the same prepared columns as its cell grid.</summary>
        private (double X, double Width, double ContentWidth) ResolveContainerFrame(
            ContainerBlock? container, PdfPanelStyle style, double parentLeft, double parentWidth, double? fontSize = null) {
            if (container?.FrameTable is not { } table)
                return ResolveContainerFrame(style, parentLeft, parentWidth);
            PdfTableStyle tableStyle = PreparePairedTableBorders(table, table.Style!);
            double available = parentWidth - container.FrameTableIndent - 2D * style.PaddingX;
            int columns = GetTableColumnCount(table);
            int headers = Math.Min(tableStyle.HeaderRowCount, table.Rows.Count);
            int footerStart = table.Rows.Count - Math.Min(tableStyle.FooterRowCount, table.Rows.Count - headers);
            double contentWidth = ResolveTableColumnLayout(table, currentOpts, tableStyle, columns, available,
                GetTableBodyFontSize(tableStyle, fontSize ?? currentOpts.DefaultFontSize), headers, footerStart).Width;
            double outerWidth = contentWidth + 2D * style.PaddingX;
            ValidatePanelStyle(style, outerWidth);
            PdfTableStyle placement = tableStyle.Clone();
            placement.LeftIndent = container.FrameTableIndent;
            return (ResolveTableX(table.Align, placement, parentLeft, parentWidth, outerWidth), outerWidth, contentWidth);
        }

        private static double ResolveContainerTableFrameBottom(ContainerRenderScope scope, double bottom, bool continues) {
            if (continues && scope.Container?.FrameTable != null && scope.TableContentBottom.HasValue) {
                // Inter-row spacing does not belong below the fragment's last
                // cell. Word closes this perimeter with half its cell gap.
                return scope.TableContentBottom.Value - scope.Container.FrameTableContinuationBottomPadding;
            }
            return bottom;
        }

        private void DrawContainerTableFrame(StringBuilder output, ContainerRenderScope scope,
            double bottom, double height, bool drawTop, bool drawBottom) {
            if (scope.Container?.FrameTableBorder is not { } source) return;
            PdfCellBorder border = source.Clone();
            border.Top &= drawTop;
            border.Bottom &= drawBottom;
            double inset = scope.Container.FrameTableBorderInset;
            double topInset = border.Top ? border.TopBorderSnapshot?.Width / 2D ?? 0D : 0D;
            double bottomInset = border.Bottom ? border.BottomBorderSnapshot?.Width / 2D ?? 0D : 0D;
            DrawCellBorder(output, border, scope.OuterX + inset, bottom + bottomInset,
                scope.OuterWidth - 2D * inset, height - topInset - bottomInset, emitGeneratedStructure);
        }
    }
}
