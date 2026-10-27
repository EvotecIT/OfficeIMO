namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly Dictionary<TableBlock, ContainerRenderScope> tableFrameScopes = new();

        private void PrepareTableFrameContinuation(TableBlock table, int rowIndex) {
            if (!tableFrameScopes.TryGetValue(table, out ContainerRenderScope? scope) || scope.Container == null) return;
            double? padding = ResolveTableFrameContinuationTopPadding(scope.Container, rowIndex);
            if (padding.HasValue) scope.Style.ContinuationPaddingTopOverride = padding;
        }

        /// <summary>Resolves a continued native grid's own outward paint without mutating captured frame styles during preview.</summary>
        private static double? ResolveTableFrameContinuationTopPadding(ContainerBlock container, int rowIndex) {
            if (container.FrameTableTopOutsets is not { } outsets || outsets.Count == 0 || container.FrameTable is not { } table)
                return null;
            PdfTableStyle style = table.Style!;
            bool repeatsHeader = GetTableRepeatHeaderRowCount(style) > 0 && rowIndex >= style.HeaderRowCount;
            int firstRow = repeatsHeader ? 0 : Math.Min(rowIndex, outsets.Count - 1);
            return outsets[firstRow];
        }

        private void RecordTableFrameContentBottom(TableBlock table, double bottom) {
            if (tableFrameScopes.TryGetValue(table, out ContainerRenderScope? scope))
                scope.TableContentBottom = bottom;
        }

        /// <summary>Measures an independent table perimeter using the same prepared columns as its cell grid.</summary>
        private (double X, double Width, double ContentWidth) ResolveContainerFrame(
            ContainerBlock? container, PdfPanelStyle style, double parentLeft, double parentWidth, double? fontSize = null) {
            if (container?.FrameTable is not { } table)
                return ResolveContainerFrame(style, parentLeft, parentWidth);
            parentLeft += style.LeftInset.ToPoints();
            parentWidth = ResolveContainerAvailableWidth(style, parentWidth);
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
            placement.HorizontalOffset = container.FrameTableHorizontalOffset;
            return (ResolveTableX(table.Align, placement, parentLeft, parentWidth, outerWidth), outerWidth, contentWidth);
        }

        private static double ResolveContainerTableFrameBottom(ContainerRenderScope scope, double bottom, bool continues) {
            return ResolveContainerTableFrameBottom(scope.Container, scope.TableContentBottom, bottom, continues);
        }

        private static double ResolveContainerTableFrameBottom(ContainerBlock? container, double? tableContentBottom,
            double bottom, bool continues) {
            if (continues && container?.FrameTable != null && tableContentBottom.HasValue) {
                // Inter-row spacing does not belong below the fragment's last
                // cell. Word closes this perimeter with half its cell gap.
                return tableContentBottom.Value - container.FrameTableContinuationBottomPadding;
            }
            return bottom;
        }

        private void DrawContainerTableFrame(StringBuilder output, ContainerRenderScope scope,
            double bottom, double height, bool drawTop, bool drawBottom) {
            DrawContainerTableFrame(output, scope.Container, scope.OuterX, scope.OuterWidth, bottom, height, drawTop, drawBottom);
        }

        /// <summary>Paints the existing independent table perimeter for body and retained note fragments.</summary>
        private void DrawContainerTableFrame(StringBuilder output, ContainerBlock? container, double outerX, double outerWidth,
            double bottom, double height, bool drawTop, bool drawBottom) {
            if (container?.FrameTableBorder is not { } source) return;
            PdfCellBorder border = source.Clone();
            border.Top &= drawTop;
            border.Bottom &= drawBottom;
            double inset = container.FrameTableBorderInset;
            double topInset = border.Top ? border.TopBorderSnapshot?.Width / 2D ?? 0D : 0D;
            double bottomInset = border.Bottom ? border.BottomBorderSnapshot?.Width / 2D ?? 0D : 0D;
            DrawCellBorder(output, border, outerX + inset, bottom + bottomInset,
                outerWidth - 2D * inset, height - topInset - bottomInset, emitGeneratedStructure);
        }
    }
}
