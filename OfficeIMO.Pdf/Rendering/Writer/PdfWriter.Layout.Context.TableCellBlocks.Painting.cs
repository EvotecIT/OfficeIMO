namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints an already accepted contained fragment inside the outer cell's physical frame.</summary>
        private void PaintContainedCell(PreparedTableGeometry geometry, TableRowSlice slice, TableCellLayout cell,
            double left, double top, double width, double height, int? rowStructure, bool header) {
            if (!slice.CellFragments!.TryGetValue(cell.Column, out NoteFlowFragment? fragment) || fragment == null) return;
            PreparedCellFlow prepared = geometry.CellFlows[slice.Row].First(flow => flow.Cell.Column == cell.Column);
            double padLeft = GetTableCellPaddingLeft(geometry.Style, slice.Row, cell.Column);
            double padTop = GetTableCellPaddingTop(geometry.Style, slice.Row, cell.Column);
            double padBottom = GetTableCellPaddingBottom(geometry.Style, slice.Row, cell.Column);
            double available = Math.Max(0D, height - padTop - padBottom);
            double unused = Math.Max(0D, available - fragment.Height);
            // Continuing stories start at the top of each new physical fragment;
            // vertical alignment applies only when the complete source cell fits.
            PdfCellVerticalAlign align = slice.StartLine == 0 && slice.CompletesCellRow
                ? GetTableCellVerticalAlignment(geometry.Style, slice.Row, cell.Column) : PdfCellVerticalAlign.Top;
            double offset = align == PdfCellVerticalAlign.Middle ? unused / 2D : align == PdfCellVerticalAlign.Bottom ? unused : 0D;
            double contentTop = top - padTop - offset + geometry.Style.RowBaselineOffset.ToPoints();
            if (slice.StartLine == 0) AddTableCellNamedDestinationName(cell.NamedDestinationName, top);
            int? structure = RegisterStructureContainer(header ? "TH" : "TD", rowStructure,
                header ? "Column" : string.Empty, cell.ColumnSpan, 1, logicalOrder: cell.Column);
            TableCellContentFrame clip = GetTableCellInkClipFrame(geometry.Style, slice.Row, cell.Column, left, top, width, height);
            BeginRoundedClip(sb, clip.Left, clip.Top - clip.Height, clip.Width, clip.Height, 0D, false, false, false, false);
            PaintFootnoteUnits(prepared.Units, fragment.Start, fragment.Count, ref contentTop, fragment.Height,
                fragment.Tables, fragment.StartTable != null, left + padLeft, structure, "P");
            EndRoundedClip(sb);
            string? uri = cell.LinkUri;
            string? destination = cell.LinkDestinationName;
            if (geometry.Table.Links.TryGetValue((slice.Row, cell.Column), out string? tableUri)) { uri = tableUri; destination = null; }
            if (HasCellLinkTarget(uri, destination)) {
                currentPage!.Annotations.Add(new LinkAnnotation {
                    X1 = left + padLeft, X2 = left + padLeft + prepared.Width, Y1 = top - height, Y2 = top,
                    Uri = uri, DestinationName = destination, Contents = cell.LinkContents ?? cell.Text,
                    StructElementIndex = structure
                });
            }
            pageDirty = true;
        }
    }
}
