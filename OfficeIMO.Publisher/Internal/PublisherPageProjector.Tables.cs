using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private void PrepareTables() {
        foreach (PublisherTableState state in _publication.Tables.Values) {
            _context.Token.ThrowIfCancellationRequested();
            _context.AccountTableModel(checked(1 + state.Columns.Count + state.Rows.Count + state.Cells.Count));
            PublisherTable table = state.Snapshot(_publication.Text, _publication.CellTextEditLimitation(state));
            _tables.Add(table.Id, table);
            foreach (PublisherTableCell cell in table.Cells) {
                if (!state.IsAuthored) _context.Record();
                long characters = Math.Max(0, cell.Paragraphs.Count - 1);
                foreach (OfficeRichTextParagraph paragraph in cell.Paragraphs) {
                    _context.Token.ThrowIfCancellationRequested();
                    characters += paragraph.Runs.Sum(run => (long)run.Text.Length);
                    if (paragraph.Label != null) characters += paragraph.Label.Run.Text.Length
                        + (paragraph.Label.FollowedBy == OfficeTextParagraphLabelFollowedBy.Nothing ? 0 : 1);
                }
                _context.AccountTableText(characters);
            }
            string location = PublisherEscherReader.ShapeLocation(table.Id);
            if (!table.StoryId.HasValue || !_publication.Text.Stories.ContainsKey(table.StoryId.Value))
                _context.Add("PUB_TEXT_REFERENCE_UNRESOLVED", "A native table has no available text story. Its grid is retained without cell text.", OfficeConversionLossKind.Omission, location);
            else if (!table.HasTextMapping)
                _context.Add("PUB_TABLE_TEXT_MAPPING_UNRESOLVED", "Recovered table text remains available in TextStories, but cannot be assigned to native cells.", OfficeConversionLossKind.Omission, location);
            if (!table.HasPlacement)
                _context.Add("PUB_TABLE_PLACEMENT_UNRESOLVED", "The native table grid and mapped cell text are retained without available page placement.", OfficeConversionLossKind.Omission, location);
            if (state.HasPlacement && !state.Hidden && !state.HasSemanticPaint
                && (state.CellFills.Any(color => color.HasValue) || state.BorderSegments.Count != 0))
                _context.Add("PUB_TABLE_CELL_STYLE_UNSUPPORTED", "Recovered cell decoration cannot be projected through this table's unsupported native graphic placement.", OfficeConversionLossKind.Omission, location);
            if (!state.IsAuthored)
                _context.Add("PUB_TABLE_CELL_STYLE_UNASSESSED", "Qualified solid cell fills and shared grid borders are retained. Other native cell appearance and cell-specific padding require additional qualification.", OfficeConversionLossKind.Unassessed, location);
            if (!table.CanEditCellText)
                _context.Add("PUB_TABLE_CELL_TEXT_EDIT_UNSUPPORTED", table.CellTextEditLimitation!, OfficeConversionLossKind.None, location);
            if (!table.CanEditBounds)
                _context.Add("PUB_TABLE_BOUNDS_EDIT_UNSUPPORTED", table.BoundsEditLimitation!, OfficeConversionLossKind.None, location);
        }
    }

    /// <summary>Projects current semantic table text. Unsupported native graphics retain their private adapter.</summary>
    private void ProjectTableText(uint id, OfficeDrawing drawing) {
        if (!_tables.TryGetValue(id, out PublisherTable? table) || !table.HasPlacement || !table.HasTextMapping) return;
        PublisherTableState state = _publication.Tables[id];
        for (int index = 0; index < table.Cells.Count; index++) {
            PublisherTableCell cell = table.Cells[index];
            if (cell.Paragraphs.Count != 0) ProjectText(_kerning.Cell(table.StoryId!.Value, index), drawing, id,
                state.Padding, cell.X, cell.Y, cell.Width, cell.Height);
        }
        _usedStories.Add(table.StoryId!.Value);
        _context.Add("PUB_TEXT_LAYOUT_APPROXIMATED", state.IsAuthored
                ? "Publication text uses shared drawing paragraph layout. Font availability, line breaking, fitting and overflow can differ from Publisher."
                : "Native text stories use the shared drawing paragraph layout. Font availability, line breaking, fitting, and overflow may differ from Publisher.",
            OfficeConversionLossKind.Approximation, state.IsAuthored ? "publication/story/" + table.StoryId : "Quill");
    }

    private void ProjectSemanticTable(PublisherTableState table, OfficeDrawing page) {
        if (!table.HasPlacement || table.Hidden) { _projected.Add(table.Id); return; }
        // Grid tracks can legitimately differ from the native frame size.
        // Preserve those logical extents, including scaled grids, in the scene.
        PublisherTable snapshot = _tables[table.Id];
        double width = Math.Max(table.Width, snapshot.Cells.Select(cell => cell.X + cell.Width).DefaultIfEmpty(0).Max());
        double height = Math.Max(table.Height, snapshot.Cells.Select(cell => cell.Y + cell.Height).DefaultIfEmpty(0).Max());
        var local = new OfficeDrawing(width, height);
        OfficeShape? paint = table.CreatePaint();
        if (paint != null && (HasFill(paint) || paint.StrokeColor.HasValue)) local.AddShape(paint, 0, 0);
        ProjectTableArtwork(snapshot, local);
        ProjectTableText(table.Id, local);
        TagSourceElements(local, new[] { "publisher-object-" + table.Id });
        if (table.X < 0 || table.Y < 0 || table.X + table.Width > page.Width || table.Y + table.Height > page.Height)
            _context.Add("PUB_PAGE_BLEED_CLIPPED", "Artwork outside the publication page remains in the scene but is clipped to the printable page in exports.",
                OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(table.Id));
        _context.AccountProjection(local);
        if (local.Elements.Count != 0) page.AddDrawingForClippedRendering(local, table.X, table.Y, null);
        _projected.Add(table.Id);
    }
}
