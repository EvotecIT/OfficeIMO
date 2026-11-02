using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Sets the opaque solid fill, or clears the fill, of the cell anchored at the given zero-based coordinates.</summary>
    /// <remarks>Covered merged positions are rejected. Existing snapshots remain unchanged. Native source saving currently supports recoloring an existing visible solid fill; changing visibility or inserting paint records is refused. Validation, cancellation or projection failure leaves the document unchanged.</remarks>
    public void SetTableCellFillColor(uint tableId, int rowIndex, int columnIndex, OfficeColor? color,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetTableCellFillColor(tableId, rowIndex, columnIndex, color, cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Changes the opaque solid color and positive width of one existing shared border, selected by its exact grid-boundary endpoints.</summary>
    /// <remarks>The selected range must match exactly one entry in BorderSegments. This operation retains its ownership and endpoints; adding, splitting or removing borders is not supported. Width uses points. Native saving requires an existing unambiguous local color and width field and an exact signed-EMU width. Failed edits leave the document unchanged.</remarks>
    public void SetTableBorderStyle(uint tableId, int firstRow, int firstColumn, int lastRow, int lastColumn,
        OfficeBorderSide side, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetTableBorderStyle(tableId, firstRow, firstColumn, lastRow, lastColumn, side, cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Merges a rectangular range of complete cells in an authored table, using zero-based coordinates and positive row and column counts.</summary>
    /// <remarks>Every intersecting cell must fit wholly inside the requested range. Styled paragraphs from the selected cells are retained in their existing definition order in the new anchor cell; other cells and grid tracks remain intact. A range already occupied by one complete cell is unchanged. Imported native table topology is not editable through this operation. Failed validation or cancellation leaves the document unchanged.</remarks>
    public void MergeTableCells(uint tableId, int firstRow, int firstColumn, int rowCount, int columnCount,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.MergeTableCells(tableId, firstRow, firstColumn, rowCount, columnCount, cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Replaces the styled paragraphs of the cell anchored at the given zero-based grid coordinates. Other cells, merged ownership and the complete story remain intact.</summary>
    /// <remarks>Compatible tables sharing the story update together. Covered merged coordinates, unresolved partitions and stories also used by ordinary frames are rejected. Clearing retains the cell boundary, final paragraph formatting and terminator convention. Qualified imported tables support bounded native text save-back with a fixed grid. Paragraph counts can change from uniform source definitions or by retaining unlabelled native paragraphs without changing their text/formatting. Original paragraph snapshots select their exact native definitions; copied values require unambiguous definitions within that cell. Retained empty paragraphs keep their complete source CR definition in interior and final positions. Count-changing native edits resolve exact snapshots and values first, then retained ordered run text, complete paragraph text with preserved fields at each position, then unchanged complete formatting. Correspondence must be unambiguous across complete native templates. Retained-text run splitting/merging maps source, requested and native character ranges while preserving opaque fields and destination CR typing. Reload may expose separate native ranges with identical visible formatting. Changed-text replacement with retained ordered runs requires a uniform native definition for each corresponding run. During paragraph-count changes, a genuine empty source run can supply its native CR definition for one or more nonempty styled runs. Existing correspondences keep priority; empty-source matching prefers retained paragraph formatting. All eligible templates must agree after complete body, CR and paragraph rewriting. The last requested run supplies the refilled CR typing; unspecified language and kerning retain their source fields. With the paragraph count unchanged, a nonempty multi-run paragraph can replace its complete text and supported visible formatting with a different run count when every contributing native body definition agrees after rewriting. Paragraph-count changes also support complete nonempty replacement after stronger snapshot/text/format/empty-source matches. The fallback requires agreement across rewritten body, paragraph and destination-role CR definitions. Interior CRs retain qualified interior typing; final destinations retain the original final CR, with only existing explicit nullable-attribute changes. Unspecified language and kerning retain their source fields; opaque and unsupported preserved fields must agree. Paragraph and CR ownership remain unchanged. Other formatting edits use the qualified character/paragraph codecs; requests without safe native correspondence refuse saving. Neighboring cells may have different styles. Empty paragraphs without runs retain native CR typing; an empty styled run requests its supplied typing style. Native reload exposes a styled empty CR run. Unsupported native definitions still refuse output with loss-policy allowance. Existing objects remain snapshots. Validation, cancellation or projection failure leaves the document unchanged; concurrent mutation is unsupported.</remarks>
    public void ReplaceTableCellParagraphs(uint tableId, int rowIndex, int columnIndex,
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, CancellationToken cancellationToken = default) {
        if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.ReplaceTableCell(tableId, rowIndex, columnIndex, paragraphs, cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Moves and proportionally resizes a qualified table in page-local points, scaling its column and row tracks by the frame's width and height ratios.</summary>
    /// <remarks>Dimensions must be positive and finite; finite negative offsets allow clipped bleed. Placement eligibility is independent of cell-text eligibility. Unsupported native graphics and transformed, hidden or unresolved placement reject bounds edits. Qualified native source saving supports direct movement and proportional fixed-track resizing with exact signed-EMU coordinates and dimensions. Resizing requires complete matching nominal and laid-out tracks; unsupported wrap relationships refuse saving. Native reflow remains unassessed and edited output requires explicit Allow. Complete text is retained on overflow. Failed edits leave the publication unchanged.</remarks>
    public void SetTableBounds(uint tableId, double x, double y, double width, double height, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetTableBounds(tableId, x, y, width, height).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Appends an empty, unmerged, borderless table to an existing qualified document or master page. Track sizes and offsets use points. Returns the current immutable snapshot.</summary>
    /// <remarks>Track lists are copied and must be nonempty, positive and finite. Populate cells through ReplaceTableCellParagraphs. A new complete story and the table use the publication's shared allocator. Validation, cancellation or projection failure consumes no identifiers. This is in-memory creation and does not provide native PUB saving or new publication/page creation.</remarks>
    public PublisherTable AddTable(uint pageId, double x, double y, IReadOnlyList<double> columnWidths, IReadOnlyList<double> rowHeights,
        CancellationToken cancellationToken = default) {
        if (columnWidths == null) throw new ArgumentNullException(nameof(columnWidths));
        if (rowHeights == null) throw new ArgumentNullException(nameof(rowHeights));
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        PublisherDocument candidate = model.AddTable(pageId, x, y, columnWidths, rowHeights, cancellationToken, out uint id).Project(cancellationToken);
        PublisherTable result = candidate.Pages.Concat(candidate.MasterPages).SelectMany(page => page.Tables).Single(table => table.Id == id);
        Commit(candidate, cancellationToken);
        return result;
    }
}
