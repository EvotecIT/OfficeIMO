using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherModelState {
    internal PublisherModelState SetTableCellFillColor(uint id, int row, int column, OfficeColor? color, CancellationToken token) =>
        WithPublication(_publication.SetTableCellFillColor(id, row, column, color, token));

    internal PublisherModelState SetTableBorderStyle(uint id, int firstRow, int firstColumn, int lastRow, int lastColumn,
        OfficeBorderSide? side, CancellationToken token) =>
        WithPublication(_publication.SetTableBorderStyle(id, firstRow, firstColumn, lastRow, lastColumn, side, _decoded.Options.Limits.MaxItems, token));

    internal PublisherModelState MergeTableCells(uint id, int row, int column, int rowCount, int columnCount, CancellationToken token) =>
        WithPublication(_publication.MergeTableCells(id, row, column, rowCount, columnCount, _decoded.Options, token));

    internal PublisherModelState ReplaceTableCell(uint tableId, int row, int column,
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, CancellationToken token) =>
        WithPublication(_publication.ReplaceTableCell(tableId, row, column, paragraphs, _decoded.Options, token));

    internal PublisherModelState SetTableBounds(uint id, double x, double y, double width, double height) =>
        WithPublication(_publication.SetTableBounds(id, x, y, width, height));

    internal PublisherModelState AddTable(uint pageId, double x, double y, IReadOnlyList<double> columns, IReadOnlyList<double> rows,
        CancellationToken token, out uint tableId) {
        token.ThrowIfCancellationRequested();
        if ((long)SourceObjectCount + _publication.AuthoredObjectCount >= _decoded.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher current object limit exceeded.");
        if (columns.Count > _decoded.Options.Limits.MaxItems || rows.Count > _decoded.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher table track limit exceeded.");
        return WithPublication(_publication.AddTable(pageId, x, y, columns.ToArray(), rows.ToArray(), _decoded.Options, token, out tableId));
    }
}
