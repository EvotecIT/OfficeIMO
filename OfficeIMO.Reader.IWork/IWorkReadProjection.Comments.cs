using OfficeIMO.IWork;
using OfficeIMO.Spreadsheet;

namespace OfficeIMO.Reader.IWork;

internal sealed partial class IWorkReadProjection {
    /// <summary>Projects qualified comments only for retained grid cells, keeping source coordinates across flattened headers.</summary>
    private void AddTableComments(OfficeDocumentPage page, IWorkTable source, int? tableIndex,
        int columnCount, int materializedHeaderRows, int headerRows, int dataRows) {
        int omitted = 0;
        foreach (IWorkTableCell cell in source.Cells) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (cell.Comment is not { } comment) continue;
            if (cell.Column > columnCount || cell.Row > headerRows + dataRows
                || cell.Row > materializedHeaderRows && cell.Row <= headerRows) {
                omitted++;
                continue;
            }

            // Bound source-controlled metadata before formatting or chunking; repeated references are charged separately.
            _projectionBudget.AddTextItem();
            _projectionBudget.AddTextCharacters(comment.Text.Length);
            _projectionBudget.AddTextCharacters(comment.Author.Length);
            _projectionBudget.AddTextCharacters(comment.SourceIdentity.EntryPath.Length);
            _projectionBudget.AddTextCharacters(comment.SourceAuthorIdentity.EntryPath.Length);
            string address = SpreadsheetRangeReference.FromCell(column: cell.Column, row: cell.Row)
                .Format(SpreadsheetAddressDialect.UnboundedA1);
            string timestamp = comment.CreationDateUtc.ToString("O", CultureInfo.InvariantCulture);
            string heading = $"Comment on table {tableIndex!.Value + 1}, cell {address} by {comment.Author} ({timestamp})";
            string prefix = "**" + EscapeMarkdown(heading, _cancellationToken) + "**\n\n";
            ReaderLocation location = AddBlock(page, "comment", comment.Text,
                prefix + EscapeMarkdown(comment.Text, _cancellationToken), null, null,
                sourceKind: "table-cell-comment", tableIndex: tableIndex, a1Range: address,
                splitMarkdownIndependently: true);
            _projectionMetadata.Add(new OfficeDocumentMetadataEntry {
                Id = location.BlockAnchor + "-comment",
                Category = "table.comment",
                Name = "RootComment",
                Value = comment.Text,
                ValueType = "string",
                SourceObjectId = comment.SourceIdentity.RecordIdentifier.ToString(CultureInfo.InvariantCulture),
                Location = location,
                Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["author"] = comment.Author,
                    ["creationDateUtc"] = timestamp,
                    ["sourceEntryPath"] = comment.SourceIdentity.EntryPath,
                    ["sourcePayloadIndex"] = comment.SourceIdentity.PayloadIndex.ToString(CultureInfo.InvariantCulture),
                    ["authorRecordIdentifier"] = comment.SourceAuthorIdentity.RecordIdentifier.ToString(CultureInfo.InvariantCulture),
                    ["authorEntryPath"] = comment.SourceAuthorIdentity.EntryPath,
                    ["authorPayloadIndex"] = comment.SourceAuthorIdentity.PayloadIndex.ToString(CultureInfo.InvariantCulture)
                }
            });
        }
        if (omitted == 0) return;
        ReaderLocation omittedLocation = Location(page);
        omittedLocation.TableIndex = tableIndex;
        _diagnostics.Add(new OfficeDocumentDiagnostic {
            Category = OfficeDocumentDiagnosticCategory.Limit,
            Code = "IWORK_READER_TABLE_COMMENTS_OMITTED",
            Message = $"Table '{source.Name}' has {omitted} qualified cell comments outside the projected grid; their cells exceed Reader table limits.",
            Source = "OfficeIMO.Reader.IWork",
            Location = omittedLocation,
            Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["tableName"] = source.Name,
                ["omittedCommentCount"] = omitted.ToString(CultureInfo.InvariantCulture)
            }
        });
    }
}
