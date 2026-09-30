using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    /// <summary>Checks the selected tile envelope before parsing row payloads; ordinary malformed bytes remain recoverable.</summary>
    private static IWorkWireMessage? ReadTileMessage(IWorkSourceDocument source, IWorkArchiveRecord tile,
        int maximumRows, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        int declaredRows;
        int totalFields;
        try {
            declaredRows = IWorkProtobuf.CountFields(tile.Payload, 5,
                source.Options.MaximumProtobufFieldCount, out totalFields);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(tile, "$", null);
            MarkTileUnsupported(tile, "IWORK_TABLE_TILE_UNSUPPORTED",
                "A selected iWork table tile payload is malformed; editable reconstruction is incomplete.",
                diagnostics, ref supportsEditableReconstruction);
            return null;
        }
        if (totalFields - declaredRows > MaximumTileMetadataFields) {
            references.Declarations.Record(tile, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            MarkTileUnsupported(tile, "IWORK_TABLE_TILE_FIELDS_UNSUPPORTED",
                "An iWork table tile contains more metadata fields than the supported tile envelope; editable reconstruction is incomplete.",
                diagnostics, ref supportsEditableReconstruction);
            return null;
        }
        if (declaredRows > maximumRows) {
            references.Declarations.Record(tile, "5", declaredRows, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            MarkTileUnsupported(tile, "IWORK_TABLE_TILE_ROW_COUNT_UNSUPPORTED",
                "An iWork table tile declares more row messages than can fit in its logical table range; editable reconstruction is incomplete.",
                diagnostics, ref supportsEditableReconstruction);
            return null;
        }
        // CountFields already checked the complete wire envelope. Parse still enforces
        // its own configured limits, and those failures must not become partial content.
        IWorkWireMessage message = source.Index.Message(tile);
        if (HasUnsupportedTileMetadata(message, totalFields - declaredRows)) {
            references.Declarations.Record(tile, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            MarkTileUnsupported(tile, "IWORK_TABLE_TILE_FIELDS_UNSUPPORTED",
                "An iWork table tile contains duplicate, unknown, or malformed metadata fields; editable reconstruction is incomplete.",
                diagnostics, ref supportsEditableReconstruction);
            return null;
        }
        return message;
    }

    /// <summary>Recovers readable row messages while retaining one-based physical positions across unreadable siblings.</summary>
    private static IReadOnlyList<(IWorkWireMessage Message, int Position)> ReadTileRows(
        IWorkSourceDocument source, IWorkArchiveRecord tile, IWorkWireMessage message,
        IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        var rows = new List<(IWorkWireMessage, int)>();
        bool malformed = false;
        int position = 0;
        foreach (IWorkWireValue value in message.EnumerateValues(5)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            position++;
            IWorkWireMessage? row = null;
            if (value.Kind == IWorkWireKind.Bytes && value.Bytes != null) {
                try {
                    row = message.ParseNestedMessage(value.Bytes);
                } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) { }
            }
            if (row == null) {
                malformed = true;
                references.Declarations.Record(tile, TileRowPath(position), 1);
            } else {
                rows.Add((row, position));
            }
        }
        if (malformed) MarkTileUnsupported(tile, "IWORK_TABLE_TILE_ROWS_UNSUPPORTED",
            "An iWork table tile contains malformed row metadata; readable sibling rows are retained, but editable reconstruction is incomplete.",
            diagnostics, ref supportsEditableReconstruction);
        return rows;
    }

    private static string TileRowPath(int position) => "5[" + position.ToString(CultureInfo.InvariantCulture) + "]";

    private static void RecordInvalidTileRow(IWorkArchiveRecord tile, int position,
        IWorkSourceReferenceIssueCollector references) =>
        references.Declarations.Record(tile, TileRowPath(position), 1,
            IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);

    private static void RecordInvalidTileEntry(IWorkArchiveRecord model, int position,
        IWorkSourceReferenceIssueCollector references) =>
        references.Declarations.Record(model, "4/3/1[" + position.ToString(CultureInfo.InvariantCulture) + "]", 1,
            IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);

    private static void MarkTileUnsupported(IWorkArchiveRecord tile, string code, string message,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, code, message,
            tile.EntryPath, tile.Identifier));
    }
}
