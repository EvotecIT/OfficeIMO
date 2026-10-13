using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Builds the Numbers document, sheet/sidebar, table and editable empty header/footer graph.</summary>
internal sealed partial class IWorkNumbersArchiveBuilder : IWorkNativeCreationGraph {
    private const ulong DocumentId = 1, MetadataId = 2, SidebarId = 3, ThemeId = 4;
    private const ulong CalculationId = 6, NamesId = 7, TrackerId = 8, CustomFormatsId = 9, FunctionsId = 10;
    private const ulong SheetStyleId = 11, CharacterId = 12, ParagraphId = 13, FrameId = 15, PresetId = 16, ImageStyleId = 17;

    internal IWorkNumbersArchiveBuilder(int maximumBytes, CancellationToken token, byte[] hash)
        : base(maximumBytes, token, hash, IWorkDocumentKind.Numbers, 5, 14) { }

    internal string DocumentIdentifier => IWorkCreationIdentity.Create(_modelHash, "document", _kind).Text;

    internal (byte[] Document, byte[] Metadata) Build(IReadOnlyList<(string Name, IWorkNativeTableData[] Tables)> sheets) {
        var root = P().Reference(4, StylesheetId).Reference(5, SidebarId).Reference(6, ThemeId)
            .Message(8, P().Message(1, P().String(4, "en_US")).String(3, "en")
                .Reference(4, CalculationId).Reference(6, FunctionsId).Reference(12, CustomFormatsId))
            .String(11, "na-letter").Message(12, Size(612, 792));
        Add(NamesId, 4003, P().Reference(1, TrackerId));
        var tracker = IWorkCreationIdentity.Create(_modelHash, "reference-tracker", _kind);
        Add(TrackerId, 4004, P().Message(1, P().UInt(1, tracker.Upper >> 32).UInt(2, (uint)tracker.Upper)
            .UInt(3, tracker.Lower >> 32).UInt(4, (uint)tracker.Lower)));
        Add(CustomFormatsId, 222, P()); Add(FunctionsId, 601, P().UInt(4, 0));
        AddTextStyles(CharacterId, ParagraphId, FrameId, "Arial", 12, new IWorkColor(0, 0, 0, 255));
        var list = P().Message(1, Style(ListId, "None")).UInt(10, 3);
        for (int index = 0; index < 9; index++) list.UInt(11, 0).Float(12, 0).Float(13, 0);
        Add(ListId, 2023, list);
        Add(PresetId, 2050, P().String(1, "officeimo-text").Reference(2, ParagraphId).Reference(3, ListId));
        Add(SheetStyleId, 12050, P().Message(1, Style(SheetStyleId, "OfficeIMO Sheet")).UInt(2, 1)
            .Message(3, P().Message(1, P().Message(1, Color(new IWorkColor(255, 255, 255, 255))))));
        Add(ImageStyleId, 3016, P().Message(1, Style(ImageStyleId, "OfficeIMO Image")).UInt(10, 2)
            .Message(11, P().Float(2, 1).Message(1, P().Message(6, P().UInt(1, 2)))));
        var sidebar = P();
        foreach (var input in sheets) {
            _cancellationToken.ThrowIfCancellationRequested();
            ulong sheetId = Allocate(), sheetNodeId = Allocate();
            var sheet = P().String(1, input.Name).Bool(3, true).Bool(4, false).Bool(5, false).Bool(6, false)
                .Float(7, 1).UInt(8, 0).Message(10, P().Float(1, 36).Float(2, 36).Float(3, 36).Float(4, 36))
                .Bool(11, false).UInt(12, 1).Float(13, 0).Float(14, 0).Bool(20, false)
                .UInt(21, 0).Reference(22, SheetStyleId).Bool(23, true).Bool(24, false);
            var sheetNode = P().Reference(3, sheetId);
            foreach (IWorkNativeTableData table in input.Tables) {
                ulong tableId = AddTable(sheetId, table, out var identifier), tableNodeId = Allocate();
                _calculationTables.Add((tableId, identifier.Lower, identifier.Upper, table.RowCount, table.ColumnCount));
                sheet.Reference(2, tableId); sheetNode.Reference(2, tableNodeId);
                Add(tableNodeId, 205, P().Reference(3, tableId));
            }
            // All three header/footer positions are native storages, even though print page numbering is disabled.
            for (int position = 0; position < 6; position++) {
                ulong storageId = Allocate(); sheet.Reference(position < 3 ? 18 : 19, storageId);
                Add(storageId, 2001, P().UInt(1, 1).Reference(2, StylesheetId).String(3, "\n").Bool(10, true)
                    .Message(5, Attribute(ParagraphId)).Message(7, Attribute(ListId)).Message(8, Attribute(CharacterId))
                    .Message(6, EmptyParagraphData()).Message(14, EmptyParagraphData()).Message(24, EmptyParagraphData())
                    .Message(28, P().Message(1, P().UInt(1, 0))));
            }
            Add(sheetId, 2, sheet); Add(sheetNodeId, 205, sheetNode);
            root.Reference(1, sheetId); sidebar.Reference(2, sheetNodeId);
        }
        AddCalculationGraph();
        Add(DocumentId, 1, root); Add(SidebarId, 205, sidebar);
        var theme = P().String(3, "OfficeIMO").Reference(4, StylesheetId)
            .Message(100, P().Reference(4, FrameId).Reference(5, FrameId).Reference(6, FrameId)
                .Reference(7, ImageStyleId).Reference(9, FrameId))
            .Message(110, P().Reference(1, ListId).Reference(2, PresetId).Reference(6, CharacterId).Reference(7, ParagraphId));
        AddTableThemePresets(theme); IWorkNativeThemeWriter.AddColorPresets(theme, _maximumBytes, _cancellationToken);
        Add(ThemeId, 12009, P().Message(1, theme));
        var stylesheet = P().Bool(4, false);
        foreach (var style in _styles)
            stylesheet.Reference(1, style.Id).Message(2, P().String(1, style.Identifier).Reference(2, style.Id));
        Add(StylesheetId, 401, stylesheet);
        var component = P().UInt(1, DocumentId).String(2, "Document");
        Version(component, 4, 2, 4, 0); Version(component, 5, 11, 2, 4);
        var metadata = P().UInt(1, _nextId - 1).Message(3, component).UInt(9, 2);
        Version(metadata, 5, 2, 4, 0); Version(metadata, 6, 11, 2, 4); Version(metadata, 7, 14, 4, 1);
        return (Compress(_records.OrderBy(record => record.Id)), Compress(new[] { (MetadataId, 11006u, metadata) }));
    }

    private IWorkProtoWriter Attribute(ulong id) => P().Message(1, P().UInt(1, 0).Reference(2, id));
    private IWorkProtoWriter EmptyParagraphData() => P().Message(1, P().UInt(1, 0).UInt(2, 0).UInt(3, 0));
    private byte[] Compress(IEnumerable<(ulong Id, uint Type, IWorkProtoWriter Message)> records) =>
        IWorkSnappy.EncodeIwa(IWorkNativeArchiveWriter.Encode(records, _maximumBytes, _cancellationToken), _maximumBytes, _cancellationToken);
}
