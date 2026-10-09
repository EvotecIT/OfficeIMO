using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkKeynoteArchiveBuilder {
    private readonly List<ulong> _tablePresets = new();

    private ulong AddTable(ulong slideId, IWorkKeynoteTable table) {
        ulong infoId = Allocate(), modelId = Allocate(), tableStyleId = Allocate(), cellStyleId = Allocate();
        ulong characterId = Allocate(), paragraphId = Allocate(), frameStyleId = Allocate();
        ulong rowHeaderId = Allocate(), columnHeaderId = Allocate(), tileId = Allocate(), stringsId = Allocate();
        AddTextStyles(characterId, paragraphId, frameStyleId, table.FontName, table.FontSizePoints, table.Color);
        AddTableStyles(tableStyleId, cellStyleId, table.FontName);
        ulong presetId = AddTablePreset(tableStyleId, cellStyleId, paragraphId, frameStyleId);
        var drawable = P().Message(1, P().Message(1, Point(table.Geometry.LeftPoints, table.Geometry.TopPoints))
            .Message(2, Size(table.Geometry.WidthPoints, table.Geometry.HeightPoints)).UInt(3, 3).Float(4, 0)).Reference(2, slideId);
        Add(infoId, 6000, P().Message(1, drawable).Reference(2, modelId).UInt(10, 1));

        var stringKeys = new Dictionary<string, uint>(StringComparer.Ordinal);
        var counts = new Dictionary<string, uint>(StringComparer.Ordinal);
        foreach (IWorkTableCell cell in table.Cells) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (cell.Kind != IWorkCellKind.Text) continue;
            string value = (string)cell.Value!;
            if (!stringKeys.ContainsKey(value)) { stringKeys.Add(value, (uint)stringKeys.Count + 1); counts.Add(value, 0); }
            counts[value]++;
        }
        var strings = P().UInt(1, 1).UInt(2, (uint)stringKeys.Count + 1);
        foreach (var item in stringKeys.OrderBy(item => item.Value))
            strings.Message(3, P().UInt(1, item.Value).UInt(2, counts[item.Key]).String(3, item.Key));
        Add(stringsId, 6005, strings);

        Add(rowHeaderId, 6006, Headers(table.RowCount, (float)(table.Geometry.HeightPoints / table.RowCount), table.ColumnCount));
        Add(columnHeaderId, 6006, Headers(table.ColumnCount, (float)(table.Geometry.WidthPoints / table.ColumnCount), table.RowCount));
        Add(tileId, 6002, ScalarTile(table, stringKeys));
        var store = P().Message(1, P().UInt(1, 1).Reference(2, rowHeaderId)).Reference(2, columnHeaderId)
            .Message(3, P().Message(1, P().UInt(1, 0).Reference(2, tileId)).UInt(2, 256))
            .Reference(4, stringsId).UInt(7, 1).UInt(8, 0)
            .Message(9, P().Message(1, P().UInt(1, 0).UInt(2, 0))).Message(10, P()).UInt(14, 4);
        // Native editing allocates from these typed catalogs even when the table has no corresponding feature.
        foreach (var catalog in new[] { (5, 4u), (6, 3u), (11, 2u), (12, 5u), (16, 7u),
                     (17, 8u), (18, 9u), (19, 10u), (20, 11u), (21, 12u), (22, 2u) }) {
            ulong id = Allocate();
            Add(id, 6005, P().UInt(1, catalog.Item2).UInt(2, 1));
            store.Reference(catalog.Item1, id);
        }
        var model = P().String(1, IWorkKeynoteIdentity.Create(_modelHash, "table-" + modelId.ToString(CultureInfo.InvariantCulture)).Text)
            .Reference(3, tableStyleId).Message(4, store).UInt(6, (uint)table.RowCount).UInt(7, (uint)table.ColumnCount)
            .String(8, table.Name).UInt(9, 0).UInt(10, 0).UInt(11, 0).Bool(12, false).Bool(13, false)
            .Double(16, table.Geometry.HeightPoints / table.RowCount).Double(17, table.Geometry.WidthPoints / table.ColumnCount)
            .Bool(22, false).Bool(29, false).Bool(31, false).Bool(32, false).Double(33, 0)
            .Reference(30, paragraphId).Reference(36, frameStyleId).Reference(48, presetId);
        foreach (int field in new[] { 18, 19, 20, 21 }) model.Reference(field, cellStyleId);
        foreach (int field in new[] { 24, 25, 26, 27 }) model.Reference(field, paragraphId);
        Add(modelId, 6001, model);
        return infoId;
    }

    private ulong AddTablePreset(ulong tableStyleId, ulong cellStyleId, ulong paragraphId, ulong frameStyleId) {
        ulong networkId = Allocate(), presetId = Allocate();
        var network = P().Reference(9, tableStyleId).Reference(10, paragraphId).Reference(11, frameStyleId)
            .UInt(12, (uint)_tablePresets.Count);
        for (int field = 1; field <= 4; field++) network.Reference(field, paragraphId);
        for (int field = 5; field <= 8; field++) network.Reference(field, cellStyleId);
        Add(networkId, 6247, network);
        Add(presetId, 6008, P().UInt(1, (uint)_tablePresets.Count).Reference(3, networkId));
        _tablePresets.Add(presetId);
        return presetId;
    }

    private void AddTableThemePresets(IWorkProtoWriter theme) {
        if (_tablePresets.Count == 0) return;
        var presets = P();
        foreach (ulong id in _tablePresets) presets.Reference(1, id);
        // Keynote's native table presets use extension 200; extension 120 owns chart presets.
        theme.Message(200, presets);
    }

    private IWorkProtoWriter Headers(int count, float size, int cellsPerHeader) {
        var headers = P().UInt(1, 1);
        for (int index = 0; index < count; index++)
            headers.Message(2, P().UInt(1, (uint)index).Float(2, size).UInt(3, 0).UInt(4, (uint)cellsPerHeader));
        return headers;
    }

    private IWorkProtoWriter ScalarTile(IWorkKeynoteTable table, IReadOnlyDictionary<string, uint> strings) {
        var tile = P().UInt(1, (uint)table.ColumnCount - 1).UInt(2, (uint)table.RowCount - 1)
            .UInt(3, (uint)(table.RowCount * table.ColumnCount)).UInt(4, (uint)table.RowCount)
            .UInt(6, 5).Bool(7, true).Bool(8, false);
        var cells = table.Cells.ToDictionary(cell => (cell.Row - 1) * table.ColumnCount + cell.Column - 1);
        for (int row = 0; row < table.RowCount; row++) {
            _cancellationToken.ThrowIfCancellationRequested();
            var offsets = Enumerable.Repeat((byte)255, 512).ToArray();
            var legacyOffsets = (byte[])offsets.Clone();
            using var buffer = new MemoryStream();
            var legacy = new byte[table.ColumnCount * 12];
            for (int column = 0; column < table.ColumnCount; column++) {
                _cancellationToken.ThrowIfCancellationRequested();
                IWorkTableCellStorageWriter.WriteOffset(offsets, column, (int)buffer.Length);
                IWorkTableCellStorageWriter.WriteOffset(legacyOffsets, column, column * 12);
                // Apple marks pre-BNC cells empty while the modern buffer owns their values.
                legacy[column * 12] = 4;
                cells.TryGetValue(row * table.ColumnCount + column, out var cell);
                byte[] bytes = IWorkTableCellStorageWriter.Encode(cell, strings);
                buffer.Write(bytes, 0, bytes.Length);
            }
            tile.Message(5, P().UInt(1, (uint)row).UInt(2, (uint)table.ColumnCount).Bytes(3, legacy)
                .Bytes(4, legacyOffsets).UInt(5, 5).Bytes(6, buffer.ToArray()).Bytes(7, offsets).Bool(8, false));
        }
        return tile;
    }

    private void AddTableStyles(ulong tableStyleId, ulong cellStyleId, string fontName) {
        var properties = P().Bool(1, false).Bool(21, false).Bool(22, false).Bool(33, true).Bool(34, true)
            .Bool(35, true).Bool(36, true).Bool(37, true).Bool(38, true).Bool(39, true).String(41, fontName);
        var stroke = P().Message(1, Color(new IWorkColor(0, 0, 0, 255))).Float(2, 1)
            .UInt(3, 0).UInt(4, 0).Float(5, 10).Message(6, P().UInt(1, 1));
        for (int field = 46; field <= 61; field++) properties.Message(field, stroke);
        Add(tableStyleId, 6003, P().Message(1, Style(tableStyleId, "OfficeIMO Table")).UInt(10, 10).Message(11, properties));
        var cell = P().Message(1, P().Message(1, Color(new IWorkColor(255, 255, 255, 255))))
            .Bool(3, true).UInt(8, 0).Message(9, P().Float(1, 4).Float(2, 4).Float(3, 4).Float(4, 4));
        Add(cellStyleId, 6004, P().Message(1, Style(cellStyleId, "OfficeIMO Cell")).UInt(10, 4).Message(11, cell));
    }
}
