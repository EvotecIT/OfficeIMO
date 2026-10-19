using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    /// <summary>Uses the native writer's cell and style encoders while retaining all existing reference indices.</summary>
    internal static LegacyXlsEditEncoding EncodePreservedEdits(ExcelDocument document, LegacyXlsWorkbook source,
        IReadOnlyList<BiffRecord> records, IReadOnlyDictionary<string, int> nativeSheets,
        IReadOnlyList<LegacyXlsCellEdit> edits, uint firstStringIndex, CancellationToken token, LegacyXlsSheetStructure? structure = null) {
        var sheets = document.Sheets.ToArray();
        var fonts = LegacyXlsFontTable.CreatePreserving(document, source);
        if (source.Fonts.Count != records.Count(record => record.Type == 0x0031))
            throw new NotSupportedException("The original font table is incomplete.");
        var styles = LegacyXlsStyleTable.CreatePreserving(document, sheets, fonts, records);
        var externs = LegacyXlsExternSheetTable.CreatePreserving(document, sheets, records, nativeSheets);
        var globals = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
        var locals = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
        uint nameIndex = 0;
        foreach (var record in records.Where(record => record.Type == 0x0018)) {
            nameIndex++;
            if (!LegacyBiffWorkbookParser.TryReadDefinedNameTableEntry(record.Payload, out string? name)) continue;
            int nativeScope = BiffRecordReader.ReadUInt16(record.Payload, 8) - 1;
            int projectedScope = Array.FindIndex(sheets, sheet => nativeSheets[sheet.Name] == nativeScope);
            if (nativeScope < 0) globals[name!] = nameIndex;
            else if (projectedScope >= 0) locals[LegacyXlsFormulaNameIndex.CreateLocalKey(projectedScope, name!)] = nameIndex;
        }
        var names = new LegacyXlsFormulaNameIndex(globals, locals, externs);
        var result = new LegacyXlsEditEncoding(externs);
        var dummyStrings = new LegacyXlsSharedStringTable(Array.Empty<LegacyXlsSharedStringEntry>(), 0);
        edits = ExpandFormulaGroups(document, source, records, edits, token, structure);
        var convertedBySheet = new Dictionary<string, List<LegacyXlsCell>>(StringComparer.Ordinal);
        foreach (var group in edits.GroupBy(edit => edit.Sheet)) {
            int index = Array.FindIndex(sheets, sheet => sheet.Name == group.Key);
            if (index < 0) throw new NotSupportedException("The edited worksheet is missing.");
            var sheet = sheets[index];
            convertedBySheet[group.Key] = new();
            var shared = LegacyXlsSharedFormulaTable.Create(sheet, index, names);
            var arrays = LegacyXlsArrayFormulaTable.Create(sheet, index, names);
            foreach (var edit in group) {
                token.ThrowIfCancellationRequested();
                if (edit.Cell == null) { result.Cells.Add((edit.Sheet, edit.Address), Array.Empty<BiffRecord>()); continue; }
                ParseCellReference(edit.Address, out uint row, out int column);
                if (row == 0 || row > 65536 || column < 1 || column > 256)
                    throw new NotSupportedException("Native XLS cells require rows 1–65536 and columns 1–256.");
                var converted = ConvertCell(sheet, document.WorkbookPartRoot, index, edit.Cell,
                    checked((ushort)(row - 1)), checked((ushort)(column - 1)), styles, document.DateSystem, names, shared, arrays, fonts);
                if (!converted.HasValue) { result.Cells.Add((edit.Sheet, edit.Address), Array.Empty<BiffRecord>()); continue; }
                var cell = converted.Value;
                if (cell.Kind == LegacyXlsCellKind.Text) {
                    cell = cell.WithSharedStringIndex(checked(firstStringIndex + (uint)result.Strings.Count));
                    result.Strings.Add(new LegacyXlsPreservedString(cell.TextValue ?? string.Empty, cell.TextFormattingRuns));
                }
                convertedBySheet[group.Key].Add(cell);
                using var output = new MemoryStream();
                WriteCellRecords(output, index, new[] { cell }, dummyStrings);
                result.Cells.Add((edit.Sheet, edit.Address), BiffRecordReader.ReadRecords(output.ToArray(), new(), token));
            }
        }
        if (structure != null) foreach (var binding in structure.Bindings.Where(binding => binding.Original == null)) {
            if (binding.Part is not DocumentFormat.OpenXml.Packaging.WorksheetPart) throw new NotSupportedException("Only new worksheets can be added through native legacy editing.");
            int index = Array.FindIndex(sheets, sheet => sheet.Name == binding.Name);
            var sheet = sheets[index];
            LegacyXlsWritePreflight.ThrowIfUnsupported(sheet, index, document.WorkbookPartRoot, fonts, names);
            if (sheet.WorksheetPart.Parts.Any()) throw new NotSupportedException("New native worksheets cannot introduce relationship-backed content into a retained workbook graph.");
            var cells = convertedBySheet.TryGetValue(binding.Name, out var existing) ? existing : new List<LegacyXlsCell>();
            cells.Sort((left, right) => left.Row == right.Row ? left.Column.CompareTo(right.Column) : left.Row.CompareTo(right.Row));
            using var output = new MemoryStream();
            WriteWorksheet(output, document, sheet, index, cells, dummyStrings, styles, document.DateSystem, names, fonts);
            var generated = BiffRecordReader.ReadRecords(output.ToArray(), new(), token).ToList();
            if (structure.NewCodeNames.TryGetValue(binding.Name, out string? codeName) && !generated.Any(record => record.Type == 0x01ba))
                generated.Insert(generated.Count - 1, new BiffRecord(0x01ba, -1, BuildCodeNamePayload(codeName)));
            result.NewWorksheets.Add(binding.Name, generated);
        }
        result.Fonts.AddRange(fonts.FontRecords.Skip(source.Fonts.Count));
        result.Formats.AddRange(styles.FormatRecords);
        result.Styles.AddRange(styles.CellFormatRecords);
        return result;
    }

    private static IReadOnlyList<LegacyXlsCellEdit> ExpandFormulaGroups(ExcelDocument document, LegacyXlsWorkbook source,
        IReadOnlyList<BiffRecord> records, IReadOnlyList<LegacyXlsCellEdit> edits, CancellationToken token, LegacyXlsSheetStructure? structure) {
        var result = edits.ToDictionary(edit => (edit.Sheet, edit.Address));
        foreach (var sheet in source.Worksheets) {
            int start = records.ToList().FindIndex(record => record.Offset == sheet.StreamOffset);
            int depth = 0;
            var ranges = new List<(int FirstRow, int LastRow, int FirstCol, int LastCol)>();
            for (int i = start; i < records.Count; i++) {
                var record = records[i];
                if (record.Type == 0x0809) { depth++; continue; }
                if (record.Type == 0x000a) { if (--depth == 0) break; continue; }
                if (depth == 1 && record.Type is 0x04bc or 0x0221 && record.Payload.Length >= 6)
                    ranges.Add((BiffRecordReader.ReadUInt16(record.Payload, 0) + 1, BiffRecordReader.ReadUInt16(record.Payload, 2) + 1, record.Payload[4] + 1, record.Payload[5] + 1));
            }
            string currentName = structure == null ? sheet.Name : structure.CurrentName(sheet.Name) ?? string.Empty;
            var current = document.Sheets.FirstOrDefault(item => item.Name == currentName);
            if (current == null) continue;
            foreach (var anchor in current.WorksheetPart.Worksheet!.Descendants<Cell>().Where(cell => cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array)) {
                if (A1.TryParseRange(anchor.CellFormula!.Reference?.Value ?? string.Empty, out int r1, out int c1, out int r2, out int c2))
                    ranges.Add((r1, r2, c1, c2));
            }
            foreach (var range in ranges) {
                bool Within(string address) { var (row, col) = A1.ParseCellRef(address); return row >= range.FirstRow && row <= range.LastRow && col >= range.FirstCol && col <= range.LastCol; }
                if (!edits.Any(edit => edit.Sheet == currentName && Within(edit.Address))) continue;
                foreach (var cell in current.WorksheetPart.Worksheet.Descendants<Cell>()) {
                    token.ThrowIfCancellationRequested();
                    string address = cell.CellReference?.Value ?? string.Empty;
                    if (Within(address) && !result.ContainsKey((currentName, address)))
                        result.Add((currentName, address), new LegacyXlsCellEdit(currentName, address, (Cell)cell.CloneNode(true)));
                }
            }
        }
        return result.Values.ToArray();
    }

    internal static byte[] EncodePreservedRow(ushort row, IReadOnlyList<ushort> columns, byte[]? original) {
        byte[] payload = original == null ? BuildRowPayload(new LegacyXlsWorksheetRow(row, Array.Empty<LegacyXlsCell>(), null)) : (byte[])original.Clone();
        WriteUInt16(payload, 2, columns.Count == 0 ? (ushort)0 : columns.Min());
        WriteUInt16(payload, 4, columns.Count == 0 ? (ushort)0 : checked((ushort)(columns.Max() + 1)));
        return payload;
    }

    internal static byte[] EncodePreservedDbCell(long firstRow, IReadOnlyList<long> firstCells, long position) =>
        BuildDbCellPayload(firstRow, firstCells, position);
}

internal sealed record LegacyXlsCellEdit(string Sheet, string Address, Cell? Cell);
internal sealed record LegacyXlsPreservedString(string Text, IReadOnlyList<LegacyXlsTextFormattingRun> Runs);
internal sealed class LegacyXlsEditEncoding {
    internal LegacyXlsEditEncoding(LegacyXlsExternSheetTable externs) { Externs = externs; }
    internal LegacyXlsExternSheetTable Externs { get; }
    internal Dictionary<(string Sheet, string Address), IReadOnlyList<BiffRecord>> Cells { get; } = new();
    internal Dictionary<string, IReadOnlyList<BiffRecord>> NewWorksheets { get; } = new(StringComparer.Ordinal);
    internal List<LegacyXlsPreservedString> Strings { get; } = new();
    internal List<byte[]> Fonts { get; } = new();
    internal List<byte[]> Formats { get; } = new();
    internal List<byte[]> Styles { get; } = new();
}
