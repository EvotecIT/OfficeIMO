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
        foreach (var edit in edits.Where(edit => edit.CacheOnly)) result.FormulaCaches.Add((edit.Sheet, edit.Address), edit.Cell!);
        edits = ExpandFormulaGroups(document, source, records, edits.Where(edit => !edit.CacheOnly).ToArray(), result, token, structure);
        var originalStyles = new Dictionary<(string Sheet, string Address), ushort>();
        foreach (var sheet in source.Worksheets) {
            string currentName = structure?.CurrentName(sheet.Name) ?? (structure == null ? sheet.Name : string.Empty);
            foreach (var cell in sheet.Cells) {
                token.ThrowIfCancellationRequested();
                originalStyles[(currentName, A1.CellReference(cell.Row, cell.Column))] = cell.StyleIndex;
            }
        }
        ushort? PreservedStyle(LegacyXlsCellEdit edit) => edit.PreserveStyle
            && originalStyles.TryGetValue((edit.Sheet, edit.Address), out ushort index) ? index : null;
        // Unchanged native XFs can carry opaque theme/font metadata. Only encode
        // styles needed by new cells or actual style edits; retain native indices
        // per cell, including distinct native XFs with identical projections.
        var styles = LegacyXlsStyleTable.CreatePreserving(document, sheets, fonts, records,
            edits.Where(edit => edit.Cell != null && !PreservedStyle(edit).HasValue)
                .Select(edit => edit.Cell!.StyleIndex?.Value ?? 0U));
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
                if (result.NativeFormulaCells.TryGetValue((edit.Sheet, edit.Address), out var native)) {
                    var retained = native.ToArray(); byte[] payload = (byte[])retained[0].Payload.Clone();
                    WriteUInt16(payload, 4, PreservedStyle(edit) ?? styles.GetBiffStyleIndex(edit.Cell.StyleIndex?.Value));
                    retained[0] = new BiffRecord(retained[0].Type, retained[0].Offset, payload);
                    result.Cells.Add((edit.Sheet, edit.Address), EncodePreservedFormulaCache(sheet, edit.Cell, retained, document.DateSystem, structure?.NameIndexes));
                    continue;
                }
                var converted = ConvertCell(sheet, document.WorkbookPartRoot, index, edit.Cell,
                    checked((ushort)(row - 1)), checked((ushort)(column - 1)), styles, document.DateSystem, names, shared, arrays, fonts,
                    PreservedStyle(edit));
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
        if (fonts.PaletteRecord is byte[] palette) {
            byte[]? originalPalette = records.Where(record => record.Type == 0x0092).Select(record => record.Payload).FirstOrDefault();
            if (originalPalette == null || !palette.SequenceEqual(originalPalette)) result.Palette = palette;
        }
        result.Fonts.AddRange(fonts.FontRecords.Skip(source.Fonts.Count));
        result.Formats.AddRange(styles.FormatRecords);
        result.Styles.AddRange(styles.CellFormatRecords);
        return result;
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

internal sealed record LegacyXlsCellEdit(string Sheet, string Address, Cell? Cell, bool CacheOnly = false,
    bool PreserveFormula = false, bool PreserveStyle = false);
internal sealed record LegacyXlsPreservedString(string Text, IReadOnlyList<LegacyXlsTextFormattingRun> Runs);
internal sealed class LegacyXlsEditEncoding {
    internal LegacyXlsEditEncoding(LegacyXlsExternSheetTable externs) { Externs = externs; }
    internal LegacyXlsExternSheetTable Externs { get; }
    internal Dictionary<(string Sheet, string Address), IReadOnlyList<BiffRecord>> NativeFormulaCells { get; } = new();
    internal Dictionary<(string Sheet, string Address), Cell> FormulaCaches { get; } = new();
    internal Dictionary<(string Sheet, string Address), IReadOnlyList<BiffRecord>> Cells { get; } = new();
    internal Dictionary<string, IReadOnlyList<BiffRecord>> NewWorksheets { get; } = new(StringComparer.Ordinal);
    internal List<LegacyXlsPreservedString> Strings { get; } = new();
    internal byte[]? Palette { get; set; }
    internal List<byte[]> Fonts { get; } = new();
    internal List<byte[]> Formats { get; } = new();
    internal List<byte[]> Styles { get; } = new();
}
