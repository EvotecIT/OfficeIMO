using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    /// <summary>Materializes edited shared groups from their native tokens, including formulas not projected as text.</summary>
    private static IReadOnlyList<LegacyXlsCellEdit> ExpandFormulaGroups(ExcelDocument document, LegacyXlsWorkbook source,
        IReadOnlyList<BiffRecord> records, IReadOnlyList<LegacyXlsCellEdit> edits, LegacyXlsEditEncoding encoding,
        CancellationToken token, LegacyXlsSheetStructure? structure) {
        var result = edits.ToDictionary(edit => (edit.Sheet, edit.Address));
        var recordIndexes = records.Select((record, index) => (record.Offset, Index: index)).ToDictionary(item => item.Offset, item => item.Index);
        foreach (var sheet in source.Worksheets) {
            string currentName = structure?.CurrentName(sheet.Name) ?? (structure == null ? sheet.Name : string.Empty);
            var current = document.Sheets.FirstOrDefault(item => item.Name == currentName);
            if (current == null) continue;
            var currentCells = current.WorksheetPart.Worksheet!.Descendants<Cell>()
                .Where(cell => cell.CellReference?.Value != null).ToDictionary(cell => cell.CellReference!.Value!, StringComparer.Ordinal);
            var nativeCells = new Dictionary<string, List<BiffRecord>>(StringComparer.Ordinal);
            var groups = new List<(BiffRecord Record, string Anchor)>();
            int start = recordIndexes[sheet.StreamOffset], depth = 0;
            List<BiffRecord>? previous = null; string? previousAddress = null;
            for (int i = start; i < records.Count; i++) {
                token.ThrowIfCancellationRequested(); var record = records[i];
                if (record.Type == 0x0809) { depth++; continue; }
                if (record.Type == 0x000a) { if (--depth == 0) break; continue; }
                if (depth != 1) continue;
                if (record.Type == 0x0006 && record.Payload.Length >= 22) {
                    previousAddress = ExcelA1.CellReference(BiffRecordReader.ReadUInt16(record.Payload, 0) + 1, BiffRecordReader.ReadUInt16(record.Payload, 2) + 1);
                    nativeCells.Add(previousAddress, previous = new() { record });
                } else if (previous != null && record.Type is 0x0207 or 0x04bc or 0x0221 or 0x0236 or 0x003c) {
                    previous.Add(record);
                    if (record.Type is 0x04bc or 0x0221 && record.Payload.Length >= 6) groups.Add((record, previousAddress!));
                } else { previous = null; previousAddress = null; }
            }
            if (groups.Count == 0) {
                foreach (var edit in edits.Where(edit => edit.Sheet == currentName && edit.PreserveFormula))
                    if (nativeCells.TryGetValue(edit.Address, out var native)) encoding.NativeFormulaCells[(currentName, edit.Address)] = native;
                continue;
            }
            var cellIndex = new ExcelCellRangeIndex<KeyValuePair<string, Cell>>(currentCells.Select(entry => {
                var position = ExcelA1.ParseCellRef(entry.Key);
                return (position.Row, position.Col, position.Row, position.Col, entry);
            }));
            var nativeIndex = new ExcelCellRangeIndex<KeyValuePair<string, List<BiffRecord>>>(nativeCells.Select(entry => {
                var position = ExcelA1.ParseCellRef(entry.Key);
                return (position.Row, position.Col, position.Row, position.Col, entry);
            }));
            var sheetEdits = edits.Where(edit => edit.Sheet == currentName).ToArray();
            var editIndex = new ExcelCellRangeIndex<LegacyXlsCellEdit>(sheetEdits.Select(edit => {
                var position = ExcelA1.ParseCellRef(edit.Address);
                return (position.Row, position.Col, position.Row, position.Col, edit);
            }));
            foreach (var group in groups) {
                var payload = group.Record.Payload;
                int top = BiffRecordReader.ReadUInt16(payload, 0) + 1, bottom = BiffRecordReader.ReadUInt16(payload, 2) + 1;
                int left = payload[4] + 1, right = payload[5] + 1;
                var affected = editIndex.Intersect(top, left, bottom, right).ToArray();
                var nativeMembers = nativeIndex.Intersect(top, left, bottom, right).ToArray();
                if (affected.Length == 0) continue;
                bool shared = group.Record.Type == 0x04bc;
                bool preserveArray = !shared && affected.All(edit => edit.PreserveFormula);
                foreach (var entry in nativeMembers) {
                    if (!currentCells.TryGetValue(entry.Key, out var cell)) continue;
                    var key = (currentName, entry.Key);
                    if (result.TryGetValue(key, out var edit) && !edit.PreserveFormula) continue;
                    if (shared || preserveArray) {
                        encoding.NativeFormulaCells[key] = shared ? ExpandNativeSharedFormula(entry.Value, group.Record) : entry.Value;
                        if (!result.ContainsKey(key)) result.Add(key, new LegacyXlsCellEdit(currentName, entry.Key, (Cell)cell.CloneNode(true), PreserveFormula: true, PreserveStyle: true));
                    }
                }
                if (!shared && !preserveArray) {
                    // A replacement/clear of the array owner rewrites every member through the typed array encoder.
                    // Unknown arrays cannot be split into cached literals by a partial edit.
                    bool typedArray = currentCells.TryGetValue(group.Anchor, out var anchor)
                        && anchor.CellFormula?.FormulaType?.Value == CellFormulaValues.Array;
                    if (!typedArray && !sheet.ArrayFormulaRecords.Any(array => array.FirstRow == top && array.FirstColumn == left && array.FormulaTextProjected)
                        && nativeMembers.Any(entry => !result.ContainsKey((currentName, entry.Key))))
                        throw new NotSupportedException("Clear or replace the complete native array formula range before editing its members.");
                    foreach (var entry in cellIndex.Intersect(top, left, bottom, right))
                        if (!result.ContainsKey((currentName, entry.Key)))
                            result.Add((currentName, entry.Key), new LegacyXlsCellEdit(currentName, entry.Key, (Cell)entry.Value.CloneNode(true), PreserveStyle: true));
                }
            }
            foreach (var edit in sheetEdits.Where(edit => edit.PreserveFormula))
                if (!encoding.NativeFormulaCells.ContainsKey((currentName, edit.Address)) && nativeCells.TryGetValue(edit.Address, out var native))
                    encoding.NativeFormulaCells[(currentName, edit.Address)] = native;
        }
        // An array's followers may be existing blank/style-only cells, so a newly
        // authored owner can be the only snapshot delta. Emit its entire range.
        foreach (var sheet in document.Sheets) {
            var cells = sheet.WorksheetPart.Worksheet!.Descendants<Cell>()
                .Where(cell => cell.CellReference?.Value != null).ToDictionary(cell => cell.CellReference!.Value!, StringComparer.Ordinal);
            foreach (var owner in cells.Values.Where(cell => cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array)) {
                var ownerKey = (sheet.Name, owner.CellReference!.Value!);
                if (!result.TryGetValue(ownerKey, out var ownerEdit) || ownerEdit.PreserveFormula) continue;
                string[] range = (owner.CellFormula!.Reference?.Value ?? string.Empty).Split(':');
                if (range.Length is < 1 or > 2) throw new NotSupportedException("Invalid native array formula range.");
                var first = ExcelA1.ParseCellRef(range[0]); var last = ExcelA1.ParseCellRef(range[range.Length - 1]);
                if (first.Row < 1 || last.Row > 65536 || first.Col < 1 || last.Col > 256 || first.Row > last.Row || first.Col > last.Col)
                    throw new NotSupportedException("Native XLS array formulas require an ordered range within rows 1–65536 and columns 1–256.");
                for (int row = first.Row; row <= last.Row; row++) {
                    token.ThrowIfCancellationRequested();
                    for (int column = first.Col; column <= last.Col; column++) {
                        string address = ExcelA1.CellReference(row, column);
                        if (!cells.TryGetValue(address, out var member)) throw new NotSupportedException("Every native array formula member requires a cell.");
                        var key = (sheet.Name, address);
                        encoding.NativeFormulaCells.Remove(key); encoding.FormulaCaches.Remove(key);
                        bool preserveStyle = !result.TryGetValue(key, out var previousEdit) || previousEdit.PreserveStyle;
                        result[key] = new LegacyXlsCellEdit(sheet.Name, address, (Cell)member.CloneNode(true), PreserveStyle: preserveStyle);
                    }
                }
            }
        }
        return result.Values.ToArray();
    }

    private static IReadOnlyList<BiffRecord> ExpandNativeSharedFormula(IReadOnlyList<BiffRecord> cellRecords, BiffRecord definition) {
        byte[] formula = cellRecords[0].Payload, shared = definition.Payload;
        int count = BiffRecordReader.ReadUInt16(shared, 8);
        if (count == 0 || count > shared.Length - 10) throw new InvalidDataException("The native shared formula is truncated.");
        byte[] tokens = shared.Skip(10).Take(count).ToArray();
        var scanned = new List<LegacyXlsFormulaTokenRecord>();
        BiffFormulaTokenScanner.ScanTokens(tokens, "SharedFormula", null, null, definition.Offset, definition.Type, scanned);
        if (scanned.Count == 0 || scanned.Any(item => item.OperandByteCount == null)
            || scanned.Last().TokenOffset + 1 + scanned.Last().OperandByteCount != tokens.Length)
            throw new NotSupportedException("The shared formula token layout cannot be materialized safely.");
        int row = BiffRecordReader.ReadUInt16(formula, 0), column = BiffRecordReader.ReadUInt16(formula, 2);
        foreach (var item in scanned) {
            byte kind = (byte)(item.Token & 0x1f);
            if (item.Token < 0x20 || kind is not 0x0c and not 0x0d and not 0x1a and not 0x1b) continue;
            int offset = item.TokenOffset + 1;
            if (kind is 0x0c or 0x0d) {
                tokens[item.TokenOffset] = (byte)(item.Token - 8); // RefN/AreaN -> Ref/Area, retaining the operand class.
            } else {
                offset += 2; // 3D operands retain their token and external-sheet index.
            }
            if (kind is 0x0c or 0x1a) Resolve(offset, offset + 2);
            else { Resolve(offset, offset + 4); Resolve(offset + 2, offset + 6); }
        }
        byte[] expanded = new byte[checked(22 + shared.Length - 10)];
        Buffer.BlockCopy(formula, 0, expanded, 0, 22); expanded[14] &= 0xf7; // no longer shared
        WriteUInt16(expanded, 20, checked((ushort)tokens.Length)); Buffer.BlockCopy(tokens, 0, expanded, 22, tokens.Length);
        Buffer.BlockCopy(shared, 10 + count, expanded, 22 + count, shared.Length - 10 - count);
        return new[] { new BiffRecord(0x0006, cellRecords[0].Offset, expanded) }
            .Concat(cellRecords.Skip(1).Where(record => record.Type != 0x04bc)).ToArray();

        void Resolve(int rowOffset, int columnOffset) {
            int nativeRow = BiffRecordReader.ReadUInt16(tokens, rowOffset), bits = BiffRecordReader.ReadUInt16(tokens, columnOffset);
            int nativeColumn = bits & 0x3fff;
            if ((bits & 0x8000) != 0) nativeRow = (row + unchecked((short)nativeRow)) & 0xffff;
            if ((bits & 0x4000) != 0) nativeColumn = (column + (nativeColumn >= 0x2000 ? nativeColumn - 0x4000 : nativeColumn)) & 0xff;
            WriteUInt16(tokens, rowOffset, checked((ushort)nativeRow));
            WriteUInt16(tokens, columnOffset, checked((ushort)((bits & 0xc000) | nativeColumn)));
        }
    }
}
