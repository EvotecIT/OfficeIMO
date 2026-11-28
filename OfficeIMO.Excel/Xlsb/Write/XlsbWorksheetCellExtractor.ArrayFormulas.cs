using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbWorksheetCellExtractor {
    /// <summary>Writes one array declaration and native group references for every declared member.</summary>
    private static List<XlsbWriteCell> ApplyArrayFormulaGroups(ExcelSheet sheet, OfficeXmlNode sheetData,
        List<XlsbWriteCell> cells, IXlsbFormulaEncodingContext? formulas) {
        IEnumerable<CellMarkup> ReadCells() {
            foreach (OfficeXmlNode row in WorksheetChildren(sheetData, "row")) {
                foreach (OfficeXmlNode cell in WorksheetChildren(row, "c")) yield return new CellMarkup(cell);
            }
        }
        var arrays = ReadCells().Where(cell => cell.Formula.Attribute("t") == "array").ToArray();
        if (arrays.Length == 0) return cells;
        var result = cells.ToDictionary(cell => (cell.Row, cell.Column));
        var occupied = new HashSet<(int, int)>();
        var memberIndex = new ExcelCellRangeIndex<CellMarkup>(ReadCells()
            .Where(cell => cell.Reference != null).Select(cell => {
                var position = ExcelA1.ParseCellRef(cell.Reference!);
                return (position.Row, position.Col, position.Row, position.Col, cell);
            }));
        foreach (var anchor in arrays) {
            string referenceText = anchor.Formula.Attribute("ref") ?? string.Empty;
            if (!ExcelA1.TryParseRange(referenceText, out int top, out int left, out int bottom, out int right)) {
                // Native Excel writes single-cell CSE groups as e.g. H1, not H1:H1.
                if (!ExcelA1.TryParseCellReferenceFast(referenceText, out top, out left))
                    throw new NotSupportedException("The native XLSB array formula requires a rectangular range.");
                bottom = top; right = left;
            }
            var address = ExcelA1.ParseCellRef(anchor.Reference!);
            if (address.Row != top || address.Col != left || bottom > 1_048_576 || right > 16_384)
                throw new NotSupportedException("The native XLSB array owner must be the top-left cell of a valid worksheet range.");
            string text = anchor.Formula.InnerText; byte[] payload; string? reason;
            bool encoded = formulas == null ? XlsbFormulaEncoder.TryEncode(text, out payload, out reason, arrayFormula: true)
                : formulas.TryEncode(text, sheet.Name, out payload, out reason, arrayFormula: true);
            if (!encoded) throw new NotSupportedException("The native XLSB array formula cannot be encoded: " + reason);
            byte[] definition = new byte[checked(17 + payload.Length - 2)];
            XlsbFormulaGroup.Put(definition, 0, checked((uint)(top - 1))); XlsbFormulaGroup.Put(definition, 4, checked((uint)(bottom - 1)));
            XlsbFormulaGroup.Put(definition, 8, checked((uint)(left - 1))); XlsbFormulaGroup.Put(definition, 12, checked((uint)(right - 1)));
            definition[16] = XlsbWorkbookMarkup.ReadBoolean(anchor.Formula, "ca") == true ? (byte)1 : (byte)0;
            Buffer.BlockCopy(payload, 2, definition, 17, payload.Length - 2);
            byte[] reference = new byte[19]; reference[0] = payload[0]; reference[1] = payload[1]; reference[2] = 5; reference[6] = 1;
            XlsbFormulaGroup.Put(reference, 7, checked((uint)(top - 1))); XlsbFormulaGroup.Put(reference, 11, 4); XlsbFormulaGroup.Put(reference, 15, checked((uint)(left - 1)));
            long expected = (long)(bottom - top + 1) * (right - left + 1);
            var members = new List<XlsbWriteCell>();
            foreach (var member in memberIndex.Intersect(top, left, bottom, right)) {
                var position = ExcelA1.ParseCellRef(member.Reference!);
                if ((position.Row != top || position.Col != left) && member.Formula.Exists)
                    throw new NotSupportedException("The native XLSB array cannot contain another formula inside its range.");
                members.Add(result.TryGetValue((position.Row, position.Col), out var existing) ? existing
                    : new XlsbWriteCell(position.Row, position.Col, member.StyleIndex ?? 0, XlsbWriteCellKind.Blank, null));
            }
            if (members.Count != expected) throw new NotSupportedException("The native XLSB array requires every member cell in its declared range.");
            foreach (var member in members) {
                if (!occupied.Add((member.Row, member.Column))) throw new NotSupportedException("Native XLSB array formula ranges cannot overlap.");
                var converted = XlsbWorksheetPartWriter.WithFormula(member, reference);
                if (member.Row == top && member.Column == left) converted.FormulaGroup = new XlsbGeneratedRecord(426, definition);
                result[(member.Row, member.Column)] = converted;
            }
        }
        return result.Values.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToList();
    }
}
