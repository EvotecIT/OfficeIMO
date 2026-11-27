using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private static IEnumerable<ReaderRow> EnumerateRowsWithCoordinates(OfficeXmlNode sheetData, CancellationToken ct) {
            int previousIndex = 0;
            foreach (OfficeXmlNode row in sheetData.Children(SpreadsheetXml, "row")) {
                ct.ThrowIfCancellationRequested();
                yield return new ReaderRow(row, ResolveRowIndex(row, ref previousIndex));
            }
        }

        private static int ResolveRowIndex(OfficeXmlNode row, ref int previousIndex) {
            if (row.UInt32Attribute("r") is uint declared && declared > 0) {
                return previousIndex = checked((int)declared);
            }
            foreach (OfficeXmlNode cell in row.Children(SpreadsheetXml, "c")) {
                if (ExcelA1.TryParseCellReferenceFast(cell.Attribute("r").AsSpan(), out int referencedRow, out _)) {
                    return previousIndex = referencedRow;
                }
            }
            return ++previousIndex;
        }

        private readonly struct ReaderRow {
            private readonly OfficeXmlNode _row;

            internal ReaderRow(OfficeXmlNode row, int index) {
                _row = row;
                Index = index;
            }

            internal int Index { get; }

            internal CellsEnumerable Cells() => new(_row.Children(SpreadsheetXml, "c"));

            internal readonly struct CellsEnumerable {
                private readonly OfficeXmlNode.ChildList _cells;

                internal CellsEnumerable(OfficeXmlNode.ChildList cells) => _cells = cells;

                public Enumerator GetEnumerator() => new(_cells.GetEnumerator());
            }

            internal struct Enumerator {
                private OfficeXmlNode.ChildList.Enumerator _cells;

                internal Enumerator(OfficeXmlNode.ChildList.Enumerator cells) => _cells = cells;

                public ReaderCell Current => new(_cells.Current);

                public bool MoveNext() => _cells.MoveNext();
            }
        }

        private readonly struct ReaderCell {
            private readonly OfficeXmlNode _cell;

            internal ReaderCell(OfficeXmlNode cell) => _cell = cell;

            internal string? Reference => _cell.Attribute("r");

            internal ExcelCellValueType? Type => ToCellValueType(ParseXmlCellKind(_cell.Attribute("t")));

            internal string? StyleAttribute => _cell.Attribute("s");

            internal uint? Style => _cell.UInt32Attribute("s");

            internal uint? ValueMetaIndex => _cell.UInt32Attribute("vm");

            internal uint? RowStyle {
                get {
                    OfficeXmlNode row = _cell.Parent;
                    return row.BooleanAttribute("customFormat") == true ? row.UInt32Attribute("s") : null;
                }
            }

            internal string? Value => TextOf(_cell.Child(SpreadsheetXml, "v"));

            internal OfficeXmlNode Formula => _cell.Child(SpreadsheetXml, "f");

            internal int GetColumnIndex(ref int nextColumn) {
                string? reference = Reference;
                int column = string.IsNullOrEmpty(reference)
                    ? nextColumn
                    : ExcelA1.ParseColumnIndexFromCellReferenceFast(reference);
                if (column <= 0) column = nextColumn;
                nextColumn = column + 1;
                return column;
            }

            internal void GetChildren(out string? formula, out string? value, out OfficeXmlNode inline) {
                formula = null;
                value = null;
                inline = default;
                for (OfficeXmlNode child = _cell.FirstChild; child.Exists; child = child.NextSibling) {
                    if (child.NamespaceId != SpreadsheetXml) continue;
                    switch (child.LocalName) {
                        case "f": formula ??= child.InnerText; break;
                        case "v": value ??= child.InnerText; break;
                        case "is" when inline.IsNone: inline = child; break;
                    }
                }
            }

            private static string? TextOf(OfficeXmlNode node) => node.IsNone ? null : node.InnerText;
        }
    }
}
