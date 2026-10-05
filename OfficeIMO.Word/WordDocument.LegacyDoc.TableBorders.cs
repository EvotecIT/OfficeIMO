using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void ApplyLegacyDocTableBorderDefaults(WordTable table, IEnumerable<LegacyDocTableBorders> rowBorders) {
            LegacyDocTableBorders[] rows = rowBorders.ToArray();
            if (rows.Length == 0) return;
            // A DOC operand belongs to a row. Only common edges can become table-wide defaults.
            LegacyDocTableCellBorder Common(Func<LegacyDocTableBorders, LegacyDocTableCellBorder> select) {
                LegacyDocTableCellBorder first = select(rows[0]);
                return rows.All(row => select(row).Equals(first)) ? first : default;
            }
            var source = new LegacyDocTableBorders(rows[0].Top, Common(row => row.Left),
                rows[rows.Length - 1].Bottom, Common(row => row.Right),
                Common(row => row.InsideHorizontal), Common(row => row.InsideVertical));
            var borders = new TableBorders();
            AppendLegacyDocTableBorder<TopBorder>(borders, source.Top);
            AppendLegacyDocTableBorder<LeftBorder>(borders, source.Left);
            AppendLegacyDocTableBorder<BottomBorder>(borders, source.Bottom);
            AppendLegacyDocTableBorder<RightBorder>(borders, source.Right);
            AppendLegacyDocTableBorder<InsideHorizontalBorder>(borders, source.InsideHorizontal);
            AppendLegacyDocTableBorder<InsideVerticalBorder>(borders, source.InsideVertical);
            table.StyleDetails!.TableBorders = borders;
            // A borderless row does not erase the visible edge supplied by its neighbor.
            LegacyDocTableCellBorder SharedHorizontal(LegacyDocTableCellBorder actual, LegacyDocTableCellBorder adjacent) =>
                (!actual.HasAny || actual.Style == LegacyDocTableCellBorderStyle.ExplicitNone)
                    && adjacent.HasAny && adjacent.Style != LegacyDocTableCellBorderStyle.ExplicitNone
                    ? adjacent : actual;
            for (int row = 0; row < rows.Length && row < table.Rows.Count; row++) {
                WordTableRow targetRow = table.Rows[row];
                for (int column = 0; column < targetRow.Cells.Count; column++) {
                    LegacyDocTableCellBorder Override(LegacyDocTableCellBorder actual, LegacyDocTableCellBorder common) {
                        if (actual.Equals(common)) return default;
                        return actual.HasAny ? actual
                            : new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0);
                    }
                    ApplyLegacyDocTableCellBorders(targetRow.Cells[column], new LegacyDocTableCellBorders(
                        Override(row == 0 ? rows[row].Top
                            : SharedHorizontal(rows[row].InsideHorizontal, rows[row - 1].InsideHorizontal),
                            row == 0 ? source.Top : source.InsideHorizontal),
                        Override(column == 0 ? rows[row].Left : rows[row].InsideVertical,
                            column == 0 ? source.Left : source.InsideVertical),
                        Override(row == rows.Length - 1 ? rows[row].Bottom
                            : SharedHorizontal(rows[row].InsideHorizontal, rows[row + 1].InsideHorizontal),
                            row == rows.Length - 1 ? source.Bottom : source.InsideHorizontal),
                        Override(column == targetRow.Cells.Count - 1 ? rows[row].Right : rows[row].InsideVertical,
                            column == targetRow.Cells.Count - 1 ? source.Right : source.InsideVertical)));
                }
            }
        }

        private static void AppendLegacyDocTableBorder<T>(TableBorders borders, LegacyDocTableCellBorder source)
            where T : BorderType, new() {
            BorderValues? style = MapLegacyDocTableCellBorderStyle(source.Style);
            if (!style.HasValue) return;
            var border = new T { Val = style.Value };
            if (style.Value != BorderValues.Nil) border.Color = source.ColorHex ?? "auto";
            if (source.SizeEighthPoints > 0) border.Size = (uint)source.SizeEighthPoints;
            if (source.SpacePoints > 0) border.Space = (uint)source.SpacePoints;
            borders.Append(border);
        }
    }
}
