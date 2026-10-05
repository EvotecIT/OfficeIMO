using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void ApplyLegacyDocTableBorderDefaults(WordTable table, IEnumerable<LegacyDocTableBorders> rowBorders) {
            LegacyDocTableBorders source = rowBorders.FirstOrDefault(borders => borders.HasAny);
            if (!source.HasAny) return;
            var borders = new TableBorders();
            AppendLegacyDocTableBorder<TopBorder>(borders, source.Top);
            AppendLegacyDocTableBorder<LeftBorder>(borders, source.Left);
            AppendLegacyDocTableBorder<BottomBorder>(borders, source.Bottom);
            AppendLegacyDocTableBorder<RightBorder>(borders, source.Right);
            AppendLegacyDocTableBorder<InsideHorizontalBorder>(borders, source.InsideHorizontal);
            AppendLegacyDocTableBorder<InsideVerticalBorder>(borders, source.InsideVertical);
            table.StyleDetails!.TableBorders = borders;
        }

        private static void AppendLegacyDocTableBorder<T>(TableBorders borders, LegacyDocTableCellBorder source)
            where T : BorderType, new() {
            BorderValues? style = MapLegacyDocTableCellBorderStyle(source.Style);
            if (!style.HasValue) return;
            var border = new T { Val = style.Value, Color = source.ColorHex ?? "auto" };
            if (source.SizeEighthPoints > 0) border.Size = (uint)source.SizeEighthPoints;
            if (source.SpacePoints > 0) border.Space = (uint)source.SpacePoints;
            borders.Append(border);
        }
    }
}
