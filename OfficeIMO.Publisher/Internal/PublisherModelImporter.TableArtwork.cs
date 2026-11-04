using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    /// <summary>Maps solid utility fills and shared grid edges without treating utility shapes as publication objects.</summary>
    private static void ImportTableArtwork(uint id, PublisherSourceTable grid, IEnumerable<PublisherTableArtwork> records,
        IReadOnlyList<OfficeColor?> palette, PublisherEscherData artwork, PublisherParseContext context,
        out OfficeColor?[] fills, out PublisherTableBorderSegment[] borders) {
        var colors = new OfficeColor?[grid.Cells.Count];
        var lines = new List<PublisherTableBorderSegment>();
        var fillOwners = new HashSet<int>();
        double totalWidth = Total(grid.Columns), totalHeight = Total(grid.Rows);
        foreach (PublisherTableArtwork record in records) {
            context.Token.ThrowIfCancellationRequested();
            uint kind = record.Value(0x2001);
            if (kind == 0) {
                uint index = record.Value(0x2003);
                if (index >= colors.Length) { Unsupported(); continue; }
                if (!fillOwners.Add((int)index)) { colors[index] = null; Unsupported(); continue; }
                if (record.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2003))) { Unsupported(); continue; }
                if (!TryColor(record, out OfficeColor? color)) { Unsupported(); continue; }
                colors[index] = color;
            } else if (kind is 1 or 2) {
                uint row = record.Value(0x2004), column = record.Value(0x2005), endRow = record.Value(0x2006), endColumn = record.Value(0x2007);
                if (row > grid.Rows.Length || endRow > grid.Rows.Length || column > grid.Columns.Length || endColumn > grid.Columns.Length
                    || (kind == 1 ? row != endRow || endColumn <= column : column != endColumn || endRow <= row)
                    || record.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2004 or 0x2005 or 0x2006 or 0x2007))) { Unsupported(); continue; }
                if (!TryColor(record, out OfficeColor? color)) { Unsupported(); continue; }
                if (!color.HasValue) continue;
                int? nativeWidth = record.Shape.Style.LineWidthEmus;
                double width = nativeWidth.GetValueOrDefault() / 12700D;
                if (!nativeWidth.HasValue || width <= 0
                    || width > (kind == 1 ? totalHeight : totalWidth)) { Unsupported(); continue; }
                lines.Add(new((int)row, (int)column, (int)endRow, (int)endColumn, new OfficeBorderSide(color.Value, width)));
            } else Unsupported();
        }
        fills = colors; borders = lines.ToArray();

        double Total(IReadOnlyList<double> tracks) {
            double result = 0;
            foreach (double track in tracks) { context.Token.ThrowIfCancellationRequested(); result += track; }
            return result;
        }

        bool TryColor(PublisherTableArtwork record, out OfficeColor? color) {
            color = null;
            var style = record.Shape.Style;
            if (record.Shape.Type != 1 || record.Shape.Hidden || record.Shape.Transform.RotationDegrees.GetValueOrDefault() != 0
                || style.ShadowEnabled == true) return false;
            if (style.FillEnabled == false) return true;
            if (style.FillType.GetValueOrDefault() != 0 || style.FillOpacity.GetValueOrDefault(1) != 1) return false;
            if (record.Value(0x2001) == 0 && artwork.HasTableCellDefault
                && record.Shape.NativeBinding?.Properties.Any(property => property.Id == 0x0181) != true) {
                if (!artwork.TableCellDefaultColor.HasValue) return false;
                // The native table-cell default uses automatic white, distinct
                // from the document drawing's theme accent inherited by Core.
                if (artwork.TableCellDefaultColor == uint.MaxValue) { color = OfficeColor.White; return true; }
                if (!new OfficeIMO.Drawing.Binary.OfficeArtColorReference(artwork.TableCellDefaultColor.Value)
                    .TryResolve(index => index < palette.Count ? palette[index] : null, out OfficeColor tableColor)) return false;
                color = tableColor; return true;
            }
            if (!style.FillColor.HasValue) return false;
            if (!style.FillColor.Value.TryResolve(index => index < palette.Count ? palette[index] : null, out OfficeColor resolved)) return false;
            color = resolved; return true;
        }
        void Unsupported() => context.Add("PUB_TABLE_CELL_STYLE_UNSUPPORTED",
            "A native table utility fill or border has unsupported paint, ownership or grid coordinates. Its bytes remain preserved; this decoration is not projected.",
            OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(id));
    }
}
