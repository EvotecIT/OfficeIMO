using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    /// <summary>Imports grid and placement once. Unsupported graphic features stay on the private native adapter; no mutable native records are authored.</summary>
    private static IReadOnlyDictionary<uint, PublisherTableState> ImportTables(PublisherContents source, PublisherEscherData artwork,
        IReadOnlyDictionary<uint, PublisherPageState> pages, IReadOnlyDictionary<uint, uint?> owners, PublisherParseContext context) {
        var result = new Dictionary<uint, PublisherTableState>();
        var shapes = artwork.Shapes.ToDictionary(shape => shape.Id);
        var mapper = new PublisherNativeShapeMapper(source.Palette, context);
        foreach (PublisherSourceShape native in source.Shapes.Values.Where(native => native.Table != null)) {
            context.Token.ThrowIfCancellationRequested();
            if (!owners.TryGetValue(native.Chunk.Id, out uint? owner) || !owner.HasValue
                || !pages.TryGetValue(owner.Value, out PublisherPageState? page)) continue;
            shapes.TryGetValue(native.Chunk.Id, out PublisherEscherShape? shape);
            PublisherSourceTable grid = native.Table!;
            bool space = shape != null && shape.LayoutTransform.TryInvert(out _) && shape.GroupTransform.TryInvert(out _);
            var bounds = shape?.Bounds.HasValue == true ? PublisherNativeGeometry.Bounds(shape, page.RegistrationX, page.RegistrationY) : default;
            bool placement = shape?.Bounds.HasValue == true && space && Finite(bounds.X) && Finite(bounds.Y)
                && Finite(bounds.Width) && Finite(bounds.Height) && bounds.Width > 0 && bounds.Height > 0;
            bool nativeTable = native.Chunk.Kind == 0x10;
            string? limitation = shape == null ? "The table has no recovered graphic placement."
                : ShapeEditLimitation(shape, page, source.Palette, placement, allowTableFrame: nativeTable);
            if (limitation == null && shape!.Type is not (1 or 75 or 201 or 202))
                limitation = "The table's native graphic is not a qualified rectangular frame.";
            if (limitation == null && shape!.Properties.Any(property => property.PropertyId == 0x0104 && property.Value != 0))
                limitation = "The table's native graphic has an unsupported picture binding.";
            OfficeShape? paint = limitation == null ? mapper.Map(shape!, bounds.Width, bounds.Height, bounds.Width, bounds.Height,
                out _, tableFrame: nativeTable) : null;
            var padding = new OfficeTextPadding((shape?.Property(0x81) ?? 36576) / 12700D, (shape?.Property(0x82) ?? 36576) / 12700D,
                (shape?.Property(0x83) ?? 36576) / 12700D, (shape?.Property(0x84) ?? 36576) / 12700D);
            result.Add(native.Chunk.Id, new PublisherTableState(native.Chunk.Id, native.Value(0x27), page.Id,
                bounds.X, bounds.Y, bounds.Width, bounds.Height, placement
                    ? OfficeTransform.Translate(bounds.X, bounds.Y).Then(PublisherNativeGeometry.PageTransform(shape!, bounds.X, bounds.Y,
                        bounds.Width, bounds.Height, page.RegistrationX, page.RegistrationY)) : OfficeTransform.Identity,
                placement, shape?.Hidden == true, grid.Columns, grid.Rows,
                grid.Cells.Select(cell => new PublisherTableCellSlot(cell.FirstRow, cell.FirstColumn,
                    cell.LastRow - cell.FirstRow + 1, cell.LastColumn - cell.FirstColumn + 1)).ToArray(), padding, paint, limitation));
        }
        return result;
    }
}
