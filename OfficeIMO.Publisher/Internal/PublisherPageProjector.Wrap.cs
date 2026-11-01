using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private IReadOnlyList<OfficeTextFlowRegion> TextExclusions(PublisherFrameState frame) {
        var result = new List<OfficeTextFlowRegion>();
        if (frame.WrapObjectIds.Count == 0) return result;
        OfficeTransform toFrame = frame.PageTransform.Invert();
        foreach (uint id in frame.WrapObjectIds) {
            _context.Record();
            PublisherImportedWrapObject? obstacle = CurrentWrapObject(id);
            if (obstacle == null || !obstacle.Available || obstacle.PageId != frame.PageId) {
                _context.Add("PUB_TEXT_WRAP_REFERENCE_UNRESOLVED", "A native text-wrap object has no resolved anchor on the frame's page. Its exclusion could not be applied.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Id));
                continue;
            }
            if (obstacle.Hidden) continue;
            _publication.Text.WrapInfo.TryGetValue(id, out PublisherWrapInfo? info);
            OfficeTextPadding margin = info?.Margins ?? OfficeTextPadding.Empty;
            if (info?.NegativeMargins == true)
                _context.Add("PUB_NEGATIVE_WRAP_DISTANCE_APPROXIMATED", "A negative native wrap distance was bounded to the object's outline.",
                    OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(id));
            var local = obstacle.PageTransform.Then(toFrame).TransformRectangleBounds(-margin.Left, -margin.Top,
                obstacle.Width + margin.Horizontal, obstacle.Height + margin.Vertical);
            result.Add(new OfficeTextFlowRegion(local.Left, local.Top, local.Right - local.Left, local.Bottom - local.Top));
            if (info?.Tight == true)
                _context.Add("PUB_TIGHT_TEXT_WRAP_APPROXIMATED", "A native wrap polygon uses its transformed bounding rectangle. Tight and through-outline text paths are not reproduced.",
                    OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(id));
        }
        _context.Add("PUB_TEXT_WRAP_APPROXIMATED", UsesBothSideWrap(frame)
            ? "Native rectangular both-side wrap references and distances exclude transformed object rectangles. Text flows left to right through free intervals on a shared baseline; native line fitting and outline wrapping can differ."
            : "Native wrap references and distances exclude transformed object rectangles. Unqualified wrap modes use the widest available interval per horizontal band; native side selection and outline wrapping can differ.",
            OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(frame.Id));
        return result;
    }

    private bool UsesBothSideWrap(PublisherFrameState frame) => frame.VerticalAlignment == OfficeTextVerticalAlignment.Top &&
        frame.WrapObjectIds.Count != 0 && frame.WrapObjectIds.All(id =>
        _publication.Text.WrapInfo.TryGetValue(id, out PublisherWrapInfo? wrap) &&
        wrap.NativeMode == PublisherNativeWrapMode.RectangularBoth && !wrap.Tight && !wrap.NegativeMargins);

    private PublisherImportedWrapObject? CurrentWrapObject(uint id) {
        if (_publication.Tables.TryGetValue(id, out PublisherTableState? table))
            return new PublisherImportedWrapObject(table.PageId, table.X, table.Y, table.Width, table.Height, table.PageTransform,
                table.HasPlacement, table.Hidden);
        if (_publication.Text.Frames.TryGetValue(id, out PublisherFrameState? frame))
            return new PublisherImportedWrapObject(frame.PageId, frame.X, frame.Y, frame.Width, frame.Height, frame.PageTransform,
                frame.HasAnchor && frame.TransformAvailable, frame.Hidden);
        if (_publication.Pictures.TryGetValue(id, out PublisherPicture? picture) && picture.CanEdit)
            return new PublisherImportedWrapObject(picture.PageId, picture.X, picture.Y, picture.Width, picture.Height, picture.PageTransform,
                picture.HasPlacement, false);
        if (_publication.Shapes.TryGetValue(id, out PublisherShape? shape) && shape.CanEdit)
            return new PublisherImportedWrapObject(shape.PageId, shape.X, shape.Y, shape.Width, shape.Height, shape.PageTransform,
                shape.HasPlacement, false);
        return _publication.Text.ImportedObstacles.TryGetValue(id, out PublisherImportedWrapObject? imported) ? imported : null;
    }
}
