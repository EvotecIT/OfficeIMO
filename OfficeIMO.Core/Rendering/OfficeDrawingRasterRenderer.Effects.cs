namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderEffectGroup(
        OfficeRasterCanvas canvas,
        OfficeDrawingEffectGroup effectGroup,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken) {
        if (effectGroup.Opacity <= 0D) return;
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterImage layer = Render(effectGroup.InnerDrawing, new OfficeDrawingRasterRenderOptions {
            Scale = scale,
            ImageCodec = imageCodec,
            TextShapingProvider = canvas.TextShapingProvider,
            TextShapingLanguage = canvas.TextShapingLanguage,
            DiagnosticSink = canvas.DiagnosticSink,
            DiagnosticSource = canvas.DiagnosticSource,
            MaximumRasterPixels = maximumRasterPixels,
            CancellationToken = cancellationToken
        });
        if (effectGroup.SoftMask != null) {
            layer = ApplySoftMask(
                layer,
                effectGroup.SoftMask,
                scale,
                imageCodec,
                canvas.TextShapingProvider,
                canvas.TextShapingLanguage,
                canvas.DiagnosticSink,
                canvas.DiagnosticSource,
                maximumRasterPixels,
                cancellationToken);
        }
        OfficeTransform transform = effectGroup.Transform;
        var pixelTransform = new OfficeTransform(transform.M11, transform.M12, transform.M21, transform.M22, transform.OffsetX * scale, transform.OffsetY * scale);
        var surfaceBounds = (Left: 0D, Top: 0D, Right: effectGroup.InnerDrawing.Width, Bottom: effectGroup.InnerDrawing.Height);
        bool interpolate =
            !ContainsNonInterpolatedImage(effectGroup.InnerDrawing, surfaceBounds) &&
            (effectGroup.SoftMask == null || !ContainsVisibleNonInterpolatedImage(effectGroup.SoftMask, surfaceBounds));
        canvas.DrawAffineImage(layer, pixelTransform, effectGroup.Opacity, effectGroup.BlendMode, interpolate);
    }

    private static bool ContainsNonInterpolatedImage(OfficeDrawing drawing) {
        return ContainsNonInterpolatedImage(drawing, visibleBounds: null);
    }

    private static bool ContainsNonInterpolatedImage(
        OfficeDrawing drawing,
        (double Left, double Top, double Right, double Bottom)? visibleBounds) {
        for (int index = 0; index < drawing.Elements.Count; index++) {
            OfficeDrawingElement element = drawing.Elements[index];
            if (element is OfficeDrawingImage { Interpolate: false, Opacity: > 0D } image &&
                IntersectsVisibleBounds(image.Projection.GetDestinationBounds(), visibleBounds)) return true;
            if (element is OfficeDrawingGroup group && ContainsVisibleNonInterpolatedImage(group, visibleBounds)) return true;
            if (element is OfficeDrawingEffectGroup { Opacity: > 0D } effectGroup &&
                ContainsVisibleNonInterpolatedImage(effectGroup, visibleBounds)) return true;
            if (element is OfficeDrawingTilingPattern { Opacity: > 0D } pattern &&
                ContainsVisibleNonInterpolatedImage(pattern, visibleBounds)) return true;
        }

        return false;
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingGroup group,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds) {
        var groupBounds = (
            Left: group.X,
            Top: group.Y,
            Right: group.X + group.ClipPath.Width,
            Bottom: group.Y + group.ClipPath.Height);
        (double Left, double Top, double Right, double Bottom)? effectiveParentBounds = parentVisibleBounds;
        if (parentVisibleBounds.HasValue && group.FrameTransform.HasValue && group.FrameTransform.Value.HasTransform) {
            OfficeTransform frameTransform = group.FrameTransform.Value.CreateDestinationTransform();
            var transformedGroupBounds = frameTransform.TransformRectangleBounds(
                group.X,
                group.Y,
                group.ClipPath.Width,
                group.ClipPath.Height);
            if (!TryIntersectBounds(transformedGroupBounds, parentVisibleBounds, out var transformedVisibleBounds) ||
                !frameTransform.TryInvert(out OfficeTransform inverseFrameTransform)) return false;
            effectiveParentBounds = inverseFrameTransform.TransformRectangleBounds(
                transformedVisibleBounds.Left,
                transformedVisibleBounds.Top,
                transformedVisibleBounds.Right - transformedVisibleBounds.Left,
                transformedVisibleBounds.Bottom - transformedVisibleBounds.Top);
        }

        if (!TryIntersectBounds(groupBounds, effectiveParentBounds, out var visibleGroupBounds)) return false;

        double contentX = group.X + group.ContentOffsetX;
        double contentY = group.Y + group.ContentOffsetY;
        var childVisibleBounds = (
            Left: visibleGroupBounds.Left - contentX,
            Top: visibleGroupBounds.Top - contentY,
            Right: visibleGroupBounds.Right - contentX,
            Bottom: visibleGroupBounds.Bottom - contentY);
        return ContainsNonInterpolatedImage(group.InnerDrawing, childVisibleBounds);
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingEffectGroup effectGroup,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds) {
        if (!parentVisibleBounds.HasValue) {
            return ContainsNonInterpolatedImage(effectGroup.InnerDrawing) ||
                (effectGroup.SoftMask != null && ContainsNonInterpolatedImage(effectGroup.SoftMask.InnerDrawing));
        }

        OfficeDrawing inner = effectGroup.InnerDrawing;
        var transformedBounds = effectGroup.Transform.TransformRectangleBounds(0D, 0D, inner.Width, inner.Height);
        if (!TryIntersectBounds(transformedBounds, parentVisibleBounds, out var transformedVisibleBounds) ||
            !effectGroup.Transform.TryInvert(out OfficeTransform inverseTransform)) return false;
        var childVisibleBounds = inverseTransform.TransformRectangleBounds(
            transformedVisibleBounds.Left,
            transformedVisibleBounds.Top,
            transformedVisibleBounds.Right - transformedVisibleBounds.Left,
            transformedVisibleBounds.Bottom - transformedVisibleBounds.Top);
        return ContainsNonInterpolatedImage(inner, childVisibleBounds) ||
            (effectGroup.SoftMask != null && ContainsVisibleNonInterpolatedImage(effectGroup.SoftMask, childVisibleBounds));
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingSoftMask softMask,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds) {
        if (!parentVisibleBounds.HasValue) return ContainsNonInterpolatedImage(softMask.InnerDrawing);
        OfficeDrawing inner = softMask.InnerDrawing;
        var transformedBounds = softMask.Transform.TransformRectangleBounds(0D, 0D, inner.Width, inner.Height);
        if (!TryIntersectBounds(transformedBounds, parentVisibleBounds, out var transformedVisibleBounds) ||
            !softMask.Transform.TryInvert(out OfficeTransform inverseTransform)) return false;
        var childVisibleBounds = inverseTransform.TransformRectangleBounds(
            transformedVisibleBounds.Left,
            transformedVisibleBounds.Top,
            transformedVisibleBounds.Right - transformedVisibleBounds.Left,
            transformedVisibleBounds.Bottom - transformedVisibleBounds.Top);
        return ContainsNonInterpolatedImage(inner, childVisibleBounds);
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingTilingPattern pattern,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds) {
        if (!parentVisibleBounds.HasValue) return ContainsNonInterpolatedImage(pattern.InnerTile);
        var patternBounds = (
            Left: pattern.Area.X,
            Top: pattern.Area.Y,
            Right: pattern.Area.X + pattern.Area.Width,
            Bottom: pattern.Area.Y + pattern.Area.Height);
        if (!TryIntersectBounds(patternBounds, parentVisibleBounds, out var visiblePatternBounds)) return false;

        OfficeDrawing tile = pattern.InnerTile;
        System.Collections.Generic.IReadOnlyList<OfficeTransform> transforms = pattern.GetTileTransforms(pattern.MaximumTileCount);
        for (int index = 0; index < transforms.Count; index++) {
            OfficeTransform transform = transforms[index];
            var transformedTileBounds = transform.TransformRectangleBounds(0D, 0D, tile.Width, tile.Height);
            if (!TryIntersectBounds(transformedTileBounds, visiblePatternBounds, out var transformedVisibleBounds) ||
                !transform.TryInvert(out OfficeTransform inverseTransform)) continue;
            var tileVisibleBounds = inverseTransform.TransformRectangleBounds(
                transformedVisibleBounds.Left,
                transformedVisibleBounds.Top,
                transformedVisibleBounds.Right - transformedVisibleBounds.Left,
                transformedVisibleBounds.Bottom - transformedVisibleBounds.Top);
            if (ContainsNonInterpolatedImage(tile, tileVisibleBounds)) return true;
        }

        return false;
    }

    private static bool IntersectsVisibleBounds(
        (double Left, double Top, double Right, double Bottom) bounds,
        (double Left, double Top, double Right, double Bottom)? visibleBounds) =>
        !visibleBounds.HasValue ||
        bounds.Right > visibleBounds.Value.Left &&
        bounds.Left < visibleBounds.Value.Right &&
        bounds.Bottom > visibleBounds.Value.Top &&
        bounds.Top < visibleBounds.Value.Bottom;

    private static bool TryIntersectBounds(
        (double Left, double Top, double Right, double Bottom) bounds,
        (double Left, double Top, double Right, double Bottom)? visibleBounds,
        out (double Left, double Top, double Right, double Bottom) intersection) {
        if (!visibleBounds.HasValue) {
            intersection = bounds;
            return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
        }

        intersection = (
            System.Math.Max(bounds.Left, visibleBounds.Value.Left),
            System.Math.Max(bounds.Top, visibleBounds.Value.Top),
            System.Math.Min(bounds.Right, visibleBounds.Value.Right),
            System.Math.Min(bounds.Bottom, visibleBounds.Value.Bottom));
        return intersection.Right > intersection.Left && intersection.Bottom > intersection.Top;
    }

    private static OfficeRasterImage ApplySoftMask(
        OfficeRasterImage source,
        OfficeDrawingSoftMask softMask,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        IOfficeTextShapingProvider? textShapingProvider,
        string? textShapingLanguage,
        System.Collections.Generic.ICollection<OfficeImageExportDiagnostic>? diagnosticSink,
        string? diagnosticSource,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken) {
        var maskScene = new OfficeDrawing(source.Width / scale, source.Height / scale);
        maskScene.AddEffectDrawing(softMask.InnerDrawing, softMask.Transform);
        OfficeRasterImage mask = Render(maskScene, new OfficeDrawingRasterRenderOptions {
            Scale = scale,
            ImageCodec = imageCodec,
            TextShapingProvider = textShapingProvider,
            TextShapingLanguage = textShapingLanguage,
            DiagnosticSink = diagnosticSink,
            DiagnosticSource = diagnosticSource,
            MaximumRasterPixels = maximumRasterPixels,
            CancellationToken = cancellationToken
        });
        var result = new OfficeRasterImage(source.Width, source.Height);
        double backdrop = GetMaskFactor(softMask.BackdropColor, softMask.Mode, softMask.LuminosityStandard);
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                OfficeColor sourcePixel = source.GetPixel(x, y);
                OfficeColor maskPixel = mask.GetPixel(x, y);
                double maskAlpha = maskPixel.A / 255D;
                double coverage = GetMaskFactor(maskPixel, softMask.Mode, softMask.LuminosityStandard) + ((1D - maskAlpha) * backdrop);
                result.SetPixel(x, y, OfficeColor.FromRgba(sourcePixel.R, sourcePixel.G, sourcePixel.B, (byte)System.Math.Round(sourcePixel.A * coverage)));
            }
        }
        return result;
    }

    private static double GetMaskFactor(OfficeColor color, OfficeSoftMaskMode mode, OfficeSoftMaskLuminosityStandard luminosityStandard) {
        double alpha = color.A / 255D;
        if (mode == OfficeSoftMaskMode.Alpha) return alpha;
        double redWeight = luminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb ? 0.3D : 0.2126D;
        double greenWeight = luminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb ? 0.59D : 0.7152D;
        double blueWeight = luminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb ? 0.11D : 0.0722D;
        return alpha * (((redWeight * color.R) + (greenWeight * color.G) + (blueWeight * color.B)) / 255D);
    }
}
