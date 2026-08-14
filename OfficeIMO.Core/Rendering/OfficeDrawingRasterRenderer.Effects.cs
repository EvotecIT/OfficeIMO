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
        var samplingInspection = new SamplingInspectionContext(cancellationToken);
        bool interpolate =
            !ContainsNonInterpolatedImage(effectGroup.InnerDrawing, surfaceBounds, samplingInspection) &&
            (effectGroup.SoftMask == null || !ContainsVisibleNonInterpolatedImage(effectGroup.SoftMask, surfaceBounds, samplingInspection));
        canvas.DrawAffineImage(layer, pixelTransform, effectGroup.Opacity, effectGroup.BlendMode, interpolate);
    }

    private static bool ContainsNonInterpolatedImage(OfficeDrawing drawing) {
        return ContainsNonInterpolatedImage(
            drawing,
            visibleBounds: null,
            new SamplingInspectionContext(System.Threading.CancellationToken.None));
    }

    private static bool ContainsNonInterpolatedImage(
        OfficeDrawing drawing,
        (double Left, double Top, double Right, double Bottom)? visibleBounds,
        SamplingInspectionContext inspection) {
        if (!visibleBounds.HasValue) return ContainsAnyNonInterpolatedImage(drawing, inspection);
        if (!ContainsAnyNonInterpolatedImage(drawing, inspection)) return false;
        for (int index = 0; index < drawing.Elements.Count; index++) {
            if (!inspection.TryConsume()) return true;
            OfficeDrawingElement element = drawing.Elements[index];
            if (element is OfficeDrawingImage { Interpolate: false, Opacity: > 0D } image &&
                IntersectsVisibleBounds(image.Projection.GetDestinationBounds(), visibleBounds)) return true;
            if (element is OfficeDrawingGroup group && ContainsVisibleNonInterpolatedImage(group, visibleBounds, inspection)) return true;
            if (element is OfficeDrawingEffectGroup { Opacity: > 0D } effectGroup &&
                ContainsVisibleNonInterpolatedImage(effectGroup, visibleBounds, inspection)) return true;
            if (element is OfficeDrawingTilingPattern { Opacity: > 0D } pattern &&
                ContainsVisibleNonInterpolatedImage(pattern, visibleBounds, inspection)) return true;
        }

        return false;
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingGroup group,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds,
        SamplingInspectionContext inspection) {
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
        return ContainsNonInterpolatedImage(group.InnerDrawing, childVisibleBounds, inspection);
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingEffectGroup effectGroup,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds,
        SamplingInspectionContext inspection) {
        if (!parentVisibleBounds.HasValue) {
            return ContainsAnyNonInterpolatedImage(effectGroup.InnerDrawing, inspection) ||
                (effectGroup.SoftMask != null && ContainsAnyNonInterpolatedImage(effectGroup.SoftMask.InnerDrawing, inspection));
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
        return ContainsNonInterpolatedImage(inner, childVisibleBounds, inspection) ||
            (effectGroup.SoftMask != null && ContainsVisibleNonInterpolatedImage(effectGroup.SoftMask, childVisibleBounds, inspection));
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingSoftMask softMask,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds,
        SamplingInspectionContext inspection) {
        if (!parentVisibleBounds.HasValue) return ContainsAnyNonInterpolatedImage(softMask.InnerDrawing, inspection);
        OfficeDrawing inner = softMask.InnerDrawing;
        var transformedBounds = softMask.Transform.TransformRectangleBounds(0D, 0D, inner.Width, inner.Height);
        if (!TryIntersectBounds(transformedBounds, parentVisibleBounds, out var transformedVisibleBounds) ||
            !softMask.Transform.TryInvert(out OfficeTransform inverseTransform)) return false;
        var childVisibleBounds = inverseTransform.TransformRectangleBounds(
            transformedVisibleBounds.Left,
            transformedVisibleBounds.Top,
            transformedVisibleBounds.Right - transformedVisibleBounds.Left,
            transformedVisibleBounds.Bottom - transformedVisibleBounds.Top);
        return ContainsNonInterpolatedImage(inner, childVisibleBounds, inspection);
    }

    private static bool ContainsVisibleNonInterpolatedImage(
        OfficeDrawingTilingPattern pattern,
        (double Left, double Top, double Right, double Bottom)? parentVisibleBounds,
        SamplingInspectionContext inspection) {
        if (!ContainsAnyNonInterpolatedImage(pattern.InnerTile, inspection)) return false;
        if (!parentVisibleBounds.HasValue) return true;
        var patternBounds = (
            Left: pattern.Area.X,
            Top: pattern.Area.Y,
            Right: pattern.Area.X + pattern.Area.Width,
            Bottom: pattern.Area.Y + pattern.Area.Height);
        if (!TryIntersectBounds(patternBounds, parentVisibleBounds, out var visiblePatternBounds)) return false;

        OfficeDrawing tile = pattern.InnerTile;
        System.Collections.Generic.IReadOnlyList<OfficeTransform> transforms = pattern.GetTileTransforms(pattern.MaximumTileCount);
        for (int index = 0; index < transforms.Count; index++) {
            if (!inspection.TryConsume()) return true;
            OfficeTransform transform = transforms[index];
            var transformedTileBounds = transform.TransformRectangleBounds(0D, 0D, tile.Width, tile.Height);
            if (!TryIntersectBounds(transformedTileBounds, visiblePatternBounds, out var transformedVisibleBounds) ||
                !transform.TryInvert(out OfficeTransform inverseTransform)) continue;
            var tileVisibleBounds = inverseTransform.TransformRectangleBounds(
                transformedVisibleBounds.Left,
                transformedVisibleBounds.Top,
                transformedVisibleBounds.Right - transformedVisibleBounds.Left,
                transformedVisibleBounds.Bottom - transformedVisibleBounds.Top);
            if (ContainsNonInterpolatedImage(tile, tileVisibleBounds, inspection)) return true;
        }

        return false;
    }

    private static bool ContainsAnyNonInterpolatedImage(
        OfficeDrawing drawing,
        SamplingInspectionContext inspection) {
        if (inspection.TryGetCached(drawing, out bool cached)) return cached;
        inspection.Cache(drawing, value: true);
        for (int index = 0; index < drawing.Elements.Count; index++) {
            if (!inspection.TryConsume()) return true;
            OfficeDrawingElement element = drawing.Elements[index];
            if (element is OfficeDrawingImage { Interpolate: false, Opacity: > 0D }) return true;
            if (element is OfficeDrawingGroup group && ContainsAnyNonInterpolatedImage(group.InnerDrawing, inspection)) return true;
            if (element is OfficeDrawingEffectGroup { Opacity: > 0D } effectGroup &&
                (ContainsAnyNonInterpolatedImage(effectGroup.InnerDrawing, inspection) ||
                 effectGroup.SoftMask != null && ContainsAnyNonInterpolatedImage(effectGroup.SoftMask.InnerDrawing, inspection))) return true;
            if (element is OfficeDrawingTilingPattern { Opacity: > 0D } pattern &&
                ContainsAnyNonInterpolatedImage(pattern.InnerTile, inspection)) return true;
        }

        inspection.Cache(drawing, value: false);
        return false;
    }

    private sealed class SamplingInspectionContext {
        private const long MaximumWork = 1_000_000L;
        private readonly System.Collections.Generic.Dictionary<OfficeDrawing, bool> _drawingResults = new System.Collections.Generic.Dictionary<OfficeDrawing, bool>();
        private readonly System.Threading.CancellationToken _cancellationToken;
        private long _work;

        internal SamplingInspectionContext(System.Threading.CancellationToken cancellationToken) {
            _cancellationToken = cancellationToken;
        }

        internal bool TryConsume() {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_work >= MaximumWork) return false;
            _work++;
            return true;
        }

        internal bool TryGetCached(OfficeDrawing drawing, out bool value) =>
            _drawingResults.TryGetValue(drawing, out value);

        internal void Cache(OfficeDrawing drawing, bool value) => _drawingResults[drawing] = value;
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
