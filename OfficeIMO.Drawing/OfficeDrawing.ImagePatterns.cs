using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>Adds a clipped repeating image pattern and returns this drawing.</summary>
    public OfficeDrawing AddImagePattern(
        byte[] bytes,
        string? contentType,
        OfficeImagePatternLayout layout,
        int maximumTileCount = 16384,
        double opacity = 1D) {
        OfficeImagePlacement area = layout.Area;
        if (area.X < 0D || area.Y < 0D || area.X + area.Width > Width || area.Y + area.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(layout), "Image-pattern area must fit inside the drawing bounds.");
        }

        var pattern = new OfficeDrawingImagePattern(bytes, contentType, layout, maximumTileCount, opacity);
        _imagePatterns.Add(pattern);
        _elements.Add(pattern);
        return this;
    }

    internal OfficeDrawing AddImagePatternShared(
        byte[] bytes,
        string contentType,
        OfficeImagePatternLayout layout,
        int maximumTileCount = 16384,
        double opacity = 1D) {
        OfficeImagePlacement area = layout.Area;
        if (area.X < 0D || area.Y < 0D || area.X + area.Width > Width || area.Y + area.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(layout), "Image-pattern area must fit inside the drawing bounds.");
        }

        var pattern = new OfficeDrawingImagePattern(bytes, contentType, layout, maximumTileCount, opacity,
            useSnapshot: true);
        _imagePatterns.Add(pattern);
        _elements.Add(pattern);
        return this;
    }

    private void AddNestedImagePattern(OfficeDrawingImagePattern pattern, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
        OfficeImagePatternLayout translated = pattern.Layout.Translate(offsetX, offsetY);
        OfficeImagePlacement area = translated.Area;
        if (!allowOverflow && (area.X < 0D || area.Y < 0D || area.X + area.Width > Width || area.Y + area.Height > Height)) {
            throw new ArgumentOutOfRangeException(nameof(pattern), "Nested image-pattern area must fit inside the drawing bounds.");
        }

        if (frameTransform.HasValue && frameTransform.Value.HasTransform) {
            OfficeImageFrameTransform frame = frameTransform.Value;
            foreach (OfficeImagePlacement tile in translated.GetTilePlacements(pattern.MaximumTileCount)) {
                var projection = new OfficeImageProjection(
                    tile,
                    rotationDegrees: frame.RotationDegrees,
                    rotationCenterX: frame.CenterX,
                    rotationCenterY: frame.CenterY,
                    flipHorizontal: frame.FlipHorizontal,
                    flipVertical: frame.FlipVertical);
                var image = new OfficeDrawingImage(pattern.EncodedBytes, pattern.ContentType, projection, alternativeText: null, opacity: pattern.Opacity, useDataSnapshot: true);
                _images.Add(image);
                _elements.Add(image);
            }

            return;
        }

        var translatedPattern = new OfficeDrawingImagePattern(
            pattern.EncodedBytes,
            pattern.ContentType,
            translated,
            pattern.MaximumTileCount,
            pattern.Opacity,
            useSnapshot: true);
        _imagePatterns.Add(translatedPattern);
        _elements.Add(translatedPattern);
    }
}
