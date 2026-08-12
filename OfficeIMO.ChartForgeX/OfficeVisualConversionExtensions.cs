using System;
using System.Collections.Generic;
using System.Text;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Raster;
using global::ChartForgeX.SvgRaster;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;

namespace OfficeIMO.ChartForgeX;

/// <summary>Converts ChartForgeX visual artifacts into reusable Office drawing and SVG payloads.</summary>
public static class OfficeVisualConversionExtensions {
    /// <summary>Renders and converts a visual artifact using vector-preserving defaults.</summary>
    public static OfficeVisualConversionResult ToOfficeVisual(
        this VisualArtifact artifact,
        OfficeVisualConversionOptions? options = null) {
        if (artifact == null) throw new ArgumentNullException(nameof(artifact));
        options ??= new OfficeVisualConversionOptions();
        string svg = artifact.ToSvg(options.RenderOptions);
        byte[] svgBytes = Encoding.UTF8.GetBytes(svg);
        OfficeSvgDrawingReaderOptions readerOptions = CreateReaderOptions(options);
        bool imported = OfficeSvgDrawingReader.TryRead(
            svgBytes,
            readerOptions,
            out OfficeDrawing? importedDrawing,
            out int unsupportedFeatureCount);
        ThrowIfConfiguredImportLimitsWereExceeded(svgBytes, imported, readerOptions);
        double sourceWidth = importedDrawing?.Width ?? artifact.NaturalSize?.Width ?? 1D;
        double sourceHeight = importedDrawing?.Height ?? artifact.NaturalSize?.Height ?? 1D;
        (double widthPoints, double heightPoints) = options.ResolveSize(sourceWidth, sourceHeight);
        var report = new OfficeVisualConversionReport {
            UnsupportedSvgFeatureCount = unsupportedFeatureCount
        };

        OfficeDrawing drawing;
        byte[] placementBytes = svgBytes;
        OfficeVisualMediaFormat placementFormat = OfficeVisualMediaFormat.Svg;
        if (!imported || importedDrawing == null) {
            if (options.SvgPolicy == OfficeVisualSvgPolicy.RequireVector) {
                throw new InvalidOperationException("ChartForgeX SVG could not be imported as an Office drawing.");
            }
            report.Warn("The SVG payload could not be imported as an Office drawing; the OfficeDrawing result uses PNG fallback.");
            placementBytes = RasterizeRenderedSvg(svgBytes, options);
            placementFormat = OfficeVisualMediaFormat.Png;
            drawing = CreateRasterDrawing(placementBytes, widthPoints, heightPoints, ResolveAlternativeText(artifact));
            report.UsedRasterFallback = true;
        } else if (unsupportedFeatureCount > 0 && options.SvgPolicy == OfficeVisualSvgPolicy.RequireVector) {
            throw new NotSupportedException("ChartForgeX SVG contains " + unsupportedFeatureCount + " feature(s) that OfficeIMO.Drawing cannot preserve.");
        } else if (unsupportedFeatureCount > 0 && options.SvgPolicy == OfficeVisualSvgPolicy.RasterizeWhenNeeded) {
            report.Warn("The SVG importer reported " + unsupportedFeatureCount + " unsupported feature(s); the OfficeDrawing result uses PNG fallback.");
            placementBytes = RasterizeRenderedSvg(svgBytes, options);
            placementFormat = OfficeVisualMediaFormat.Png;
            drawing = CreateRasterDrawing(placementBytes, widthPoints, heightPoints, ResolveAlternativeText(artifact));
            report.UsedRasterFallback = true;
        } else {
            drawing = ScaleDrawing(importedDrawing, widthPoints, heightPoints);
            report.IsVector = true;
            if (unsupportedFeatureCount > 0) {
                report.Warn("The vector scene was preserved with " + unsupportedFeatureCount + " unsupported SVG feature(s). Inspect the Office output or choose RasterizeWhenNeeded.");
            }
        }

        if (HasLinkedRegions(artifact)) {
            report.Warn("Artifact region links are retained in the conversion result; image-based Office placements do not create per-region hyperlinks.");
        }

        return new OfficeVisualConversionResult(
            artifact,
            artifact.Id,
            artifact.Title,
            svgBytes,
            placementBytes,
            placementFormat,
            drawing,
            widthPoints,
            heightPoints,
            ResolveAlternativeText(artifact),
            artifact.Accessibility.IsDecorative,
            options.SvgPolicy,
            ConvertRegions(artifact, sourceWidth, sourceHeight, widthPoints, heightPoints),
            report);
    }

    /// <summary>Converts a portable SVG source into reusable Office drawing and placement payloads.</summary>
    public static OfficeVisualConversionResult ToOfficeVisual(
        this OfficeVisualSource source,
        OfficeVisualConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new OfficeVisualConversionOptions();
        byte[] svgBytes = source.GetSvgBytes();
        OfficeSvgDrawingReaderOptions readerOptions = CreateReaderOptions(options);
        bool imported = OfficeSvgDrawingReader.TryRead(svgBytes, readerOptions, out OfficeDrawing? importedDrawing, out int unsupportedFeatureCount);
        ThrowIfConfiguredImportLimitsWereExceeded(svgBytes, imported, readerOptions);
        double sourceWidth = importedDrawing?.Width ?? 1D;
        double sourceHeight = importedDrawing?.Height ?? 1D;
        (double widthPoints, double heightPoints) = options.ResolveSize(sourceWidth, sourceHeight);
        var report = new OfficeVisualConversionReport { UnsupportedSvgFeatureCount = unsupportedFeatureCount };
        OfficeDrawing drawing;
        byte[] placementBytes = svgBytes;
        OfficeVisualMediaFormat placementFormat = OfficeVisualMediaFormat.Svg;
        bool rasterize = !imported || importedDrawing == null || (unsupportedFeatureCount > 0 && options.SvgPolicy == OfficeVisualSvgPolicy.RasterizeWhenNeeded);
        if ((!imported || importedDrawing == null || unsupportedFeatureCount > 0) && options.SvgPolicy == OfficeVisualSvgPolicy.RequireVector) {
            throw new NotSupportedException("SVG content cannot be fully preserved as an Office drawing.");
        }
        if (rasterize) {
            placementBytes = SvgRasterizer.ToPng(svgBytes, options: options.RenderOptions?.Raster);
            if (!imported || importedDrawing == null) {
                RgbaImage raster = RasterImageDecoder.Decode(placementBytes);
                (widthPoints, heightPoints) = options.ResolveSize(raster.Width, raster.Height);
            }
            placementFormat = OfficeVisualMediaFormat.Png;
            drawing = CreateRasterDrawing(placementBytes, widthPoints, heightPoints, ResolveAlternativeText(source));
            report.UsedRasterFallback = true;
            report.Warn(imported
                ? "The SVG importer reported " + unsupportedFeatureCount + " unsupported feature(s); the Office result uses PNG fallback."
                : "The SVG payload could not be imported as an Office drawing; the Office result uses PNG fallback.");
        } else {
            drawing = ScaleDrawing(importedDrawing!, widthPoints, heightPoints);
            report.IsVector = true;
            if (unsupportedFeatureCount > 0) report.Warn("The vector scene was preserved with " + unsupportedFeatureCount + " unsupported SVG feature(s). Inspect the Office output or choose RasterizeWhenNeeded.");
        }
        return new OfficeVisualConversionResult(
            null,
            source.Id,
            source.Title,
            svgBytes,
            placementBytes,
            placementFormat,
            drawing,
            widthPoints,
            heightPoints,
            ResolveAlternativeText(source),
            source.IsDecorative,
            options.SvgPolicy,
            Array.Empty<OfficeVisualRegion>(),
            report);
    }

    /// <summary>Renders a visual artifact and returns the Office drawing plus a typed fidelity report.</summary>
    public static OfficeDrawing ToOfficeDrawing(
        this VisualArtifact artifact,
        out OfficeVisualConversionReport report,
        OfficeVisualConversionOptions? options = null) {
        OfficeVisualConversionResult result = artifact.ToOfficeVisual(options);
        report = result.Report;
        return result.Drawing;
    }

    private static OfficeDrawing ScaleDrawing(OfficeDrawing source, double width, double height) {
        if (Math.Abs(source.Width - width) < 0.000001D && Math.Abs(source.Height - height) < 0.000001D) return source;
        var target = new OfficeDrawing(width, height);
        target.AddEffectDrawing(source, OfficeTransform.Scale(width / source.Width, height / source.Height));
        return target;
    }

    private static OfficeSvgDrawingReaderOptions CreateReaderOptions(OfficeVisualConversionOptions options) => new OfficeSvgDrawingReaderOptions {
        MaximumElements = options.MaximumSvgElements,
        MaximumViewportDimension = options.MaximumSvgViewportDimension,
        MaximumViewportPixels = options.MaximumSvgViewportPixels
    };

    private static void ThrowIfConfiguredImportLimitsWereExceeded(
        byte[] svgBytes,
        bool imported,
        OfficeSvgDrawingReaderOptions configuredOptions) {
        if (imported ||
            configuredOptions.MaximumElements == OfficeSvgDrawingReaderOptions.MaximumAllowedElements &&
            configuredOptions.MaximumViewportDimension == OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension &&
            configuredOptions.MaximumViewportPixels == OfficeSvgDrawingReaderOptions.MaximumAllowedViewportPixels) {
            return;
        }

        var hardLimits = new OfficeSvgDrawingReaderOptions {
            MaximumElements = OfficeSvgDrawingReaderOptions.MaximumAllowedElements,
            MaximumViewportDimension = OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension,
            MaximumViewportPixels = OfficeSvgDrawingReaderOptions.MaximumAllowedViewportPixels
        };
        if (OfficeSvgDrawingReader.TryRead(svgBytes, hardLimits, out _)) {
            throw new InvalidOperationException("SVG content exceeds the configured import limits. Increase the matching limit only for trusted input.");
        }
    }

    private static byte[] RasterizeRenderedSvg(byte[] svgBytes, OfficeVisualConversionOptions options) =>
        SvgRasterizer.ToPng(svgBytes, options: options.RenderOptions?.Raster);

    private static OfficeDrawing CreateRasterDrawing(
        byte[] png,
        double width,
        double height,
        string alternativeText) {
        var drawing = new OfficeDrawing(width, height);
        drawing.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, width, height)),
            alternativeText);
        return drawing;
    }

    private static string ResolveAlternativeText(VisualArtifact artifact) {
        if (artifact.Accessibility.IsDecorative) return string.Empty;
        if (!string.IsNullOrWhiteSpace(artifact.Accessibility.Description)) return artifact.Accessibility.Description!;
        if (!string.IsNullOrWhiteSpace(artifact.Accessibility.Name)) return artifact.Accessibility.Name!;
        if (!string.IsNullOrWhiteSpace(artifact.Title) && !string.IsNullOrWhiteSpace(artifact.Subtitle)) return artifact.Title + ". " + artifact.Subtitle;
        if (!string.IsNullOrWhiteSpace(artifact.Title)) return artifact.Title;
        return artifact.Id;
    }

    private static string ResolveAlternativeText(OfficeVisualSource source) {
        if (source.IsDecorative) return string.Empty;
        if (!string.IsNullOrWhiteSpace(source.AlternativeText)) return source.AlternativeText;
        if (!string.IsNullOrWhiteSpace(source.Title)) return source.Title;
        return source.Id;
    }

    private static bool HasLinkedRegions(VisualArtifact artifact) {
        for (int i = 0; i < artifact.Regions.Count; i++) {
            if (!string.IsNullOrWhiteSpace(artifact.Regions[i].Href)) return true;
        }
        return false;
    }

    private static IReadOnlyList<OfficeVisualRegion> ConvertRegions(
        VisualArtifact artifact,
        double sourceWidth,
        double sourceHeight,
        double targetWidth,
        double targetHeight) {
        var regions = new List<OfficeVisualRegion>(artifact.Regions.Count);
        double scaleX = targetWidth / sourceWidth;
        double scaleY = targetHeight / sourceHeight;
        for (int i = 0; i < artifact.Regions.Count; i++) {
            VisualArtifactRegion region = artifact.Regions[i];
            ChartRect? bounds = region.Bounds;
            regions.Add(new OfficeVisualRegion(
                region.Id,
                region.Kind,
                region.Label,
                region.AlternativeText,
                region.Href,
                bounds?.X * scaleX,
                bounds?.Y * scaleY,
                bounds?.Width * scaleX,
                bounds?.Height * scaleY,
                new Dictionary<string, string>(region.Metadata, StringComparer.Ordinal)));
        }
        return regions;
    }
}
