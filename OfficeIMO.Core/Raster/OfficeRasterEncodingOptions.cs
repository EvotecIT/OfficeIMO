using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Format-specific settings used by the shared raster encoder.
/// </summary>
public sealed class OfficeRasterEncodingOptions {
    internal double ResolvedDpiX { get; private set; } = 96D;
    internal double ResolvedDpiY { get; private set; } = 96D;

    /// <summary>Writes format-appropriate resolution metadata when supported.</summary>
    public bool WriteResolutionMetadata { get; set; } = true;

    /// <summary>
    /// Explicit native resolution override shared by raster encoders. Null retains the selected format's own settings.
    /// </summary>
    /// <remarks>TIFF retains native units, including aspect ratio. Other plain encoders require physical resolution;
    /// use EncodeWithMetadata to preserve a unitless ratio in PNG, JPEG, or GIF container metadata.</remarks>
    public OfficeImageResolution? Resolution { get; set; }

    /// <summary>PNG encoding settings.</summary>
    public OfficePngEncodeOptions Png { get; set; } = new OfficePngEncodeOptions();

    /// <summary>JPEG encoding settings.</summary>
    public OfficeJpegEncodeOptions Jpeg { get; set; } = new OfficeJpegEncodeOptions();

    /// <summary>TIFF encoding settings.</summary>
    public OfficeTiffEncodeOptions Tiff { get; set; } = new OfficeTiffEncodeOptions();

    /// <summary>WebP compression mode, quality, and density settings.</summary>
    public OfficeWebpEncodeOptions Webp { get; set; } = new OfficeWebpEncodeOptions { WritePhysicalResolution = true };

    /// <summary>Creates an independent copy of these settings.</summary>
    public OfficeRasterEncodingOptions Clone() {
        OfficePngEncodeOptions png = Png ?? throw new InvalidOperationException("PNG encoding options cannot be null.");
        OfficeJpegEncodeOptions jpeg = Jpeg ?? throw new InvalidOperationException("JPEG encoding options cannot be null.");
        OfficeTiffEncodeOptions tiff = Tiff ?? throw new InvalidOperationException("TIFF encoding options cannot be null.");
        OfficeWebpEncodeOptions webp = Webp ?? throw new InvalidOperationException("WebP encoding options cannot be null.");
        var clone = new OfficeRasterEncodingOptions {
            Png = new OfficePngEncodeOptions {
                Compression = png.Compression,
                DpiX = png.DpiX,
                DpiY = png.DpiY,
                WritePhysicalResolution = png.WritePhysicalResolution
            },
            Jpeg = new OfficeJpegEncodeOptions {
                Quality = jpeg.Quality,
                Subsampling = jpeg.Subsampling,
                Progressive = jpeg.Progressive,
                OptimizeHuffman = jpeg.OptimizeHuffman,
                Metadata = jpeg.Metadata,
                WriteJfifHeader = jpeg.WriteJfifHeader,
                Background = jpeg.Background,
                DpiX = jpeg.DpiX,
                DpiY = jpeg.DpiY,
                RetainedManagedBytes = jpeg.RetainedManagedBytes
            },
            Tiff = new OfficeTiffEncodeOptions {
                Compression = tiff.Compression,
                Predictor = tiff.Predictor,
                DpiX = tiff.DpiX,
                DpiY = tiff.DpiY,
                Resolution = tiff.Resolution,
                WriteResolution = tiff.WriteResolution
            },
            Webp = new OfficeWebpEncodeOptions {
                Mode = webp.Mode,
                Quality = webp.Quality,
                DpiX = webp.DpiX,
                DpiY = webp.DpiY,
                WritePhysicalResolution = webp.WritePhysicalResolution,
                RetainedManagedBytes = webp.RetainedManagedBytes
            }
        };
        clone.WriteResolutionMetadata = WriteResolutionMetadata;
        clone.Resolution = Resolution;
        clone.ResolvedDpiX = ResolvedDpiX;
        clone.ResolvedDpiY = ResolvedDpiY;
        return clone;
    }

    internal OfficeRasterEncodingOptions Resolve(
        OfficeImageExportFormat format,
        double scaleRatio = 1D) {
        if (!format.IsRaster()) {
            throw new ArgumentException("A raster output format is required.", nameof(format));
        }
        if (scaleRatio <= 0D || double.IsNaN(scaleRatio) || double.IsInfinity(scaleRatio)) {
            throw new ArgumentOutOfRangeException(nameof(scaleRatio));
        }

        OfficeRasterEncodingOptions resolved = Clone();
        OfficeImageResolution? resolution = Resolution;
        if (resolution != null && resolution.Unit == OfficeImageResolutionUnit.AspectRatio &&
            format != OfficeImageExportFormat.Tiff && resolved.WriteResolutionMetadata) {
            throw new ArgumentException("A unitless resolution override requires TIFF or the metadata-aware encoding operation.", nameof(Resolution));
        }
        double dpiX;
        double dpiY;
        switch (format) {
            case OfficeImageExportFormat.Png:
                dpiX = resolution?.PhysicalDpiX ?? resolved.Png.DpiX;
                dpiY = resolution?.PhysicalDpiY ?? resolved.Png.DpiY;
                resolved.Png.DpiX = dpiX * scaleRatio;
                resolved.Png.DpiY = dpiY * scaleRatio;
                resolved.Png.WritePhysicalResolution &= resolved.WriteResolutionMetadata;
                break;
            case OfficeImageExportFormat.Jpeg:
                dpiX = resolution?.PhysicalDpiX ?? resolved.Jpeg.DpiX;
                dpiY = resolution?.PhysicalDpiY ?? resolved.Jpeg.DpiY;
                resolved.Jpeg.DpiX = dpiX * scaleRatio;
                resolved.Jpeg.DpiY = dpiY * scaleRatio;
                resolved.Jpeg.WriteJfifHeader &= resolved.WriteResolutionMetadata;
                break;
            case OfficeImageExportFormat.Tiff:
                dpiX = resolution?.PhysicalDpiX ?? resolved.Tiff.Resolution?.PhysicalDpiX ?? resolved.Tiff.DpiX;
                dpiY = resolution?.PhysicalDpiY ?? resolved.Tiff.Resolution?.PhysicalDpiY ?? resolved.Tiff.DpiY;
                // Native-derived values belong to Resolution, whose rational storage
                // bounds differ from authored legacy DPI. Keep them out of DpiX/Y;
                // explicit shared overrides use native rational storage as well.
                if (resolved.Tiff.Resolution == null && resolution == null) {
                    resolved.Tiff.DpiX = dpiX * scaleRatio;
                }
                if (resolved.Tiff.Resolution == null && resolution == null) {
                    resolved.Tiff.DpiY = dpiY * scaleRatio;
                }
                if (resolution != null || resolved.Tiff.Resolution != null) {
                    OfficeImageResolution native = resolution ?? resolved.Tiff.Resolution!;
                    resolved.Tiff.Resolution = new OfficeImageResolution(native.Horizontal * scaleRatio, native.Vertical * scaleRatio, native.Unit);
                }
                resolved.Tiff.WriteResolution &= resolved.WriteResolutionMetadata;
                break;
            case OfficeImageExportFormat.Webp:
                dpiX = resolution?.PhysicalDpiX ?? resolved.Webp.DpiX;
                dpiY = resolution?.PhysicalDpiY ?? resolved.Webp.DpiY;
                resolved.Webp.DpiX = dpiX * scaleRatio;
                resolved.Webp.DpiY = dpiY * scaleRatio;
                resolved.Webp.WritePhysicalResolution &= resolved.WriteResolutionMetadata;
                break;
            case OfficeImageExportFormat.Bmp:
            case OfficeImageExportFormat.Pbm:
            case OfficeImageExportFormat.Tga:
            case OfficeImageExportFormat.Icon:
                dpiX = resolution?.PhysicalDpiX ?? 96D;
                dpiY = resolution?.PhysicalDpiY ?? 96D;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }

        resolved.ResolvedDpiX = dpiX * scaleRatio;
        resolved.ResolvedDpiY = dpiY * scaleRatio;
        if (resolution != null) resolved.Resolution = new OfficeImageResolution(resolution.Horizontal * scaleRatio, resolution.Vertical * scaleRatio, resolution.Unit);
        // Preserve native units through nested export/streaming resolution.
        return resolved;
    }
}
