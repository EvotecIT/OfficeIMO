using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Format-specific settings used by the shared raster encoder.
/// </summary>
public sealed class OfficeRasterEncodingOptions {
    private double _dpiX = 96D;
    private double _dpiY = 96D;
    private bool _hasExplicitDpiX;
    private bool _hasExplicitDpiY;

    /// <summary>
    /// Horizontal output resolution shared by raster encoders when explicitly assigned.
    /// Otherwise the selected format's DPI setting remains authoritative.
    /// </summary>
    public double DpiX {
        get => _dpiX;
        set {
            _dpiX = value;
            _hasExplicitDpiX = true;
        }
    }

    /// <summary>
    /// Vertical output resolution shared by raster encoders when explicitly assigned.
    /// Otherwise the selected format's DPI setting remains authoritative.
    /// </summary>
    public double DpiY {
        get => _dpiY;
        set {
            _dpiY = value;
            _hasExplicitDpiY = true;
        }
    }

    /// <summary>PNG encoding settings.</summary>
    public OfficePngEncodeOptions Png { get; set; } = new OfficePngEncodeOptions();

    /// <summary>JPEG encoding settings.</summary>
    public OfficeJpegEncodeOptions Jpeg { get; set; } = new OfficeJpegEncodeOptions();

    /// <summary>TIFF encoding settings.</summary>
    public OfficeTiffEncodeOptions Tiff { get; set; } = new OfficeTiffEncodeOptions();

    /// <summary>Creates an independent copy of these settings.</summary>
    public OfficeRasterEncodingOptions Clone() {
        OfficePngEncodeOptions png = Png ?? throw new InvalidOperationException("PNG encoding options cannot be null.");
        OfficeJpegEncodeOptions jpeg = Jpeg ?? throw new InvalidOperationException("JPEG encoding options cannot be null.");
        OfficeTiffEncodeOptions tiff = Tiff ?? throw new InvalidOperationException("TIFF encoding options cannot be null.");
        var clone = new OfficeRasterEncodingOptions {
            Png = new OfficePngEncodeOptions {
                Compression = png.Compression,
                DpiX = png.DpiX,
                DpiY = png.DpiY
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
                DpiY = jpeg.DpiY
            },
            Tiff = new OfficeTiffEncodeOptions {
                Compression = tiff.Compression,
                DpiX = tiff.DpiX,
                DpiY = tiff.DpiY
            }
        };
        clone._dpiX = _dpiX;
        clone._dpiY = _dpiY;
        clone._hasExplicitDpiX = _hasExplicitDpiX;
        clone._hasExplicitDpiY = _hasExplicitDpiY;
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
        double dpiX;
        double dpiY;
        switch (format) {
            case OfficeImageExportFormat.Png:
                dpiX = _hasExplicitDpiX ? _dpiX : resolved.Png.DpiX;
                dpiY = _hasExplicitDpiY ? _dpiY : resolved.Png.DpiY;
                resolved.Png.DpiX = dpiX * scaleRatio;
                resolved.Png.DpiY = dpiY * scaleRatio;
                break;
            case OfficeImageExportFormat.Jpeg:
                dpiX = _hasExplicitDpiX ? _dpiX : resolved.Jpeg.DpiX;
                dpiY = _hasExplicitDpiY ? _dpiY : resolved.Jpeg.DpiY;
                resolved.Jpeg.DpiX = dpiX * scaleRatio;
                resolved.Jpeg.DpiY = dpiY * scaleRatio;
                break;
            case OfficeImageExportFormat.Tiff:
                dpiX = _hasExplicitDpiX ? _dpiX : resolved.Tiff.DpiX;
                dpiY = _hasExplicitDpiY ? _dpiY : resolved.Tiff.DpiY;
                resolved.Tiff.DpiX = dpiX * scaleRatio;
                resolved.Tiff.DpiY = dpiY * scaleRatio;
                break;
            case OfficeImageExportFormat.Webp:
                dpiX = _dpiX;
                dpiY = _dpiY;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }

        resolved._dpiX = dpiX * scaleRatio;
        resolved._dpiY = dpiY * scaleRatio;
        resolved._hasExplicitDpiX = true;
        resolved._hasExplicitDpiY = true;
        return resolved;
    }
}
