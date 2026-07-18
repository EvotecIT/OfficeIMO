using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Format-specific settings used by the shared raster encoder.
/// </summary>
public sealed class OfficeRasterEncodingOptions {
    /// <summary>Horizontal output resolution shared by raster encoders.</summary>
    public double DpiX { get; set; } = 96D;

    /// <summary>Vertical output resolution shared by raster encoders.</summary>
    public double DpiY { get; set; } = 96D;

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
        return new OfficeRasterEncodingOptions {
            DpiX = DpiX,
            DpiY = DpiY,
            Png = new OfficePngEncodeOptions {
                Compression = png.Compression,
                DpiX = DpiX,
                DpiY = DpiY
            },
            Jpeg = new OfficeJpegEncodeOptions {
                Quality = jpeg.Quality,
                Subsampling = jpeg.Subsampling,
                Progressive = jpeg.Progressive,
                OptimizeHuffman = jpeg.OptimizeHuffman,
                Metadata = jpeg.Metadata,
                WriteJfifHeader = jpeg.WriteJfifHeader,
                Background = jpeg.Background,
                DpiX = DpiX,
                DpiY = DpiY
            },
            Tiff = new OfficeTiffEncodeOptions {
                Compression = tiff.Compression,
                DpiX = DpiX,
                DpiY = DpiY
            }
        };
    }
}
