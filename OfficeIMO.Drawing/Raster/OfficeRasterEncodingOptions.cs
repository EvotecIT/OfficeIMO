using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Format-specific settings used by the shared raster encoder.
/// </summary>
public sealed class OfficeRasterEncodingOptions {
    /// <summary>JPEG encoding settings.</summary>
    public OfficeJpegEncodeOptions Jpeg { get; set; } = new OfficeJpegEncodeOptions();

    /// <summary>TIFF encoding settings.</summary>
    public OfficeTiffEncodeOptions Tiff { get; set; } = new OfficeTiffEncodeOptions();

    /// <summary>Creates an independent copy of these settings.</summary>
    public OfficeRasterEncodingOptions Clone() {
        OfficeJpegEncodeOptions jpeg = Jpeg ?? throw new InvalidOperationException("JPEG encoding options cannot be null.");
        OfficeTiffEncodeOptions tiff = Tiff ?? throw new InvalidOperationException("TIFF encoding options cannot be null.");
        return new OfficeRasterEncodingOptions {
            Jpeg = new OfficeJpegEncodeOptions {
                Quality = jpeg.Quality,
                Subsampling = jpeg.Subsampling,
                Progressive = jpeg.Progressive,
                OptimizeHuffman = jpeg.OptimizeHuffman,
                Metadata = jpeg.Metadata,
                WriteJfifHeader = jpeg.WriteJfifHeader,
                Background = jpeg.Background
            },
            Tiff = new OfficeTiffEncodeOptions {
                Compression = tiff.Compression,
                DpiX = tiff.DpiX,
                DpiY = tiff.DpiY
            }
        };
    }
}
