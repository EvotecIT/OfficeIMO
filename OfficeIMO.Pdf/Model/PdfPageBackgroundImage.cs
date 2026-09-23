using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable page background image rendered behind page content.
/// </summary>
public sealed class PdfPageBackgroundImage {
    private readonly OfficeImageInfo _info;
    private readonly PdfWriter.PdfImageStream? _preparedStream;
    private byte[] _data;
    private double _opacity = 1D;
    private OfficeImageFit _fit = OfficeImageFit.Cover;

    /// <summary>Creates a page background image from raster bytes supported by OfficeIMO.Drawing.</summary>
    public PdfPageBackgroundImage(byte[] data) {
        Guard.NotNullOrEmpty(data, nameof(data));
        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(data);
        _info = prepared.Info;
        _data = prepared.Data;
        _preparedStream = prepared.PreparedStream;
    }

    /// <summary>How the image is fitted into the page box.</summary>
    public OfficeImageFit Fit {
        get => _fit;
        set {
            PdfDocument.ValidateImageFit(value, nameof(Fit));
            _fit = value;
        }
    }

    /// <summary>Image opacity from 0 to 1. Defaults to 1.</summary>
    public double Opacity {
        get => _opacity;
        set {
            if (value < 0D || value > 1D || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentOutOfRangeException(nameof(Opacity), "PDF page background image opacity must be a finite number between 0 and 1.");
            }

            _opacity = value;
        }
    }

    // _data is the prepared image's private copy and is never written, so snapshots and clones share it
    // instead of copying a (large-object-heap) cover photo several times per page.
    internal byte[] DataSnapshot => _data;
    internal OfficeImageInfo ImageInfo => _info;
    internal PdfWriter.PdfImageStream? PreparedStream => _preparedStream;

    private PdfPageBackgroundImage(PdfPageBackgroundImage source) {
        _info = source._info;
        _data = source._data;
        _preparedStream = source._preparedStream;
        _fit = source._fit;
        _opacity = source._opacity;
    }

    /// <summary>Creates a copy of this page background image. The image bytes are never modified, so the copy shares them.</summary>
    public PdfPageBackgroundImage Clone() => new(this);
}
