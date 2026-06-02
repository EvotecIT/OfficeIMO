using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable page background image rendered behind page content.
/// </summary>
public sealed class PdfPageBackgroundImage {
    private readonly OfficeImageInfo _info;
    private byte[] _data;
    private double _opacity = 1D;
    private OfficeImageFit _fit = OfficeImageFit.Cover;

    /// <summary>Creates a page background image from supported JPEG or PNG bytes.</summary>
    public PdfPageBackgroundImage(byte[] data) {
        Guard.NotNullOrEmpty(data, nameof(data));
        _info = PdfDoc.ValidateImageBytes(data);
        _data = (byte[])data.Clone();
    }

    /// <summary>How the image is fitted into the page box.</summary>
    public OfficeImageFit Fit {
        get => _fit;
        set {
            PdfDoc.ValidateImageFit(value, nameof(Fit));
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

    internal byte[] DataSnapshot => (byte[])_data.Clone();
    internal OfficeImageInfo ImageInfo => _info;

    /// <summary>Creates a deep copy of this page background image.</summary>
    public PdfPageBackgroundImage Clone() => new PdfPageBackgroundImage(_data) {
        Fit = Fit,
        Opacity = Opacity
    };
}
