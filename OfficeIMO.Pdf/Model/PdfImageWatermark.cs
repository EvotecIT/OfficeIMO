using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable image watermark rendered behind page content.
/// </summary>
public sealed class PdfImageWatermark {
    private readonly OfficeImageInfo _info;
    private byte[] _data;
    private double _width;
    private double _height;
    private double _opacity = 0.16D;
    private double _rotationAngle;

    /// <summary>Creates an image watermark from raster bytes supported by OfficeIMO.Drawing.</summary>
    public PdfImageWatermark(byte[] data, double width, double height) {
        Guard.NotNullOrEmpty(data, nameof(data));
        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(data);
        _info = prepared.Info;
        _data = prepared.Data;
        Width = width;
        Height = height;
    }

    /// <summary>Target watermark width in points.</summary>
    public double Width {
        get => _width;
        set {
            Guard.Positive(value, nameof(Width));
            _width = value;
        }
    }

    /// <summary>Target watermark height in points.</summary>
    public double Height {
        get => _height;
        set {
            Guard.Positive(value, nameof(Height));
            _height = value;
        }
    }

    /// <summary>Fill opacity from 0 to 1. Defaults to 0.16.</summary>
    public double Opacity {
        get => _opacity;
        set {
            if (value < 0D || value > 1D || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentOutOfRangeException(nameof(Opacity), "PDF image watermark opacity must be a finite number between 0 and 1.");
            }

            _opacity = value;
        }
    }

    /// <summary>Rotation angle in degrees. Defaults to 0.</summary>
    public double RotationAngle {
        get => _rotationAngle;
        set {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentOutOfRangeException(nameof(RotationAngle), "PDF image watermark rotation angle must be finite.");
            }

            _rotationAngle = value;
        }
    }

    internal byte[] DataSnapshot => (byte[])_data.Clone();
    internal OfficeImageInfo ImageInfo => _info;

    /// <summary>Creates a deep copy of this watermark.</summary>
    public PdfImageWatermark Clone() => new PdfImageWatermark(_data, Width, Height) {
        Opacity = Opacity,
        RotationAngle = RotationAngle
    };
}
