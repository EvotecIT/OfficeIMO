namespace OfficeIMO.Pdf;

/// <summary>
/// Optional dependency boundary for image formats, such as JPEG, that the dependency-free PDF core cannot decode itself.
/// </summary>
public interface IPdfRedactionImageDecoder {
    /// <summary>Attempts to decode one PDF image payload into an 8-bit RGBA raster.</summary>
    bool TryDecode(PdfRedactionImageDecodeRequest request, out PdfRedactionDecodedImage? image);
}

/// <summary>Describes an encoded PDF image passed to an optional redaction image decoder.</summary>
public sealed class PdfRedactionImageDecodeRequest {
    private readonly byte[] _encodedBytes;

    internal PdfRedactionImageDecodeRequest(PdfExtractedImage image) {
        ResourceName = image.ResourceName;
        ObjectNumber = image.ObjectNumber;
        Width = image.Width;
        Height = image.Height;
        BitsPerComponent = image.BitsPerComponent;
        ColorSpace = image.ColorSpace;
        Filter = image.Filter;
        _encodedBytes = image.Bytes;
    }

    /// <summary>PDF resource name used by the image placement.</summary>
    public string ResourceName { get; }

    /// <summary>Indirect PDF object number, or zero for a direct image stream.</summary>
    public int ObjectNumber { get; }

    /// <summary>Declared image width.</summary>
    public int Width { get; }

    /// <summary>Declared image height.</summary>
    public int Height { get; }

    /// <summary>Declared component bit depth.</summary>
    public int BitsPerComponent { get; }

    /// <summary>Resolved or declared PDF color-space name.</summary>
    public string ColorSpace { get; }

    /// <summary>PDF stream filter chain.</summary>
    public string Filter { get; }

    /// <summary>Returns a snapshot of the encoded image-file bytes.</summary>
    public byte[] EncodedBytes => (byte[])_encodedBytes.Clone();
}

/// <summary>Validated 8-bit RGBA output returned by an optional redaction image decoder.</summary>
public sealed class PdfRedactionDecodedImage {
    private readonly byte[] _rgbaPixels;

    /// <summary>Creates decoded image data. Pixels are ordered RGBA, top row first.</summary>
    public PdfRedactionDecodedImage(int width, int height, byte[] rgbaPixels) {
        Guard.PositiveInteger(width, nameof(width));
        Guard.PositiveInteger(height, nameof(height));
        Guard.NotNull(rgbaPixels, nameof(rgbaPixels));
        if ((long)width * height * 4 != rgbaPixels.Length) throw new ArgumentException("RGBA pixel length must equal width * height * 4.", nameof(rgbaPixels));
        Width = width;
        Height = height;
        _rgbaPixels = (byte[])rgbaPixels.Clone();
    }

    /// <summary>Decoded image width.</summary>
    public int Width { get; }

    /// <summary>Decoded image height.</summary>
    public int Height { get; }

    /// <summary>Returns a copy of the RGBA pixels.</summary>
    public byte[] GetRgbaPixels() => (byte[])_rgbaPixels.Clone();
}
