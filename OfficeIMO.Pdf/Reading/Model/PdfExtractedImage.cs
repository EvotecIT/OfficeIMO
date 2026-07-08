using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an image XObject extracted from a parsed PDF page.
/// </summary>
public sealed class PdfExtractedImage {
    internal PdfExtractedImage(
        int pageNumber,
        string resourceName,
        int objectNumber,
        int width,
        int height,
        int bitsPerComponent,
        string colorSpace,
        string filter,
        byte[] bytes,
        string? fileExtension,
        string? mimeType,
        bool isImageFile,
        string? transparencyMaskKind = null,
        bool transparencyMaskResolved = false,
        int directStreamIdentity = 0,
        bool isImageMask = false,
        OfficeColor? imageMaskColor = null) {
        PageNumber = pageNumber;
        ResourceName = resourceName;
        ObjectNumber = objectNumber;
        Width = width;
        Height = height;
        BitsPerComponent = bitsPerComponent;
        ColorSpace = colorSpace;
        Filter = filter;
        Bytes = bytes;
        FileExtension = fileExtension;
        MimeType = mimeType;
        IsImageFile = isImageFile;
        TransparencyMaskKind = transparencyMaskKind;
        TransparencyMaskResolved = transparencyMaskResolved;
        DirectStreamIdentity = directStreamIdentity;
        IsImageMask = isImageMask;
        ImageMaskColor = imageMaskColor ?? OfficeColor.Black;
    }

    /// <summary>One-based page number containing the image resource.</summary>
    public int PageNumber { get; }

    /// <summary>Image resource name from the page XObject dictionary.</summary>
    public string ResourceName { get; }

    /// <summary>PDF object number for the image stream, or 0 when the image is direct.</summary>
    public int ObjectNumber { get; }

    /// <summary>Runtime identity for a direct image stream, or 0 when the image is indirect.</summary>
    internal int DirectStreamIdentity { get; }

    /// <summary>Image width in pixels.</summary>
    public int Width { get; }

    /// <summary>Image height in pixels.</summary>
    public int Height { get; }

    /// <summary>Bits per color component.</summary>
    public int BitsPerComponent { get; }

    /// <summary>PDF color space name when available.</summary>
    public string ColorSpace { get; }

    /// <summary>PDF filter name or names when available.</summary>
    public string Filter { get; }

    /// <summary>
    /// Extracted bytes. JPEG images are returned as JPEG files. Simple PNG-predictor Flate
    /// images, supported ImageMask stencil streams, color-key masked simple and Indexed streams, Decode-aware simple 8-bit DeviceGray/DeviceRGB/basic-converted DeviceCMYK streams, basic ICCBased N=1/3/4 streams, and Decode-aware Indexed palette streams are returned as PNG files when their filters are supported.
    /// JPEG images with PDF transparency masks can expose unresolved mask metadata when the JPEG payload is passed through without alpha conversion. Other supported image streams return their original encoded bytes.
    /// </summary>
    public byte[] Bytes { get; }

    /// <summary>Suggested file extension, such as jpg or png, when the bytes are a complete image file.</summary>
    public string? FileExtension { get; }

    /// <summary>Suggested MIME type when the bytes are a complete image file.</summary>
    public string? MimeType { get; }

    /// <summary>True when <see cref="Bytes"/> is a complete image file rather than a raw PDF image stream.</summary>
    public bool IsImageFile { get; }

    /// <summary>PDF transparency mask kind attached to the image, such as soft-mask, color-key-mask, or explicit-mask-image.</summary>
    public string? TransparencyMaskKind { get; }

    /// <summary>True when the extracted image payload includes the PDF transparency mask semantics.</summary>
    public bool TransparencyMaskResolved { get; }

    /// <summary>True when the source PDF image declared a transparency mask.</summary>
    public bool HasTransparencyMask => !string.IsNullOrWhiteSpace(TransparencyMaskKind);

    /// <summary>True when the source PDF image declared a transparency mask that is not represented by the extracted payload.</summary>
    public bool HasUnresolvedTransparencyMask => HasTransparencyMask && !TransparencyMaskResolved;

    /// <summary>True when this image came from a PDF /ImageMask stencil XObject.</summary>
    public bool IsImageMask { get; }

    internal OfficeColor ImageMaskColor { get; }
}
