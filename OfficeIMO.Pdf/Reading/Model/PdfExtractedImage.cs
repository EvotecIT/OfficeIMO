using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an image XObject extracted from a parsed PDF page.
/// </summary>
public sealed partial class PdfExtractedImage {
    private readonly PdfExtractedImagePayload? _payload;
    private readonly PdfDeferredImagePayloads? _deferredPayloads;
    private readonly int _deferredPayloadIndex;
    private PdfExtractedImagePayloadFacts? _deferredFacts;

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
        OfficeColor? imageMaskColor = null,
        OfficeIccRenderingIntent renderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        bool hasExplicitDecode = false,
        bool hasDecodeParameters = false,
        bool interpolate = false,
        bool hasAuthoredRenderingIntent = false,
        bool requiresScanDecode = false,
        bool hasUnsafePassThroughDecode = false,
        bool ownsBytes = false)
        : this(pageNumber, resourceName, objectNumber, width, height, bitsPerComponent, colorSpace, filter,
            new PdfExtractedImagePayload(bytes, ownsBytes, fileExtension, mimeType, isImageFile, transparencyMaskKind, transparencyMaskResolved),
            deferredPayloads: null, deferredPayloadIndex: 0, directStreamIdentity, isImageMask, imageMaskColor, renderingIntent,
            hasExplicitDecode, hasDecodeParameters, interpolate, hasAuthoredRenderingIntent, requiresScanDecode, hasUnsafePassThroughDecode) {
    }

    internal PdfExtractedImage(
        int pageNumber,
        string resourceName,
        int objectNumber,
        int width,
        int height,
        int bitsPerComponent,
        string colorSpace,
        string filter,
        PdfExtractedImagePayload? payload,
        PdfDeferredImagePayloads? deferredPayloads,
        int deferredPayloadIndex,
        int directStreamIdentity,
        bool isImageMask,
        OfficeColor? imageMaskColor,
        OfficeIccRenderingIntent renderingIntent,
        bool hasExplicitDecode,
        bool hasDecodeParameters,
        bool interpolate,
        bool hasAuthoredRenderingIntent,
        bool requiresScanDecode,
        bool hasUnsafePassThroughDecode) {
        if (payload is null && deferredPayloads is null) throw new ArgumentNullException(nameof(payload));
        _payload = payload;
        _deferredPayloads = payload is null ? deferredPayloads : null;
        _deferredPayloadIndex = deferredPayloadIndex;
        PageNumber = pageNumber;
        ResourceName = resourceName;
        ObjectNumber = objectNumber;
        Width = width;
        Height = height;
        BitsPerComponent = bitsPerComponent;
        ColorSpace = colorSpace;
        Filter = filter;
        DirectStreamIdentity = directStreamIdentity;
        IsImageMask = isImageMask;
        ImageMaskColor = imageMaskColor ?? OfficeColor.Black;
        RenderingIntent = renderingIntent;
        HasAuthoredRenderingIntent = hasAuthoredRenderingIntent;
        HasExplicitDecode = hasExplicitDecode;
        HasUnsafePassThroughDecode = hasUnsafePassThroughDecode;
        HasDecodeParameters = hasDecodeParameters;
        Interpolate = interpolate;
        RequiresScanDecode = requiresScanDecode;
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

    /// <summary>Whether the PDF requests interpolation while scaling the image.</summary>
    public bool Interpolate { get; }

    /// <summary>PDF color space name when available.</summary>
    public string ColorSpace { get; }

    /// <summary>PDF filter name or names when available.</summary>
    public string Filter { get; }

    // Derived from resolved PDF names; Filter is a lossy display string, not a parseable filter list.
    internal bool RequiresScanDecode { get; }

    /// <summary>
    /// Extracted bytes. JPEG images are returned as JPEG files. Simple PNG-predictor Flate
    /// images, supported ImageMask stencil streams, color-key masked simple and Indexed streams, Decode-aware simple 8-bit DeviceGray/DeviceRGB/basic-converted DeviceCMYK streams, basic ICCBased N=1/3/4 streams, and Decode-aware Indexed palette streams are returned as PNG files when their filters are supported.
    /// JPEG images with PDF transparency masks can expose unresolved mask metadata when the JPEG payload is passed through without alpha conversion. Other supported image streams return their original encoded bytes.
    /// </summary>
    public byte[] Bytes => (byte[])GetPayloadBytes(CancellationToken.None).Clone();

    /// <summary>Copies extracted image bytes without creating another in-memory payload.</summary>
    public void CopyTo(Stream destination, CancellationToken cancellationToken = default) {
        Guard.NotNull(destination, nameof(destination));
        byte[] bytes = GetPayloadBytes(cancellationToken);
        if (bytes.Length == 0) return;
        // Repeated placements can share one payload, so the destination only ever sees a chunk copy it may change.
        byte[] chunk = new byte[Math.Min(64 * 1024, bytes.Length)];
        for (int offset = 0; offset < bytes.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = Math.Min(chunk.Length, bytes.Length - offset);
            Buffer.BlockCopy(bytes, offset, chunk, 0, length);
            destination.Write(chunk, 0, length);
            offset += length;
        }
    }

    internal byte[] EncodedBytes => GetPayloadBytes(CancellationToken.None);

    /// <summary>Suggested file extension, such as jpg or png, when the bytes are a complete image file.</summary>
    public string? FileExtension => GetPayloadFacts().FileExtension;

    /// <summary>Suggested MIME type when the bytes are a complete image file.</summary>
    public string? MimeType => GetPayloadFacts().MimeType;

    /// <summary>True when <see cref="Bytes"/> is a complete image file rather than a raw PDF image stream.</summary>
    public bool IsImageFile => GetPayloadFacts().IsImageFile;

    /// <summary>PDF transparency mask kind attached to the image, such as soft-mask, color-key-mask, or explicit-mask-image.</summary>
    public string? TransparencyMaskKind => GetPayloadFacts().TransparencyMaskKind;

    /// <summary>True when the extracted image payload includes the PDF transparency mask semantics.</summary>
    public bool TransparencyMaskResolved => GetPayloadFacts().TransparencyMaskResolved;

    // Images read through the logical reading model build their payload on first use; the page's deferred payloads
    // keep the bytes only while somebody holds them, while the facts about the payload stay with the image.
    private byte[] GetPayloadBytes(CancellationToken cancellationToken) {
        if (_payload is not null) return _payload.Bytes;
        PdfExtractedImagePayload payload = _deferredPayloads!.Get(_deferredPayloadIndex, cancellationToken);
        Volatile.Write(ref _deferredFacts, payload.Facts);
        return payload.Bytes;
    }

    private PdfExtractedImagePayloadFacts GetPayloadFacts() {
        if (_payload is not null) return _payload.Facts;
        PdfExtractedImagePayloadFacts? facts = Volatile.Read(ref _deferredFacts);
        if (facts is not null) return facts;
        facts = _deferredPayloads!.Get(_deferredPayloadIndex, CancellationToken.None).Facts;
        Volatile.Write(ref _deferredFacts, facts);
        return facts;
    }

    /// <summary>True when the source PDF image declared a transparency mask.</summary>
    public bool HasTransparencyMask => !string.IsNullOrWhiteSpace(TransparencyMaskKind);

    /// <summary>True when the source PDF image declared a transparency mask that is not represented by the extracted payload.</summary>
    public bool HasUnresolvedTransparencyMask => HasTransparencyMask && !TransparencyMaskResolved;

    /// <summary>True when this image came from a PDF /ImageMask stencil XObject.</summary>
    public bool IsImageMask { get; }

    /// <summary>Paint color applied to a stencil image mask.</summary>
    public OfficeColor ImageMaskColor { get; }

    /// <summary>Effective ICC rendering intent used while normalizing extracted samples.</summary>
    public OfficeIccRenderingIntent RenderingIntent { get; }

    /// <summary>True when the source image or active graphics state explicitly selected a rendering intent.</summary>
    public bool HasAuthoredRenderingIntent { get; }

    /// <summary>True when the image declared a Decode array.</summary>
    public bool HasExplicitDecode { get; }

    /// <summary>True when passing through the encoded image would omit a non-identity or unrecognized Decode mapping.</summary>
    internal bool HasUnsafePassThroughDecode { get; }

    /// <summary>True when the image filter declared DecodeParms.</summary>
    public bool HasDecodeParameters { get; }
}
