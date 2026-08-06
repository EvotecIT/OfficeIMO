using System;

namespace OfficeIMO.Drawing;

/// <summary>Outcome of a placement-aware encoded-image optimization request.</summary>
public enum OfficeImageOptimizationStatus {
    /// <summary>The encoded image was resized or converted.</summary>
    Optimized,
    /// <summary>The original already fit the requested placement.</summary>
    AlreadySuitable,
    /// <summary>The source format is intentionally not rewritten by the managed optimizer.</summary>
    UnsupportedFormat,
    /// <summary>The source bytes could not be decoded safely.</summary>
    DecodeFailed,
    /// <summary>The candidate was not smaller, so the original bytes were retained.</summary>
    OriginalWasSmaller
}

/// <summary>Placement-derived policy for resizing and re-encoding an image.</summary>
public sealed class OfficeImageOptimizationRequest {
    private int _targetPixelWidth;
    private int _targetPixelHeight;
    private int _jpegQuality = 85;

    /// <summary>Creates an optimization request for a target pixel bounding box.</summary>
    public OfficeImageOptimizationRequest(int targetPixelWidth, int targetPixelHeight) {
        TargetPixelWidth = targetPixelWidth;
        TargetPixelHeight = targetPixelHeight;
    }

    /// <summary>Maximum required output width in pixels.</summary>
    public int TargetPixelWidth {
        get => _targetPixelWidth;
        set {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(TargetPixelWidth));
            _targetPixelWidth = value;
        }
    }

    /// <summary>Maximum required output height in pixels.</summary>
    public int TargetPixelHeight {
        get => _targetPixelHeight;
        set {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(TargetPixelHeight));
            _targetPixelHeight = value;
        }
    }

    /// <summary>Allows enlarging source pixels when the placement is larger than the source.</summary>
    public bool AllowUpscaling { get; set; }

    /// <summary>Preserves the source aspect ratio within the requested pixel bounds.</summary>
    public bool PreserveAspectRatio { get; set; } = true;

    /// <summary>Sampling mode used when dimensions change.</summary>
    public OfficeRasterResamplingMode ResamplingMode { get; set; } = OfficeRasterResamplingMode.Bilinear;

    /// <summary>Optional PNG or JPEG output override. Null preserves JPEG and otherwise emits PNG.</summary>
    public OfficeImageFormat? OutputFormat { get; set; }

    /// <summary>JPEG quality from 1 through 100.</summary>
    public int JpegQuality {
        get => _jpegQuality;
        set {
            if (value < 1 || value > 100) throw new ArgumentOutOfRangeException(nameof(JpegQuality));
            _jpegQuality = value;
        }
    }

    /// <summary>JPEG chroma subsampling used for optimized output.</summary>
    public OfficeJpegSubsampling JpegSubsampling { get; set; } = OfficeJpegSubsampling.Y420;

    /// <summary>Background used when explicit JPEG output flattens alpha.</summary>
    public OfficeColor JpegBackground { get; set; } = OfficeColor.White;

    /// <summary>Keeps original bytes when the candidate would be the same size or larger.</summary>
    public bool KeepOriginalWhenNotSmaller { get; set; } = true;
}

/// <summary>Immutable result of encoded-image optimization.</summary>
public sealed class OfficeImageOptimizationResult {
    private readonly byte[] _bytes;

    internal OfficeImageOptimizationResult(byte[] bytes, OfficeImageOptimizationStatus status, OfficeImageInfo original, OfficeImageInfo final) {
        _bytes = (byte[])bytes.Clone();
        Status = status;
        Original = original;
        Final = final;
    }

    /// <summary>Resulting encoded bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();
    /// <summary>Optimization outcome.</summary>
    public OfficeImageOptimizationStatus Status { get; }
    /// <summary>Original image metadata.</summary>
    public OfficeImageInfo Original { get; }
    /// <summary>Final encoded image metadata.</summary>
    public OfficeImageInfo Final { get; }
    /// <summary>Whether the result contains newly encoded bytes.</summary>
    public bool Changed => Status == OfficeImageOptimizationStatus.Optimized;
    /// <summary>Signed encoded-byte reduction.</summary>
    public long BytesSaved => OriginalEncodedLength - FinalEncodedLength;
    /// <summary>Original encoded byte length.</summary>
    public long OriginalEncodedLength { get; internal set; }
    /// <summary>Final encoded byte length.</summary>
    public long FinalEncodedLength => _bytes.LongLength;
}

/// <summary>Shared dependency-free placement-aware encoded-image optimizer.</summary>
public static class OfficeImageOptimizer {
    /// <summary>Resizes PNG, JPEG, or uncompressed BMP bytes for a known placement and emits PNG or JPEG.</summary>
    public static OfficeImageOptimizationResult Optimize(byte[] encodedBytes, OfficeImageOptimizationRequest request, string? fileName = null) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (!OfficeImageReader.TryIdentify(encodedBytes, fileName, out OfficeImageInfo original)) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.UnsupportedFormat, new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0), new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0));
        }

        if (original.Format != OfficeImageFormat.Png && original.Format != OfficeImageFormat.Jpeg && original.Format != OfficeImageFormat.Bmp) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.UnsupportedFormat, original, original);
        }

        if (!OfficeRasterImageDecoder.TryDecode(encodedBytes, out OfficeRasterImage? decoded) || decoded == null) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.DecodeFailed, original, original);
        }

        ResolveDimensions(decoded.Width, decoded.Height, request, out int width, out int height);
        OfficeImageFormat outputFormat = ResolveOutputFormat(original.Format, request.OutputFormat);
        if (width == decoded.Width && height == decoded.Height && outputFormat == original.Format) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.AlreadySuitable, original, original);
        }

        OfficeRasterImage candidateImage = width == decoded.Width && height == decoded.Height
            ? decoded
            : OfficeRasterResampler.Resize(decoded, width, height, request.ResamplingMode);
        byte[] candidate = Encode(candidateImage, outputFormat, request);
        OfficeImageInfo final = new OfficeImageInfo(outputFormat, width, height, original.DpiX, original.DpiY);
        if (request.KeepOriginalWhenNotSmaller && candidate.LongLength >= encodedBytes.LongLength) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.OriginalWasSmaller, original, original);
        }

        return Result(candidate, OfficeImageOptimizationStatus.Optimized, original, final, encodedBytes.LongLength);
    }

    private static OfficeImageOptimizationResult Result(byte[] bytes, OfficeImageOptimizationStatus status, OfficeImageInfo original, OfficeImageInfo final, long? originalLength = null) =>
        new OfficeImageOptimizationResult(bytes, status, original, final) { OriginalEncodedLength = originalLength ?? bytes.LongLength };

    private static void ResolveDimensions(int sourceWidth, int sourceHeight, OfficeImageOptimizationRequest request, out int width, out int height) {
        if (!request.PreserveAspectRatio) {
            width = request.AllowUpscaling ? request.TargetPixelWidth : Math.Min(sourceWidth, request.TargetPixelWidth);
            height = request.AllowUpscaling ? request.TargetPixelHeight : Math.Min(sourceHeight, request.TargetPixelHeight);
            return;
        }

        double scale = Math.Min(request.TargetPixelWidth / (double)sourceWidth, request.TargetPixelHeight / (double)sourceHeight);
        if (!request.AllowUpscaling) scale = Math.Min(scale, 1D);
        width = Math.Max(1, (int)Math.Round(sourceWidth * scale));
        height = Math.Max(1, (int)Math.Round(sourceHeight * scale));
    }

    private static OfficeImageFormat ResolveOutputFormat(OfficeImageFormat source, OfficeImageFormat? requested) {
        OfficeImageFormat format = requested ?? (source == OfficeImageFormat.Jpeg ? OfficeImageFormat.Jpeg : OfficeImageFormat.Png);
        if (format != OfficeImageFormat.Png && format != OfficeImageFormat.Jpeg) {
            throw new ArgumentOutOfRangeException(nameof(requested), "Managed optimization output must be PNG or JPEG.");
        }
        return format;
    }

    private static byte[] Encode(OfficeRasterImage image, OfficeImageFormat format, OfficeImageOptimizationRequest request) {
        if (format == OfficeImageFormat.Png) return OfficePngWriter.Encode(image);
        return OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Quality = request.JpegQuality,
            Subsampling = request.JpegSubsampling,
            Background = request.JpegBackground
        });
    }
}
