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
    private double? _outputDpiX;
    private double? _outputDpiY;

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

    /// <summary>Optional PNG, JPEG, TIFF, or WebP output override. Null preserves JPEG and otherwise emits PNG.</summary>
    public OfficeImageFormat? OutputFormat { get; set; }

    /// <summary>
    /// Optional horizontal output resolution. Null preserves source DPI. Values outside the selected
    /// destination format's metadata range are clamped to its nearest representable limit.
    /// </summary>
    public double? OutputDpiX {
        get => _outputDpiX;
        set {
            ValidateOutputDpi(value, nameof(OutputDpiX));
            _outputDpiX = value;
        }
    }

    /// <summary>
    /// Optional vertical output resolution. Null preserves source DPI. Values outside the selected
    /// destination format's metadata range are clamped to its nearest representable limit.
    /// </summary>
    public double? OutputDpiY {
        get => _outputDpiY;
        set {
            ValidateOutputDpi(value, nameof(OutputDpiY));
            _outputDpiY = value;
        }
    }

    /// <summary>PNG compression used when optimized output is PNG.</summary>
    public OfficePngCompression PngCompression { get; set; } = OfficePngCompression.Optimal;

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

    /// <summary>Writes progressive JPEG scans when optimized output is JPEG.</summary>
    public bool JpegProgressive { get; set; }

    /// <summary>Builds image-specific Huffman tables when optimized output is JPEG.</summary>
    public bool JpegOptimizeHuffman { get; set; }

    /// <summary>Background used when explicit JPEG output flattens alpha.</summary>
    public OfficeColor JpegBackground { get; set; } = OfficeColor.White;

    /// <summary>TIFF strip compression used when optimized output is TIFF.</summary>
    public OfficeTiffCompression TiffCompression { get; set; } = OfficeTiffCompression.PackBits;

    /// <summary>Keeps original bytes when the candidate would be the same size or larger.</summary>
    public bool KeepOriginalWhenNotSmaller { get; set; } = true;

    private static void ValidateOutputDpi(double? value, string paramName) {
        if (value.HasValue && (value.Value <= 0D || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new ArgumentOutOfRangeException(paramName, "Output DPI must be finite and greater than zero.");
        }
    }
}

/// <summary>Immutable result of encoded-image optimization.</summary>
public sealed class OfficeImageOptimizationResult {
    private readonly byte[] _bytes;

    internal OfficeImageOptimizationResult(
        byte[] bytes,
        OfficeImageOptimizationStatus status,
        OfficeImageInfo original,
        OfficeImageInfo final,
        bool takeOwnership) {
        _bytes = takeOwnership ? bytes : (byte[])bytes.Clone();
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
    /// <summary>
    /// Resizes managed static raster input for a known placement and emits PNG, JPEG, TIFF, or WebP.
    /// Animated input is rejected so optimization never silently discards frames.
    /// </summary>
    public static OfficeImageOptimizationResult Optimize(byte[] encodedBytes, OfficeImageOptimizationRequest request, string? fileName = null) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (!OfficeImageReader.TryIdentify(encodedBytes, fileName, out OfficeImageInfo original)) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.UnsupportedFormat, new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0), new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0));
        }

        if (!IsSupportedInputFormat(original.Format)) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.UnsupportedFormat, original, original);
        }

        var decodeOptions = new OfficeRasterDecodeOptions {
            AnimationPolicy = OfficeRasterAnimationPolicy.RejectAnimated
        };
        if (!OfficeRasterImageDecoder.TryDecode(encodedBytes, decodeOptions, out OfficeRasterImage? decoded, out _) || decoded == null) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.DecodeFailed, original, original);
        }

        ResolveDimensions(decoded.Width, decoded.Height, request, out int width, out int height);
        OfficeImageFormat outputFormat = ResolveOutputFormat(original.Format, request.OutputFormat);
        if (width == decoded.Width && height == decoded.Height && outputFormat == original.Format &&
            !request.OutputDpiX.HasValue && !request.OutputDpiY.HasValue) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.AlreadySuitable, original, original);
        }

        OfficeRasterImage candidateImage = width == decoded.Width && height == decoded.Height
            ? decoded
            : OfficeRasterResampler.Resize(decoded, width, height, request.ResamplingMode);
        byte[] candidate = Encode(candidateImage, outputFormat, original, request);
        OfficeImageInfo final = OfficeImageReader.Identify(candidate);
        if (request.KeepOriginalWhenNotSmaller && candidate.LongLength >= encodedBytes.LongLength) {
            return Result(encodedBytes, OfficeImageOptimizationStatus.OriginalWasSmaller, original, original);
        }

        return Result(candidate, OfficeImageOptimizationStatus.Optimized, original, final, encodedBytes.LongLength, takeOwnership: true);
    }

    private static OfficeImageOptimizationResult Result(
        byte[] bytes,
        OfficeImageOptimizationStatus status,
        OfficeImageInfo original,
        OfficeImageInfo final,
        long? originalLength = null,
        bool takeOwnership = false) =>
        new OfficeImageOptimizationResult(bytes, status, original, final, takeOwnership) {
            OriginalEncodedLength = originalLength ?? bytes.LongLength
        };

    private static bool IsSupportedInputFormat(OfficeImageFormat format) =>
        format == OfficeImageFormat.Png ||
        format == OfficeImageFormat.Jpeg ||
        format == OfficeImageFormat.Gif ||
        format == OfficeImageFormat.Bmp ||
        format == OfficeImageFormat.Tiff ||
        format == OfficeImageFormat.Webp;

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
        if (format != OfficeImageFormat.Png &&
            format != OfficeImageFormat.Jpeg &&
            format != OfficeImageFormat.Tiff &&
            format != OfficeImageFormat.Webp) {
            throw new ArgumentOutOfRangeException(nameof(requested), "Managed optimization output must be PNG, JPEG, TIFF, or WebP.");
        }
        return format;
    }

    private static byte[] Encode(
        OfficeRasterImage image,
        OfficeImageFormat format,
        OfficeImageInfo original,
        OfficeImageOptimizationRequest request) {
        OfficeImageExportFormat exportFormat = ToExportFormat(format);
        double dpiX = OfficeRasterImageEncoder.NormalizeDpi(exportFormat, request.OutputDpiX ?? original.DpiX);
        double dpiY = OfficeRasterImageEncoder.NormalizeDpi(exportFormat, request.OutputDpiY ?? original.DpiY);
        var options = new OfficeRasterEncodingOptions {
            DpiX = dpiX,
            DpiY = dpiY,
            Png = new OfficePngEncodeOptions {
                Compression = request.PngCompression
            },
            Jpeg = new OfficeJpegEncodeOptions {
                Quality = request.JpegQuality,
                Subsampling = request.JpegSubsampling,
                Progressive = request.JpegProgressive,
                OptimizeHuffman = request.JpegOptimizeHuffman,
                Background = request.JpegBackground
            },
            Tiff = new OfficeTiffEncodeOptions {
                Compression = request.TiffCompression
            }
        };
        return OfficeRasterImageEncoder.Encode(image, exportFormat, options);
    }

    private static OfficeImageExportFormat ToExportFormat(OfficeImageFormat format) => format switch {
        OfficeImageFormat.Png => OfficeImageExportFormat.Png,
        OfficeImageFormat.Jpeg => OfficeImageExportFormat.Jpeg,
        OfficeImageFormat.Tiff => OfficeImageExportFormat.Tiff,
        OfficeImageFormat.Webp => OfficeImageExportFormat.Webp,
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };
}
