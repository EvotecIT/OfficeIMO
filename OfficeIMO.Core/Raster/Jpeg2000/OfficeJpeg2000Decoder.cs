namespace OfficeIMO.Drawing;

/// <summary>Colour interpretation declared by a JP2 colour specification box.</summary>
internal enum OfficeJpeg2000ColorSpace {
    /// <summary>A raw codestream or a JP2 file without a usable colour specification.</summary>
    Unspecified,
    Gray,
    Srgb,
    Sycc,
    Cmyk,
    /// <summary>A restricted or any ICC profile, available through <see cref="OfficeJpeg2000Image.IccProfile"/>.</summary>
    Icc,
    /// <summary>Another enumerated colour space; <see cref="OfficeJpeg2000Image.EnumeratedColorSpace"/> holds its value.</summary>
    Other
}

/// <summary>Limits and output shape for <see cref="OfficeJpeg2000Decoder"/>.</summary>
internal sealed class OfficeJpeg2000DecodeOptions {
    /// <summary>Largest accepted reference-grid image area.</summary>
    internal long MaximumPixels { get; set; } = OfficeRasterGuards.MaximumPixels;

    /// <summary>Budget for every buffer the decoder holds at once: code-block data, wavelet rows, planes and output.</summary>
    internal long MaximumWorkingBytes { get; set; } = OfficeRasterGuards.MaximumDecodedBytes;

    /// <summary>
    /// When false, samples are returned unscaled, which callers use for palette indices. Components wider than eight
    /// bits are then rejected.
    /// </summary>
    internal bool ScaleToEightBits { get; set; } = true;

    /// <summary>When false, JP2 palette, component-mapping and channel-definition boxes are ignored.</summary>
    internal bool ApplyContainerChannels { get; set; } = true;

    internal CancellationToken CancellationToken { get; set; }
}

/// <summary>Decoded JPEG 2000 samples, interleaved with eight bits per channel and alpha last.</summary>
internal sealed class OfficeJpeg2000Image {
    internal OfficeJpeg2000Image(int width, int height, int colorChannels, bool hasAlpha, bool premultipliedAlpha, byte[] samples) {
        Width = width;
        Height = height;
        ColorChannels = colorChannels;
        HasAlpha = hasAlpha;
        PremultipliedAlpha = premultipliedAlpha;
        Samples = samples;
    }

    internal int Width { get; }
    internal int Height { get; }

    /// <summary>Colour channels that precede the optional alpha channel.</summary>
    internal int ColorChannels { get; }
    internal bool HasAlpha { get; }
    internal bool PremultipliedAlpha { get; }
    internal int Channels => ColorChannels + (HasAlpha ? 1 : 0);
    internal byte[] Samples { get; }
    internal OfficeJpeg2000ColorSpace ColorSpace { get; set; }
    internal int EnumeratedColorSpace { get; set; } = -1;
    internal byte[]? IccProfile { get; set; }

    /// <summary>True when a JP2 palette expanded the codestream components.</summary>
    internal bool UsedPalette { get; set; }

    /// <summary>Codestream component precisions before scaling, in output channel order where known.</summary>
    internal int[] Precisions { get; set; } = Array.Empty<int>();
    internal bool HasSignedComponents { get; set; }
}

/// <summary>
/// Managed JPEG 2000 Part 1 decoder (ISO/IEC 15444-1): JP2 container and raw codestreams, all progression orders and
/// POC, precincts, SOP/EPH, PPM/PPT, every code-block style, 5/3 and 9/7 wavelets, RCT/ICT, ROI max-shift and
/// subsampled components. Inputs are untrusted: every allocation is charged to the working budget first.
/// </summary>
internal static partial class OfficeJpeg2000Decoder {
    internal static bool HasSignature(byte[] bytes) =>
        OfficeJpeg2000Header.IsJp2Container(bytes) ||
        bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0x4F && bytes[2] == 0xFF && bytes[3] == 0x51;

    /// <summary>Decodes the image; false for malformed, unsupported or over-budget input.</summary>
    internal static bool TryDecode(byte[] bytes, OfficeJpeg2000DecodeOptions options, out OfficeJpeg2000Image? image) =>
        TryDecode(bytes, options, out image, out _);

    internal static bool TryDecode(byte[] bytes, OfficeJpeg2000DecodeOptions options, out OfficeJpeg2000Image? image, out string? failure) {
        image = null;
        failure = null;
        options.CancellationToken.ThrowIfCancellationRequested();
        try {
            image = Decode(bytes, options);
            return true;
        } catch (Jpeg2000Exception exception) {
            failure = exception.Message;
            return false;
        } catch (OverflowException) {
            failure = "Arithmetic overflow while decoding the JPEG 2000 image.";
            return false;
        }
    }

    private static OfficeJpeg2000Image Decode(byte[] bytes, OfficeJpeg2000DecodeOptions options) {
        var budget = new Budget(options.MaximumWorkingBytes, options.CancellationToken);
        budget.Reserve(bytes.LongLength);
        Container container = ReadContainer(bytes, options.CancellationToken);
        Codestream codestream = ReadCodestream(bytes, container.CodestreamStart, container.CodestreamEnd, budget);
        long area = (long)codestream.Width * codestream.Height;
        if (area <= 0 || area > options.MaximumPixels) throw new Jpeg2000Exception("The JPEG 2000 image exceeds the pixel limit.");
        ComponentPlane[] planes = CreatePlanes(codestream);
        OutputLayout? layout = PlanDirectOutput(codestream, container.Header, planes, options, budget);
        if (layout == null) {
            foreach (ComponentPlane plane in planes) AllocatePlane(plane, budget);
        }
        DecodeTiles(bytes, codestream, planes, budget);
        layout ??= ResolveOutput(codestream, container.Header, planes, options, budget);
        InterleavePlanes(layout, codestream, planes, budget);
        return CreateImage(layout, container.Header, planes, options, budget);
    }

    /// <summary>Malformed, unsupported or over-budget input; caught at the decoder boundary.</summary>
    private sealed class Jpeg2000Exception : FormatException {
        internal Jpeg2000Exception(string message) : base(message) { }
    }

    /// <summary>Tracks retained bytes and coding work against the caller's limits.</summary>
    private sealed class Budget {
        private readonly long _maximumBytes;
        private long _reserved;
        private long _work;
        private readonly long _maximumWork;

        internal Budget(long maximumBytes, CancellationToken token) {
            _maximumBytes = Math.Max(0L, maximumBytes);
            // Coding work is bounded by passes times coefficients; allow generous headroom over the byte budget.
            _maximumWork = Math.Max(1L << 24, _maximumBytes * 64L);
            Token = token;
        }

        internal CancellationToken Token { get; }

        internal void Reserve(long bytes) {
            Token.ThrowIfCancellationRequested();
            if (bytes < 0 || bytes > _maximumBytes - _reserved) throw new Jpeg2000Exception("The JPEG 2000 image exceeds the working memory budget.");
            _reserved += bytes;
        }

        internal long Reserved => _reserved;

        internal void Release(long bytes) => _reserved = Math.Max(0L, _reserved - bytes);

        internal void Charge(long work) {
            if (work < 0 || work > _maximumWork - _work) throw new Jpeg2000Exception("The JPEG 2000 image exceeds the decoding work budget.");
            _work += work;
        }
    }
}
