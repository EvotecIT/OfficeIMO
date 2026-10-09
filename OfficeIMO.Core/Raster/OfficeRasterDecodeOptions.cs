namespace OfficeIMO.Drawing;

/// <summary>Policy used when a raster container exposes more than one frame or page.</summary>
public enum OfficeRasterFrameLossPolicy {
    /// <summary>Decode the explicitly selected unit and report the units not retained in the static result.</summary>
    UseSelectedFrame,

    /// <summary>Reject multi-frame or multi-page input instead of returning a lossy static result.</summary>
    RejectMultipleFrames
}

/// <summary>Reason a managed raster request did not produce a result.</summary>
public enum OfficeRasterDecodeFailure {
    /// <summary>The request succeeded.</summary>
    None,
    /// <summary>The input is invalid or outside the supported codec subset.</summary>
    InvalidOrUnsupported,
    /// <summary>The decoder could not distinguish malformed input, an unsupported subset, or an internal resource rejection.</summary>
    Unclassified,
    /// <summary>The encoded input exceeds the requested byte limit.</summary>
    EncodedLimit,
    /// <summary>The decoded result exceeds the requested aggregate pixel limit.</summary>
    DecodedLimit,
    /// <summary>The selected frame or page does not exist.</summary>
    FrameSelectionOutOfRange,
    /// <summary>The frame-loss policy rejects the source.</summary>
    FramePolicyRejected,
    /// <summary>The complete sequence exceeds the requested frame count.</summary>
    FrameCountLimit
}

/// <summary>Shared options for deterministic raster decoding.</summary>
public sealed class OfficeRasterDecodeOptions {
    private int _frameIndex;
    private long _maximumDecodedPixels = 50_000_000L;
    private int _maximumEncodedBytes = 128 * 1024 * 1024;
    private OfficeRasterFrameLossPolicy _frameLossPolicy = OfficeRasterFrameLossPolicy.UseSelectedFrame;

    /// <summary>Zero-based frame or page to decode.</summary>
    public int FrameIndex {
        get => _frameIndex;
        set {
            if (value < 0) throw new System.ArgumentOutOfRangeException(nameof(FrameIndex));
            _frameIndex = value;
        }
    }

    /// <summary>Behavior when the source contains more than one frame or page.</summary>
    public OfficeRasterFrameLossPolicy FrameLossPolicy {
        get => _frameLossPolicy;
        set {
            if (value != OfficeRasterFrameLossPolicy.UseSelectedFrame &&
                value != OfficeRasterFrameLossPolicy.RejectMultipleFrames) {
                throw new System.ArgumentOutOfRangeException(nameof(FrameLossPolicy));
            }
            _frameLossPolicy = value;
        }
    }

    /// <summary>Maximum encoded bytes read or decoded by this request.</summary>
    public int MaximumEncodedBytes {
        get => _maximumEncodedBytes;
        set {
            if (value < 1 || value > 128 * 1024 * 1024) throw new System.ArgumentOutOfRangeException(nameof(MaximumEncodedBytes));
            _maximumEncodedBytes = value;
        }
    }

    /// <summary>Maximum decoded pixels retained by this request.</summary>
    public long MaximumDecodedPixels {
        get => _maximumDecodedPixels;
        set {
            if (value < 1L || value > 50_000_000L) throw new System.ArgumentOutOfRangeException(nameof(MaximumDecodedPixels));
            _maximumDecodedPixels = value;
        }
    }

    /// <summary>Optional trusted decoder for inspected raster payloads not decoded by the managed core.</summary>
    /// <remarks>The codec is used for frame zero only, after container and resource-policy checks.
    /// Cancellation is checked before and after the synchronous codec call.</remarks>
    public IOfficeRasterImageCodec? ImageCodec { get; set; }

    /// <summary>Cancellation observed while reading, parsing, or decoding the request.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    /// <summary>Whether orientation-aware decoders normalize EXIF display orientation; defaults to true.</summary>
    /// <remarks>Set false to retain JPEG and TIFF stored sample order for workflows that apply orientation explicitly.
    /// This does not add orientation processing to formats whose decoder already returns stored sample order.</remarks>
    public bool ApplyExifOrientation { get; set; } = true;

    /// <summary>Creates an independent settings snapshot, retaining the trusted codec and cancellation token by reference/value.</summary>
    public OfficeRasterDecodeOptions Clone() => WithAdditionalRetainedManagedBytes(0L);

    internal long RetainedManagedBytes { get; set; }
    internal long MaximumInspectionWorkPixels { get; set; } = OfficeRasterGuards.MaximumPixels;
    /// <summary>Owned fixed-page adapters use TIFF sample order instead of its optional display orientation.</summary>
    internal bool IgnoreTiffOrientation { get; set; }

    internal OfficeRasterDecodeOptions WithAdditionalRetainedManagedBytes(long bytes) {
        if (bytes < 0L) throw new System.ArgumentOutOfRangeException(nameof(bytes));
        return new OfficeRasterDecodeOptions {
            ImageCodec = ImageCodec,
            FrameIndex = FrameIndex,
            FrameLossPolicy = FrameLossPolicy,
            MaximumEncodedBytes = MaximumEncodedBytes,
            MaximumDecodedPixels = MaximumDecodedPixels,
            MaximumInspectionWorkPixels = MaximumInspectionWorkPixels,
            IgnoreTiffOrientation = IgnoreTiffOrientation,
            ApplyExifOrientation = ApplyExifOrientation,
            CancellationToken = CancellationToken,
            RetainedManagedBytes = checked(RetainedManagedBytes + bytes)
        };
    }

    internal void Validate() {
        if (_frameLossPolicy != OfficeRasterFrameLossPolicy.UseSelectedFrame &&
            _frameLossPolicy != OfficeRasterFrameLossPolicy.RejectMultipleFrames) {
            throw new System.ArgumentOutOfRangeException(nameof(FrameLossPolicy));
        }
        if (RetainedManagedBytes < 0L) throw new System.ArgumentOutOfRangeException(nameof(RetainedManagedBytes));
        if (MaximumInspectionWorkPixels < 1L || MaximumInspectionWorkPixels > OfficeRasterGuards.MaximumPixels) {
            throw new System.ArgumentOutOfRangeException(nameof(MaximumInspectionWorkPixels));
        }
    }
}

/// <summary>Typed evidence describing one shared raster decode decision.</summary>
public sealed class OfficeRasterDecodeInfo {
    internal bool UsedCallerCodec { get; set; }

    internal OfficeRasterDecodeInfo(OfficeImageFormat format, int frameCount, int selectedFrameIndex, bool succeeded, string? diagnostic, OfficeRasterContainerInfo? container = null) {
        Format = format;
        FrameCount = frameCount;
        SelectedFrameIndex = selectedFrameIndex;
        Succeeded = succeeded;
        Diagnostic = diagnostic;
        Container = container;
        Failure = succeeded ? OfficeRasterDecodeFailure.None : OfficeRasterDecodeFailure.Unclassified;
    }

    /// <summary>Detected source format, or <see cref="OfficeImageFormat.Unknown"/>.</summary>
    public OfficeImageFormat Format { get; }

    /// <summary>Known frame count. Static formats report one; zero means the count could not be established.</summary>
    public int FrameCount { get; }

    /// <summary>Requested zero-based frame index.</summary>
    public int SelectedFrameIndex { get; }

    /// <summary>True when the requested selected frame or complete sequence was decoded.</summary>
    public bool Succeeded { get; }

    /// <summary>Typed reason for failure; cancellation and invalid caller options still throw.</summary>
    public OfficeRasterDecodeFailure Failure { get; internal set; }

    /// <summary>True when every source frame or page was retained by the operation.</summary>
    public bool DecodedAllFrames { get; internal set; }

    /// <summary>True when decoding normalized a non-default EXIF orientation in at least one retained unit.</summary>
    public bool OrientationNormalized { get; internal set; }

    /// <summary>Stored orientation of the selected source unit, or Normal when absent.</summary>
    public OfficeImageOrientation SourceOrientation { get; internal set; } = OfficeImageOrientation.Normal;

    /// <summary>True when the inspected source contains timed animation frames.</summary>
    public bool IsAnimated => Container?.IsAnimated ?? FrameCount > 1;

    /// <summary>True when the inspected source contains more than one frame or page.</summary>
    public bool HasMultipleFramesOrPages => FrameCount > 1;

    /// <summary>True when the static result intentionally retained only the selected frame or page.</summary>
    public bool FramesOrPagesDiscarded => Succeeded && !DecodedAllFrames && HasMultipleFramesOrPages;

    /// <summary>True when a multi-page source was reduced to the selected page.</summary>
    public bool PagesDiscarded => FramesOrPagesDiscarded && Container?.IsMultiPage == true;

    /// <summary>True when a static result intentionally represents only one frame of an animated source.</summary>
    public bool AnimationDiscarded => Succeeded && !DecodedAllFrames && IsAnimated;

    /// <summary>Stable human-readable reason when decoding did not complete or discarded animation.</summary>
    public string? Diagnostic { get; }

    /// <summary>Format-neutral frame/page inventory when the container could be inspected safely.</summary>
    public OfficeRasterContainerInfo? Container { get; }

    /// <summary>Selected frame or page descriptor when available.</summary>
    public OfficeRasterFrameInfo? SelectedFrame => Container != null &&
        SelectedFrameIndex >= 0 && SelectedFrameIndex < Container.Count
            ? Container.Frames[SelectedFrameIndex]
            : null;
}
