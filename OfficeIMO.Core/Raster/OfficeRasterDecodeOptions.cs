namespace OfficeIMO.Drawing;

/// <summary>Policy used when a raster container exposes more than one frame or page.</summary>
public enum OfficeRasterFrameLossPolicy {
    /// <summary>Decode the explicitly selected unit and report the units not retained in the static result.</summary>
    UseSelectedFrame,

    /// <summary>Reject multi-frame or multi-page input instead of returning a lossy static result.</summary>
    RejectMultipleFrames
}

/// <summary>Policy used when a raster container exposes more than one frame.</summary>
public enum OfficeRasterAnimationPolicy {
    /// <summary>Decode the explicitly selected static frame and report animation loss.</summary>
    UseSelectedFrame,

    /// <summary>Reject multi-frame or animated input instead of silently discarding frames.</summary>
    RejectAnimated
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

    /// <summary>
    /// Compatibility alias for <see cref="FrameLossPolicy"/>. The policy also applies to TIFF pages.
    /// </summary>
    public OfficeRasterAnimationPolicy AnimationPolicy {
        get => _frameLossPolicy == OfficeRasterFrameLossPolicy.RejectMultipleFrames
            ? OfficeRasterAnimationPolicy.RejectAnimated
            : OfficeRasterAnimationPolicy.UseSelectedFrame;
        set {
            if (value != OfficeRasterAnimationPolicy.UseSelectedFrame &&
                value != OfficeRasterAnimationPolicy.RejectAnimated) {
                throw new System.ArgumentOutOfRangeException(nameof(AnimationPolicy));
            }
            _frameLossPolicy = value == OfficeRasterAnimationPolicy.RejectAnimated
                ? OfficeRasterFrameLossPolicy.RejectMultipleFrames
                : OfficeRasterFrameLossPolicy.UseSelectedFrame;
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

    /// <summary>Cancellation observed while reading, parsing, or decoding the request.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    internal void Validate() {
        if (_frameLossPolicy != OfficeRasterFrameLossPolicy.UseSelectedFrame &&
            _frameLossPolicy != OfficeRasterFrameLossPolicy.RejectMultipleFrames) {
            throw new System.ArgumentOutOfRangeException(nameof(FrameLossPolicy));
        }
    }
}

/// <summary>Typed evidence describing one shared raster decode decision.</summary>
public sealed class OfficeRasterDecodeInfo {
    internal OfficeRasterDecodeInfo(OfficeImageFormat format, int frameCount, int selectedFrameIndex, bool succeeded, string? diagnostic, OfficeRasterContainerInfo? container = null) {
        Format = format;
        FrameCount = frameCount;
        SelectedFrameIndex = selectedFrameIndex;
        Succeeded = succeeded;
        Diagnostic = diagnostic;
        Container = container;
    }

    /// <summary>Detected source format, or <see cref="OfficeImageFormat.Unknown"/>.</summary>
    public OfficeImageFormat Format { get; }

    /// <summary>Known frame count. Static formats report one; zero means the count could not be established.</summary>
    public int FrameCount { get; }

    /// <summary>Requested zero-based frame index.</summary>
    public int SelectedFrameIndex { get; }

    /// <summary>True when the requested static frame was decoded.</summary>
    public bool Succeeded { get; }

    /// <summary>True when the inspected source contains timed animation frames.</summary>
    public bool IsAnimated => Container?.IsAnimated ?? FrameCount > 1;

    /// <summary>True when the inspected source contains more than one frame or page.</summary>
    public bool HasMultipleFramesOrPages => FrameCount > 1;

    /// <summary>True when the static result intentionally retained only the selected frame or page.</summary>
    public bool FramesOrPagesDiscarded => Succeeded && HasMultipleFramesOrPages;

    /// <summary>True when a multi-page source was reduced to the selected page.</summary>
    public bool PagesDiscarded => FramesOrPagesDiscarded && Container?.IsMultiPage == true;

    /// <summary>True when a static result intentionally represents only one frame of an animated source.</summary>
    public bool AnimationDiscarded => Succeeded && IsAnimated && FrameCount > 1;

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
