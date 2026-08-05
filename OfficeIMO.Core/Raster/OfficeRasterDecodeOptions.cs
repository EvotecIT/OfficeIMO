namespace OfficeIMO.Drawing;

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

    /// <summary>Zero-based frame to decode. Managed arbitrary-frame decoding currently applies to GIF.</summary>
    public int FrameIndex {
        get => _frameIndex;
        set {
            if (value < 0) throw new System.ArgumentOutOfRangeException(nameof(FrameIndex));
            _frameIndex = value;
        }
    }

    /// <summary>Behavior when the source contains more than one frame.</summary>
    public OfficeRasterAnimationPolicy AnimationPolicy { get; set; } = OfficeRasterAnimationPolicy.UseSelectedFrame;

    internal void Validate() {
        if (AnimationPolicy != OfficeRasterAnimationPolicy.UseSelectedFrame &&
            AnimationPolicy != OfficeRasterAnimationPolicy.RejectAnimated) {
            throw new System.ArgumentOutOfRangeException(nameof(AnimationPolicy));
        }
    }
}

/// <summary>Typed evidence describing one shared raster decode decision.</summary>
public sealed class OfficeRasterDecodeInfo {
    internal OfficeRasterDecodeInfo(OfficeImageFormat format, int frameCount, int selectedFrameIndex, bool succeeded, string? diagnostic) {
        Format = format;
        FrameCount = frameCount;
        SelectedFrameIndex = selectedFrameIndex;
        Succeeded = succeeded;
        Diagnostic = diagnostic;
    }

    /// <summary>Detected source format, or <see cref="OfficeImageFormat.Unknown"/>.</summary>
    public OfficeImageFormat Format { get; }

    /// <summary>Known frame count. Static formats report one; zero means the count could not be established.</summary>
    public int FrameCount { get; }

    /// <summary>Requested zero-based frame index.</summary>
    public int SelectedFrameIndex { get; }

    /// <summary>True when the requested static frame was decoded.</summary>
    public bool Succeeded { get; }

    /// <summary>True when the source contains more than one frame.</summary>
    public bool IsAnimated => FrameCount > 1;

    /// <summary>True when a static result intentionally represents only one frame of an animated source.</summary>
    public bool AnimationDiscarded => Succeeded && IsAnimated;

    /// <summary>Stable human-readable reason when decoding did not complete or discarded animation.</summary>
    public string? Diagnostic { get; }
}
