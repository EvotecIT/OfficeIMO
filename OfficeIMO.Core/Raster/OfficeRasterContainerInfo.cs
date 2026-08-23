using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Identifies how a raster container presents a decoded unit.</summary>
public enum OfficeRasterFrameKind {
    /// <summary>A single static image.</summary>
    Image,
    /// <summary>A timed animation frame.</summary>
    AnimationFrame,
    /// <summary>An independent document page.</summary>
    Page
}

/// <summary>Describes how a frame changes the canvas after its display interval.</summary>
public enum OfficeRasterFrameDisposal {
    /// <summary>The rendered canvas remains unchanged.</summary>
    None,
    /// <summary>The frame rectangle is cleared to the container background.</summary>
    Background,
    /// <summary>The frame rectangle is restored to its previous pixels.</summary>
    Previous
}

/// <summary>Describes how a frame is combined with the current canvas.</summary>
public enum OfficeRasterFrameBlend {
    /// <summary>The frame replaces its canvas rectangle.</summary>
    Source,
    /// <summary>The frame is alpha-composited over the current canvas.</summary>
    Over
}

/// <summary>Immutable timing and composition information for one image, animation frame, or page.</summary>
public sealed class OfficeRasterFrameInfo {
    internal OfficeRasterFrameInfo(
        int index,
        OfficeRasterFrameKind kind,
        int width,
        int height,
        int x,
        int y,
        TimeSpan duration,
        OfficeRasterFrameDisposal disposal,
        OfficeRasterFrameBlend blend,
        bool isDefaultImage) {
        Index = index;
        Kind = kind;
        Width = width;
        Height = height;
        X = x;
        Y = y;
        Duration = duration;
        Disposal = disposal;
        Blend = blend;
        IsDefaultImage = isDefaultImage;
    }

    /// <summary>Zero-based display or page index.</summary>
    public int Index { get; }
    /// <summary>Whether the unit is static, timed, or paged.</summary>
    public OfficeRasterFrameKind Kind { get; }
    /// <summary>Frame or page width in pixels before orientation normalization.</summary>
    public int Width { get; }
    /// <summary>Frame or page height in pixels before orientation normalization.</summary>
    public int Height { get; }
    /// <summary>Horizontal frame offset in canvas pixels.</summary>
    public int X { get; }
    /// <summary>Vertical frame offset in canvas pixels.</summary>
    public int Y { get; }
    /// <summary>Display duration. Static images and TIFF pages report <see cref="TimeSpan.Zero"/>.</summary>
    public TimeSpan Duration { get; }
    /// <summary>Canvas disposal requested after the display interval.</summary>
    public OfficeRasterFrameDisposal Disposal { get; }
    /// <summary>Canvas blend requested while rendering the unit.</summary>
    public OfficeRasterFrameBlend Blend { get; }
    /// <summary>Whether this frame is also the container's backwards-compatible static image.</summary>
    public bool IsDefaultImage { get; }
}

/// <summary>Bounded format-neutral inventory of the images, frames, or pages in a raster container.</summary>
public sealed class OfficeRasterContainerInfo {
    private readonly OfficeRasterFrameInfo[] _frames;
    private readonly IReadOnlyList<OfficeRasterFrameInfo> _readOnlyFrames;

    internal OfficeRasterContainerInfo(
        OfficeImageFormat format,
        int canvasWidth,
        int canvasHeight,
        OfficeRasterFrameInfo[] frames,
        int loopCount,
        OfficeColor background) {
        Format = format;
        CanvasWidth = canvasWidth;
        CanvasHeight = canvasHeight;
        _frames = (OfficeRasterFrameInfo[])frames.Clone();
        _readOnlyFrames = Array.AsReadOnly(_frames);
        LoopCount = loopCount;
        Background = background;
    }

    /// <summary>Detected container format.</summary>
    public OfficeImageFormat Format { get; }
    /// <summary>Logical canvas width in pixels.</summary>
    public int CanvasWidth { get; }
    /// <summary>Logical canvas height in pixels.</summary>
    public int CanvasHeight { get; }
    /// <summary>Ordered frame or page descriptors.</summary>
    public IReadOnlyList<OfficeRasterFrameInfo> Frames => _readOnlyFrames;
    /// <summary>Number of frames or pages.</summary>
    public int Count => _frames.Length;
    /// <summary>Animation loop count; zero means infinite when <see cref="IsAnimated"/> is true.</summary>
    public int LoopCount { get; }
    /// <summary>Container canvas background, when defined.</summary>
    public OfficeColor Background { get; }
    /// <summary>Whether the container has timed animation frames.</summary>
    public bool IsAnimated {
        get {
            for (int index = 0; index < _frames.Length; index++) {
                if (_frames[index].Kind == OfficeRasterFrameKind.AnimationFrame) return true;
            }
            return false;
        }
    }
    /// <summary>Whether the container has more than one independent page.</summary>
    public bool IsMultiPage => _frames.Length > 1 && _frames[0].Kind == OfficeRasterFrameKind.Page;
}
