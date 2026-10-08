using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>An ordered snapshot of raster frame references with format-neutral animation timing.</summary>
public sealed class OfficeRasterFrames : IReadOnlyList<OfficeRasterFrame> {
    private const int MaximumFrameCount = 4096;
    private readonly OfficeRasterFrame[] _frames;

    /// <summary>Creates a nonempty frame sequence, retaining the frames and their mutable image buffers.</summary>
    /// <param name="frames">One through 4096 frames in display or page order, with at most 50 million aggregate pixels.</param>
    /// <param name="playCount">Total animation plays; zero means infinite. Static images and pages use one.</param>
    public OfficeRasterFrames(IEnumerable<OfficeRasterFrame> frames, int playCount = 1) {
        if (frames == null) throw new ArgumentNullException(nameof(frames));
        if (playCount < 0) throw new ArgumentOutOfRangeException(nameof(playCount));
        var collected = new List<OfficeRasterFrame>();
        long pixels = 0;
        foreach (OfficeRasterFrame frame in frames) {
            if (frame == null) throw new ArgumentException("Frames cannot contain null entries.", nameof(frames));
            if (collected.Count == MaximumFrameCount) throw new ArgumentException("The frame sequence exceeds the supported frame count.", nameof(frames));
            pixels += (long)frame.Image.Width * frame.Image.Height;
            if (pixels > OfficeRasterGuards.MaximumPixels) throw new ArgumentException("The frame sequence exceeds the aggregate decoded-pixel limit.", nameof(frames));
            collected.Add(frame);
        }
        if (collected.Count == 0) throw new ArgumentException("At least one frame is required.", nameof(frames));
        _frames = collected.ToArray();
        PlayCount = playCount;
    }

    /// <summary>Number of frames or pages.</summary>
    public int Count => _frames.Length;
    /// <summary>Frame or page at the requested zero-based index.</summary>
    public OfficeRasterFrame this[int index] => _frames[index];
    /// <summary>Total animation plays; zero means infinite.</summary>
    public int PlayCount { get; }

    /// <summary>Maps every frame into a new sequence, preserving duration and playback count.</summary>
    /// <remarks>
    /// The complete source, planned result buffers and caller-declared retained storage must fit a combined 256 MiB retained-memory
    /// budget. Planning and validation finish before any mapper call. The mapper must return the
    /// planned dimensions and should leave its input unchanged so failures cannot partially alter
    /// the source sequence. Temporary memory allocated by arbitrary callbacks is the callback's
    /// responsibility; this budget covers retained frame buffers and collection storage.
    /// </remarks>
    /// <param name="transform">Operation returning an image for each source frame.</param>
    /// <param name="outputSize">Result dimensions, evaluated before mapping; omitted for operations that preserve dimensions.</param>
    /// <param name="cancellationToken">Observes cancellation during planning and between frames. Callbacks can also observe this token.</param>
    /// <param name="additionalRetainedBytes">Nonnegative bytes retained alongside the source and result, such as an auxiliary image used by every callback. Callers account for this storage before mapping begins.</param>
    /// <returns>A complete new sequence. Failure or cancellation does not publish a partial sequence.</returns>
    public OfficeRasterFrames Transform(Func<OfficeRasterImage, OfficeRasterImage> transform,
        Func<OfficeRasterImage, (int Width, int Height)>? outputSize = null,
        CancellationToken cancellationToken = default, long additionalRetainedBytes = 0) {
        if (transform == null) throw new ArgumentNullException(nameof(transform));
        if (additionalRetainedBytes < 0) throw new ArgumentOutOfRangeException(nameof(additionalRetainedBytes));
        cancellationToken.ThrowIfCancellationRequested();
        if (additionalRetainedBytes > OfficeRasterGuards.MaximumDecodedBytes) {
            throw new ArgumentException("The additional retained storage exceeds the aggregate retained-memory limit.", nameof(additionalRetainedBytes));
        }
        var sizes = new (int Width, int Height)[Count];
        long sourcePixels = 0, resultPixels = 0;
        for (int index = 0; index < Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRasterImage source = _frames[index].Image;
            var size = outputSize == null ? (source.Width, source.Height) : outputSize(source);
            long pixels = OfficeRasterGuards.EnsureOutputPixels(size.Item1, size.Item2, "Transformed frame dimensions exceed the managed image limit.");
            sourcePixels += (long)source.Width * source.Height;
            resultPixels += pixels;
            if (resultPixels > OfficeRasterGuards.MaximumPixels ||
                (sourcePixels + resultPixels) * 4L + Count * 192L + 65536L > OfficeRasterGuards.MaximumDecodedBytes - additionalRetainedBytes) {
                throw new ArgumentException("The transformed frame sequence exceeds the aggregate retained-memory limit.", nameof(outputSize));
            }
            sizes[index] = size;
        }
        var result = new OfficeRasterFrame[Count];
        for (int index = 0; index < Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRasterImage image = transform(_frames[index].Image);
            if (image == null || image.Width != sizes[index].Width || image.Height != sizes[index].Height) {
                throw new ArgumentException("The transform must return the planned frame dimensions.", nameof(transform));
            }
            result[index] = new OfficeRasterFrame(image, _frames[index].Duration);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new OfficeRasterFrames(result, PlayCount);
    }
    /// <summary>Enumerates frames in display or page order.</summary>
    public IEnumerator<OfficeRasterFrame> GetEnumerator() => ((IEnumerable<OfficeRasterFrame>)_frames).GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => _frames.GetEnumerator();
}
