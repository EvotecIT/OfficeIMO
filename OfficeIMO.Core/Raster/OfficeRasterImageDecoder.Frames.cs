using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterImageDecoder {
    /// <summary>Attempts to decode every image, rendered animation frame, or page into independent managed buffers.</summary>
    /// <remarks>
    /// This eager operation succeeds only when every frame is supported and the complete sequence
    /// fits the frame-count, aggregate decoded-pixel, and retained working-memory limits. Cancellation
    /// propagates. FrameIndex must be zero; FrameLossPolicy does not reject animation because all frames
    /// are retained. GIF/APNG frames contain the rendered canvas after blending and before disposal.
    /// Source timing is retained without substituting a playback delay for zero-duration frames.
    /// </remarks>
    /// <param name="bytes">Encoded raster container.</param>
    /// <param name="options">Decode policy. MaximumDecodedPixels limits the entire decoded sequence.</param>
    /// <param name="frames">Complete decoded sequence on success, otherwise null.</param>
    /// <param name="maximumFrames">Maximum frame or page count, from one to 4096; the default is 256.</param>
    public static bool TryDecodeFrames(byte[]? bytes, OfficeRasterDecodeOptions? options,
        out OfficeRasterFrames? frames, int maximumFrames = 256) {
        frames = null;
        if (maximumFrames < 1 || maximumFrames > 4096) throw new ArgumentOutOfRangeException(nameof(maximumFrames));
        var effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (effective.FrameIndex != 0) throw new ArgumentException("All-frame decoding requires FrameIndex zero.", nameof(options));
        var selected = effective.WithAdditionalRetainedManagedBytes(0L);
        selected.FrameLossPolicy = OfficeRasterFrameLossPolicy.UseSelectedFrame;
        if (!TryDecode(bytes, selected, out OfficeRasterImage? first, out OfficeRasterDecodeInfo info) || first == null) return false;
        int count = info.FrameCount;
        if (count < 1 || count > maximumFrames) return false;
        var decoded = new OfficeRasterFrame[count];
        long pixels = (long)first.Width * first.Height;
        if (pixels > effective.MaximumDecodedPixels) return false;
        decoded[0] = new OfficeRasterFrame(first, GetDuration(info, 0));
        long retainedBytes = pixels * 4L + count * 64L;
        for (int index = 1; index < count; index++) {
            effective.CancellationToken.ThrowIfCancellationRequested();
            selected = effective.WithAdditionalRetainedManagedBytes(retainedBytes);
            selected.FrameIndex = index;
            selected.FrameLossPolicy = OfficeRasterFrameLossPolicy.UseSelectedFrame;
            if (!TryDecode(bytes, selected, out OfficeRasterImage? image, out OfficeRasterDecodeInfo next) ||
                image == null || next.FrameCount != count) return false;
            pixels += (long)image.Width * image.Height;
            if (pixels > effective.MaximumDecodedPixels) return false;
            decoded[index] = new OfficeRasterFrame(image, GetDuration(next, index));
            retainedBytes = pixels * 4L + count * 64L;
        }
        effective.CancellationToken.ThrowIfCancellationRequested();
        frames = new OfficeRasterFrames(decoded, info.Container?.PlayCount ?? 1);
        return true;
    }

    private static TimeSpan GetDuration(OfficeRasterDecodeInfo info, int index) =>
        info.Container != null && index < info.Container.Count ? info.Container.Frames[index].Duration : TimeSpan.Zero;
}
