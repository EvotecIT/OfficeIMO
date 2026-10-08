using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterFrames {
    /// <summary>Resizes every frame through shared fit planning, retaining timing and playback count.</summary>
    /// <remarks>All source, result and largest operation-temporary storage is validated before mapping starts.
    /// A planning failure or cancellation leaves this sequence's pixels unchanged.</remarks>
    public OfficeRasterFrames Resize(OfficeRasterResizeOptions options, CancellationToken cancellationToken = default,
        long additionalRetainedBytes = 0) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        var captured = new OfficeRasterResizeOptions {
            Width = options.Width, Height = options.Height, Fit = options.Fit,
            ResamplingMode = options.ResamplingMode, ColorSpace = options.ColorSpace
        };
        return ResizeFrames(source => OfficeRasterResampler.PlanResize(source.Width, source.Height, captured, cancellationToken),
            cancellationToken, additionalRetainedBytes);
    }

    /// <summary>Resizes each frame by a positive percentage using Bicubic sampling, preserving timing and playback count.</summary>
    /// <remarks>Each dimension is multiplied by the percentage, rounded down and bounded to at least one pixel.
    /// Frames may have different source dimensions. All dimensions and combined retained and temporary storage are validated before mapping.</remarks>
    public OfficeRasterFrames Resize(int percentage, CancellationToken cancellationToken = default,
        long additionalRetainedBytes = 0) {
        if (percentage <= 0) throw new ArgumentOutOfRangeException(nameof(percentage));
        return ResizeFrames(source => {
            long width = Math.Max(1L, (long)source.Width * percentage / 100L);
            long height = Math.Max(1L, (long)source.Height * percentage / 100L);
            if (width > int.MaxValue || height > int.MaxValue) {
                throw new ArgumentOutOfRangeException(nameof(percentage), "Scaled dimensions are not representable.");
            }
            return OfficeRasterResampler.PlanResize(source.Width, source.Height, new OfficeRasterResizeOptions {
                Width = (int)width, Height = (int)height, Fit = OfficeImageFit.Stretch
            }, cancellationToken);
        }, cancellationToken, additionalRetainedBytes);
    }

    private OfficeRasterFrames ResizeFrames(Func<OfficeRasterImage, OfficeRasterResizePlan> createPlan,
        CancellationToken cancellationToken, long additionalRetainedBytes) {
        if (additionalRetainedBytes < 0) throw new ArgumentOutOfRangeException(nameof(additionalRetainedBytes));
        if (additionalRetainedBytes > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Additional retained storage exceeds the managed image limit.", nameof(additionalRetainedBytes));
        cancellationToken.ThrowIfCancellationRequested();
        var plans = new OfficeRasterResizePlan[Count];
        long temporaryBytes = 0;
        for (int i = 0; i < Count; i++) {
            OfficeRasterImage source = this[i].Image;
            plans[i] = createPlan(source);
            temporaryBytes = Math.Max(temporaryBytes, plans[i].AdditionalWorkingBytes);
        }
        int planningIndex = 0, mappingIndex = 0;
        return Transform(source => OfficeRasterResampler.Resize(source, plans[mappingIndex++], cancellationToken),
            _ => { OfficeRasterResizePlan plan = plans[planningIndex++]; return (plan.Width, plan.Height); },
            cancellationToken, checked(additionalRetainedBytes + temporaryBytes + Count * 128L));
    }
}
