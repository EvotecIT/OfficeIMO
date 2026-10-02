using System;

namespace OfficeIMO.Drawing;

/// <summary>One bounded still-item decode, preserving color/alpha ownership through final RGBA publication.</summary>
internal static class OfficeAvifCodec {
    internal static bool TryDecode(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image) =>
        TryDecode(bytes, options, out image, out _);

    internal static bool TryDecode(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image,
        out bool callerCodecEligible) {
        image = null;
        callerCodecEligible = false;
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        if (options.FrameIndex != 0 || !OfficeAvifContainerReader.TryRead(bytes, options, out var container)) return false;
        try {
            var colorItem = container!.Color;
            var retained = options.WithAdditionalRetainedManagedBytes(bytes.LongLength);
            var color = DecodeItem(bytes, colorItem, options, out var sequence);
            OfficeAv1ReconstructedFrame? alpha = null;
            if (container.Alpha != null)
                alpha = DecodeItem(bytes, container.Alpha, options.WithAdditionalRetainedManagedBytes(color.StorageBytes), out _);
            try {
                image = OfficeAvifColorConverter.Compose(color, colorItem.ColorDescription ?? sequence.Color, alpha, retained, sequence.ChromaSamplePosition);
            } catch (NotSupportedException) {
                // Both selected items are validated and all composition bounds passed.
                // Only an unsupported color conversion may reach the trusted codec.
                callerCodecEligible = true;
                return false;
            }
            return true;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static OfficeAv1ReconstructedFrame DecodeItem(byte[] bytes, OfficeAvifImageItem item,
        OfficeRasterDecodeOptions options, out OfficeAv1StillSequence sequence) {
        if (!OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out var parsed) ||
            !OfficeAv1StillFrameReader.TryRead(bytes, item, parsed!, options, out var frame))
            throw new FormatException("Unsupported or invalid AVIF still payload.");
        sequence = parsed!;
        return OfficeAv1FrameReconstructor.Decode(bytes, sequence, frame!, options, OfficeAv1ReconstructionStage.Restored);
    }
}
