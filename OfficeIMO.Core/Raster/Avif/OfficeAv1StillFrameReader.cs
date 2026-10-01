using System;

namespace OfficeIMO.Drawing;

/// <summary>Reads reduced Main 8/10-bit still-frame syntax and a complete combined-OBU tile group.</summary>
internal static partial class OfficeAv1StillFrameReader {
    internal static bool TryRead(byte[] bytes, OfficeAvifImageItem item, OfficeAv1StillSequence sequence,
        OfficeRasterDecodeOptions options, out OfficeAv1StillFrame? frame) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        options.CancellationToken.ThrowIfCancellationRequested();
        frame = null;
        if (bytes == null || item == null || sequence == null || bytes.Length > options.MaximumEncodedBytes) return false;
        try {
            Require(sequence.BitDepth is 8 or 10 && sequence.BitDepth == item.BitDepth);
            Require(item.Offset >= 0 && item.Length > 0 && item.Offset <= bytes.Length - item.Length);
            Require(sequence.FrameOffset >= item.Offset && sequence.FrameLength > 0 &&
                sequence.FrameOffset <= item.Offset + item.Length - sequence.FrameLength);
            var bits = new OfficeAv1Bits(bytes, sequence.FrameOffset, sequence.FrameLength, options.CancellationToken);
            var result = new OfficeAv1StillFrame {
                BitDepth = sequence.BitDepth,
                DisableCdfUpdate = bits.Flag(),
                AllowScreenContentTools = bits.Flag(),
                UpscaledWidth = sequence.MaximumWidth,
                Height = sequence.MaximumHeight
            };
            // Reduced headers imply KEY_FRAME, show_frame, error resilience, PRIMARY_REF_NONE and no size override.
            if (result.AllowScreenContentTools) bits.Flag(); // Signaled force_integer_mv; FrameIsIntra subsequently forces it to 1.
            Require(result.UpscaledWidth == item.Width && result.Height == item.Height);
            Require(OfficeRasterGuards.TryEnsurePixelCount(result.UpscaledWidth, result.Height, options.MaximumDecodedPixels, out _));
            if (sequence.SuperResolution && bits.Flag()) result.SuperResolutionDenominator = bits.Read(3) + 9;
            result.Width = (result.UpscaledWidth * 8 + result.SuperResolutionDenominator / 2) / result.SuperResolutionDenominator;
            result.MiCols = 2 * ((result.Width + 7) >> 3);
            result.MiRows = 2 * ((result.Height + 7) >> 3);
            bool differentRenderSize = bits.Flag();
            result.RenderWidth = differentRenderSize ? bits.Read(16) + 1 : result.UpscaledWidth;
            result.RenderHeight = differentRenderSize ? bits.Read(16) + 1 : result.Height;
            if (result.AllowScreenContentTools && result.Width == result.UpscaledWidth) result.AllowIntraBlockCopy = bits.Flag();
            ReadTileInfo(bits, sequence, result);
            ReadQuantization(bits, sequence, result);
            ReadSegmentation(bits, result);
            ReadDeltaParameters(bits, result);
            ReadFilters(bits, sequence, result);
            result.TransformMode = result.CodedLossless ? 0 : (bits.Flag() ? 2 : 1);
            result.ReducedTransformSet = bits.Flag();
            // Applied grain needs synthesis as well as syntax; never silently produce an ungrained image.
            Require(!sequence.FilmGrain || !bits.Flag());
            bits.AlignZero();
            result.HeaderBytes = bits.ByteOffset - sequence.FrameOffset;
            ReadTileGroup(bytes, bits, sequence, result, options);
            frame = result;
            return true;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static void Require(bool condition) {
        if (!condition) throw new FormatException("Invalid or unsupported AV1 still frame.");
    }
}
