using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Normative Main-8 horizontal superresolution with immutable input and bounded owned output.</summary>
/// <remarks>AV1 section 7.16. Samples clamp to the reconstructed MI extent, including its coded-edge padding.</remarks>
internal sealed class OfficeAv1Upscaler {
    private static readonly int[,] Filters = {
        { 0, 0, 0, 128, 0, 0, 0, 0 },
        { 0, 0, -1, 128, 2, -1, 0, 0 },
        { 0, 1, -3, 127, 4, -2, 1, 0 },
        { 0, 1, -4, 127, 6, -3, 1, 0 },
        { 0, 2, -6, 126, 8, -3, 1, 0 },
        { 0, 2, -7, 125, 11, -4, 1, 0 },
        { -1, 2, -8, 125, 13, -5, 2, 0 },
        { -1, 3, -9, 124, 15, -6, 2, 0 },
        { -1, 3, -10, 123, 18, -6, 2, -1 },
        { -1, 3, -11, 122, 20, -7, 3, -1 },
        { -1, 4, -12, 121, 22, -8, 3, -1 },
        { -1, 4, -13, 120, 25, -9, 3, -1 },
        { -1, 4, -14, 118, 28, -9, 3, -1 },
        { -1, 4, -15, 117, 30, -10, 4, -1 },
        { -1, 5, -16, 116, 32, -11, 4, -1 },
        { -1, 5, -16, 114, 35, -12, 4, -1 },
        { -1, 5, -17, 112, 38, -12, 4, -1 },
        { -1, 5, -18, 111, 40, -13, 5, -1 },
        { -1, 5, -18, 109, 43, -14, 5, -1 },
        { -1, 6, -19, 107, 45, -14, 5, -1 },
        { -1, 6, -19, 105, 48, -15, 5, -1 },
        { -1, 6, -19, 103, 51, -16, 5, -1 },
        { -1, 6, -20, 101, 53, -16, 6, -1 },
        { -1, 6, -20, 99, 56, -17, 6, -1 },
        { -1, 6, -20, 97, 58, -17, 6, -1 },
        { -1, 6, -20, 95, 61, -18, 6, -1 },
        { -2, 7, -20, 93, 64, -18, 6, -2 },
        { -2, 7, -20, 91, 66, -19, 6, -1 },
        { -2, 7, -20, 88, 69, -19, 6, -1 },
        { -2, 7, -20, 86, 71, -19, 6, -1 },
        { -2, 7, -20, 84, 74, -20, 7, -2 },
        { -2, 7, -20, 81, 76, -20, 7, -1 },
        { -2, 7, -20, 79, 79, -20, 7, -2 },
        { -1, 7, -20, 76, 81, -20, 7, -2 },
        { -2, 7, -20, 74, 84, -20, 7, -2 },
        { -1, 6, -19, 71, 86, -20, 7, -2 },
        { -1, 6, -19, 69, 88, -20, 7, -2 },
        { -1, 6, -19, 66, 91, -20, 7, -2 },
        { -2, 6, -18, 64, 93, -20, 7, -2 },
        { -1, 6, -18, 61, 95, -20, 6, -1 },
        { -1, 6, -17, 58, 97, -20, 6, -1 },
        { -1, 6, -17, 56, 99, -20, 6, -1 },
        { -1, 6, -16, 53, 101, -20, 6, -1 },
        { -1, 5, -16, 51, 103, -19, 6, -1 },
        { -1, 5, -15, 48, 105, -19, 6, -1 },
        { -1, 5, -14, 45, 107, -19, 6, -1 },
        { -1, 5, -14, 43, 109, -18, 5, -1 },
        { -1, 5, -13, 40, 111, -18, 5, -1 },
        { -1, 4, -12, 38, 112, -17, 5, -1 },
        { -1, 4, -12, 35, 114, -16, 5, -1 },
        { -1, 4, -11, 32, 116, -16, 5, -1 },
        { -1, 4, -10, 30, 117, -15, 4, -1 },
        { -1, 3, -9, 28, 118, -14, 4, -1 },
        { -1, 3, -9, 25, 120, -13, 4, -1 },
        { -1, 3, -8, 22, 121, -12, 4, -1 },
        { -1, 3, -7, 20, 122, -11, 3, -1 },
        { -1, 2, -6, 18, 123, -10, 3, -1 },
        { 0, 2, -6, 15, 124, -9, 3, -1 },
        { 0, 2, -5, 13, 125, -8, 2, -1 },
        { 0, 1, -4, 11, 125, -7, 2, 0 },
        { 0, 1, -3, 8, 126, -6, 2, 0 },
        { 0, 1, -3, 6, 127, -4, 1, 0 },
        { 0, 1, -2, 4, 127, -3, 1, 0 },
        { 0, 0, -1, 2, 128, -1, 0, 0 }
    };
    private readonly int _width, _upscaledWidth, _height, _miWidth, _planes;
    private readonly CancellationToken _cancellation;
    private readonly long _workLimit;
    private long _work;
    internal int Stride { get; }
    internal int Rows { get; }

    internal static long ContextBytes(OfficeAv1StillFrame frame, bool monochrome, int sb) {
        if (frame.UpscaledWidth < 1 || frame.UpscaledWidth > 65536 || frame.Height < 1 || frame.Height > 65536 ||
            (sb != 64 && sb != 128)) throw new FormatException("Invalid AV1 upscaled storage dimensions.");
        return (long)((frame.UpscaledWidth + sb - 1) / sb * sb) * ((frame.Height + sb - 1) / sb * sb) *
            (monochrome ? 2 : 3) / 2 + 4096;
    }
    internal OfficeAv1Upscaler(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence, int sb,
        OfficeRasterDecodeOptions options) {
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        if (!sequence.SuperResolution || frame.Width < 1 || frame.Width >= frame.UpscaledWidth ||
            frame.SuperResolutionDenominator < 9 || frame.SuperResolutionDenominator > 16 ||
            frame.Width != ((long)frame.UpscaledWidth * 8 + frame.SuperResolutionDenominator / 2) / frame.SuperResolutionDenominator ||
            frame.MiCols != 2 * ((frame.Width + 7) / 8) || frame.MiRows != 2 * ((frame.Height + 7) / 8) ||
            (long)frame.UpscaledWidth * frame.Height > options.MaximumDecodedPixels ||
            options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes - ContextBytes(frame, sequence.Monochrome, sb))
            throw new FormatException("Invalid or unbounded AV1 superresolution dimensions.");
        _width = frame.Width; _upscaledWidth = frame.UpscaledWidth; _height = frame.Height;
        _miWidth = frame.MiCols * 4; _planes = sequence.Monochrome ? 1 : 3;
        Stride = (_upscaledWidth + sb - 1) / sb * sb; Rows = (_height + sb - 1) / sb * sb;
        _cancellation = options.CancellationToken; _workLimit = options.MaximumInspectionWorkPixels;
    }

    /// <summary>Upscales a complete immutable coded frame. Each invocation is charged to the owner's work bound.</summary>
    internal byte[][] Apply(byte[][] input, int stride) {
        _cancellation.ThrowIfCancellationRequested();
        if (input.Length != _planes || stride < _miWidth || (stride & 1) != 0)
            throw new FormatException("Invalid AV1 superresolution source planes.");
        long pixels = (long)_upscaledWidth * _height;
        if (_planes == 3) pixels += 2L * ((_upscaledWidth + 1) / 2) * ((_height + 1) / 2);
        if (_work > _workLimit - pixels) throw new FormatException("AV1 superresolution exceeds filter work.");
        _work += pixels;
        for (int p = 0; p < _planes; p++) {
            int sub = p == 0 ? 0 : 1, height = (_height + (1 << sub) - 1) >> sub;
            if (input[p] == null || (long)(stride >> sub) * height > input[p].Length)
                throw new FormatException("Incomplete AV1 superresolution source plane.");
        }
        var output = new byte[_planes][];
        for (int p = 0; p < _planes; p++) {
            _cancellation.ThrowIfCancellationRequested();
            int sub = p == 0 ? 0 : 1, width = (_width + (1 << sub) - 1) >> sub;
            int upscaled = (_upscaledWidth + (1 << sub) - 1) >> sub, height = (_height + (1 << sub) - 1) >> sub;
            int pitch = stride >> sub, outPitch = Stride >> sub, maximumX = (_miWidth >> sub) - 1;
            output[p] = new byte[checked(outPitch * (Rows >> sub))];
            int step = (int)(((long)width * 16384 + upscaled / 2) / upscaled);
            int error = upscaled * step - width * 16384;
            int initial = (int)((-((long)(upscaled - width) * 8192) + upscaled / 2) / upscaled) + 128 - error / 2;
            initial &= 16383;
            for (int y = 0; y < height; y++) {
                _cancellation.ThrowIfCancellationRequested();
                for (int x = 0; x < upscaled; x++) {
                    int position = -16384 + initial + x * step, sample = position >> 14, phase = (position & 16383) >> 8;
                    int sum = 0;
                    for (int k = 0; k < 8; k++) sum += input[p][y * pitch + Math.Max(0, Math.Min(maximumX, sample + k - 3))] * Filters[phase, k];
                    output[p][y * outPitch + x] = (byte)Math.Max(0, Math.Min(255, (sum + 64) >> 7));
                }
            }
        }
        _cancellation.ThrowIfCancellationRequested(); return output;
    }
}
