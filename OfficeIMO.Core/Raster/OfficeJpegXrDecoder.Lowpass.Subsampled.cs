using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private sealed partial class LpContext {
        private static readonly int[] Transpose420 = { 0, 2, 1, 3 };
        private static readonly int[] Transpose422 = { 0, 2, 1, 3, 4, 6, 5, 7 };
        private static readonly int[] ChromaRemap = { 4, 1, 2, 3, 5, 6, 7 };

        // Subsampled LP codes jointly describe alternating U/V coefficients.
        private int ReadReducedChroma(Bits bits, int[] output, int offset, bool present) {
            int size = _color == 1 ? 4 : 8, count = 0;
            int[] transpose = _color == 1 ? Transpose420 : Transpose422;
            if (present) {
                count = _blocks.Read(bits, true, _color == 1 ? 10 : 2);
                int position = 0;
                for (int i = 0; i < count; i++) {
                    position += _blocks.Runs[i];
                    if (position >= 2 * (size - 1)) throw new FormatException("JPEG-XR chroma LP run exceeds its plane.");
                    int coefficient = transpose[ChromaRemap[(position >> 1) + (_color == 1 ? 1 : 0)]];
                    output[offset + ((position & 1) + 1) * 16 + coefficient] = _blocks.Levels[i];
                    position++;
                }
            }
            int refinement = _model.Bits[1];
            if (refinement != 0) for (int k = 1; k < size; k++) for (int c = 1; c <= 2; c++) {
                int index = offset + c * 16 + transpose[k], original = output[index];
                uint tail = bits.Read(refinement);
                long value = original == 0 ? tail : (Math.Abs((long)original) << refinement) + tail;
                if (value > int.MaxValue) throw new FormatException("JPEG-XR LP coefficient exceeds the managed range.");
                bool negative = original < 0 || original == 0 && value != 0 && bits.Flag();
                output[index] = negative ? -(int)value : (int)value;
            }
            return count;
        }
    }
}
