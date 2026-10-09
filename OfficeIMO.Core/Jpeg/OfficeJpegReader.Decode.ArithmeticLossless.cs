using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // T.81 H.1.2.3: two-dimensional difference contexts share 158 probability
    // bins per conditioning destination. A ring retains the preceding sample
    // row even when an interleaved MCU visits several rows before moving right.
    private sealed class LosslessArithmeticScan {
        private readonly ArithmeticReader _reader;
        private readonly ArithmeticConditioning _conditioning;
        private readonly JpegFrame _frame;
        private readonly byte[][] _bins = new byte[16][];
        private readonly short[][] _differences;
        private readonly int[] _strides;

        public static long WorkingBytes(JpegFrame frame) {
            long bytes = 8192 + frame.ComponentCount * 96L;
            int columns = (frame.Width + frame.MaxH * 8 - 1) / (frame.MaxH * 8);
            foreach (var component in frame.Components)
                bytes = checked(bytes + columns * component.H * 8L * (component.V + 1) * 2);
            return bytes;
        }

        public LosslessArithmeticScan(OfficeByteView data, ScanHeader scan, JpegFrame frame,
            BaselineState state, ArithmeticConditioning conditioning, CancellationToken token) {
            _reader = new ArithmeticReader(data, token);
            _conditioning = conditioning; _frame = frame;
            _differences = new short[frame.ComponentCount][];
            _strides = new int[frame.ComponentCount];
            foreach (int index in scan.ComponentIndices) {
                var component = frame.Components[index];
                _bins[component.DcTable] ??= new byte[158];
                _strides[index] = state.Components[index].Stride;
                _differences[index] = new short[checked(_strides[index] * (component.V + 1))];
            }
        }

        public int Read(int componentIndex, int x, int y, int firstRow) {
            var component = _frame.Components[componentIndex];
            short[] differences = _differences[componentIndex];
            int stride = _strides[componentIndex], ring = component.V + 1;
            int at = (y % ring) * stride + x;
            int left = x == 0 ? 0 : Classify(differences[at - 1], component.DcTable);
            int above = y == firstRow ? 0 : Classify(differences[((y - 1) % ring) * stride + x], component.DcTable);
            byte[] bins = _bins[component.DcTable];
            int context = left * 20 + above * 4, difference = 0;
            if (_reader.Decode(ref bins[context]) != 0) {
                int sign = _reader.Decode(ref bins[context + 1]);
                int magnitudeContext = above <= 2 ? 100 : 129;
                int magnitude = DecodeArithmeticMagnitude(_reader, bins, context + 2 + sign,
                    magnitudeContext, magnitudeContext + 1);
                difference = sign == 0 ? magnitude : -magnitude;
            }
            differences[at] = unchecked((short)difference);
            return difference;
        }

        private int Classify(int difference, int table) {
            int magnitude = Math.Abs(difference);
            if (magnitude <= (1 << _conditioning.Lower[table]) / 2) return 0;
            return (magnitude > (1 << _conditioning.Upper[table]) ? 3 : 1) + (difference < 0 ? 1 : 0);
        }

        public void Restart() {
            _reader.Restart();
            foreach (byte[] bins in _bins) if (bins != null) Array.Clear(bins, 0, bins.Length);
        }

        public void Finish() => _reader.Finish();
    }
}
