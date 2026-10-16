using System.Threading;

namespace OfficeIMO.Pdf.Filters;

internal static partial class Jbig2Decoder {
    private sealed class Bitmap {
        private readonly byte[] _pixels;
        private readonly Action<long> _charge;
        private readonly CancellationToken _token;

        internal Bitmap(int width, int height, Action<long> charge, CancellationToken token) {
            _charge = charge; _token = token;
            if (width < 0 || height < 0 || (long)width * height > MaximumBitmapPixels) throw new InvalidDataException("JBIG2 bitmap exceeds the pixel limit.");
            Width = width;
            Height = height;
            _pixels = new byte[width * height];
        }

        internal int Width { get; }
        internal int Height { get; }

        internal int Get(int x, int y) => (uint)x < (uint)Width && (uint)y < (uint)Height ? _pixels[y * Width + x] : 0;

        internal void Set(int x, int y, int value) => _pixels[y * Width + x] = (byte)value;

        internal void Fill(int value) {
            _charge(_pixels.LongLength);
            for (int index = 0; index < _pixels.Length; index++) {
                if ((index & 1023) == 0) _token.ThrowIfCancellationRequested();
                _pixels[index] = (byte)value;
            }
        }

        internal void CopyRow(int from, int to) {
            _token.ThrowIfCancellationRequested();
            _charge(Width);
            Buffer.BlockCopy(_pixels, from * Width, _pixels, to * Width, Width);
        }

        // Combination operators (7.4.3.1.1): 0 OR, 1 AND, 2 XOR, 3 XNOR, 4 REPLACE.
        internal void Compose(Bitmap source, int left, int top, int operation) {
            _charge((long)source.Width * source.Height);
            for (int y = 0; y < source.Height; y++) {
                _token.ThrowIfCancellationRequested();
                int targetY = top + y;
                if ((uint)targetY >= (uint)Height) continue;
                for (int x = 0; x < source.Width; x++) {
                    if ((x & 1023) == 0) _token.ThrowIfCancellationRequested();
                    int targetX = left + x;
                    if ((uint)targetX >= (uint)Width) continue;
                    int s = source._pixels[y * source.Width + x];
                    int index = targetY * Width + targetX;
                    int d = _pixels[index];
                    _pixels[index] = (byte)(operation switch { 0 => d | s, 1 => d & s, 2 => d ^ s, 3 => 1 - (d ^ s), _ => s });
                }
            }
        }
    }

}
