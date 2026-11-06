using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeDeflateStreamValidator {
        private static int[] RentHeaderArray(int length, DeflateWorkspace? workspace) =>
            workspace == null ? new int[length] : workspace.RentCleared(length);

        /// <summary>Owns only integer Huffman metadata until its block has been fully decoded.</summary>
        private sealed class DeflateWorkspace : IDisposable {
            private readonly int[][] _arrays;
            private int _count;
            private int _highWater;

            internal DeflateWorkspace() {
#if NET8_0_OR_GREATER
                _arrays = ArrayPool<int[]>.Shared.Rent(24);
                Array.Clear(_arrays, 0, _arrays.Length);
#else
                _arrays = new int[24][];
#endif
            }

            internal int[] RentCleared(int length) {
                if (_count == 24) throw new InvalidOperationException("The Huffman workspace is full.");
                int[]? buffer = _count < _highWater ? _arrays[_count] : null;
                if (buffer == null || buffer.Length < length) {
#if NET8_0_OR_GREATER
                    if (buffer != null) ArrayPool<int>.Shared.Return(buffer, clearArray: true);
                    buffer = ArrayPool<int>.Shared.Rent(Math.Max(1, length));
#else
                    buffer = new int[length];
#endif
                    _arrays[_count] = buffer;
                }
                Array.Clear(buffer, 0, length);
                _count++;
                _highWater = Math.Max(_highWater, _count);
                return buffer;
            }

            internal void Reset() => _count = 0;

            public void Dispose() {
#if NET8_0_OR_GREATER
                for (int index = 0; index < _highWater; index++)
                    ArrayPool<int>.Shared.Return(_arrays[index], clearArray: true);
                ArrayPool<int[]>.Shared.Return(_arrays, clearArray: true);
#endif
            }
        }
    }
}
