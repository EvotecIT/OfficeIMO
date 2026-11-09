#if NET8_0_OR_GREATER
using System.Buffers;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeDeflateStreamValidator {
        /// <summary>Decodes within a logical part length and publishes only stable produced prefixes.</summary>
        internal static bool TryInflateExact(byte[] bytes, int offset, int count, byte[] output,
            int outputLength, Action<int>? publish, CancellationToken cancellationToken = default) {
            if (bytes == null || output == null || ReferenceEquals(bytes, output) || offset < 0 || count < 0 || offset > bytes.Length - count
                || outputLength < 0 || outputLength > output.Length) return false;
            var reader = new DeflateBitReader(bytes, offset, count);
            int written = 0;
            int nextPublication = publish == null ? int.MaxValue : 65_536;
            int[]? literalScratch = null;
            int[]? distanceScratch = null;
            try {
                literalScratch = ArrayPool<int>.Shared.Rent(1 << 15);
                distanceScratch = ArrayPool<int>.Shared.Rent(1 << 15);
                using var workspace = new DeflateWorkspace();
                bool final;
                do {
                    workspace.Reset();
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!reader.TryReadBits(1, out int finalValue) || !reader.TryReadBits(2, out int kind)) return false;
                    final = finalValue != 0;
                    if (kind == 0) {
                        reader.AlignToByte();
                        if (!reader.TryReadBits(16, out int length) || !reader.TryReadBits(16, out int complement)
                            || (ushort)length != unchecked((ushort)~complement) || length > outputLength - written
                            || !reader.TryCopyBytes(output, written, length)) return false;
                        written += length;
                    } else {
                        HuffmanTable literals, distances;
                        if (kind == 1) { literals = FixedLiteralLength; distances = FixedDistance; }
                        else {
                            // BCL readers reject reserved dynamic distance slots even when unused.
                            if (kind != 2 || !TryReadDynamicTables(reader, out literals, out distances,
                                out int distanceSymbols, cancellationToken, workspace, buildValueLookups: false)
                                || distanceSymbols > 30) return false;
                        }
                        if (!reader.TryInflateBlockFast(literals, distances, output, outputLength, ref written, literalScratch, distanceScratch, publish, ref nextPublication, cancellationToken)) return false;
                    }
                    if (written >= nextPublication) PublishInflatedPrefix(written, publish!, ref nextPublication, cancellationToken);
                } while (!final);
                cancellationToken.ThrowIfCancellationRequested();
                return written == outputLength && reader.ConsumedBytes == count;
            } finally {
                if (literalScratch != null) ArrayPool<int>.Shared.Return(literalScratch, clearArray: true);
                if (distanceScratch != null) ArrayPool<int>.Shared.Return(distanceScratch, clearArray: true);
            }
        }

        private static void PublishInflatedPrefix(int written, Action<int> publish,
            ref int nextPublication, CancellationToken cancellationToken) {
            publish(written);
            cancellationToken.ThrowIfCancellationRequested();
            nextPublication = (int)Math.Min(int.MaxValue, (long)written + 65_536);
        }

        private static bool TryInflateCompressedBlock(DeflateBitReader reader, HuffmanTable literals,
            HuffmanTable distances, byte[] output, int outputLength, ref int written, Action<int>? publish,
            ref int nextPublication, CancellationToken cancellationToken) {
            int symbols = 0;
            while (literals.TryDecode(reader, out int symbol)) {
                if (written >= nextPublication) PublishInflatedPrefix(written, publish!, ref nextPublication, cancellationToken);
                if ((symbols++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (symbol < 256) {
                    if (written == outputLength) return false;
                    output[written++] = (byte)symbol;
                    continue;
                }
                if (symbol == 256) return true;
                if (symbol < 257 || symbol > 285) return false;
                int lengthIndex = symbol - 257;
                if (!reader.TryReadBits(LengthExtraBits[lengthIndex], out int lengthExtra)
                    || !distances.TryDecode(reader, out int distanceSymbol)
                    || (uint)distanceSymbol >= (uint)DistanceBases.Length
                    || !reader.TryReadBits(DistanceExtraBits[distanceSymbol], out int distanceExtra)) return false;
                int length = LengthBases[lengthIndex] + lengthExtra;
                int distance = DistanceBases[distanceSymbol] + distanceExtra;
                if (distance > written || length > outputLength - written) return false;
                if (distance == 1) {
                    output.AsSpan(written, length).Fill(output[written - 1]);
                } else {
                    int copied = 0;
                    while (copied < length) {
                        int chunk = Math.Min(distance + copied, length - copied);
                        output.AsSpan(written - distance, chunk).CopyTo(output.AsSpan(written + copied, chunk));
                        copied += chunk;
                    }
                }
                written += length;
            }
            return false;
        }
    }
}

#endif
