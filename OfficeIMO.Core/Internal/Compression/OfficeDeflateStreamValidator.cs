using System;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    /// <summary>Validates RFC 1951 block structure and requires one exact Deflate payload.</summary>
    internal static class OfficeDeflateStreamValidator {
        private static readonly int[] LengthBases = {
            3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
            35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258
        };

        private static readonly int[] LengthExtraBits = {
            0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2,
            3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0
        };

        private static readonly int[] DistanceBases = {
            1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
            257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193,
            12289, 16385, 24577
        };

        private static readonly int[] DistanceExtraBits = {
            0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6,
            7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13
        };

        private static readonly int[] CodeLengthOrder = {
            16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15
        };

        private static readonly HuffmanTable FixedLiteralLength = CreateFixedLiteralLengthTable();
        private static readonly HuffmanTable FixedDistance = CreateFixedDistanceTable();

        internal static bool TryValidateExact(
            byte[] bytes,
            int offset,
            int count,
            long maximumOutputBytes,
            CancellationToken cancellationToken = default) =>
            TryValidateExact(bytes, offset, count, maximumOutputBytes, out _, cancellationToken);

        internal static bool TryValidateExact(
            byte[] bytes,
            int offset,
            int count,
            long maximumOutputBytes,
            out bool outputLimitExceeded,
            CancellationToken cancellationToken = default) {
            outputLimitExceeded = false;
            if (bytes == null || offset < 0 || count < 0 || maximumOutputBytes < 0 ||
                offset > bytes.Length - count) return false;
            var reader = new DeflateBitReader(bytes, offset, count);
            long outputCount = 0;
            bool final;
            do {
                cancellationToken.ThrowIfCancellationRequested();
                if (!reader.TryReadBits(1, out int finalValue) ||
                    !reader.TryReadBits(2, out int blockType)) {
                    return false;
                }
                final = finalValue != 0;
                switch (blockType) {
                    case 0:
                        if (!TryValidateStoredBlock(reader, ref outputCount, maximumOutputBytes,
                                ref outputLimitExceeded)) return false;
                        break;
                    case 1:
                        if (!TryValidateCompressedBlock(reader, FixedLiteralLength, FixedDistance, ref outputCount,
                                maximumOutputBytes, ref outputLimitExceeded, cancellationToken)) {
                            return false;
                        }
                        break;
                    case 2:
                        if (!TryReadDynamicTables(reader, out HuffmanTable dynamicLiteralLength,
                                out HuffmanTable dynamicDistance, cancellationToken) ||
                            !TryValidateCompressedBlock(reader, dynamicLiteralLength, dynamicDistance,
                                ref outputCount, maximumOutputBytes, ref outputLimitExceeded,
                                cancellationToken)) {
                            return false;
                        }
                        break;
                    default:
                        return false;
                }
            } while (!final);

            return reader.ConsumedBytes == count;
        }

        private static bool TryValidateStoredBlock(
            DeflateBitReader reader,
            ref long outputCount,
            long maximumOutputBytes,
            ref bool outputLimitExceeded) {
            reader.AlignToByte();
            if (!reader.TryReadBits(16, out int length) ||
                !reader.TryReadBits(16, out int complement) ||
                (ushort)length != unchecked((ushort)~complement) ||
                !reader.TrySkipBytes(length)) {
                return false;
            }
            if (outputCount > maximumOutputBytes - length) {
                outputLimitExceeded = true;
                return false;
            }
            outputCount += length;
            return true;
        }

        private static HuffmanTable CreateFixedLiteralLengthTable() {
            var literalLengths = new int[288];
            for (int symbol = 0; symbol <= 143; symbol++) literalLengths[symbol] = 8;
            for (int symbol = 144; symbol <= 255; symbol++) literalLengths[symbol] = 9;
            for (int symbol = 256; symbol <= 279; symbol++) literalLengths[symbol] = 7;
            for (int symbol = 280; symbol <= 287; symbol++) literalLengths[symbol] = 8;
            if (!HuffmanTable.TryCreate(literalLengths, out HuffmanTable table)) {
                throw new InvalidOperationException("Invalid fixed Deflate literal table.");
            }
            return table;
        }

        private static HuffmanTable CreateFixedDistanceTable() {
            var distanceLengths = new int[32];
            for (int symbol = 0; symbol < distanceLengths.Length; symbol++) distanceLengths[symbol] = 5;
            if (!HuffmanTable.TryCreate(distanceLengths, out HuffmanTable table)) {
                throw new InvalidOperationException("Invalid fixed Deflate distance table.");
            }
            return table;
        }

        private static bool TryReadDynamicTables(
            DeflateBitReader reader,
            out HuffmanTable literalLength,
            out HuffmanTable distance,
            CancellationToken cancellationToken) {
            literalLength = default;
            distance = default;
            if (!reader.TryReadBits(5, out int literalCountValue) ||
                !reader.TryReadBits(5, out int distanceCountValue) ||
                !reader.TryReadBits(4, out int codeLengthCountValue)) {
                return false;
            }

            int literalCount = literalCountValue + 257;
            int distanceCount = distanceCountValue + 1;
            int codeLengthCount = codeLengthCountValue + 4;
            if (literalCount > 286 || distanceCount > 32) return false;
            var codeLengthLengths = new int[19];
            for (int index = 0; index < codeLengthCount; index++) {
                if (!reader.TryReadBits(3, out int length)) return false;
                codeLengthLengths[CodeLengthOrder[index]] = length;
            }
            if (!HuffmanTable.TryCreate(codeLengthLengths, out HuffmanTable codeLengthTable)) return false;

            var lengths = new int[literalCount + distanceCount];
            int output = 0;
            while (output < lengths.Length) {
                if ((output & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!codeLengthTable.TryDecode(reader, out int symbol)) return false;
                if (symbol <= 15) {
                    lengths[output++] = symbol;
                    continue;
                }

                int repeat;
                int value;
                if (symbol == 16) {
                    if (output == 0 || !reader.TryReadBits(2, out int extra)) return false;
                    repeat = extra + 3;
                    value = lengths[output - 1];
                } else if (symbol == 17) {
                    if (!reader.TryReadBits(3, out int extra)) return false;
                    repeat = extra + 3;
                    value = 0;
                } else if (symbol == 18) {
                    if (!reader.TryReadBits(7, out int extra)) return false;
                    repeat = extra + 11;
                    value = 0;
                } else {
                    return false;
                }
                if (repeat > lengths.Length - output) return false;
                for (int index = 0; index < repeat; index++) lengths[output++] = value;
            }

            var literalLengths = new int[literalCount];
            var distanceLengths = new int[distanceCount];
            Array.Copy(lengths, 0, literalLengths, 0, literalCount);
            Array.Copy(lengths, literalCount, distanceLengths, 0, distanceCount);
            return literalLengths[256] != 0 &&
                   HuffmanTable.TryCreate(literalLengths, out literalLength) &&
                   HuffmanTable.TryCreate(distanceLengths, out distance, allowEmpty: true);
        }

        private static bool TryValidateCompressedBlock(
            DeflateBitReader reader,
            HuffmanTable literalLength,
            HuffmanTable distance,
            ref long outputCount,
            long maximumOutputBytes,
            ref bool outputLimitExceeded,
            CancellationToken cancellationToken) {
            int decodedSymbols = 0;
            while (literalLength.TryDecode(reader, out int symbol)) {
                if ((decodedSymbols++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (symbol < 256) {
                    if (outputCount >= maximumOutputBytes) {
                        outputLimitExceeded = true;
                        return false;
                    }
                    outputCount++;
                    continue;
                }
                if (symbol == 256) return true;
                if (symbol < 257 || symbol > 285) return false;

                int lengthIndex = symbol - 257;
                if (!reader.TryReadBits(LengthExtraBits[lengthIndex], out int lengthExtra) ||
                    !distance.TryDecode(reader, out int distanceSymbol) ||
                    distanceSymbol < 0 || distanceSymbol >= DistanceBases.Length ||
                    !reader.TryReadBits(DistanceExtraBits[distanceSymbol], out int distanceExtra)) {
                    return false;
                }
                int matchLength = LengthBases[lengthIndex] + lengthExtra;
                int matchDistance = DistanceBases[distanceSymbol] + distanceExtra;
                if (matchDistance > outputCount) return false;
                if (outputCount > maximumOutputBytes - matchLength) {
                    outputLimitExceeded = true;
                    return false;
                }
                outputCount += matchLength;
            }
            return false;
        }

        private readonly struct HuffmanTable {
            private readonly int[]? _counts;
            private readonly int[]? _firstCodes;
            private readonly int[]? _firstSymbols;
            private readonly int[]? _symbols;
            private readonly int _maximumLength;
            private readonly int[]? _lookup;
            private readonly int _lookupBits;

            private HuffmanTable(
                int[] counts,
                int[] firstCodes,
                int[] firstSymbols,
                int[] symbols,
                int maximumLength,
                int[] lookup,
                int lookupBits) {
                _counts = counts;
                _firstCodes = firstCodes;
                _firstSymbols = firstSymbols;
                _symbols = symbols;
                _maximumLength = maximumLength;
                _lookup = lookup;
                _lookupBits = lookupBits;
            }

            internal static bool TryCreate(int[] lengths, out HuffmanTable table, bool allowEmpty = false) {
                table = default;
                var counts = new int[16];
                int symbolCount = 0;
                int maximumLength = 0;
                for (int symbol = 0; symbol < lengths.Length; symbol++) {
                    int length = lengths[symbol];
                    if (length < 0 || length > 15) return false;
                    if (length == 0) continue;
                    counts[length]++;
                    symbolCount++;
                    maximumLength = Math.Max(maximumLength, length);
                }
                if (symbolCount == 0) return allowEmpty;

                int remaining = 1;
                for (int length = 1; length <= 15; length++) {
                    remaining = (remaining << 1) - counts[length];
                    if (remaining < 0) return false;
                }

                var firstCodes = new int[16];
                var firstSymbols = new int[16];
                int code = 0;
                int firstSymbol = 0;
                for (int length = 1; length <= 15; length++) {
                    code = (code + counts[length - 1]) << 1;
                    firstCodes[length] = code;
                    firstSymbols[length] = firstSymbol;
                    firstSymbol += counts[length];
                }

                var symbols = new int[symbolCount];
                var nextSymbol = (int[])firstSymbols.Clone();
                for (int symbol = 0; symbol < lengths.Length; symbol++) {
                    int length = lengths[symbol];
                    if (length == 0) continue;
                    symbols[nextSymbol[length]++] = symbol;
                }
                int lookupBits = Math.Min(9, maximumLength);
                var lookup = new int[1 << lookupBits];
                for (int length = 1; length <= lookupBits; length++) {
                    for (int codeOffset = 0; codeOffset < counts[length]; codeOffset++) {
                        int canonicalCode = firstCodes[length] + codeOffset;
                        int reversedCode = 0;
                        for (int bit = 0; bit < length; bit++) {
                            reversedCode = (reversedCode << 1) | ((canonicalCode >> bit) & 1);
                        }
                        int entry = (symbols[firstSymbols[length] + codeOffset] << 4) | length;
                        for (int index = reversedCode; index < lookup.Length; index += 1 << length) {
                            lookup[index] = entry;
                        }
                    }
                }
                table = new HuffmanTable(counts, firstCodes, firstSymbols, symbols, maximumLength, lookup, lookupBits);
                return true;
            }

            internal bool TryDecode(DeflateBitReader reader, out int symbol) {
                symbol = -1;
                // Peeking may read ahead, but only the matched code's bits are
                // consumed. Exact payload accounting must ignore lookahead.
                if (_lookup != null && reader.TryPeekBits(_lookupBits, out int prefix)) {
                    int entry = _lookup[prefix];
                    if (entry != 0) {
                        reader.DropBits(entry & 15);
                        symbol = entry >> 4;
                        return true;
                    }
                }
                if (_symbols == null || _counts == null || _firstCodes == null || _firstSymbols == null) return false;
                int code = 0;
                for (int length = 1; length <= _maximumLength; length++) {
                    if (!reader.TryReadBits(1, out int bit)) return false;
                    code = (code << 1) | bit;
                    int codeOffset = code - _firstCodes[length];
                    if (codeOffset >= 0 && codeOffset < _counts[length]) {
                        symbol = _symbols[_firstSymbols[length] + codeOffset];
                        return true;
                    }
                }
                return false;
            }
        }

        private sealed class DeflateBitReader {
            private readonly byte[] _bytes;
            private readonly int _start;
            private readonly int _end;
            private int _byteOffset;
            private int _bitOffset;
            private int _readOffset;
            private uint _buffer;
            private int _bufferedBits;

            internal DeflateBitReader(byte[] bytes, int offset, int count) {
                _bytes = bytes;
                _start = offset;
                _end = offset + count;
                _byteOffset = offset;
                _readOffset = offset;
            }

            internal int ConsumedBytes => _byteOffset - _start + (_bitOffset == 0 ? 0 : 1);

            internal bool TryReadBits(int count, out int value) {
                if (!TryPeekBits(count, out value)) return false;
                DropBits(count);
                return true;
            }

            internal bool TryPeekBits(int count, out int value) {
                value = 0;
                if (count < 0 || count > 16) return false;
                while (_bufferedBits < count) {
#if NET8_0_OR_GREATER
                    // Refilling at most 15 buffered bits with 16 more fits the
                    // uint reservoir. Lookahead does not advance consumption.
                    if (_end - _readOffset >= 2) {
                        _buffer |= (uint)System.Buffers.Binary.BinaryPrimitives.ReadUInt16LittleEndian(
                            _bytes.AsSpan(_readOffset, 2)) << _bufferedBits;
                        _readOffset += 2;
                        _bufferedBits += 16;
                        continue;
                    }
#endif
                    if (_readOffset >= _end) return false;
                    _buffer |= (uint)_bytes[_readOffset++] << _bufferedBits;
                    _bufferedBits += 8;
                }
                value = (int)(_buffer & ((1U << count) - 1));
                return true;
            }

            internal void DropBits(int count) {
                _buffer >>= count;
                _bufferedBits -= count;
                _bitOffset += count;
                _byteOffset += _bitOffset >> 3;
                _bitOffset &= 7;
            }

            internal void AlignToByte() {
                if (_bitOffset == 0) return;
                DropBits(8 - _bitOffset);
            }

            internal bool TrySkipBytes(int count) {
                if (_bitOffset != 0 || count < 0 || count > _end - _byteOffset) return false;
                _byteOffset += count;
                // A compressed block can have prefetched the next stored
                // header. Restart at the logical position after the skip.
                _readOffset = _byteOffset;
                _buffer = 0;
                _bufferedBits = 0;
                return true;
            }
        }
    }
}
