using System;

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

        internal static bool TryValidateExact(byte[] bytes, int offset, int count) {
            if (bytes == null || offset < 0 || count < 0 || offset > bytes.Length - count) return false;
            var reader = new DeflateBitReader(bytes, offset, count);
            long outputCount = 0;
            bool final;
            do {
                if (!reader.TryReadBits(1, out int finalValue) ||
                    !reader.TryReadBits(2, out int blockType)) {
                    return false;
                }
                final = finalValue != 0;
                switch (blockType) {
                    case 0:
                        if (!TryValidateStoredBlock(reader, ref outputCount)) return false;
                        break;
                    case 1:
                        if (!TryCreateFixedTables(out HuffmanTable literalLength, out HuffmanTable distance) ||
                            !TryValidateCompressedBlock(reader, literalLength, distance, ref outputCount)) {
                            return false;
                        }
                        break;
                    case 2:
                        if (!TryReadDynamicTables(reader, out HuffmanTable dynamicLiteralLength, out HuffmanTable dynamicDistance) ||
                            !TryValidateCompressedBlock(reader, dynamicLiteralLength, dynamicDistance, ref outputCount)) {
                            return false;
                        }
                        break;
                    default:
                        return false;
                }
            } while (!final);

            return reader.ConsumedBytes == count;
        }

        private static bool TryValidateStoredBlock(DeflateBitReader reader, ref long outputCount) {
            reader.AlignToByte();
            if (!reader.TryReadBits(16, out int length) ||
                !reader.TryReadBits(16, out int complement) ||
                (ushort)length != unchecked((ushort)~complement) ||
                !reader.TrySkipBytes(length)) {
                return false;
            }
            outputCount += length;
            return outputCount >= 0;
        }

        private static bool TryCreateFixedTables(out HuffmanTable literalLength, out HuffmanTable distance) {
            literalLength = default;
            distance = default;
            var literalLengths = new int[288];
            for (int symbol = 0; symbol <= 143; symbol++) literalLengths[symbol] = 8;
            for (int symbol = 144; symbol <= 255; symbol++) literalLengths[symbol] = 9;
            for (int symbol = 256; symbol <= 279; symbol++) literalLengths[symbol] = 7;
            for (int symbol = 280; symbol <= 287; symbol++) literalLengths[symbol] = 8;
            var distanceLengths = new int[32];
            for (int symbol = 0; symbol < distanceLengths.Length; symbol++) distanceLengths[symbol] = 5;
            return HuffmanTable.TryCreate(literalLengths, out literalLength) &&
                   HuffmanTable.TryCreate(distanceLengths, out distance);
        }

        private static bool TryReadDynamicTables(
            DeflateBitReader reader,
            out HuffmanTable literalLength,
            out HuffmanTable distance) {
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
            ref long outputCount) {
            while (literalLength.TryDecode(reader, out int symbol)) {
                if (symbol < 256) {
                    outputCount++;
                    if (outputCount < 0) return false;
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
                if (matchDistance > outputCount || outputCount > long.MaxValue - matchLength) return false;
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

            private HuffmanTable(
                int[] counts,
                int[] firstCodes,
                int[] firstSymbols,
                int[] symbols,
                int maximumLength) {
                _counts = counts;
                _firstCodes = firstCodes;
                _firstSymbols = firstSymbols;
                _symbols = symbols;
                _maximumLength = maximumLength;
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
                table = new HuffmanTable(counts, firstCodes, firstSymbols, symbols, maximumLength);
                return true;
            }

            internal bool TryDecode(DeflateBitReader reader, out int symbol) {
                symbol = -1;
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

            internal DeflateBitReader(byte[] bytes, int offset, int count) {
                _bytes = bytes;
                _start = offset;
                _end = offset + count;
                _byteOffset = offset;
            }

            internal int ConsumedBytes => _byteOffset - _start + (_bitOffset == 0 ? 0 : 1);

            internal bool TryReadBits(int count, out int value) {
                value = 0;
                if (count < 0 || count > 16) return false;
                for (int bit = 0; bit < count; bit++) {
                    if (_byteOffset >= _end) return false;
                    value |= ((_bytes[_byteOffset] >> _bitOffset) & 1) << bit;
                    _bitOffset++;
                    if (_bitOffset == 8) {
                        _bitOffset = 0;
                        _byteOffset++;
                    }
                }
                return true;
            }

            internal void AlignToByte() {
                if (_bitOffset == 0) return;
                _bitOffset = 0;
                _byteOffset++;
            }

            internal bool TrySkipBytes(int count) {
                if (_bitOffset != 0 || count < 0 || count > _end - _byteOffset) return false;
                _byteOffset += count;
                return true;
            }
        }
    }
}
