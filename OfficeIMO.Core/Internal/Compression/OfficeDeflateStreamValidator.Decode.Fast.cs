#if NET8_0_OR_GREATER
using System;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeDeflateStreamValidator {
        private readonly partial struct HuffmanTable {
            /// <summary>Builds a bounded two-level lookup from the already validated canonical codes.</summary>
            [MethodImpl(MethodImplOptions.NoInlining)]
            internal int[] GetFastLookup(int[] scratch, bool distanceTree, out int rootBits, out int maximumBits) {
                maximumBits = _maximumLength;
                rootBits = Math.Min(distanceTree ? 6 : 9, maximumBits);
                if (_counts == null || _firstCodes == null || _firstSymbols == null || _symbols == null)
                    return Array.Empty<int>();
                int rootSize = 1 << rootBits;
                int rootMask = rootSize - 1;
                Array.Clear(scratch, 0, rootSize);
                for (int length = 1; length <= maximumBits; length++) {
                    for (int offset = 0; offset < _counts[length]; offset++) {
                        int reversed = ReverseCode(_firstCodes[length] + offset, length);
                        if (length <= rootBits) {
                            int entry = CreateInflateEntry(_symbols[_firstSymbols[length] + offset], length, distanceTree);
                            for (int index = reversed; index < rootSize; index += 1 << length) scratch[index] = entry;
                        } else {
                            int prefix = reversed & rootMask;
                            scratch[prefix] = Math.Min(scratch[prefix], -(length - rootBits));
                        }
                    }
                }
                int next = rootSize;
                for (int prefix = 0; prefix < rootSize; prefix++) {
                    if (scratch[prefix] >= 0) continue;
                    int extraBits = -scratch[prefix];
                    int size = 1 << extraBits;
                    Array.Clear(scratch, next, size);
                    scratch[prefix] = int.MinValue | (next << 4) | extraBits;
                    next += size;
                }
                for (int length = rootBits + 1; length <= maximumBits; length++) {
                    for (int offset = 0; offset < _counts[length]; offset++) {
                        int reversed = ReverseCode(_firstCodes[length] + offset, length);
                        int pointer = scratch[reversed & rootMask];
                        int start = (pointer & int.MaxValue) >> 4;
                        int size = 1 << (pointer & 15);
                        int entry = CreateInflateEntry(_symbols[_firstSymbols[length] + offset], length, distanceTree);
                        for (int index = reversed >> rootBits; index < size; index += 1 << (length - rootBits))
                            scratch[start + index] = entry;
                    }
                }
                return scratch;
            }

            private static int CreateInflateEntry(int symbol, int bits, bool distanceTree) {
                if (distanceTree) {
                    if ((uint)symbol >= (uint)DistanceBases.Length) return (3 << 24) | bits;
                    return (DistanceBases[symbol] << 8) | (DistanceExtraBits[symbol] << 4) | bits;
                }
                if (symbol < 256) return (symbol << 8) | bits;
                if (symbol == 256) return (1 << 24) | bits;
                if (symbol > 285) return (3 << 24) | bits;
                int index = symbol - 257;
                return (2 << 24) | (LengthBases[index] << 8) | (LengthExtraBits[index] << 4) | bits;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static int ReverseCode(int code, int length) {
                uint value = (uint)code;
                value = ((value & 0x5555U) << 1) | ((value >> 1) & 0x5555U);
                value = ((value & 0x3333U) << 2) | ((value >> 2) & 0x3333U);
                value = ((value & 0x0F0FU) << 4) | ((value >> 4) & 0x0F0FU);
                value = ((value & 0x00FFU) << 8) | (value >> 8);
                return (int)(value >> (16 - length));
            }

        }

        private sealed partial class DeflateBitReader {
            /// <summary>Keeps the memory reservoir local while complete source and destination windows are available.</summary>
            internal bool TryInflateBlockFast(HuffmanTable literals, HuffmanTable distances,
                byte[] output, int outputLength, ref int written, int[] literalScratch, int[] distanceScratch,
                Action<int>? publish, ref int nextPublication, CancellationToken cancellationToken) {
                if (_stream != null) return false;
                int outputPosition = written;
                int outputLimit = outputLength - 273;
                int[] literalTable = literals.GetFastLookup(literalScratch, distanceTree: false, out int literalBits, out int literalMaximumBits);
                int[] distanceTable = distances.GetFastLookup(distanceScratch, distanceTree: true, out int distanceBits, out int distanceMaximumBits);
                if (literalTable.Length == 0) return false;
                int literalMask = (1 << literalBits) - 1;
                int distanceMask = (1 << distanceBits) - 1;
                ulong reservoir = _buffer;
                int available = _bufferedBits;
                int position = checked((int)_readOffset);
                int end = checked((int)_end);
                int symbols = 0;
                ref byte input = ref MemoryMarshal.GetArrayDataReference(_bytes);
                ref byte destination = ref MemoryMarshal.GetArrayDataReference(output);

                // Eight input bytes bound each refill; 56..63 counted bits retain partial-byte lookahead.
                // A match emits at most 258 bytes; fifteen extra bytes bound the widest copy.
                while (position <= end - 8 && outputPosition <= outputLimit) {
                    if (outputPosition >= nextPublication) PublishInflatedPrefix(outputPosition, publish!, ref nextPublication, cancellationToken);
                    if ((symbols++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (available < literalMaximumBits) Refill64(ref input, ref position, ref reservoir, ref available);
                    int entry = LookupEntry(literalTable, reservoir, literalBits, literalMask);
                    if (entry == 0) return false;
                    int bits = entry & 15;
                    reservoir >>= bits;
                    available -= bits;
                    int action = entry & (3 << 24);
                    if (action == 0) {
                        Unsafe.Add(ref destination, outputPosition++) = (byte)(entry >> 8);
                        continue;
                    }
                    if (action == 1 << 24) {
                        written = outputPosition;
                        SaveInflateState(position, reservoir, available);
                        return true;
                    }
                    if (action != 2 << 24 || distanceTable.Length == 0) return false;
                    // After the length code, at most 5 + 15 + 13 bits remain in a match.
                    if (available < 33) Refill64(ref input, ref position, ref reservoir, ref available);
                    int extraBits = (entry >> 4) & 15;

                    int length = ((entry >> 8) & 0xFFFF) + (int)(reservoir & ((1UL << extraBits) - 1));
                    reservoir >>= extraBits;
                    available -= extraBits;

                    entry = LookupEntry(distanceTable, reservoir, distanceBits, distanceMask);
                    if (entry == 0 || (entry & (3 << 24)) != 0) return false;
                    bits = entry & 15;
                    reservoir >>= bits;
                    available -= bits;
                    extraBits = (entry >> 4) & 15;

                    int distance = ((entry >> 8) & 0xFFFF) + (int)(reservoir & ((1UL << extraBits) - 1));
                    reservoir >>= extraBits;
                    available -= extraBits;
                    if (distance > outputPosition) return false;
                    CopyInflateMatch(ref destination, outputPosition, length, distance);
                    outputPosition += length;
                }
                written = outputPosition;
                SaveInflateState(position, reservoir, available);
                return TryInflateCompressedBlock(this, literals, distances, output, outputLength, ref written, publish, ref nextPublication, cancellationToken);
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static int LookupEntry(int[] table, ulong reservoir, int rootBits, int mask) {
                // Construction initialized every root slot and every subtable slot,
                // with ordinary checked writes into this exclusively owned array.
                // The mask and pointer both come from that completed construction.
                ref int first = ref MemoryMarshal.GetArrayDataReference(table);
                int entry = Unsafe.Add(ref first, (int)reservoir & mask);
                if (entry < 0) {
                    int offset = (entry & int.MaxValue) >> 4;
                    int subMask = (1 << (entry & 15)) - 1;
                    entry = Unsafe.Add(ref first, offset + ((int)(reservoir >> rootBits) & subMask));
                }
                return entry;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static void Refill64(ref byte source, ref int position, ref ulong reservoir, ref int available) {
                ulong value = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref source, position));
                if (!BitConverter.IsLittleEndian) value = BinaryPrimitives.ReverseEndianness(value);
                reservoir |= value << available;
                position += 7 - (available >> 3);
                available |= 56;
            }

            private void SaveInflateState(int position, ulong reservoir, int available) {
                _readOffset = position;
                _buffer = reservoir;
                _bufferedBits = available;
                long consumedBits = ((long)position << 3) - available;
                _byteOffset = consumedBits >> 3;
                _bitOffset = (int)(consumedBits & 7);
            }

            /// <summary>Copies established history; the caller reserves fifteen bytes beyond the logical match.</summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static void CopyInflateMatch(ref byte output, int written, int length, int distance) {
                if (distance == 1) {
                    var repeated = Vector128.Create(Unsafe.Add(ref output, written - 1));
                    for (int copied = 0; copied < length; copied += 16)
                        Unsafe.WriteUnaligned(ref Unsafe.Add(ref output, written + copied), repeated);
                    return;
                }
                if (distance >= 16) {
                    for (int copied = 0; copied < length; copied += 16) {
                        var value = Unsafe.ReadUnaligned<Vector128<byte>>(
                            ref Unsafe.Add(ref output, written + copied - distance));
                        Unsafe.WriteUnaligned(ref Unsafe.Add(ref output, written + copied), value);
                    }
                    return;
                }
                int offset = 0;
                int stableDistance = distance;
                if (distance < 8) {
                    int prefix = Math.Min(8, length);
                    for (; offset < prefix; offset++)
                        Unsafe.Add(ref output, written + offset) = Unsafe.Add(ref output, written - distance + offset);
                    stableDistance = ((7 / distance) + 1) * distance;
                }
                for (; offset < length; offset += 8) {
                    ulong value = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref output, written + offset - stableDistance));
                    Unsafe.WriteUnaligned(ref Unsafe.Add(ref output, written + offset), value);
                }
            }
        }
    }
}
#endif
