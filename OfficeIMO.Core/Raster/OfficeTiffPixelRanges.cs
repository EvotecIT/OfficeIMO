using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Plans bounded pixel references before expanding ranges used by lossless metadata edits.</summary>
internal static class OfficeTiffPixelRanges {
    internal const int MaximumPixelReferences = 65535;
    internal const long PlanningBytes = 256L * 1024L;

    private readonly struct Values {
        internal Values(int offset, int count, int size) {
            Offset = offset;
            Count = count;
            Size = size;
        }
        internal int Offset { get; }
        internal int Count { get; }
        internal int Size { get; }
    }
    private readonly struct Pair {
        internal Pair(Values offsets, Values lengths) {
            Offsets = offsets;
            Lengths = lengths;
        }
        internal Values Offsets { get; }
        internal Values Lengths { get; }
    }

    internal static List<(long Start, long Length)> Read(byte[] input, CancellationToken token, long additionallyRetainedBytes = 0L) {
        token.ThrowIfCancellationRequested();
        if (checked(input.LongLength + additionallyRetainedBytes + PlanningBytes) > OfficeRasterGuards.MaximumDecodedBytes) {
            throw new ArgumentException("TIFF metadata planning exceeds the managed working-set limit.");
        }
        if (!OfficeTiffStructureValidator.TryValidate(input, 0, input.Length, token)) {
            throw new FormatException("TIFF metadata is structurally invalid.");
        }
        bool little = input[0] == 73;
        var pending = new Stack<int>();
        var visited = new HashSet<int>();
        var pairs = new List<Pair>();
        pending.Push(checked((int)OfficeExifProfileCodec.Read(input, 4, 4, little)));
        int references = 0;
        while (pending.Count != 0) {
            token.ThrowIfCancellationRequested();
            int cursor = pending.Pop();
            if (!visited.Add(cursor) || visited.Count > 1024) {
                throw new FormatException("TIFF directories are cyclic or too numerous.");
            }
            int count = checked((int)OfficeExifProfileCodec.Read(input, cursor, 2, little));
            var fields = new Dictionary<int, Values>();
            for (int index = 0; index < count; index++) {
                if ((index & 255) == 0) {
                    token.ThrowIfCancellationRequested();
                }
                int entry = cursor + 2 + index * 12;
                int id = (int)OfficeExifProfileCodec.Read(input, entry, 2, little);
                int type = (int)OfficeExifProfileCodec.Read(input, entry + 2, 2, little);
                int elements = checked((int)OfficeExifProfileCodec.Read(input, entry + 4, 4, little));
                if ((id == 34665 || id == 34853 || id == 40965) && type == 4 && elements == 1) {
                    Schedule((uint)OfficeExifProfileCodec.Read(input, entry + 8, 4, little));
                } else if ((type == 13 || id == 330 && type == 4) && elements > 0) {
                    int values = elements == 1 ? entry + 8 : checked((int)OfficeExifProfileCodec.Read(input, entry + 8, 4, little));
                    for (int value = 0; value < elements; value++) {
                        if ((value & 255) == 0) {
                            token.ThrowIfCancellationRequested();
                        }
                        Schedule((uint)OfficeExifProfileCodec.Read(input, values + value * 4, 4, little));
                    }
                }
                if (id != 273 && id != 279 && id != 324 && id != 325 && id != 513 && id != 514) {
                    continue;
                }
                int size = type == 3 ? 2 : type == 4 ? 4 : throw new FormatException("Unsupported TIFF pixel offset representation.");
                if (elements <= 0 || elements > MaximumPixelReferences || fields.ContainsKey(id)) {
                    throw new FormatException("TIFF pixel references are empty, duplicated, or too numerous.");
                }
                int at = (long)elements * size <= 4 ? entry + 8 : checked((int)OfficeExifProfileCodec.Read(input, entry + 8, 4, little));
                fields.Add(id, new Values(at, elements, size));
            }
            AddPair(273, 279);
            AddPair(324, 325);
            AddPair(513, 514);
            Schedule((uint)OfficeExifProfileCodec.Read(input, cursor + 2 + count * 12, 4, little));

            void AddPair(int offsetId, int lengthId) {
                bool hasOffsets = fields.TryGetValue(offsetId, out Values offsets);
                bool hasLengths = fields.TryGetValue(lengthId, out Values lengths);
                if (!hasOffsets && !hasLengths) {
                    return;
                }
                if (!hasOffsets || !hasLengths || offsets.Count != lengths.Count) {
                    throw new FormatException("TIFF pixel offset and length fields are incomplete or differ in count.");
                }
                if (offsets.Count > MaximumPixelReferences - references) {
                    throw new FormatException("TIFF exceeds the aggregate pixel-reference limit.");
                }
                references += offsets.Count;
                pairs.Add(new Pair(offsets, lengths));
            }
            void Schedule(uint offset) {
                if (offset != 0) {
                    pending.Push(checked((int)offset));
                }
            }
        }
        // Exact capacity prevents List growth from transiently retaining two large arrays.
        long retained = checked(input.LongLength + additionallyRetainedBytes + PlanningBytes + references * 16L + 24L);
        if (retained > OfficeRasterGuards.MaximumDecodedBytes) {
            throw new ArgumentException("TIFF pixel ranges exceed the managed working-set limit.");
        }
        var result = new List<(long Start, long Length)>(references);
        foreach (Pair pair in pairs) {
            for (int index = 0; index < pair.Offsets.Count; index++) {
                if ((index & 255) == 0) {
                    token.ThrowIfCancellationRequested();
                }
                long start = (long)OfficeExifProfileCodec.Read(input, pair.Offsets.Offset + index * pair.Offsets.Size, pair.Offsets.Size, little);
                long length = (long)OfficeExifProfileCodec.Read(input, pair.Lengths.Offset + index * pair.Lengths.Size, pair.Lengths.Size, little);
                if (start > input.LongLength || length > input.LongLength - start) {
                    throw new FormatException("TIFF pixel data is outside the encoded container.");
                }
                if (length != 0) {
                    result.Add((start, length));
                }
            }
        }
        result.Sort((first, second) => first.Start.CompareTo(second.Start));
        int merged = 0;
        for (int index = 0; index < result.Count; index++) {
            if ((index & 255) == 0) {
                token.ThrowIfCancellationRequested();
            }
            (long Start, long Length) range = result[index];
            if (merged != 0 && range.Start <= result[merged - 1].Start + result[merged - 1].Length) {
                (long Start, long Length) previous = result[merged - 1];
                result[merged - 1] = (previous.Start, Math.Max(previous.Start + previous.Length, range.Start + range.Length) - previous.Start);
            } else {
                result[merged++] = range;
            }
        }
        if (merged != result.Count) {
            result.RemoveRange(merged, result.Count - merged);
        }
        token.ThrowIfCancellationRequested();
        return result;
    }

    internal static long RetainedBytes(List<(long Start, long Length)> ranges) => checked(ranges.Capacity * 16L + 24L);
    internal static bool Overlaps(List<(long Start, long Length)> ranges, long start, long length) {
        int low = 0, high = ranges.Count;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (ranges[middle].Start + ranges[middle].Length <= start) {
                low = middle + 1;
            } else {
                high = middle;
            }
        }
        return low < ranges.Count && ranges[low].Start < start + length;
    }
}
