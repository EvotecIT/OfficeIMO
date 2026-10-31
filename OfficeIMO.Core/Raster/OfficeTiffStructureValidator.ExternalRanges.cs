using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeTiffStructureValidator {
    private readonly struct OwnedRange {
        internal OwnedRange(int start, int end, int ownerEntry) {
            Start = start; End = end; OwnerEntry = ownerEntry;
        }
        internal int Start { get; }
        internal int End { get; }
        internal int OwnerEntry { get; }
    }

    private readonly struct WritableRange {
        internal WritableRange(int start, int end, int ownerEntry) {
            Start = start; End = end; OwnerEntry = ownerEntry;
        }
        internal int Start { get; }
        internal int End { get; }
        internal int OwnerEntry { get; }
    }

    private readonly struct RangeEvent {
        internal RangeEvent(int position, int id, bool writable, bool starts) {
            Position = position; Id = id; Writable = writable; Starts = starts;
        }
        internal int Position { get; }
        internal int Id { get; }
        internal bool Writable { get; }
        internal bool Starts { get; }
    }

    /// <summary>
    /// Validates all external Exif values erased by one rewrite with one TIFF traversal.
    /// Each triple is the absolute value offset, byte length, and owning entry offset.
    /// </summary>
    internal static bool TryValidateExclusiveExternalWritableRanges(byte[] bytes, int offset, int count,
        CancellationToken token, IReadOnlyList<int> offsetLengthOwnerTriples) {
        token.ThrowIfCancellationRequested();
        if (offsetLengthOwnerTriples == null || offsetLengthOwnerTriples.Count == 0 ||
            offsetLengthOwnerTriples.Count % 3 != 0 || !TryValidate(bytes, offset, count, token)) return false;

        bool little = bytes[offset] == (byte)'I';
        var writable = new List<WritableRange>(offsetLengthOwnerTriples.Count / 3);
        for (int index = 0; index < offsetLengthOwnerTriples.Count; index += 3) {
            if ((index & 0x3FF) == 0) token.ThrowIfCancellationRequested();
            int start = offsetLengthOwnerTriples[index];
            int length = offsetLengthOwnerTriples[index + 1];
            int owner = offsetLengthOwnerTriples[index + 2];
            if (length <= 4 || start < offset + 8 || (long)start + length > (long)offset + count ||
                owner < offset + 8 || owner > offset + count - 12) return false;
            ushort type = ReadUInt16(bytes, owner + 2, little);
            uint valueCount = ReadUInt32(bytes, owner + 4, little);
            int typeSize = GetTypeSize(type);
            ulong valueSize = (ulong)valueCount * (uint)typeSize;
            if (typeSize == 0 || valueSize != (uint)length ||
                ReadUInt32(bytes, owner + 8, little) != (uint)(start - offset)) return false;
            writable.Add(new WritableRange(start, checked(start + length), owner));
        }

        var owned = new List<OwnedRange> { new OwnedRange(offset, offset + 8, -1) };
        var valueByEntry = new Dictionary<int, int>();
        var pending = new Stack<int>();
        var visited = new HashSet<int>();
        pending.Push((int)ReadUInt32(bytes, offset + 4, little));
        while (pending.Count > 0) {
            token.ThrowIfCancellationRequested();
            int relativeIfd = pending.Pop();
            if (!visited.Add(relativeIfd)) return false;
            int absoluteIfd = offset + relativeIfd;
            int entryCount = ReadUInt16(bytes, absoluteIfd, little);
            int entry = absoluteIfd + 2;
            owned.Add(new OwnedRange(absoluteIfd, checked(absoluteIfd + 6 + entryCount * 12), -1));
            for (int index = 0; index < entryCount; index++, entry += 12) {
                if ((index & 255) == 0) token.ThrowIfCancellationRequested();
                ushort tag = ReadUInt16(bytes, entry, little);
                ushort type = ReadUInt16(bytes, entry + 2, little);
                uint valueCount = ReadUInt32(bytes, entry + 4, little);
                int valueLength = checked((int)((ulong)valueCount * (uint)GetTypeSize(type)));
                if (valueLength > 4) {
                    int valueOffset = checked(offset + (int)ReadUInt32(bytes, entry + 8, little));
                    if (valueByEntry.ContainsKey(entry)) return false;
                    valueByEntry.Add(entry, owned.Count);
                    owned.Add(new OwnedRange(valueOffset, checked(valueOffset + valueLength), entry));
                }

                if ((tag == 34665 || tag == 34853 || tag == 40965) && type == 4 && valueCount == 1) {
                    uint child = ReadUInt32(bytes, entry + 8, little);
                    if (child != 0) pending.Push((int)child);
                } else if ((type == 13 || tag == 330 && type == 4) && valueCount > 0) {
                    int values = valueLength <= 4 ? entry + 8 : offset + (int)ReadUInt32(bytes, entry + 8, little);
                    for (uint item = 0; item < valueCount; item++) {
                        if ((item & 255U) == 0) token.ThrowIfCancellationRequested();
                        uint child = ReadUInt32(bytes, values + (int)item * 4, little);
                        if (child != 0) pending.Push((int)child);
                    }
                }
            }
            uint next = ReadUInt32(bytes, entry, little);
            if (next != 0) pending.Push((int)next);
        }

        foreach (WritableRange range in writable) {
            if (!valueByEntry.TryGetValue(range.OwnerEntry, out int id)) return false;
            OwnedRange owner = owned[id];
            if (owner.Start != range.Start || owner.End != range.End) return false;
        }

        var events = new List<RangeEvent>(checked((owned.Count + writable.Count) * 2));
        for (int index = 0; index < owned.Count; index++) {
            events.Add(new RangeEvent(owned[index].Start, index, writable: false, starts: true));
            events.Add(new RangeEvent(owned[index].End, index, writable: false, starts: false));
        }
        for (int index = 0; index < writable.Count; index++) {
            events.Add(new RangeEvent(writable[index].Start, index, writable: true, starts: true));
            events.Add(new RangeEvent(writable[index].End, index, writable: true, starts: false));
        }
        events.Sort((left, right) => left.Position.CompareTo(right.Position));
        int activeOwned = 0, activeWritable = 0, ownedId = 0, writableId = 0;
        for (int index = 0; index < events.Count;) {
            if ((index & 0x3FFF) == 0) token.ThrowIfCancellationRequested();
            int position = events[index].Position;
            do {
                RangeEvent item = events[index++];
                if (item.Writable) { activeWritable += item.Starts ? 1 : -1; writableId ^= item.Id; }
                else { activeOwned += item.Starts ? 1 : -1; ownedId ^= item.Id; }
            } while (index < events.Count && events[index].Position == position);
            if (index < events.Count && activeWritable > 0 &&
                (activeWritable != 1 || activeOwned != 1 ||
                 owned[ownedId].OwnerEntry != writable[writableId].OwnerEntry)) return false;
        }
        return true;
    }
}
