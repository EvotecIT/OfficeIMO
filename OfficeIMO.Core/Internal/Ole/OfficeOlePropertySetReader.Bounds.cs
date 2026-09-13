using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeOlePropertySetReader {
        // Validate each property's own extent before decoding any values. A valid stream-wide
        // range alone would allow a string/blob to consume a sibling property or section.
        private static void ValidateSections(byte[] bytes, CancellationToken token) {
            EnsureAvailable(bytes, 0, 28);
            if (ReadUInt16(bytes, 0) != 0xfffe) throw new InvalidDataException("Unsupported OLE byte order.");
            uint count = ReadUInt32(bytes, 24);
            if (count == 0 || count > 8) throw new InvalidDataException("Invalid OLE section count.");
            int tableEnd = 28 + (int)count * 20;
            EnsureAvailable(bytes, 0, tableEnd);
            var sections = new List<(int Start, int End)>();
            for (int index = 0; index < count; index++) {
                token.ThrowIfCancellationRequested();
                uint rawStart = ReadUInt32(bytes, 28 + index * 20 + 16);
                if (rawStart < tableEnd || rawStart > bytes.Length) throw new InvalidDataException("Invalid OLE section offset.");
                int start = (int)rawStart;
                EnsureAvailable(bytes, start, 8);
                uint size = ReadUInt32(bytes, start);
                if (size < 8 || size > bytes.Length - start) throw new InvalidDataException("Invalid OLE section size.");
                int end = start + (int)size;
                foreach (var existing in sections)
                    if (start < existing.End && end > existing.Start) throw new InvalidDataException("Overlapping OLE sections.");
                sections.Add((start, end));
            }
            foreach (var section in sections) ValidateSection(bytes, section.Start, section.End, token);
        }

        private static void ValidateSection(byte[] bytes, int start, int end, CancellationToken token) {
            uint count = ReadUInt32(bytes, start + 4);
            if (count > 1024 || count > (end - start - 8) / 8) throw new InvalidDataException("Invalid OLE property table.");
            int tableEnd = start + 8 + (int)count * 8;
            var entries = new List<(uint Id, int Offset)>();
            var ids = new HashSet<uint>(); var offsets = new HashSet<int>();
            for (int index = 0; index < count; index++) {
                token.ThrowIfCancellationRequested();
                uint id = ReadUInt32(bytes, start + 8 + index * 8);
                uint relative = ReadUInt32(bytes, start + 12 + index * 8);
                if (relative < tableEnd - start || relative > end - start - 4)
                    throw new InvalidDataException("OLE property lies outside its section payload.");
                int offset = start + (int)relative;
                if (!ids.Add(id) || !offsets.Add(offset)) throw new InvalidDataException("Duplicate OLE property ID or offset.");
                entries.Add((id, offset));
            }
            entries.Sort((left, right) => left.Offset.CompareTo(right.Offset));
            int codePage = 1252;
            // Validate ordinary values before looking at the code page used by dictionaries.
            for (int index = 0; index < entries.Count; index++) {
                token.ThrowIfCancellationRequested();
                var entry = entries[index];
                int limit = index + 1 < entries.Count ? entries[index + 1].Offset : end;
                if (entry.Id == PropertyDictionaryId) continue;
                ValidateValue(bytes, entry.Offset, limit);
                if (entry.Id == CodePagePropertyId) {
                    ushort type = ReadUInt16(bytes, entry.Offset);
                    if (type == 2) codePage = ReadUInt16(bytes, entry.Offset + 4);
                    else if (type == 3) codePage = unchecked((int)ReadUInt32(bytes, entry.Offset + 4));
                }
            }
            for (int index = 0; index < entries.Count; index++) {
                if (entries[index].Id != PropertyDictionaryId) continue;
                int limit = index + 1 < entries.Count ? entries[index + 1].Offset : end;
                ValidateDictionary(bytes, entries[index].Offset, limit, codePage, token);
            }
        }

        private static void ValidateValue(byte[] bytes, int offset, int limit) {
            RequireRange(offset, 4, limit);
            ushort type = ReadUInt16(bytes, offset);
            int payload = offset + 4;
            switch (type) {
                case 2: case 0xb: case 0x12: RequireRange(payload, 2, limit); break;
                case 3: case 4: case 0x13: case 0x16: case 0x17: RequireRange(payload, 4, limit); break;
                case 5: case 6: case 7: case 0x14: case 0x15: case 0x40: RequireRange(payload, 8, limit); break;
                case 0x10: case 0x11: RequireRange(payload, 1, limit); break;
                case 0x1e: case 0x1f: case 0x41:
                    RequireRange(payload, 4, limit);
                    BoundedLength(ReadUInt32(bytes, payload), type == 0x1f ? 2 : 1, limit - payload - 4);
                    break;
            }
        }

        private static void ValidateDictionary(byte[] bytes, int offset, int limit, int codePage, CancellationToken token) {
            RequireRange(offset, 4, limit);
            uint count = ReadUInt32(bytes, offset);
            if (count > 1024) throw new InvalidDataException("Too many OLE dictionary entries.");
            int cursor = offset + 4;
            var ids = new HashSet<uint>();
            for (int index = 0; index < count; index++) {
                token.ThrowIfCancellationRequested();
                RequireRange(cursor, 8, limit);
                if (!ids.Add(ReadUInt32(bytes, cursor))) throw new InvalidDataException("Duplicate OLE dictionary ID.");
                uint length = ReadUInt32(bytes, cursor + 4);
                cursor += 8;
                cursor += BoundedLength(length, codePage == 1200 ? 2 : 1, limit - cursor);
                if (codePage == 1200) {
                    int padding = (4 - cursor % 4) % 4;
                    RequireRange(cursor, padding, limit);
                    cursor += padding;
                }
            }
        }

        private static int BoundedLength(uint length, int width, int available) {
            if (available < 0 || length > available / width) throw new InvalidDataException("OLE value exceeds its property boundary.");
            return (int)length * width;
        }

        private static void RequireRange(int offset, int count, int limit) {
            if (offset < 0 || count < 0 || offset > limit || count > limit - offset)
                throw new InvalidDataException("Truncated OLE property value.");
        }
    }
}
