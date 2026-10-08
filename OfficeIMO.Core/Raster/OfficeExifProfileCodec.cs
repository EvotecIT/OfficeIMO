using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Reads Exif values and appends replacement directories without relocating opaque maker-note data.</summary>
internal static class OfficeExifProfileCodec {
    internal const int MaximumProfileBytes = 16 * 1024 * 1024;
    internal sealed class Field {
        internal OfficeExifTag Tag;
        internal int EntryOffset;
        internal int ValueOffset;
        internal int ValueLength;
        internal uint Count;
        internal object Value = Array.Empty<byte>();
    }
    internal sealed class Directory {
        internal int Offset;
        internal uint Next;
        internal readonly List<Field> Fields = new List<Field>();
    }
    internal sealed class Profile {
        internal byte[] Bytes = Array.Empty<byte>();
        internal bool Little;
        internal long RetainedManagedBytes;
        internal readonly Dictionary<OfficeExifDirectory, Directory> Directories = new Dictionary<OfficeExifDirectory, Directory>();
    }

    internal static Profile Parse(byte[] bytes, bool imageTiff = false, CancellationToken cancellationToken = default, long additionallyRetainedBytes = 0L) {
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (bytes.Length > (imageTiff ? OfficeRasterGuards.MaximumEncodedBytes : MaximumProfileBytes)) throw new FormatException("Exif metadata exceeds the profile-size limit.");
        if (additionallyRetainedBytes < 0L || checked(additionallyRetainedBytes + bytes.LongLength * (imageTiff ? 1L : 2L) + 24L) > OfficeRasterGuards.MaximumDecodedBytes) throw new FormatException("Exif metadata exceeds the managed working-set limit.");
        int skip = bytes.Length >= 6 && bytes[0] == 69 && bytes[1] == 120 && bytes[2] == 105 && bytes[3] == 102 && bytes[4] == 0 && bytes[5] == 0 ? 6 : 0;
        if (!OfficeTiffStructureValidator.TryValidateExif(bytes, skip, bytes.Length - skip, cancellationToken)) throw new FormatException("Exif metadata is not a valid bounded classic TIFF profile.");
        // Image-TIFF parsing is operation-local and read-only. Borrow its encoded
        // container instead of retaining another complete raster-sized copy.
        var profile = new Profile { Bytes = imageTiff && skip == 0 ? bytes : new byte[bytes.Length - skip], Little = bytes[skip] == 73, RetainedManagedBytes = checked(additionallyRetainedBytes + bytes.LongLength * (imageTiff && skip == 0 ? 1 : 2)) };
        if (!ReferenceEquals(profile.Bytes, bytes)) Buffer.BlockCopy(bytes, skip, profile.Bytes, 0, profile.Bytes.Length);
        ParseDirectory(profile, OfficeExifDirectory.Image, checked((int)Read(profile.Bytes, 4, 4, profile.Little)), imageTiff, cancellationToken);
        return profile;
    }

    private static void ParseDirectory(Profile profile, OfficeExifDirectory kind, int offset, bool imageTiff, CancellationToken token) {
        if (offset == 0 || profile.Directories.ContainsKey(kind)) return;
        byte[] bytes = profile.Bytes;
        var directory = new Directory { Offset = offset };
        profile.Directories.Add(kind, directory);
        int count = checked((int)Read(bytes, offset, 2, profile.Little));
        for (int index = 0; index < count; index++) {
            if ((index & 255) == 0) token.ThrowIfCancellationRequested();
            int entry = offset + 2 + index * 12;
            ushort id = (ushort)Read(bytes, entry, 2, profile.Little);
            int type = (int)Read(bytes, entry + 2, 2, profile.Little);
            uint length = (uint)Read(bytes, entry + 4, 4, profile.Little);
            int size = Size(type);
            int byteLength = checked((int)length * size);
            int valueOffset = byteLength <= 4 ? entry + 8 : checked((int)Read(bytes, entry + 8, 4, profile.Little));
            if (imageTiff && kind == OfficeExifDirectory.Image && !IsTiffMetadataTag(id)) continue;
            if (byteLength > MaximumProfileBytes) throw new FormatException("An Exif field exceeds the metadata-size limit.");
            profile.RetainedManagedBytes = checked(profile.RetainedManagedBytes + byteLength * 2L + 96L);
            if (profile.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes) throw new FormatException("Exif metadata exceeds the managed working-set limit.");
            // TIFF IFD pointers (type 13) and private unknown fields remain intact but are not exposed as ordinary values.
            if (type < 1 || type > 12) continue;
            var tag = new OfficeExifTag(id, (OfficeExifDataType)type, kind, KnownName(kind, id));
            var field = new Field { Tag = tag, EntryOffset = entry, ValueOffset = valueOffset, ValueLength = byteLength, Count = length };
            field.Value = Decode(bytes, valueOffset, checked((int)length), tag.DataType, profile.Little, token);
            directory.Fields.Add(field);
            if (type == 4 && length == 1) {
                int child = checked((int)Read(bytes, valueOffset, 4, profile.Little));
                if (kind == OfficeExifDirectory.Image && id == 34665) ParseDirectory(profile, OfficeExifDirectory.Exif, child, imageTiff, token);
                if (kind == OfficeExifDirectory.Image && id == 34853) ParseDirectory(profile, OfficeExifDirectory.Gps, child, imageTiff, token);
                if (kind == OfficeExifDirectory.Exif && id == 40965) ParseDirectory(profile, OfficeExifDirectory.Interoperability, child, imageTiff, token);
            }
        }
        directory.Next = (uint)Read(bytes, offset + 2 + count * 12, 4, profile.Little);
    }

    internal static byte[] Encode(Profile? source, IDictionary<OfficeExifTag, OfficeExifValue> changes, ISet<OfficeExifTag> removed, bool imageTiff = false, CancellationToken cancellationToken = default, long additionallyRetainedBytes = 0L) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] original = source?.Bytes ?? new byte[] { 73, 73, 42, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        bool little = source?.Little ?? true;
        long retained = checked((source?.RetainedManagedBytes ?? original.LongLength) + additionallyRetainedBytes);
        if (changes.Count == 0 && removed.Count == 0) {
            if (checked(retained + original.LongLength + 24L) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Metadata rewriting exceeds the managed working-set limit.");
            return (byte[])original.Clone();
        }
        long growth = 128;
        foreach (KeyValuePair<OfficeExifTag, OfficeExifValue> change in changes) {
            long bytes = change.Value.EncodedByteLength;
            retained = checked(retained + bytes * 4L + 128L);
            growth = checked(growth + bytes + 16L);
        }
        if (source != null) foreach (Directory directory in source.Directories.Values) {
            long count = (long)Read(original, directory.Offset, 2, little);
            growth = checked(growth + 6L + count * 12L);
            retained = checked(retained + count * 96L);
        }
        long capacityHint = Math.Min(imageTiff ? OfficeRasterGuards.MaximumEncodedBytes : MaximumProfileBytes, checked(original.LongLength + growth));
        using var output = new OfficeMetadataRewriteStream(retained, checked((int)capacityHint), cancellationToken, imageTiff ? OfficeRasterGuards.MaximumEncodedBytes : MaximumProfileBytes);
        output.Write(original, 0, original.Length);
        var changedDirectories = new HashSet<OfficeExifDirectory>();
        foreach (OfficeExifTag tag in changes.Keys) { if (!imageTiff) ValidateUserTag(tag); changedDirectories.Add(tag.Directory); }
        foreach (OfficeExifTag tag in removed) { if (!imageTiff) ValidateUserTag(tag); changedDirectories.Add(tag.Directory); }
        if (changedDirectories.Contains(OfficeExifDirectory.Interoperability)) changedDirectories.Add(OfficeExifDirectory.Exif);
        if (changedDirectories.Count > 0) changedDirectories.Add(OfficeExifDirectory.Image);
        if (source != null) ValidateErasedDirectories(source, changedDirectories, cancellationToken);
        var newOffsets = new Dictionary<OfficeExifDirectory, uint>();
        foreach (OfficeExifDirectory kind in new[] { OfficeExifDirectory.Interoperability, OfficeExifDirectory.Exif, OfficeExifDirectory.Gps, OfficeExifDirectory.Image }) {
            if (!changedDirectories.Contains(kind)) continue;
            Directory? directory = source != null && source.Directories.TryGetValue(kind, out Directory? found) ? found : null;
            var entries = new SortedDictionary<ushort, byte[]>();
            if (directory != null) {
                int count = (int)Read(original, directory.Offset, 2, little);
                for (int i = 0; i < count; i++) {
                    if ((i & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                    int entry = directory.Offset + 2 + i * 12;
                    ushort id = (ushort)Read(original, entry, 2, little);
                    var identity = new OfficeExifTag(id, OfficeExifDataType.Undefined, kind);
                    if (removed.Contains(identity) || changes.ContainsKey(identity)) continue;
                    if (entries.ContainsKey(id)) throw new FormatException("Duplicate Exif fields cannot be edited safely.");
                    var raw = new byte[12]; Buffer.BlockCopy(original, entry, raw, 0, 12); entries.Add(id, raw);
                }
                // Erase obsolete values only when no other reachable field owns those bytes.
                foreach (Field field in directory.Fields) {
                    if (!removed.Contains(field.Tag) && !changes.ContainsKey(field.Tag)) continue;
                    if (field.ValueLength > 4 && !OfficeTiffStructureValidator.TryValidateExclusiveWritableRanges(original, 0, original.Length, cancellationToken,
                            new[] { field.ValueOffset, field.ValueLength, field.EntryOffset })) throw new FormatException("An edited Exif value overlaps another field.");
                    output.Position = field.ValueOffset; output.Write(new byte[field.ValueLength], 0, field.ValueLength);
                    output.Position = field.EntryOffset; output.Write(new byte[12], 0, 12);
                }
                // Every live entry was copied above. Erase the superseded table as well as
                // changed payloads so repeated edits cannot retain stale inline values.
                int tableLength = checked(6 + count * 12);
                output.Position = directory.Offset; output.Write(new byte[tableLength], 0, tableLength);
            }
            output.Position = output.Length;
            foreach (KeyValuePair<OfficeExifTag, OfficeExifValue> change in changes) {
                if (change.Key.Directory != kind) continue;
                byte[] payload = change.Value.EncodeValue(little, out uint count, cancellationToken);
                byte[] entry = MakeEntry(change.Key.Id, (ushort)change.Key.DataType, count, little);
                if (payload.Length <= 4) Buffer.BlockCopy(payload, 0, entry, 8, payload.Length);
                else { Align(output); Write(entry, 8, (uint)output.Length, 4, little); output.Write(payload, 0, payload.Length); }
                entries[change.Key.Id] = entry;
            }
            SetChild(kind, OfficeExifDirectory.Interoperability, 40965);
            SetChild(kind, OfficeExifDirectory.Exif, 34665);
            SetChild(kind, OfficeExifDirectory.Gps, 34853);
            Align(output);
            newOffsets[kind] = checked((uint)output.Length);
            byte[] table = new byte[checked(6 + entries.Count * 12)];
            if (entries.Count > ushort.MaxValue) throw new FormatException("Exif directory contains too many fields.");
            Write(table, 0, (uint)entries.Count, 2, little);
            int cursor = 2; foreach (byte[] entry in entries.Values) { Buffer.BlockCopy(entry, 0, table, cursor, 12); cursor += 12; }
            Write(table, cursor, directory?.Next ?? 0, 4, little);
            output.Write(table, 0, table.Length);
            if (output.Length > (imageTiff ? OfficeRasterGuards.MaximumEncodedBytes : MaximumProfileBytes)) throw new FormatException("Edited Exif metadata exceeds the profile-size limit.");

            void SetChild(OfficeExifDirectory parent, OfficeExifDirectory child, ushort id) {
                bool belongs = parent == OfficeExifDirectory.Image && child != OfficeExifDirectory.Interoperability || parent == OfficeExifDirectory.Exif && child == OfficeExifDirectory.Interoperability;
                if (removed.Contains(new OfficeExifTag(id, OfficeExifDataType.Long, parent))) return;
                if (belongs && newOffsets.TryGetValue(child, out uint childOffset)) {
                    byte[] entry = MakeEntry(id, 4, 1, little); Write(entry, 8, childOffset, 4, little); entries[id] = entry;
                }
            }
        }
        byte[] result = output.ToArray(); Write(result, 4, newOffsets[OfficeExifDirectory.Image], 4, little);
        if (!OfficeTiffStructureValidator.TryValidateExif(result, 0, result.Length, cancellationToken)) throw new FormatException("The edited Exif profile is structurally invalid.");
        return result;
    }

    private static void ValidateErasedDirectories(Profile source, ISet<OfficeExifDirectory> changed, CancellationToken token) {
        var erased = new List<(int Start, int Length)>();
        foreach (OfficeExifDirectory kind in changed) if (source.Directories.TryGetValue(kind, out Directory? directory)) { int count = (int)Read(source.Bytes, directory.Offset, 2, source.Little); erased.Add((directory.Offset, checked(6 + count * 12))); }
        var pending = new Stack<int>(); var visited = new HashSet<int>(); pending.Push(checked((int)Read(source.Bytes, 4, 4, source.Little)));
        while (pending.Count != 0) {
            token.ThrowIfCancellationRequested();
            int at = pending.Pop(); if (at == 0 || !visited.Add(at)) continue; int count = (int)Read(source.Bytes, at, 2, source.Little);
            for (int i = 0; i < count; i++) {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                int entry = at + 2 + i * 12; int id = (int)Read(source.Bytes, entry, 2, source.Little); int type = (int)Read(source.Bytes, entry + 2, 2, source.Little); int elements = checked((int)Read(source.Bytes, entry + 4, 4, source.Little));
                int length = checked(elements * Size(type)); int value = length <= 4 ? entry + 8 : checked((int)Read(source.Bytes, entry + 8, 4, source.Little));
                if (length > 4) foreach ((int Start, int Length) range in erased) if (value < range.Start + range.Length && range.Start < (long)value + length) throw new FormatException("An Exif value aliases a directory table that must be rewritten.");
                bool pointer = type == 13 || type == 4 && (id == 34665 || id == 34853 || id == 40965 || id == 330);
                if (pointer) for (int item = 0; item < elements; item++) { if ((item & 255) == 0) token.ThrowIfCancellationRequested(); pending.Push(checked((int)Read(source.Bytes, value + item * 4, 4, source.Little))); }
            }
            pending.Push(checked((int)Read(source.Bytes, at + 2 + count * 12, 4, source.Little)));
        }
    }

    internal static bool IsTiffMetadataTag(ushort id) => id == 270 || id == 271 || id == 272 || id == 274 || id == 282 || id == 283 || id == 296 || id == 305 || id == 306 || id == 315 || id == 33432 || id == 34665 || id == 34853 || id == 700 || id == 34675 || id == 33723;

    private static void ValidateUserTag(OfficeExifTag tag) {
        if (tag.Id == 34665 || tag.Id == 34853 || tag.Id == 40965 || tag.Id == 330 || tag.Id == 513 || tag.Id == 514) throw new ArgumentException("Exif structural pointers cannot be edited as ordinary values.", nameof(tag));
    }
    private static string? KnownName(OfficeExifDirectory directory, ushort id) => (directory, id) switch {
        (OfficeExifDirectory.Image, 256) => nameof(OfficeExifTag.ImageWidth), (OfficeExifDirectory.Image, 257) => nameof(OfficeExifTag.ImageLength),
        (OfficeExifDirectory.Image, 270) => nameof(OfficeExifTag.ImageDescription), (OfficeExifDirectory.Image, 271) => nameof(OfficeExifTag.Make), (OfficeExifDirectory.Image, 272) => nameof(OfficeExifTag.Model),
        (OfficeExifDirectory.Image, 274) => nameof(OfficeExifTag.Orientation), (OfficeExifDirectory.Image, 282) => nameof(OfficeExifTag.XResolution), (OfficeExifDirectory.Image, 283) => nameof(OfficeExifTag.YResolution),
        (OfficeExifDirectory.Image, 296) => nameof(OfficeExifTag.ResolutionUnit), (OfficeExifDirectory.Image, 305) => nameof(OfficeExifTag.Software), (OfficeExifDirectory.Image, 306) => nameof(OfficeExifTag.DateTime),
        (OfficeExifDirectory.Image, 315) => nameof(OfficeExifTag.Artist), (OfficeExifDirectory.Image, 33432) => nameof(OfficeExifTag.Copyright),
        (OfficeExifDirectory.Exif, 33434) => nameof(OfficeExifTag.ExposureTime), (OfficeExifDirectory.Exif, 33437) => nameof(OfficeExifTag.FNumber), (OfficeExifDirectory.Exif, 34855) => nameof(OfficeExifTag.ISOSpeedRatings),
        (OfficeExifDirectory.Exif, 36864) => nameof(OfficeExifTag.ExifVersion), (OfficeExifDirectory.Exif, 36867) => nameof(OfficeExifTag.DateTimeOriginal), (OfficeExifDirectory.Exif, 36868) => nameof(OfficeExifTag.DateTimeDigitized),
        (OfficeExifDirectory.Exif, 37510) => nameof(OfficeExifTag.UserComment), (OfficeExifDirectory.Gps, 1) => nameof(OfficeExifTag.GPSLatitudeRef), (OfficeExifDirectory.Gps, 2) => nameof(OfficeExifTag.GPSLatitude),
        (OfficeExifDirectory.Gps, 3) => nameof(OfficeExifTag.GPSLongitudeRef), (OfficeExifDirectory.Gps, 4) => nameof(OfficeExifTag.GPSLongitude), _ => null
    };
    private static void Align(MemoryStream stream) { if ((stream.Length & 1) != 0) stream.WriteByte(0); }
    private static byte[] MakeEntry(ushort id, ushort type, uint count, bool little) {
        var entry = new byte[12]; Write(entry, 0, id, 2, little); Write(entry, 2, type, 2, little); Write(entry, 4, count, 4, little); return entry;
    }
    internal static int Size(int type) => type == 3 || type == 8 ? 2 : type == 4 || type == 9 || type == 11 || type == 13 ? 4 : type == 5 || type == 10 || type == 12 ? 8 : 1;
    internal static object Decode(byte[] bytes, int offset, int count, OfficeExifDataType type, bool little, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (type == OfficeExifDataType.Ascii) return Encoding.ASCII.GetString(bytes, offset, count).TrimEnd('\0');
        if (type == OfficeExifDataType.Byte || type == OfficeExifDataType.Undefined) { var copied = new byte[count]; Buffer.BlockCopy(bytes, offset, copied, 0, count); return count == 1 && type == OfficeExifDataType.Byte ? (object)copied[0] : copied; }
        Type element = type switch { OfficeExifDataType.Byte or OfficeExifDataType.Undefined => typeof(byte), OfficeExifDataType.SignedByte => typeof(sbyte), OfficeExifDataType.Short => typeof(ushort), OfficeExifDataType.SignedShort => typeof(short), OfficeExifDataType.Long => typeof(uint), OfficeExifDataType.SignedLong => typeof(int), OfficeExifDataType.Rational => typeof(OfficeRational), OfficeExifDataType.SignedRational => typeof(OfficeSignedRational), OfficeExifDataType.Float => typeof(float), OfficeExifDataType.Double => typeof(double), _ => throw new FormatException("Unsupported Exif field type.") };
        Array values = Array.CreateInstance(element, count);
        for (int i = 0; i < count; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int position = offset + i * Size((int)type);
            ulong value = Read(bytes, position, Math.Min(8, Size((int)type)), little);
            object decoded = type switch { OfficeExifDataType.Byte or OfficeExifDataType.Undefined => (object)(byte)value, OfficeExifDataType.SignedByte => (sbyte)value, OfficeExifDataType.Short => (ushort)value, OfficeExifDataType.SignedShort => (short)value, OfficeExifDataType.Long => (uint)value, OfficeExifDataType.SignedLong => (int)value, OfficeExifDataType.Rational => new OfficeRational((uint)Read(bytes, position, 4, little), (uint)Read(bytes, position + 4, 4, little)), OfficeExifDataType.SignedRational => new OfficeSignedRational((int)Read(bytes, position, 4, little), (int)Read(bytes, position + 4, 4, little)), OfficeExifDataType.Float => BitConverter.ToSingle(BitConverter.GetBytes((uint)value), 0), OfficeExifDataType.Double => BitConverter.Int64BitsToDouble((long)value), _ => throw new FormatException() };
            values.SetValue(decoded, i);
        }
        return count == 1 && type != OfficeExifDataType.Undefined ? values.GetValue(0)! : values;
    }
    internal static byte[] EncodeValue(OfficeExifDataType type, object value, bool little, out uint count, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (type == OfficeExifDataType.Ascii) {
            if (!(value is string text) || text.IndexOf('\0') >= 0) throw new ArgumentException("An ASCII Exif field requires text without embedded NUL characters.", nameof(value));
            foreach (char ch in text) if (ch > 127) throw new ArgumentException("An ASCII Exif field cannot contain non-ASCII text.", nameof(value));
            byte[] ascii = Encoding.ASCII.GetBytes(text + "\0"); count = (uint)ascii.Length; return ascii;
        }
        Array values = value as Array ?? new object[] { value };
        count = checked((uint)values.Length);
        if (values.Rank != 1 || values.GetLowerBound(0) != 0 || values.Length == 0 || (long)values.Length * Size((int)type) > MaximumProfileBytes) throw new ArgumentException("An Exif value requires a nonempty bounded one-dimensional array.", nameof(value));
        var bytes = new byte[checked(values.Length * Size((int)type))];
        for (int i = 0; i < values.Length; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            object item = values.GetValue(i)!; int offset = i * Size((int)type);
            if (type == OfficeExifDataType.Rational) { if (!(item is OfficeRational r)) throw new ArgumentException("Rational fields require OfficeRational values."); Write(bytes, offset, r.Numerator, 4, little); Write(bytes, offset + 4, r.Denominator, 4, little); }
            else if (type == OfficeExifDataType.SignedRational) { if (!(item is OfficeSignedRational r)) throw new ArgumentException("Signed rational fields require OfficeSignedRational values."); Write(bytes, offset, unchecked((uint)r.Numerator), 4, little); Write(bytes, offset + 4, unchecked((uint)r.Denominator), 4, little); }
            else {
                ulong encoded = type switch { OfficeExifDataType.Byte or OfficeExifDataType.Undefined => Convert.ToByte(item), OfficeExifDataType.SignedByte => unchecked((byte)Convert.ToSByte(item)), OfficeExifDataType.Short => Convert.ToUInt16(item), OfficeExifDataType.SignedShort => unchecked((ushort)Convert.ToInt16(item)), OfficeExifDataType.Long => Convert.ToUInt32(item), OfficeExifDataType.SignedLong => unchecked((uint)Convert.ToInt32(item)), OfficeExifDataType.Float => BitConverter.ToUInt32(BitConverter.GetBytes(Convert.ToSingle(item)), 0), OfficeExifDataType.Double => unchecked((ulong)BitConverter.DoubleToInt64Bits(Convert.ToDouble(item))), _ => throw new ArgumentException("Unsupported Exif data type.", nameof(type)) };
                Write(bytes, offset, encoded, Size((int)type), little);
            }
        }
        return bytes;
    }
    internal static ulong Read(byte[] bytes, int offset, int size, bool little) { ulong value = 0; for (int i = 0; i < size; i++) value |= (ulong)bytes[offset + i] << (8 * (little ? i : size - 1 - i)); return value; }
    internal static void Write(byte[] bytes, int offset, ulong value, int size, bool little) { for (int i = 0; i < size; i++) bytes[offset + i] = (byte)(value >> (8 * (little ? i : size - 1 - i))); }
}
