using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Edits scalar FOPT rows without re-encoding unrelated rows or their complex-data tail.</summary>
internal static class OfficeArtPropertyTableEditor {
    internal static byte[]? RewriteScalars(byte[] source, int offset, int length, ushort count,
        IReadOnlyDictionary<ushort, uint> changes, out ushort resultCount) {
        resultCount = count;
        if (offset < 0 || length < 0 || offset > source.Length - length || count > 4095 || count * 6 > length) return null;
        var properties = OfficeArtPropertyTableReader.Read(source, offset, length, count);
        if (properties.Count != count) return null;
        int complexOffset = offset + count * 6;
        foreach (OfficeArtProperty property in properties.Where(property => property.IsComplex)) {
            uint required = OfficeArtPropertyTableReader.ExpectedComplexDataLength(property.PropertyId, property.Value,
                source, complexOffset, offset + length - complexOffset);
            if (required > offset + length - complexOffset || property.AvailableComplexDataLength != required) return null;
            complexOffset += checked((int)required);
        }
        foreach (ushort id in changes.Keys) {
            var matches = properties.Where(property => property.PropertyId == id).ToArray();
            if (id > 0x3FFF || matches.Length > 1 || matches.Any(property => property.IsComplex || property.IsBlipId)) return null;
        }
        ushort[] missing = changes.Keys.Where(id => properties.All(property => property.PropertyId != id)).OrderBy(id => id).ToArray();
        // Growth must keep the native scalar ordering without moving existing
        // complex rows relative to the opaque tail they address.
        if (missing.Length > 0 && properties.Zip(properties.Skip(1), (a, b) => a.PropertyId <= b.PropertyId).Any(sorted => !sorted)) return null;
        if (count + missing.Length > 4095) return null;
        resultCount = checked((ushort)(count + missing.Length));
        var output = new byte[checked(length + missing.Length * 6)];
        int position = 0, next = 0;
        void Write(ushort operation, uint value) {
            output[position++] = (byte)operation; output[position++] = (byte)(operation >> 8);
            for (int index = 0; index < 4; index++) output[position++] = (byte)(value >> (index * 8));
        }
        foreach (OfficeArtProperty property in properties) {
            while (next < missing.Length && missing[next] < property.PropertyId) {
                ushort id = missing[next++]; Write(id, changes[id]);
            }
            Write(property.RawOperationId, changes.TryGetValue(property.PropertyId, out uint value) ? value : property.Value);
        }
        while (next < missing.Length) { ushort id = missing[next++]; Write(id, changes[id]); }
        Buffer.BlockCopy(source, offset + count * 6, output, position, length - count * 6);
        return output;
    }
}
