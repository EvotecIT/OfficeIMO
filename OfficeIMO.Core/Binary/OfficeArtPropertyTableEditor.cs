using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Edits bounded FOPT rows while retaining unrelated operations, complex payloads and trailing bytes.</summary>
internal static class OfficeArtPropertyTableEditor {
    internal static byte[]? RewriteScalars(byte[] source, int offset, int length, ushort count,
        IReadOnlyDictionary<ushort, uint> changes, out ushort resultCount) => Rewrite(source, offset, length, count,
            changes.ToDictionary(pair => pair.Key, pair => new OfficeArtProperty(0, pair.Key, pair.Value)), out resultCount);

    /// <summary>Explicit replacements must retain an existing row's scalar/complex and BLIP-ID roles. Missing rows use the supplied roles.</summary>
    internal static byte[]? Rewrite(byte[] source, int offset, int length, ushort count,
        IReadOnlyDictionary<ushort, OfficeArtProperty> changes, out ushort resultCount) {
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
            OfficeArtProperty replacement = changes[id];
            if (id > 0x3FFF || replacement.PropertyId != id || matches.Length > 1
                || matches.Any(property => property.RawOperationId != replacement.RawOperationId)
                || replacement.IsComplex && !replacement.HasCompleteComplexData) return null;
        }
        ushort[] missing = changes.Keys.Where(id => properties.All(property => property.PropertyId != id)).OrderBy(id => id).ToArray();
        // Growth must keep the native scalar ordering without moving existing
        // complex rows relative to the opaque tail they address.
        if (missing.Length > 0 && properties.Zip(properties.Skip(1), (a, b) => a.PropertyId <= b.PropertyId).Any(sorted => !sorted)) return null;
        if (count + missing.Length > 4095) return null;
        resultCount = checked((ushort)(count + missing.Length));
        var rows = new List<OfficeArtProperty>(resultCount); int next = 0;
        foreach (OfficeArtProperty property in properties) {
            while (next < missing.Length && missing[next] < property.PropertyId) rows.Add(changes[missing[next++]]);
            rows.Add(changes.TryGetValue(property.PropertyId, out OfficeArtProperty? replacement) ? replacement : property);
        }
        while (next < missing.Length) rows.Add(changes[missing[next++]]);
        int suffixLength = offset + length - complexOffset;
        var output = new byte[checked(resultCount * 6 + rows.Where(row => row.IsComplex).Sum(row => row.AvailableComplexDataLength!.Value) + suffixLength)];
        int position = 0;
        void Write(ushort operation, uint value) {
            output[position++] = (byte)operation; output[position++] = (byte)(operation >> 8);
            for (int index = 0; index < 4; index++) output[position++] = (byte)(value >> (index * 8));
        }
        foreach (OfficeArtProperty property in rows) Write(property.RawOperationId, property.Value);
        foreach (OfficeArtProperty property in rows.Where(row => row.IsComplex)) {
            byte[] bytes = property.CopyComplexData() ?? Array.Empty<byte>();
            Buffer.BlockCopy(bytes, 0, output, position, bytes.Length); position += bytes.Length;
        }
        Buffer.BlockCopy(source, complexOffset, output, position, suffixLength);
        return output;
    }
}
