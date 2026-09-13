using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    /// <summary>Replaces selected string properties while retaining every other raw section and property value.</summary>
    internal static class OfficeOlePropertySetEditor {
        internal static byte[] RewriteStrings(byte[]? source, Guid sectionId, IReadOnlyDictionary<uint, string?> replacements, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            var sections = new List<(Guid FormatId, byte[] Section)>(); bool found = false;
            if (source != null) {
                _ = OfficeOlePropertySetReader.ReadSections(source, token);
                int count = BitConverter.ToInt32(source, 24);
                for (int index = 0; index < count; index++) {
                    token.ThrowIfCancellationRequested(); int slot = 28 + index * 20;
                    var guidBytes = new byte[16]; Buffer.BlockCopy(source, slot, guidBytes, 0, 16); var id = new Guid(guidBytes);
                    int offset = BitConverter.ToInt32(source, slot + 16); int size = BitConverter.ToInt32(source, offset);
                    var section = new byte[size]; Buffer.BlockCopy(source, offset, section, 0, size);
                    if (id == sectionId) { section = Replace(section, replacements, token); found = true; }
                    sections.Add((id, section));
                }
            }
            if (!found) sections.Add((sectionId, Replace(null, replacements, token)));
            byte[] result = OfficeOlePropertySetWriter.CreatePropertySet(sections.ToArray());
            if (source != null) Buffer.BlockCopy(source, 0, result, 0, 24);
            return result;
        }
        private static byte[] Replace(byte[]? section, IReadOnlyDictionary<uint, string?> replacements, CancellationToken token) {
            var values = new Dictionary<uint, byte[]>();
            if (section != null) {
                int count = BitConverter.ToInt32(section, 4);
                var entries = new List<(uint Id, int Offset)>();
                for (int i = 0; i < count; i++) entries.Add((BitConverter.ToUInt32(section, 8 + i * 8), BitConverter.ToInt32(section, 12 + i * 8)));
                entries.Sort((a, b) => a.Offset.CompareTo(b.Offset));
                for (int i = 0; i < entries.Count; i++) {
                    token.ThrowIfCancellationRequested(); int end = i + 1 == entries.Count ? section.Length : entries[i + 1].Offset;
                    var bytes = new byte[end - entries[i].Offset]; Buffer.BlockCopy(section, entries[i].Offset, bytes, 0, bytes.Length);
                    values.Add(entries[i].Id, bytes);
                }
            } else values.Add(1, OfficeOleProperty.Integer(1, (short)1200).ValueBytes);
            foreach (var pair in replacements) {
                if (pair.Key < 2) throw new ArgumentException("Dictionary and code-page properties cannot be replaced as strings.", nameof(replacements));
                if (pair.Value == null) values.Remove(pair.Key); else values[pair.Key] = OfficeOleProperty.String(pair.Key, pair.Value).ValueBytes;
            }
            using var buffer = new MemoryStream(); using var writer = new BinaryWriter(buffer);
            int length = checked(8 + values.Count * 8 + values.Values.Sum(v => checked((v.Length + 3) & ~3)));
            writer.Write(length); writer.Write(values.Count); int offset = 8 + values.Count * 8;
            foreach (var pair in values) { writer.Write(pair.Key); writer.Write(offset); offset = checked(offset + ((pair.Value.Length + 3) & ~3)); }
            foreach (var pair in values) { token.ThrowIfCancellationRequested(); writer.Write(pair.Value); while ((buffer.Position & 3) != 0) writer.Write((byte)0); }
            return buffer.ToArray();
        }
    }
}
