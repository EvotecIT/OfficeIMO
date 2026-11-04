using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Bounded metadata access shared by substitution and font line-metric selection.</summary>
internal static class OfficeOpenTypeLayoutFeatures {
    private const int MaximumFeatures = 4096;

    internal static IReadOnlyList<string>? GetLatinDefaultTags(OfficeOpenTypeReader reader, string tableTag) {
        if (!reader.TryGetTable(tableTag, out int table, out int length)) return Array.Empty<string>();
        int[]? indexes = GetLatinDefaultIndexes(reader, table, length, out _);
        if (indexes == null) return null;
        int features = table + reader.ReadUInt16(table + 6);
        var tags = new List<string>(indexes.Length);
        foreach (int index in indexes) {
            int record = features + 2 + index * 6;
            tags.Add(Tag(reader, record));
        }
        return tags;
    }

    private static string Tag(OfficeOpenTypeReader reader, int offset) {
        int first = reader.ReadUInt16(offset), second = reader.ReadUInt16(offset + 2);
        return new string(new[] { (char)(first >> 8), (char)(first & 255), (char)(second >> 8), (char)(second & 255) });
    }

    internal static int[]? GetLatinDefaultIndexes(OfficeOpenTypeReader reader, int table, int length, out int requiredFeature) {
        requiredFeature = -1;
        try {
            int end = checked(table + length);
            void Ensure(int offset, int count) {
                if (offset < table || count < 0 || offset > end - count) throw new InvalidDataException("Invalid OpenType layout metadata.");
            }
            int Relative(int basis, int delta, int count) { int offset = checked(basis + delta); Ensure(offset, count); return offset; }
            Ensure(table, 10);
            int scripts = Relative(table, reader.ReadUInt16(table + 4), 2);
            int count = reader.ReadUInt16(scripts);
            if (count > MaximumFeatures) return null;
            Ensure(scripts + 2, checked(count * 6));
            int script = 0;
            for (int index = 0; index < count; index++) {
                int record = scripts + 2 + index * 6;
                string tag = Tag(reader, record);
                bool latin = tag == "latn", fallback = tag == "DFLT";
                if (latin || fallback && script == 0) {
                    script = Relative(scripts, reader.ReadUInt16(record + 4), 4);
                    if (latin) break;
                }
            }
            if (script == 0 || reader.ReadUInt16(script) == 0) return Array.Empty<int>();
            int language = Relative(script, reader.ReadUInt16(script), 6);
            int featureCount = reader.ReadUInt16(language + 4);
            if (featureCount > MaximumFeatures) return null;
            Ensure(language + 6, checked(featureCount * 2));
            int features = Relative(table, reader.ReadUInt16(table + 6), 2);
            int total = reader.ReadUInt16(features);
            if (total > MaximumFeatures) return null;
            Ensure(features + 2, checked(total * 6));
            var selected = new SortedSet<int>();
            int required = reader.ReadUInt16(language + 2);
            if (required != ushort.MaxValue) { selected.Add(required); requiredFeature = required; }
            for (int index = 0; index < featureCount; index++) selected.Add(reader.ReadUInt16(language + 6 + index * 2));
            var result = new int[selected.Count];
            int output = 0;
            foreach (int index in selected) { if (index >= total) return null; result[output++] = index; }
            return result;
        } catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return null; }
    }
}
