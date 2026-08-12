using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceText {
    private static readonly byte[] BeginDelimiter = Encoding.ASCII.GetBytes("-----BEGIN C2PA MANIFEST-----");
    private static readonly byte[] EndDelimiter = Encoding.ASCII.GetBytes("-----END C2PA MANIFEST-----");
    private static readonly byte[] WrapperMagic = Encoding.ASCII.GetBytes("C2PATXT\0");
    private static readonly byte[] DataUriPrefix = Encoding.ASCII.GetBytes("data:application/c2pa;base64,");

    internal static bool HasStructuredDelimiter(byte[] data) => IndexOf(data, BeginDelimiter, 0) >= 0;

    internal static bool HasUnstructuredWrapperPrefix(byte[] data) {
        int offset = 0;
        while (offset < data.Length) {
            if (!TryReadCodePoint(data, offset, out int codePoint, out int prefixBytes) || codePoint != 0xFEFF) { offset++; continue; }
            int cursor = offset + prefixBytes;
            bool matches = true;
            for (int index = 0; index < WrapperMagic.Length; index++) {
                if (!TryReadCodePoint(data, cursor, out int selector, out int selectorBytes) ||
                    !TrySelectorToByte(selector, out byte value) || value != WrapperMagic[index]) {
                    matches = false;
                    break;
                }
                cursor += selectorBytes;
            }
            if (matches) return true;
            offset += prefixBytes;
        }
        return false;
    }

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        int structuredCount = 0;
        foreach (StructuredBlock block in FindStructuredBlocks(data, options.MaxManifestBytes)) {
            context.Add(new OfficeProvenanceEvidence(
                block.IsExternal ? OfficeProvenanceCarrierKind.C2paExternalManifest : OfficeProvenanceCarrierKind.C2paManifest,
                $"Text/C2PA@{block.Start}",
                block.IsValid,
                block.ManifestLength,
                block.ExternalUri));
            structuredCount++;
        }
        int wrapperCount = 0;
        foreach (TextWrapper wrapper in FindWrappers(data, options.MaxManifestBytes, includeInvalid: true)) {
            context.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"Text/C2PATextManifestWrapper@{wrapper.Start}",
                wrapper.IsValid,
                wrapper.ManifestLength));
            wrapperCount++;
        }
        if (structuredCount > 1) context.Diagnostics.Add("The structured text contains multiple C2PA manifest blocks.");
        if (wrapperCount > 1) context.Diagnostics.Add("The unstructured text contains multiple C2PA manifest wrappers.");
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes) {
        if (!options.RemoveC2paManifests && !options.RemoveExternalC2paReferences) return (byte[])data.Clone();
        var ranges = new List<RemovalRange>();
        foreach (StructuredBlock block in FindStructuredBlocks(data, options.Limits.MaxManifestBytes)) {
            bool requested = block.IsExternal ? options.RemoveExternalC2paReferences : options.RemoveC2paManifests;
            if (!requested || (!block.IsValid && options.RequireStructurallyValidCarrier)) continue;
            ranges.Add(new RemovalRange(block.LineStart, block.LineEnd - block.LineStart));
            changes.Add(new OfficeProvenanceChange(
                block.IsExternal ? OfficeProvenanceCarrierKind.C2paExternalManifest : OfficeProvenanceCarrierKind.C2paManifest,
                $"Text/C2PA@{block.Start}",
                block.LineEnd - block.LineStart));
        }
        if (options.RemoveC2paManifests) {
            foreach (TextWrapper wrapper in FindWrappers(data, options.Limits.MaxManifestBytes, includeInvalid: true)) {
                if (!wrapper.IsValid && options.RequireStructurallyValidCarrier) continue;
                ranges.Add(new RemovalRange(wrapper.Start, wrapper.End - wrapper.Start));
                changes.Add(new OfficeProvenanceChange(
                    OfficeProvenanceCarrierKind.C2paManifest,
                    $"Text/C2PATextManifestWrapper@{wrapper.Start}",
                    wrapper.End - wrapper.Start));
            }
        }
        if (ranges.Count == 0) return (byte[])data.Clone();
        ranges.Sort((left, right) => left.Start.CompareTo(right.Start));
        using var output = new MemoryStream(data.Length);
        int offset = 0;
        foreach (RemovalRange range in ranges) {
            if (range.Start < offset) continue;
            output.Write(data, offset, range.Start - offset);
            offset = range.Start + range.Length;
        }
        output.Write(data, offset, data.Length - offset);
        return output.ToArray();
    }

    private static IEnumerable<StructuredBlock> FindStructuredBlocks(byte[] data, long maximumManifestBytes) {
        int search = 0;
        while (search < data.Length) {
            int begin = IndexOf(data, BeginDelimiter, search);
            if (begin < 0) yield break;
            if (!IsStandaloneDelimiter(data, begin, BeginDelimiter.Length)) {
                search = begin + BeginDelimiter.Length;
                continue;
            }
            int contentStart = begin + BeginDelimiter.Length;
            int end = IndexOf(data, EndDelimiter, contentStart);
            if (end < 0) yield break;
            int newerBegin = FindStandaloneDelimiter(data, BeginDelimiter, contentStart);
            if (newerBegin >= 0 && newerBegin < end) {
                search = newerBegin;
                continue;
            }
            if (!IsStandaloneDelimiter(data, end, EndDelimiter.Length)) {
                search = end + EndDelimiter.Length;
                continue;
            }
            int contentEnd = end;
            while (contentStart < contentEnd && IsAsciiWhitespace(data[contentStart])) contentStart++;
            while (contentEnd > contentStart && IsAsciiWhitespace(data[contentEnd - 1])) contentEnd--;
            int valueLength = contentEnd - contentStart;
            bool external = false;
            bool valid = false;
            long manifestLength = 0;
            string? externalUri = null;
            if (StartsWith(data, contentStart, valueLength, DataUriPrefix)) {
                int base64Offset = contentStart + DataUriPrefix.Length;
                int base64Length = contentEnd - base64Offset;
                if (base64Length <= maximumManifestBytes * 2L) {
                    try {
                        string encoded = Encoding.ASCII.GetString(data, base64Offset, base64Length);
                        byte[] manifest = Convert.FromBase64String(encoded);
                        manifestLength = manifest.Length;
                        valid = manifest.LongLength <= maximumManifestBytes &&
                            OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, maximumManifestBytes, out _);
                    } catch (FormatException) {
                        valid = false;
                    }
                }
            } else if (valueLength > 0) {
                string value = Encoding.UTF8.GetString(data, contentStart, valueLength).Trim();
                if (Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)) {
                    external = true;
                    valid = true;
                    externalUri = uri.AbsoluteUri;
                }
            }
            int lineStart = FindLineStart(data, begin);
            int lineEnd = FindLineEndIncludingTerminator(data, end + EndDelimiter.Length);
            yield return new StructuredBlock(begin, lineStart, lineEnd, external, valid, manifestLength, externalUri);
            search = end + EndDelimiter.Length;
        }
    }

    private static IEnumerable<TextWrapper> FindWrappers(byte[] data, long maximumManifestBytes, bool includeInvalid) {
        int offset = 0;
        while (offset < data.Length) {
            if (!TryReadCodePoint(data, offset, out int codePoint, out int prefixBytes) || codePoint != 0xFEFF) { offset++; continue; }
            int selectorOffset = offset + prefixBytes;
            var decoded = new List<byte>();
            int cursor = selectorOffset;
            while (TryReadCodePoint(data, cursor, out int selector, out int selectorBytes) && TrySelectorToByte(selector, out byte value)) {
                decoded.Add(value);
                cursor += selectorBytes;
                if (decoded.Count > maximumManifestBytes + 13L || decoded.Count > int.MaxValue) break;
            }
            bool hasMagic = decoded.Count >= WrapperMagic.Length;
            for (int index = 0; hasMagic && index < WrapperMagic.Length; index++) hasMagic = decoded[index] == WrapperMagic[index];
            if (!hasMagic) { offset += prefixBytes; continue; }
            bool valid = false;
            long manifestLength = 0;
            if (decoded.Count >= 13 && decoded[8] == 1) {
                uint declared = ((uint)decoded[9] << 24) | ((uint)decoded[10] << 16) | ((uint)decoded[11] << 8) | decoded[12];
                manifestLength = declared;
                if (declared <= maximumManifestBytes && decoded.Count == 13L + declared) {
                    byte[] manifest = decoded.GetRange(13, (int)declared).ToArray();
                    valid = OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, maximumManifestBytes, out _);
                }
            }
            if (valid || includeInvalid) yield return new TextWrapper(offset, cursor, valid, manifestLength);
            offset = Math.Max(cursor, offset + prefixBytes);
        }
    }

    private static bool TryReadCodePoint(byte[] data, int offset, out int codePoint, out int bytes) {
        codePoint = 0; bytes = 0;
        if (offset < 0 || offset >= data.Length) return false;
        byte first = data[offset];
        if (first < 0x80) { codePoint = first; bytes = 1; return true; }
        int count; int value;
        if ((first & 0xE0) == 0xC0) { count = 2; value = first & 0x1F; }
        else if ((first & 0xF0) == 0xE0) { count = 3; value = first & 0x0F; }
        else if ((first & 0xF8) == 0xF0) { count = 4; value = first & 0x07; }
        else return false;
        if (count > data.Length - offset) return false;
        for (int index = 1; index < count; index++) {
            byte continuation = data[offset + index];
            if ((continuation & 0xC0) != 0x80) return false;
            value = (value << 6) | (continuation & 0x3F);
        }
        if ((count == 2 && value < 0x80) || (count == 3 && value < 0x800) || (count == 4 && value < 0x10000) ||
            value > 0x10FFFF || (value >= 0xD800 && value <= 0xDFFF)) return false;
        codePoint = value; bytes = count; return true;
    }

    private static bool TrySelectorToByte(int codePoint, out byte value) {
        if (codePoint >= 0xFE00 && codePoint <= 0xFE0F) { value = (byte)(codePoint - 0xFE00); return true; }
        if (codePoint >= 0xE0100 && codePoint <= 0xE01EF) { value = (byte)(codePoint - 0xE0100 + 16); return true; }
        value = 0; return false;
    }

    private static int IndexOf(byte[] data, byte[] pattern, int start) {
        for (int offset = Math.Max(0, start); offset <= data.Length - pattern.Length; offset++) {
            int index = 0;
            while (index < pattern.Length && data[offset + index] == pattern[index]) index++;
            if (index == pattern.Length) return offset;
        }
        return -1;
    }

    private static int FindStandaloneDelimiter(byte[] data, byte[] delimiter, int start) {
        int search = start;
        while (search < data.Length) {
            int offset = IndexOf(data, delimiter, search);
            if (offset < 0) return -1;
            if (IsStandaloneDelimiter(data, offset, delimiter.Length)) return offset;
            search = offset + delimiter.Length;
        }
        return -1;
    }

    private static bool StartsWith(byte[] data, int offset, int available, byte[] pattern) {
        if (pattern.Length > available || offset < 0 || offset > data.Length - pattern.Length) return false;
        for (int index = 0; index < pattern.Length; index++) if (data[offset + index] != pattern[index]) return false;
        return true;
    }

    private static bool IsAsciiWhitespace(byte value) => value is 0x09 or 0x0A or 0x0D or 0x20;
    private static int FindLineStart(byte[] data, int offset) {
        while (offset > 0 && data[offset - 1] != 0x0A && data[offset - 1] != 0x0D) offset--;
        return offset;
    }
    private static int FindLineEndIncludingTerminator(byte[] data, int offset) {
        while (offset < data.Length && data[offset] != 0x0A && data[offset] != 0x0D) offset++;
        if (offset >= data.Length) return offset;
        if (data[offset] == 0x0D && offset + 1 < data.Length && data[offset + 1] == 0x0A) return offset + 2;
        return offset + 1;
    }
    private static bool IsStandaloneDelimiter(byte[] data, int offset, int length) {
        int lineStart = FindLineStart(data, offset);
        for (int index = lineStart; index < offset; index++) if (!IsHorizontalWhitespace(data[index])) return false;
        int cursor = offset + length;
        while (cursor < data.Length && data[cursor] != 0x0A && data[cursor] != 0x0D) {
            if (!IsHorizontalWhitespace(data[cursor])) return false;
            cursor++;
        }
        return true;
    }
    private static bool IsHorizontalWhitespace(byte value) => value is 0x09 or 0x20;

    private sealed class StructuredBlock {
        internal StructuredBlock(int start, int lineStart, int lineEnd, bool isExternal, bool isValid, long manifestLength, string? externalUri) {
            Start = start; LineStart = lineStart; LineEnd = lineEnd; IsExternal = isExternal; IsValid = isValid;
            ManifestLength = manifestLength; ExternalUri = externalUri;
        }
        internal int Start { get; }
        internal int LineStart { get; }
        internal int LineEnd { get; }
        internal bool IsExternal { get; }
        internal bool IsValid { get; }
        internal long ManifestLength { get; }
        internal string? ExternalUri { get; }
    }
    private sealed class TextWrapper {
        internal TextWrapper(int start, int end, bool isValid, long manifestLength) { Start = start; End = end; IsValid = isValid; ManifestLength = manifestLength; }
        internal int Start { get; }
        internal int End { get; }
        internal bool IsValid { get; }
        internal long ManifestLength { get; }
    }
    private readonly struct RemovalRange {
        internal RemovalRange(int start, int length) { Start = start; Length = length; }
        internal int Start { get; }
        internal int Length { get; }
    }
}
