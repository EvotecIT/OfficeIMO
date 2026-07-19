namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static void ExpandObjectStreams(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        HashSet<int>? allowedObjectStreamNumbers,
        PdfReadLimits limits) {
        // Snapshot keys to avoid modifying during enumeration
        var keys = new List<int>(map.Keys);
        keys.Sort((left, right) => GetSourceOffset(left).CompareTo(GetSourceOffset(right)));
        var effectiveOffsets = new Dictionary<int, int>(parsedOffsets);
        foreach (var id in keys) {
            if (allowedObjectStreamNumbers is not null &&
                !allowedObjectStreamNumbers.Contains(id)) {
                continue;
            }

            if (!map.TryGetValue(id, out var ind)) continue;
            if (ind.Value is not PdfStream s) continue;
            var type = s.Dictionary.Get<PdfName>("Type")?.Name;
            if (!string.Equals(type, "ObjStm", StringComparison.Ordinal)) continue;
            int objectStreamOffset = GetSourceOffset(id);

            // Decode object stream bytes (flate only for now)
            var data = Filters.StreamDecoder.Decode(s.Dictionary, s.Data, map, limits.MaxDecodedStreamBytes);
            int n = (int)(s.Dictionary.Get<PdfNumber>("N")?.Value ?? 0);
            int first = (int)(s.Dictionary.Get<PdfNumber>("First")?.Value ?? 0);
            if (n <= 0 || first <= 0 || first > data.Length) continue;
            // Header: pairs of objectNumber and offset (ASCII)
            var headerBytes = new byte[first];
            Buffer.BlockCopy(data, 0, headerBytes, 0, first);
            string header = PdfEncoding.Latin1GetString(headerBytes);
            var pairs = ParsePairs(header, n);
            if (pairs.Count != n) continue;
            for (int i = 0; i < n; i++) {
                int objNum = pairs[i].Obj;
                int off = pairs[i].Off;
                if (map.ContainsKey(objNum) &&
                    effectiveOffsets.TryGetValue(objNum, out int currentOffset) &&
                    currentOffset > objectStreamOffset) {
                    continue;
                }

                int start = first + off;
                int end = (i + 1 < n) ? first + pairs[i + 1].Off : data.Length;
                if (start < 0 || end > data.Length || end <= start) continue;
                int len = end - start;
                var sliceBytes = new byte[len];
                Buffer.BlockCopy(data, start, sliceBytes, 0, len);
                var slice = PdfEncoding.Latin1GetString(sliceBytes);
                var parsed = ParseTopLevelObject(
                    slice,
                    limits,
                    trackEncodedStringSourceSpans: false);
                if (parsed is not null) {
                    map[objNum] = new PdfIndirectObject(objNum, 0, parsed);
                    effectiveOffsets[objNum] = objectStreamOffset;
                }
            }
        }

        int GetSourceOffset(int objectNumber) => parsedOffsets.TryGetValue(objectNumber, out int offset) ? offset : int.MaxValue;
    }

    private static List<(int Obj, int Off)> ParsePairs(string header, int n) {
        var list = new List<(int, int)>(n);
        int i = 0; int count = 0;
        while (i < header.Length && count < n) {
            SkipWs();
            if (!ReadInt(out int obj)) break;
            SkipWs();
            if (!ReadInt(out int off)) break;
            list.Add((obj, off)); count++;
        }
        return list;

        void SkipWs() { while (i < header.Length && char.IsWhiteSpace(header[i])) i++; }
        bool ReadInt(out int val) {
            int sign = 1; if (i < header.Length && header[i] == '-') { sign = -1; i++; }
            int start = i; long v = 0; bool any = false;
            while (i < header.Length && char.IsDigit(header[i])) { v = v * 10 + (header[i] - '0'); i++; any = true; if (i - start > 10) break; }
            val = any ? (int)(v * sign) : 0; return any;
        }
    }
}
