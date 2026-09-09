namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static void ExpandObjectStreams(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        HashSet<int>? allowedObjectStreamNumbers,
        PdfReadLimits limits,
        PdfDecodedStreamBudget decodedStreamBudget,
        Action<int> reportUnreadable) {
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
            var data = decodedStreamBudget.Decode(s, map);
            if (!TryReadObjectStreamLayout(s.Dictionary, data.Length, limits, out int n, out int first)) {
                reportUnreadable(id);
                continue;
            }
            // Header: pairs of objectNumber and offset (ASCII)
            var headerBytes = new byte[first];
            Buffer.BlockCopy(data, 0, headerBytes, 0, first);
            string header = PdfEncoding.Latin1GetString(headerBytes);
            var pairs = ParsePairs(header, n, out bool completeHeader);
            if (!completeHeader) { reportUnreadable(id); continue; }
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
                if (start < first || end > data.Length || end <= start || objNum <= 0) {
                    reportUnreadable(id);
                    continue;
                }
                int len = end - start;
                var sliceBytes = new byte[len];
                Buffer.BlockCopy(data, start, sliceBytes, 0, len);
                var slice = PdfEncoding.Latin1GetString(sliceBytes);
                var parsed = ParseTopLevelObject(
                    slice,
                    limits,
                    trackEncodedStringSourceSpans: false);
                if (parsed is not null) {
                    if (parsed.HasIncompleteSyntax) reportUnreadable(objNum);
                    map[objNum] = new PdfIndirectObject(objNum, 0, parsed);
                    effectiveOffsets[objNum] = objectStreamOffset;
                } else reportUnreadable(objNum);
            }
        }

        int GetSourceOffset(int objectNumber) => parsedOffsets.TryGetValue(objectNumber, out int offset) ? offset : int.MaxValue;
    }

    private static bool TryReadObjectStreamLayout(
        PdfDictionary dictionary,
        int decodedLength,
        PdfReadLimits limits,
        out int objectCount,
        out int firstObjectOffset) {
        objectCount = 0;
        firstObjectOffset = 0;

        if (dictionary.Get<PdfNumber>("N") is not PdfNumber countValue ||
            dictionary.Get<PdfNumber>("First") is not PdfNumber firstValue) return false;
        double rawObjectCount = countValue.Value;
        double rawFirstOffset = firstValue.Value;
        if (double.IsNaN(rawObjectCount) || double.IsInfinity(rawObjectCount)
            || rawObjectCount < 0D || rawObjectCount != Math.Truncate(rawObjectCount)) {
            return false;
        }
        if (rawObjectCount > limits.MaxIndirectObjects) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.IndirectObjects,
                limits.MaxIndirectObjects,
                rawObjectCount > long.MaxValue ? long.MaxValue : (long)rawObjectCount);
        }
        if (double.IsNaN(rawFirstOffset) || double.IsInfinity(rawFirstOffset)
            || rawFirstOffset < 0D || rawFirstOffset > decodedLength
            || rawFirstOffset != Math.Truncate(rawFirstOffset)) {
            return false;
        }

        objectCount = (int)rawObjectCount;
        firstObjectOffset = (int)rawFirstOffset;
        if (objectCount == 0) return decodedLength == firstObjectOffset;
        if (firstObjectOffset == 0) return false;
        return true;
    }

    private static List<(int Obj, int Off)> ParsePairs(string header, int n, out bool complete) {
        var list = new List<(int, int)>(n);
        int i = 0; int count = 0;
        var identifiers = new HashSet<int>();
        while (i < header.Length && count < n) {
            SkipWs();
            if (!ReadInt(out int obj)) break;
            SkipWs();
            if (!ReadInt(out int off)) break;
            if (obj <= 0 || off < 0 || !identifiers.Add(obj)) break;
            if (list.Count == 0 && off != 0) break;
            if (list.Count > 0 && off <= list[list.Count - 1].Item2) break;
            list.Add((obj, off)); count++;
        }
        SkipWs();
        complete = list.Count == n && i == header.Length;
        return list;

        void SkipWs() { i = SkipWhitespaceAndComments(header, i, header.Length); }
        bool ReadInt(out int val) {
            int start = i;
            if (i < header.Length && (header[i] == '-' || header[i] == '+')) i++;
            while (i < header.Length && header[i] >= '0' && header[i] <= '9') i++;
#if NET6_0_OR_GREATER
            return int.TryParse(header.AsSpan(start, i - start), System.Globalization.NumberStyles.AllowLeadingSign,
#else
            return int.TryParse(header.Substring(start, i - start), System.Globalization.NumberStyles.AllowLeadingSign,
#endif
                System.Globalization.CultureInfo.InvariantCulture, out val);
        }
    }
}
