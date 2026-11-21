namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static void ExpandObjectStreams(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        HashSet<int>? allowedObjectStreamNumbers,
        PdfReadLimits limits,
        PdfDecodedStreamBudget decodedStreamBudget,
        Action<int> reportUnreadable,
        System.Threading.CancellationToken cancellationToken) {
        // Most files without an xref stream have no object stream either; the sorted key and
        // offset snapshots below are only needed when there is one to expand.
        if (!HasObjectStreamCandidate(map, allowedObjectStreamNumbers, cancellationToken)) return;
        // Snapshot keys to avoid modifying during enumeration
        var keys = new List<int>(map.Count);
        foreach (int key in map.Keys) {
            cancellationToken.ThrowIfCancellationRequested();
            keys.Add(key);
        }
        try {
            keys.Sort((left, right) => {
                cancellationToken.ThrowIfCancellationRequested();
                return GetSourceOffset(left).CompareTo(GetSourceOffset(right));
            });
        } catch (InvalidOperationException error) when (error.InnerException is OperationCanceledException) {
            cancellationToken.ThrowIfCancellationRequested();
            throw;
        }
        var effectiveOffsets = new Dictionary<int, int>(parsedOffsets.Count);
        foreach (KeyValuePair<int, int> offset in parsedOffsets) {
            cancellationToken.ThrowIfCancellationRequested();
            effectiveOffsets[offset.Key] = offset.Value;
        }
        foreach (var id in keys) {
            cancellationToken.ThrowIfCancellationRequested();
            if (allowedObjectStreamNumbers is not null &&
                !allowedObjectStreamNumbers.Contains(id)) {
                continue;
            }

            if (!map.TryGetValue(id, out var ind) || IsUnbuiltWithOtherType(ind, "ObjStm")) continue;
            if (ind.Value is not PdfStream s) continue;
            var type = s.Dictionary.Get<PdfName>("Type")?.Name;
            if (!string.Equals(type, "ObjStm", StringComparison.Ordinal)) continue;
            int objectStreamOffset = GetSourceOffset(id);

            // Decode object stream bytes (flate only for now)
            var data = decodedStreamBudget.Decode(s, map, limits.MaxDecodedStreamBytes, cancellationToken);
            if (!TryReadObjectStreamLayout(s.Dictionary, data.Length, limits, out int n, out int first)) {
                reportUnreadable(id);
                continue;
            }
            // Header: pairs of objectNumber and offset (ASCII)
            var pairs = ParsePairs(data, first, n, out bool completeHeader, cancellationToken);
            if (!completeHeader) { reportUnreadable(id); continue; }
            for (int i = 0; i < n; i++) {
                cancellationToken.ThrowIfCancellationRequested();
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
                var parsed = ParseTopLevelObject(
                    data,
                    start,
                    end - start,
                    limits,
                    trackEncodedStringSourceSpans: false,
                    cancellationToken);
                if (parsed is not null) {
                    if (parsed.HasIncompleteSyntax) reportUnreadable(objNum);
                    map[objNum] = new PdfIndirectObject(objNum, 0, parsed) { SourceOffset = objectStreamOffset };
                    effectiveOffsets[objNum] = objectStreamOffset;
                } else reportUnreadable(objNum);
            }
        }

        int GetSourceOffset(int objectNumber) => parsedOffsets.TryGetValue(objectNumber, out int offset) ? offset : int.MaxValue;
    }

    // The same test the expansion loop applies before it decodes anything.
    private static bool HasObjectStreamCandidate(
        Dictionary<int, PdfIndirectObject> map,
        HashSet<int>? allowedObjectStreamNumbers,
        System.Threading.CancellationToken cancellationToken) {
        foreach (KeyValuePair<int, PdfIndirectObject> entry in map) {
            cancellationToken.ThrowIfCancellationRequested();
            if (allowedObjectStreamNumbers is not null && !allowedObjectStreamNumbers.Contains(entry.Key)) continue;
            if (IsUnbuiltWithOtherType(entry.Value, "ObjStm")) continue;
            if (entry.Value.Value is PdfStream stream &&
                string.Equals(stream.Dictionary.Get<PdfName>("Type")?.Name, "ObjStm", StringComparison.Ordinal)) return true;
        }
        return false;
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

    /// <summary>Reads the object-number and offset pairs from the first <paramref name="length"/> decoded bytes of an object stream.</summary>
    private static List<(int Obj, int Off)> ParsePairs(byte[] data, int length, int n, out bool complete,
        System.Threading.CancellationToken cancellationToken) {
        var header = new PdfByteContentText(data);
        var list = new List<(int, int)>(n);
        int i = 0; int count = 0;
        // Each pair needs at least two digits and a separator, so the header length bounds the set.
        var identifiers = PdfCollectionSizing.CreateHashSet<int>(n, length / 2 + 1);
        while (i < length && count < n) {
            cancellationToken.ThrowIfCancellationRequested();
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
        complete = list.Count == n && i == length;
        return list;

        void SkipWs() { i = SkipWhitespaceAndComments(header, i, length, cancellationToken); }
        bool ReadInt(out int val) {
            int start = i;
            if (i < length && (header[i] == '-' || header[i] == '+')) i++;
            while (i < length && header[i] >= '0' && header[i] <= '9') {
                if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                i++;
            }
            if (i - start > MaxNumericTokenCharacters) {
                val = default;
                return false;
            }
#if NET6_0_OR_GREATER
            Span<char> digits = stackalloc char[i - start];
            for (int index = 0; index < digits.Length; index++) digits[index] = header[start + index];
            return int.TryParse(digits, System.Globalization.NumberStyles.AllowLeadingSign,
#else
            return int.TryParse(header.Substring(start, i - start, cancellationToken), System.Globalization.NumberStyles.AllowLeadingSign,
#endif
                System.Globalization.CultureInfo.InvariantCulture, out val);
        }
    }
}
