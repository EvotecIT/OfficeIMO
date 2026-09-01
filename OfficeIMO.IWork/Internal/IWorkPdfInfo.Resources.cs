namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkPdfInfo {
    private readonly struct ResourceDictionary {
        internal ResourceDictionary(int start, int end) {
            IsDeclared = true;
            IsValid = start >= 0 && end >= start;
            Start = start;
            End = end;
        }

        internal bool IsDeclared { get; }
        internal bool IsValid { get; }
        internal int Start { get; }
        internal int End { get; }
    }

    private static ResourceDictionary ResolveResources(byte[] bytes,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets,
        int[] orderedObjectOffsets, int dictionaryStart, int dictionaryEnd, int limit,
        ResourceDictionary inherited) {
        int resourcesOffset = FindDictionaryName(bytes, "/Resources", dictionaryStart,
            dictionaryEnd, out int resourcesCount);
        if (resourcesCount == 0) return inherited;
        if (resourcesCount != 1 || resourcesOffset < 0) return new ResourceDictionary(-1, -1);
        resourcesOffset += 10;
        if (!SkipWhitespaceAndComments(bytes, ref resourcesOffset, dictionaryEnd)) {
            return new ResourceDictionary(-1, -1);
        }

        int resourceDictionaryStart;
        int resourceDictionaryEnd;
        if (StartsWith(bytes, resourcesOffset, "<<")) {
            resourceDictionaryStart = resourcesOffset;
            resourceDictionaryEnd = FindDictionaryEnd(bytes, resourceDictionaryStart,
                dictionaryEnd);
            if (resourceDictionaryEnd < 0) return new ResourceDictionary(-1, -1);
        } else {
            if (!TryReadReference(bytes, ref resourcesOffset, dictionaryEnd,
                    out long resourceObject, out long resourceGeneration)
                || !inUseOffsets.TryGetValue((resourceObject, resourceGeneration),
                    out int resourceObjectOffset)) return new ResourceDictionary(-1, -1);
            int resourceObjectLimit = GetObjectLimit(orderedObjectOffsets,
                resourceObjectOffset, limit);
            if (!TryGetObjectDictionary(bytes, ref resourceObjectOffset, resourceObjectLimit,
                    resourceObject, resourceGeneration,
                    out resourceDictionaryStart, out resourceDictionaryEnd)) {
                return new ResourceDictionary(-1, -1);
            }
        }

        return new ResourceDictionary(resourceDictionaryStart, resourceDictionaryEnd);
    }

    private static bool HasResolvableDictionaryReferences(byte[] bytes, int start, int end,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets) {
        int literalDepth = 0;
        bool escaped = false;
        bool inHexString = false;
        bool inComment = false;
        for (int offset = start; offset < end; offset++) {
            byte current = bytes[offset];
            if (inComment) {
                if (current is 0x0a or 0x0d) inComment = false;
                continue;
            }
            if (literalDepth > 0) {
                if (escaped) escaped = false;
                else if (current == (byte)'\\') escaped = true;
                else if (current == (byte)'(') literalDepth++;
                else if (current == (byte)')') literalDepth--;
                continue;
            }
            if (inHexString) {
                if (current == (byte)'>') inHexString = false;
                continue;
            }
            if (current == (byte)'%') {
                inComment = true;
                continue;
            }
            if (current == (byte)'(') {
                literalDepth = 1;
                continue;
            }
            if (current == (byte)'<' && offset + 1 < end) {
                if (bytes[offset + 1] == (byte)'<') offset++;
                else inHexString = true;
                continue;
            }
            if (current < (byte)'0' || current > (byte)'9') continue;

            int referenceOffset = offset;
            if (!TryReadReference(bytes, ref referenceOffset, end,
                    out long objectNumber, out long generation)) continue;
            if (objectNumber <= 0
                || !inUseOffsets.ContainsKey((objectNumber, generation))) return false;
            offset = referenceOffset - 1;
        }
        return literalDepth == 0 && !inHexString;
    }

    private static bool TryReadReference(byte[] bytes, ref int offset, int limit,
        out long objectNumber, out long generation) {
        objectNumber = 0;
        generation = 0;
        int candidate = offset;
        if (!TryReadDecimal(bytes, ref candidate, limit, out objectNumber)
            || !SkipWhitespaceAndComments(bytes, ref candidate, limit)
            || !TryReadDecimal(bytes, ref candidate, limit, out generation)
            || generation > 65534
            || !SkipWhitespaceAndComments(bytes, ref candidate, limit)
            || candidate >= limit || bytes[candidate++] != (byte)'R'
            || candidate < limit && !IsDelimiter(bytes[candidate])) return false;
        offset = candidate;
        return true;
    }
}
