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

    private sealed class ResourceValidationState {
        internal Dictionary<(int Start, int End), bool> Spans { get; } = new();
        internal Dictionary<(long Object, long Generation), bool> Objects { get; } = new();
        internal HashSet<(long Object, long Generation)> Visiting { get; } = new();
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

    private static bool HasCompleteResourceSpan(byte[] bytes, int start, int end,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets,
        int[] orderedObjectOffsets, int limit, ResourceValidationState state) {
        if (state.Spans.TryGetValue((start, end), out bool cached)) return cached;
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
            if (objectNumber <= 0 || !IsCompleteResourceObject(bytes, objectNumber,
                    generation, inUseOffsets, orderedObjectOffsets, limit, state)) {
                state.Spans[(start, end)] = false;
                return false;
            }
            offset = referenceOffset - 1;
        }
        bool complete = literalDepth == 0 && !inHexString;
        state.Spans[(start, end)] = complete;
        return complete;
    }

    private static bool IsCompleteResourceObject(byte[] bytes, long objectNumber,
        long generation,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets,
        int[] orderedObjectOffsets, int limit, ResourceValidationState state) {
        var identity = (objectNumber, generation);
        if (state.Objects.TryGetValue(identity, out bool cached)) return cached;
        if (!inUseOffsets.TryGetValue(identity, out int objectOffset)
            || !TryGetResourceObjectSpan(bytes, objectOffset, limit, objectNumber,
                generation, inUseOffsets, orderedObjectOffsets,
                out int bodyStart, out int bodyEnd)) {
            state.Objects[identity] = false;
            return false;
        }
        if (!state.Visiting.Add(identity)) return true;
        bool complete = HasCompleteResourceSpan(bytes, bodyStart, bodyEnd,
            inUseOffsets, orderedObjectOffsets, limit, state);
        state.Visiting.Remove(identity);
        state.Objects[identity] = complete;
        return complete;
    }

    private static bool TryGetResourceObjectSpan(byte[] bytes, int offset, int limit,
        long expectedObject, long expectedGeneration,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets,
        int[] orderedObjectOffsets, out int bodyStart, out int bodyEnd) {
        bodyStart = -1;
        bodyEnd = -1;
        int objectOffset = offset;
        int objectLimit = GetObjectLimit(orderedObjectOffsets, offset, limit);
        SkipWhitespace(bytes, ref offset, objectLimit);
        if (!TryReadDecimal(bytes, ref offset, objectLimit, out long objectNumber)
            || objectNumber != expectedObject) return false;
        SkipWhitespace(bytes, ref offset, objectLimit);
        if (!TryReadDecimal(bytes, ref offset, objectLimit, out long generation)
            || generation != expectedGeneration) return false;
        SkipWhitespace(bytes, ref offset, objectLimit);
        if (!StartsWith(bytes, offset, "obj")
            || offset + 3 < objectLimit && !IsDelimiter(bytes[offset + 3])) return false;
        offset += 3;
        if (!SkipWhitespaceAndComments(bytes, ref offset, objectLimit)) return false;
        bodyStart = offset;

        if (StartsWith(bytes, offset, "<<")) {
            bodyEnd = FindDictionaryEnd(bytes, offset,
                Math.Min(objectLimit, offset + 65536));
            if (bodyEnd < 0) return false;
            bodyEnd += 2;
            int trailing = bodyEnd;
            if (!SkipWhitespaceAndComments(bytes, ref trailing, objectLimit)) return false;
            if (StartsWith(bytes, trailing, "stream")) {
                return IsCompleteStreamObject(bytes, inUseOffsets, orderedObjectOffsets,
                    objectOffset, limit, expectedObject, expectedGeneration);
            }
            return StartsWith(bytes, trailing, "endobj")
                && (trailing + 6 >= objectLimit || IsDelimiter(bytes[trailing + 6]));
        }

        if (offset < objectLimit && bytes[offset] == (byte)'[') {
            int arrayEnd = FindResourceArrayEnd(bytes, offset, objectLimit);
            if (arrayEnd < 0) return false;
            bodyEnd = arrayEnd + 1;
            int trailing = bodyEnd;
            return SkipWhitespaceAndComments(bytes, ref trailing, objectLimit)
                && StartsWith(bytes, trailing, "endobj")
                && (trailing + 6 >= objectLimit || IsDelimiter(bytes[trailing + 6]));
        }

        int endObject = IndexOf(bytes, "endobj", offset, objectLimit);
        if (endObject < 0 || !IsCompleteResourceScalar(bytes, offset, endObject)) return false;
        bodyEnd = endObject;
        return true;
    }

    private static bool IsCompleteResourceScalar(byte[] bytes, int start, int end) {
        int offset = start;
        if (!SkipWhitespaceAndComments(bytes, ref offset, end)) return false;
        if (TryReadPdfNumber(bytes, ref offset, end, out _)) {
            return SkipWhitespaceAndComments(bytes, ref offset, end) && offset == end;
        }
        foreach (string keyword in new[] { "true", "false", "null" }) {
            if (!StartsWith(bytes, offset, keyword)) continue;
            offset += keyword.Length;
            return SkipWhitespaceAndComments(bytes, ref offset, end) && offset == end;
        }
        if (offset < end && bytes[offset] == (byte)'/') {
            int tokenStart = ++offset;
            while (offset < end && !IsDelimiter(bytes[offset])) offset++;
            return offset > tokenStart
                && SkipWhitespaceAndComments(bytes, ref offset, end) && offset == end;
        }
        return false;
    }

    private static int FindResourceArrayEnd(byte[] bytes, int start, int limit) {
        int arrayDepth = 0;
        int dictionaryDepth = 0;
        int literalDepth = 0;
        bool escaped = false;
        bool inHexString = false;
        bool inComment = false;
        for (int offset = start; offset < limit; offset++) {
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
            if (current == (byte)'<' && offset + 1 < limit) {
                if (bytes[offset + 1] == (byte)'<') {
                    dictionaryDepth++;
                    offset++;
                } else inHexString = true;
                continue;
            }
            if (current == (byte)'>' && offset + 1 < limit
                && bytes[offset + 1] == (byte)'>') {
                if (dictionaryDepth == 0) return -1;
                dictionaryDepth--;
                offset++;
                continue;
            }
            if (dictionaryDepth > 0) continue;
            if (current == (byte)'[') arrayDepth++;
            else if (current == (byte)']') {
                arrayDepth--;
                if (arrayDepth == 0) return offset;
                if (arrayDepth < 0) return -1;
            }
        }
        return -1;
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
