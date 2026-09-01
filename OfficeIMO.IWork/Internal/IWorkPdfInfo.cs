namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkPdfInfo {
    internal static bool IsComplete(byte[] bytes) {
        if (bytes.Length < 20 || !HasValidHeader(bytes)) return false;

        int eof = LastIndexOf(bytes, "%%EOF");
        if (eof < 0 || !ContainsOnlyTrailingWhitespace(bytes, eof + 5)) return false;
        int startXref = LastIndexOf(bytes, "startxref", eof);
        if (startXref < 0 || startXref + 9 >= eof
            || !IsWhitespace(bytes[startXref + 9])) return false;

        int offset = startXref + 9;
        SkipWhitespace(bytes, ref offset, eof);
        if (!TryReadDecimal(bytes, ref offset, eof, out long xrefOffset)
            || xrefOffset < 0 || xrefOffset >= startXref || xrefOffset > int.MaxValue) return false;

        int xref = (int)xrefOffset;
        SkipWhitespace(bytes, ref xref, startXref);
        if (StartsWith(bytes, xref, "xref")) return IsClassicXref(bytes, xref, startXref);
        // XRef streams may be filtered, use predictors, and reference compressed object
        // streams. Accepting them without decoding their entries would make separately
        // searchable object text look authoritative, so the bounded validator fails closed.
        return false;
    }

    private static bool IsClassicXref(byte[] bytes, int offset, int limit) {
        var inUseOffsets = new Dictionary<(long Object, long Generation), int>();
        var seenObjects = new HashSet<long>();
        var visitedXrefs = new HashSet<int>();
        long totalEntries = 0;
        long? size = null;
        long? rootObject = null;
        long rootGeneration = 0;
        int objectLimit = -1;
        int currentXref = offset;
        bool chainComplete = false;
        for (int depth = 0; depth < 128; depth++) {
            if (!visitedXrefs.Add(currentXref)
                || !TryReadClassicXref(bytes, currentXref, limit, inUseOffsets,
                    seenObjects, ref totalEntries, out int trailerOffset, out long sectionSize,
                    out bool hasRoot, out long sectionRootObject, out long sectionRootGeneration,
                    out bool hasPrevious, out long previousOffset)) return false;
            objectLimit = objectLimit < 0 ? trailerOffset : objectLimit;
            size ??= sectionSize;
            if (!rootObject.HasValue && hasRoot) {
                rootObject = sectionRootObject;
                rootGeneration = sectionRootGeneration;
            }
            if (!hasPrevious) {
                chainComplete = true;
                break;
            }
            if (previousOffset < 0 || previousOffset >= currentXref || previousOffset > int.MaxValue) return false;
            currentXref = (int)previousOffset;
            SkipWhitespace(bytes, ref currentXref, limit);
            if (!StartsWith(bytes, currentXref, "xref")) return false;
        }
        if (!chainComplete || !size.HasValue || !rootObject.HasValue
            || rootObject.Value <= 0 || rootObject.Value >= size.Value
            || seenObjects.Any(objectNumber => objectNumber >= size.Value)
            || !inUseOffsets.TryGetValue((rootObject.Value, rootGeneration), out int rootOffset)) return false;
        int[] orderedObjectOffsets = inUseOffsets.Values.Concat(visitedXrefs)
            .Distinct().OrderBy(value => value).ToArray();
        if (!IsCatalogObjectAt(bytes, orderedObjectOffsets, rootOffset, objectLimit,
                rootObject.Value, rootGeneration,
                out long pagesObject, out long pagesGeneration)
            || !inUseOffsets.TryGetValue((pagesObject, pagesGeneration), out int pagesOffset)) return false;
        var visited = new HashSet<(long Object, long Generation)>();
        return IsCompletePageTree(bytes, inUseOffsets, orderedObjectOffsets, pagesOffset, objectLimit,
            pagesObject, pagesGeneration, parent: null, hasInheritedMediaBox: false,
            inheritedResources: default, visited, depth: 0, out _);
    }

    private static bool HasValidHeader(byte[] bytes) {
        if (bytes.Length < 9 || !StartsWith(bytes, 0, "%PDF-")
            || bytes[5] < (byte)'0' || bytes[5] > (byte)'9'
            || bytes[6] != (byte)'.'
            || bytes[7] < (byte)'0' || bytes[7] > (byte)'9') return false;
        int offset = 8;
        return ConsumeLineEnd(bytes, ref offset, bytes.Length);
    }

    private static bool TryReadClassicXref(byte[] bytes, int offset, int limit,
        IDictionary<(long Object, long Generation), int> inUseOffsets,
        ISet<long> seenObjects, ref long totalEntries, out int trailerOffset,
        out long size, out bool hasRoot, out long rootObject, out long rootGeneration,
        out bool hasPrevious, out long previousOffset) {
        trailerOffset = -1;
        size = 0;
        hasRoot = false;
        rootObject = 0;
        rootGeneration = 0;
        hasPrevious = false;
        previousOffset = 0;
        offset += 4;
        bool hasSubsection = false;
        var tableObjects = new HashSet<long>();
        while (offset < limit) {
            SkipWhitespace(bytes, ref offset, limit);
            if (StartsWith(bytes, offset, "trailer")) break;
            if (!TryReadDecimal(bytes, ref offset, limit, out long firstObject)) return false;
            SkipWhitespace(bytes, ref offset, limit);
            if (!TryReadDecimal(bytes, ref offset, limit, out long entryCount)
                || entryCount <= 0 || entryCount > 1_000_000
                || firstObject > long.MaxValue - (entryCount - 1)
                || totalEntries > 1_000_000 - entryCount) return false;
            totalEntries += entryCount;
            hasSubsection = true;
            for (long index = 0; index < entryCount; index++) {
                SkipWhitespace(bytes, ref offset, limit);
                if (!TryReadFixedDecimal(bytes, ref offset, limit, 10, out long objectOffset)) return false;
                SkipHorizontalWhitespace(bytes, ref offset, limit);
                if (!TryReadFixedDecimal(bytes, ref offset, limit, 5, out long generation)) return false;
                SkipHorizontalWhitespace(bytes, ref offset, limit);
                if (offset >= limit || bytes[offset] != (byte)'n' && bytes[offset] != (byte)'f') return false;
                bool inUse = bytes[offset++] == (byte)'n';
                if (generation > (inUse ? 65534 : 65535)) return false;
                if (!ConsumeLineEnd(bytes, ref offset, limit)) return false;
                long objectNumber = checked(firstObject + index);
                if (!tableObjects.Add(objectNumber)) return false;
                if (seenObjects.Add(objectNumber) && inUse && objectOffset <= int.MaxValue) {
                    inUseOffsets.Add((objectNumber, generation), (int)objectOffset);
                }
            }
        }
        if (!hasSubsection || !StartsWith(bytes, offset, "trailer")) return false;
        trailerOffset = offset;
        int dictionaryStart = offset + 7;
        if (!SkipWhitespaceAndComments(bytes, ref dictionaryStart, limit)
            || !StartsWith(bytes, dictionaryStart, "<<")) return false;
        int dictionaryEnd = FindDictionaryEnd(bytes, dictionaryStart,
            Math.Min(limit, dictionaryStart + 65536));
        if (dictionaryEnd < 0
            || !TryReadDictionaryInteger(bytes, dictionaryStart, dictionaryEnd, "/Size", out size)
            || size <= 0
            || IndexOfDictionaryName(bytes, "/XRefStm", dictionaryStart, dictionaryEnd) >= 0
            || IndexOfDictionaryName(bytes, "/Encrypt", dictionaryStart, dictionaryEnd) >= 0) {
            return false;
        }
        int afterDictionary = dictionaryEnd + 2;
        if (!SkipWhitespaceAndComments(bytes, ref afterDictionary, limit)
            || !StartsWith(bytes, afterDictionary, "startxref")
            || afterDictionary + 9 >= bytes.Length
            || !IsWhitespace(bytes[afterDictionary + 9])) return false;
        int rootName = FindDictionaryName(bytes, "/Root", dictionaryStart, dictionaryEnd,
            out int rootCount);
        if (rootCount > 1) return false;
        hasRoot = rootName >= 0;
        if (hasRoot && (!TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Root",
                out rootObject, out rootGeneration) || rootObject <= 0 || rootObject >= size)) return false;
        int previousName = FindDictionaryName(bytes, "/Prev", dictionaryStart, dictionaryEnd,
            out int previousCount);
        if (previousCount > 1) return false;
        hasPrevious = previousName >= 0;
        if (hasPrevious && !TryReadDictionaryInteger(bytes, dictionaryStart, dictionaryEnd,
                "/Prev", out previousOffset)) return false;
        return true;
    }

    private static bool TryReadFixedDecimal(byte[] bytes, ref int offset, int limit, int digits, out long value) {
        value = 0;
        if (offset < 0 || offset > limit - digits) return false;
        for (int index = 0; index < digits; index++) {
            byte current = bytes[offset++];
            if (current < (byte)'0' || current > (byte)'9') return false;
            value = value * 10 + current - (byte)'0';
        }
        return true;
    }

    private static bool TryReadDictionaryInteger(byte[] bytes, int start, int end,
        string name, out long value) {
        value = 0;
        int offset = FindDictionaryName(bytes, name, start, end, out int count);
        if (offset < 0 || count != 1) return false;
        offset += name.Length;
        SkipWhitespace(bytes, ref offset, end);
        if (!TryReadDecimal(bytes, ref offset, end, out value)) return false;
        SkipWhitespace(bytes, ref offset, end);
        return offset >= end || IsDelimiter(bytes[offset]);
    }

    private static bool TryReadDictionaryReference(byte[] bytes, int start, int end,
        string name, out long objectNumber, out long generation) {
        objectNumber = 0;
        generation = 0;
        int offset = FindDictionaryName(bytes, name, start, end, out int count);
        if (offset < 0 || count != 1) return false;
        offset += name.Length;
        SkipWhitespace(bytes, ref offset, end);
        if (!TryReadDecimal(bytes, ref offset, end, out objectNumber)) return false;
        SkipWhitespace(bytes, ref offset, end);
        if (!TryReadDecimal(bytes, ref offset, end, out generation)) return false;
        if (generation > 65534) return false;
        SkipWhitespace(bytes, ref offset, end);
        if (offset >= end || bytes[offset++] != (byte)'R') return false;
        return offset >= end || IsDelimiter(bytes[offset]);
    }

    private static bool IsCatalogObjectAt(byte[] bytes, int[] orderedObjectOffsets,
        int offset, int limit,
        long expectedObject, long expectedGeneration, out long pagesObject, out long pagesGeneration) {
        pagesObject = 0;
        pagesGeneration = 0;
        int objectLimit = GetObjectLimit(orderedObjectOffsets, offset, limit);
        if (!TryGetObjectDictionary(bytes, ref offset, objectLimit, expectedObject, expectedGeneration,
                out int dictionaryStart, out int dictionaryEnd)) return false;
        return HasDictionaryNameValue(bytes, dictionaryStart, dictionaryEnd, "/Type", "/Catalog")
            && TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Pages",
                out pagesObject, out pagesGeneration);
    }

    private static bool IsCompletePageTree(byte[] bytes,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets,
        int[] orderedObjectOffsets,
        int offset, int limit, long expectedObject, long expectedGeneration,
        (long Object, long Generation)? parent,
        bool hasInheritedMediaBox,
        ResourceDictionary inheritedResources,
        ISet<(long Object, long Generation)> visited, int depth, out long pageCount) {
        pageCount = 0;
        if (depth > 256 || !visited.Add((expectedObject, expectedGeneration))) return false;
        int objectLimit = GetObjectLimit(orderedObjectOffsets, offset, limit);
        if (!TryGetObjectDictionary(bytes, ref offset, objectLimit, expectedObject, expectedGeneration,
                out int dictionaryStart, out int dictionaryEnd)) return false;
        if (!TryResolveMediaBox(bytes, dictionaryStart, dictionaryEnd,
                hasInheritedMediaBox, out bool hasMediaBox)) return false;
        ResourceDictionary resources = ResolveResources(bytes, inUseOffsets,
            orderedObjectOffsets, dictionaryStart, dictionaryEnd, limit,
            inheritedResources);
        if (HasDictionaryNameValue(bytes, dictionaryStart, dictionaryEnd, "/Type", "/Page")) {
            if (!parent.HasValue
                || !hasMediaBox
                || resources.IsDeclared && (!resources.IsValid
                    || !HasResolvableDictionaryReferences(bytes, resources.Start,
                        resources.End, inUseOffsets))
                || !TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Parent",
                    out long parentObject, out long parentGeneration)
                || parentObject != parent.Value.Object || parentGeneration != parent.Value.Generation
                || !HasCompletePageContents(bytes, inUseOffsets, orderedObjectOffsets,
                    dictionaryStart, dictionaryEnd, limit)) {
                return false;
            }
            pageCount = 1;
            return true;
        }
        if (!HasDictionaryNameValue(bytes, dictionaryStart, dictionaryEnd, "/Type", "/Pages")
            || parent.HasValue && (!TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd,
                    "/Parent", out long pagesParentObject, out long pagesParentGeneration)
                || pagesParentObject != parent.Value.Object || pagesParentGeneration != parent.Value.Generation)
            || !TryReadDictionaryInteger(bytes, dictionaryStart, dictionaryEnd, "/Count", out long declaredCount)
            || declaredCount < 0
            || !TryReadDictionaryReferenceArray(bytes, dictionaryStart, dictionaryEnd, "/Kids",
                out IReadOnlyList<(long Object, long Generation)> children)) return false;
        long total = 0;
        foreach ((long childObject, long childGeneration) in children) {
            if (!inUseOffsets.TryGetValue((childObject, childGeneration), out int childOffset)
                || !IsCompletePageTree(bytes, inUseOffsets, orderedObjectOffsets, childOffset, limit,
                    childObject, childGeneration, (expectedObject, expectedGeneration),
                    hasMediaBox, resources, visited, depth + 1, out long childCount)
                || total > long.MaxValue - childCount) return false;
            total += childCount;
        }
        pageCount = total;
        return total == declaredCount;
    }

    private static bool TryResolveMediaBox(byte[] bytes, int start, int end,
        bool inherited, out bool resolved) {
        resolved = inherited;
        int offset = FindDictionaryName(bytes, "/MediaBox", start, end, out int count);
        if (count == 0) return true;
        if (count != 1 || offset < 0) return false;
        offset += 9;
        if (!SkipWhitespaceAndComments(bytes, ref offset, end)) return false;
        if (offset >= end || bytes[offset++] != (byte)'[') return false;
        var values = new double[4];
        for (int index = 0; index < values.Length; index++) {
            if (!SkipWhitespaceAndComments(bytes, ref offset, end)) return false;
            if (!TryReadPdfNumber(bytes, ref offset, end, out values[index])) return false;
        }
        if (!SkipWhitespaceAndComments(bytes, ref offset, end)) return false;
        if (offset >= end || bytes[offset++] != (byte)']'
            || offset < end && !IsDelimiter(bytes[offset])
            || values[2] <= values[0] || values[3] <= values[1]) return false;
        resolved = true;
        return true;
    }

    private static bool TryReadPdfNumber(byte[] bytes, ref int offset, int limit,
        out double value) {
        value = 0;
        int start = offset;
        if (offset < limit && bytes[offset] is (byte)'+' or (byte)'-') offset++;
        bool hasDigit = false;
        while (offset < limit && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9') {
            hasDigit = true;
            offset++;
        }
        if (offset < limit && bytes[offset] == (byte)'.') {
            offset++;
            while (offset < limit && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9') {
                hasDigit = true;
                offset++;
            }
        }
        if (!hasDigit) return false;
        string token = System.Text.Encoding.ASCII.GetString(bytes, start, offset - start);
        return double.TryParse(token, System.Globalization.NumberStyles.AllowLeadingSign
            | System.Globalization.NumberStyles.AllowDecimalPoint,
            System.Globalization.CultureInfo.InvariantCulture, out value)
            && !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static bool HasCompletePageContents(byte[] bytes,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets,
        int[] orderedObjectOffsets,
        int dictionaryStart, int dictionaryEnd, int limit) {
        int contentsOffset = FindDictionaryName(bytes, "/Contents", dictionaryStart,
            dictionaryEnd, out int contentsCount);
        if (contentsCount == 0) return true;
        if (contentsCount != 1 || contentsOffset < 0) return false;

        IReadOnlyList<(long Object, long Generation)> references;
        if (TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Contents",
                out long contentObject, out long contentGeneration)) {
            references = new[] { (contentObject, contentGeneration) };
        } else if (!TryReadDictionaryReferenceArray(bytes, dictionaryStart, dictionaryEnd,
                       "/Contents", out references)) {
            return false;
        }

        var validated = new HashSet<(long Object, long Generation)>();
        foreach ((long referencedObject, long referencedGeneration) in references) {
            if (!inUseOffsets.TryGetValue((referencedObject, referencedGeneration), out int contentOffset)
                || validated.Add((referencedObject, referencedGeneration))
                && !IsCompleteStreamObject(bytes, inUseOffsets, orderedObjectOffsets,
                    contentOffset, limit,
                    referencedObject, referencedGeneration)) return false;
        }
        return true;
    }

    private static bool IsCompleteStreamObject(byte[] bytes,
        IReadOnlyDictionary<(long Object, long Generation), int> inUseOffsets,
        int[] orderedObjectOffsets,
        int offset, int limit, long expectedObject, long expectedGeneration) {
        int packageLimit = limit;
        limit = GetObjectLimit(orderedObjectOffsets, offset, packageLimit);
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long objectNumber)
            || objectNumber != expectedObject) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long generation)
            || generation != expectedGeneration) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!StartsWith(bytes, offset, "obj")) return false;
        offset += 3;
        SkipWhitespace(bytes, ref offset, limit);
        int dictionaryStart = StartsWith(bytes, offset, "<<") ? offset : -1;
        int dictionaryEnd = dictionaryStart < 0 ? -1 : FindDictionaryEnd(bytes,
            dictionaryStart, Math.Min(limit, dictionaryStart + 65536));
        if (dictionaryStart < 0 || dictionaryEnd < 0) return false;

        long streamLength;
        if (!TryReadDictionaryInteger(bytes, dictionaryStart, dictionaryEnd, "/Length",
                out streamLength)) {
            if (!TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Length",
                    out long lengthObject, out long lengthGeneration)
                || !inUseOffsets.TryGetValue((lengthObject, lengthGeneration), out int lengthOffset)
                || !TryReadIndirectIntegerObject(bytes, orderedObjectOffsets, lengthOffset,
                    packageLimit,
                    lengthObject, lengthGeneration, out streamLength)) return false;
        }
        if (streamLength < 0 || streamLength > int.MaxValue) return false;

        offset = dictionaryEnd + 2;
        SkipWhitespace(bytes, ref offset, limit);
        if (!StartsWith(bytes, offset, "stream")) return false;
        offset += 6;
        if (!ConsumeLineEnd(bytes, ref offset, limit)
            || offset > limit - streamLength) return false;
        offset += (int)streamLength;
        if (!ConsumeLineEnd(bytes, ref offset, limit)
            || !StartsWith(bytes, offset, "endstream")) return false;
        offset += 9;
        SkipWhitespace(bytes, ref offset, limit);
        return StartsWith(bytes, offset, "endobj");
    }

    private static bool TryReadIndirectIntegerObject(byte[] bytes, int[] orderedObjectOffsets,
        int offset, int limit,
        long expectedObject, long expectedGeneration, out long value) {
        value = 0;
        limit = GetObjectLimit(orderedObjectOffsets, offset, limit);
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long objectNumber)
            || objectNumber != expectedObject) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long generation)
            || generation != expectedGeneration) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!StartsWith(bytes, offset, "obj")) return false;
        offset += 3;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out value)) return false;
        SkipWhitespace(bytes, ref offset, limit);
        return StartsWith(bytes, offset, "endobj");
    }

    private static int GetObjectLimit(int[] orderedObjectOffsets, int offset, int defaultLimit) {
        int index = Array.BinarySearch(orderedObjectOffsets, offset);
        index = index < 0 ? ~index : index + 1;
        while (index < orderedObjectOffsets.Length && orderedObjectOffsets[index] <= offset) index++;
        return index < orderedObjectOffsets.Length && orderedObjectOffsets[index] < defaultLimit
            ? orderedObjectOffsets[index]
            : defaultLimit;
    }

    private static bool TryReadDictionaryReferenceArray(byte[] bytes, int start, int end,
        string name, out IReadOnlyList<(long Object, long Generation)> references) {
        var result = new List<(long Object, long Generation)>();
        references = result;
        int offset = FindDictionaryName(bytes, name, start, end, out int count);
        if (offset < 0 || count != 1) return false;
        offset += name.Length;
        SkipWhitespace(bytes, ref offset, end);
        if (offset >= end || bytes[offset++] != (byte)'[') return false;
        while (true) {
            SkipWhitespace(bytes, ref offset, end);
            if (offset >= end) return false;
            if (bytes[offset] == (byte)']') {
                offset++;
                references = result;
                return offset >= end || IsDelimiter(bytes[offset]);
            }
            if (!TryReadDecimal(bytes, ref offset, end, out long objectNumber)) return false;
            SkipWhitespace(bytes, ref offset, end);
            if (!TryReadDecimal(bytes, ref offset, end, out long generation)) return false;
            if (generation > 65534) return false;
            SkipWhitespace(bytes, ref offset, end);
            if (offset >= end || bytes[offset++] != (byte)'R'
                || objectNumber <= 0 || result.Count >= 1_000_000) return false;
            result.Add((objectNumber, generation));
        }
    }

    private static bool TryGetObjectDictionary(byte[] bytes, ref int offset, int limit,
        long expectedObject, long expectedGeneration, out int dictionaryStart, out int dictionaryEnd) {
        dictionaryStart = -1;
        dictionaryEnd = -1;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long objectNumber)
            || objectNumber != expectedObject) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long generation)
            || generation != expectedGeneration) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!StartsWith(bytes, offset, "obj")) return false;
        int objectEnd = IndexOf(bytes, "endobj", offset + 3, Math.Min(limit, offset + 65536));
        if (objectEnd < 0) return false;
        dictionaryStart = IndexOf(bytes, "<<", offset + 3, objectEnd);
        dictionaryEnd = dictionaryStart < 0 ? -1 : FindDictionaryEnd(bytes, dictionaryStart, objectEnd);
        if (dictionaryStart < 0 || dictionaryEnd < 0) return false;
        int trailingOffset = dictionaryEnd + 2;
        return SkipWhitespaceAndComments(bytes, ref trailingOffset, objectEnd)
            && trailingOffset == objectEnd;
    }

    private static int FindDictionaryEnd(byte[] bytes, int start, int limit) {
        if (start < 0 || start >= limit - 1
            || bytes[start] != (byte)'<' || bytes[start + 1] != (byte)'<') return -1;
        int depth = 0;
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
                if (escaped) {
                    escaped = false;
                } else if (current == (byte)'\\') {
                    escaped = true;
                } else if (current == (byte)'(') {
                    literalDepth++;
                } else if (current == (byte)')') {
                    literalDepth--;
                }
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
                    depth++;
                    offset++;
                } else {
                    inHexString = true;
                }
                continue;
            }
            if (current == (byte)'>' && offset + 1 < limit
                && bytes[offset + 1] == (byte)'>') {
                depth--;
                if (depth == 0) return offset;
                if (depth < 0) return -1;
                offset++;
            }
        }
        return -1;
    }

    private static bool HasDictionaryNameValue(byte[] bytes, int start, int end,
        string name, string value) {
        int offset = FindDictionaryName(bytes, name, start, end, out int count);
        if (offset < 0 || count != 1) return false;
        offset += name.Length;
        SkipWhitespace(bytes, ref offset, end);
        return StartsWith(bytes, offset, value)
            && (offset + value.Length >= end || IsDelimiter(bytes[offset + value.Length]));
    }

    private static int IndexOfDictionaryName(byte[] bytes, string name, int start, int end) =>
        FindDictionaryName(bytes, name, start, end, out _);

    private static int FindDictionaryName(byte[] bytes, string name, int start, int end,
        out int matchCount) {
        matchCount = 0;
        int firstMatch = -1;
        int dictionaryDepth = 0;
        int arrayDepth = 0;
        int literalDepth = 0;
        bool escaped = false;
        bool inHexString = false;
        bool inComment = false;
        bool expectingKey = true;
        bool literalCompletesValue = false;
        bool hexCompletesValue = false;
        bool dictionaryCompletesValue = false;
        bool arrayCompletesValue = false;
        for (int offset = Math.Max(0, start); offset < end; offset++) {
            byte current = bytes[offset];
            if (inComment) {
                if (current is 0x0a or 0x0d) inComment = false;
                continue;
            }
            if (literalDepth > 0) {
                if (escaped) {
                    escaped = false;
                } else if (current == (byte)'\\') {
                    escaped = true;
                } else if (current == (byte)'(') {
                    literalDepth++;
                } else if (current == (byte)')') {
                    literalDepth--;
                    if (literalDepth == 0 && literalCompletesValue) {
                        expectingKey = true;
                        literalCompletesValue = false;
                    }
                }
                continue;
            }
            if (inHexString) {
                if (current == (byte)'>') {
                    inHexString = false;
                    if (hexCompletesValue) {
                        expectingKey = true;
                        hexCompletesValue = false;
                    }
                }
                continue;
            }
            if (current == (byte)'%') {
                inComment = true;
                continue;
            }
            if (current == (byte)'(') {
                literalDepth = 1;
                literalCompletesValue = dictionaryDepth == 1 && arrayDepth == 0 && !expectingKey;
                continue;
            }
            if (current == (byte)'<' && offset + 1 < end) {
                if (bytes[offset + 1] == (byte)'<') {
                    if (dictionaryDepth == 1 && arrayDepth == 0 && !expectingKey) {
                        dictionaryCompletesValue = true;
                    }
                    dictionaryDepth++;
                    offset++;
                } else {
                    inHexString = true;
                    hexCompletesValue = dictionaryDepth == 1 && arrayDepth == 0 && !expectingKey;
                }
                continue;
            }
            if (current == (byte)'>' && offset + 1 < end && bytes[offset + 1] == (byte)'>') {
                dictionaryDepth--;
                if (dictionaryDepth < 0) return -1;
                if (dictionaryDepth == 1 && dictionaryCompletesValue) {
                    expectingKey = true;
                    dictionaryCompletesValue = false;
                }
                offset++;
                continue;
            }
            if (current == (byte)'[') {
                if (dictionaryDepth == 1 && arrayDepth == 0 && !expectingKey) {
                    arrayCompletesValue = true;
                }
                arrayDepth++;
                continue;
            }
            if (current == (byte)']') {
                if (arrayDepth == 0) return -1;
                arrayDepth--;
                if (arrayDepth == 0 && arrayCompletesValue) {
                    expectingKey = true;
                    arrayCompletesValue = false;
                }
                continue;
            }
            if (dictionaryDepth != 1 || arrayDepth != 0 || IsWhitespace(current)) continue;
            if (current == (byte)'/') {
                int nameOffset = offset;
                int tokenEnd = offset + 1;
                while (tokenEnd < end && !IsDelimiter(bytes[tokenEnd])) tokenEnd++;
                if (!expectingKey) {
                    expectingKey = true;
                    offset = tokenEnd - 1;
                    continue;
                }
                bool matches = tokenEnd - offset == name.Length
                    && StartsWith(bytes, offset, name);
                expectingKey = false;
                offset = tokenEnd - 1;
                if (matches) {
                    matchCount++;
                    if (firstMatch < 0) firstMatch = nameOffset;
                }
                continue;
            }
            if (!expectingKey) expectingKey = true;
        }
        return firstMatch;
    }

    private static bool IsDelimiter(byte value) => IsWhitespace(value)
        || value is (byte)'(' or (byte)')' or (byte)'<' or (byte)'>' or (byte)'[' or (byte)']'
            or (byte)'{' or (byte)'}' or (byte)'/' or (byte)'%';

    private static int FindObjectHeader(byte[] bytes, long objectNumber, long generation, int limit) {
        string header = objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " "
            + generation.ToString(System.Globalization.CultureInfo.InvariantCulture) + " obj";
        int offset = 0;
        while (offset < limit) {
            int found = IndexOf(bytes, header, offset, limit);
            if (found < 0) return -1;
            int after = found + header.Length;
            bool startsAtToken = found == 0 || IsDelimiter(bytes[found - 1]);
            bool endsAtToken = after >= limit || IsDelimiter(bytes[after]);
            if (startsAtToken && endsAtToken) return found;
            offset = found + 1;
        }
        return -1;
    }

    private static void SkipHorizontalWhitespace(byte[] bytes, ref int offset, int limit) {
        while (offset < limit && bytes[offset] is 0x09 or 0x20) offset++;
    }

    private static bool ConsumeLineEnd(byte[] bytes, ref int offset, int limit) {
        SkipHorizontalWhitespace(bytes, ref offset, limit);
        if (offset >= limit) return false;
        if (bytes[offset] == 0x0d) {
            offset++;
            if (offset < limit && bytes[offset] == 0x0a) offset++;
            return true;
        }
        if (bytes[offset] == 0x0a) {
            offset++;
            return true;
        }
        return false;
    }

    private static bool TryReadDecimal(byte[] bytes, ref int offset, int limit, out long value) {
        value = 0;
        int start = offset;
        while (offset < limit && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9') {
            int digit = bytes[offset++] - (byte)'0';
            if (value > (long.MaxValue - digit) / 10) return false;
            value = value * 10 + digit;
        }
        return offset > start;
    }

    private static void SkipWhitespace(byte[] bytes, ref int offset, int limit) {
        while (offset < limit && IsWhitespace(bytes[offset])) offset++;
    }

    private static bool SkipWhitespaceAndComments(byte[] bytes, ref int offset, int limit) {
        while (offset < limit) {
            SkipWhitespace(bytes, ref offset, limit);
            if (offset >= limit || bytes[offset] != (byte)'%') return true;
            while (offset < limit && bytes[offset] is not 0x0a and not 0x0d) offset++;
            if (offset >= limit) return false;
        }
        return true;
    }

    private static bool ContainsOnlyTrailingWhitespace(byte[] bytes, int offset) {
        for (int index = offset; index < bytes.Length; index++) {
            if (!IsWhitespace(bytes[index]) && bytes[index] != 0) return false;
        }
        return true;
    }

    private static bool IsWhitespace(byte value) =>
        value is 0x09 or 0x0a or 0x0c or 0x0d or 0x20;

    private static int LastIndexOf(byte[] bytes, string value, int? before = null) {
        int last = -1;
        int limit = Math.Min(before ?? bytes.Length, bytes.Length);
        for (int index = 0; index <= limit - value.Length; index++) {
            if (StartsWith(bytes, index, value)) last = index;
        }
        return last;
    }

    private static int IndexOf(byte[] bytes, string value, int start, int limit) {
        int end = Math.Min(limit, bytes.Length);
        for (int index = Math.Max(0, start); index <= end - value.Length; index++) {
            if (StartsWith(bytes, index, value)) return index;
        }
        return -1;
    }

    private static bool StartsWith(byte[] bytes, int offset, string value) {
        if (offset < 0 || offset > bytes.Length - value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            if (bytes[offset + index] != (byte)value[index]) return false;
        }
        return true;
    }
}
