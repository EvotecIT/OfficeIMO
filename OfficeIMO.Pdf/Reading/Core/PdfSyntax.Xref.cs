namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private const int MaximumXrefObjectCharacters = 1_000_000;

    private sealed class XrefObjectScanBudget {
        private readonly long _maximum;
        private long _used;

        internal XrefObjectScanBudget(PdfReadLimits limits) {
            MaximumPerObject = Math.Min(MaximumXrefObjectCharacters, limits.MaxObjectCharacters);
            _maximum = limits.MaxInputBytes;
        }

        internal int MaximumPerObject { get; }

        internal void Charge(long characters) {
            if (characters <= 0) return;
            _used = checked(_used + characters);
            if (_used > _maximum) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, _maximum, _used);
            }
        }
    }

    private static void ResolveIndirectStreamLengths(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        List<(int Id, int Generation, int DataStart)> streamLocations,
        PdfReadLimits limits,
        System.Threading.CancellationToken cancellationToken) {
        foreach (var streamLocation in streamLocations) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!map.TryGetValue(streamLocation.Id, out var indirect) || indirect.Value is not PdfStream stream) {
                continue;
            }

            if (!TryGetResolvedLength(stream.Dictionary, map, out int byteLen)) {
                continue;
            }

            int byteStart = streamLocation.DataStart;
            if (byteStart < 0 || byteLen < 0 || byteStart + byteLen > pdf.Length) {
                continue;
            }

            if (stream.DataLength == byteLen) {
                continue;
            }

            if (byteLen > limits.MaxRawStreamBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.RawStreamBytes, limits.MaxRawStreamBytes, byteLen);
            }

            byte[] data = CopyBytes(pdf, byteStart, byteLen, cancellationToken);
            map[streamLocation.Id] = new PdfIndirectObject(streamLocation.Id, streamLocation.Generation, new PdfStream(stream.Dictionary, data, stream.DecodingFailed, stream.DecodingError));
        }
    }

    private static bool ApplyClassicXrefEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        string text,
        Dictionary<int, int> parsedOffsets,
        HashSet<int> activeObjectNumbers,
        PdfReadLimits limits,
        XrefObjectScanBudget scanBudget,
        PdfDecodedStreamBudget decodedStreamBudget,
        Action reportIncompleteXref,
        out bool appliedClassicEntries,
        System.Threading.CancellationToken cancellationToken) {
        appliedClassicEntries = false;
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset, cancellationToken)) {
            return false;
        }

        var tables = GetClassicXrefTableChain(text, activeXrefOffset, cancellationToken);
        if (tables.Count == 0) {
            return false;
        }

        appliedClassicEntries = true;
        bool appliedXrefStream = false;
        var parsedObjectsByOffset = new Dictionary<int, PdfIndirectObject?>(parsedOffsets.Count);
        foreach (KeyValuePair<int, int> parsedOffset in parsedOffsets) {
            cancellationToken.ThrowIfCancellationRequested();
            if (map.TryGetValue(parsedOffset.Key, out PdfIndirectObject? parsedObject)) {
                parsedObjectsByOffset[parsedOffset.Value] = parsedObject;
            }
        }
        foreach (var table in tables) {
            cancellationToken.ThrowIfCancellationRequested();
            ApplyClassicXrefTableEntries(map, pdf, parsedOffsets, text, table.Entries, scanBudget, activeObjectNumbers, parsedObjectsByOffset, cancellationToken);
            if (table.XrefStreamOffset.HasValue) {
                appliedXrefStream = ApplyXrefStreamAtOffset(map, pdf, parsedOffsets, text, table.XrefStreamOffset.Value, limits, scanBudget, decodedStreamBudget, reportIncompleteXref, cancellationToken) || appliedXrefStream;
            }
        }

        return appliedXrefStream;
    }

    private static void ApplyClassicXrefTableEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        IReadOnlyList<(int ObjectNumber, int Offset, int Generation, bool InUse)> entries,
        XrefObjectScanBudget scanBudget,
        HashSet<int>? activeObjectNumbers = null,
        Dictionary<int, PdfIndirectObject?>? parsedObjectsByOffset = null,
        System.Threading.CancellationToken cancellationToken = default) {
        parsedObjectsByOffset ??= new Dictionary<int, PdfIndirectObject?>();
        foreach (var entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!entry.InUse) {
                if (entry.ObjectNumber != 0) {
                    map.Remove(entry.ObjectNumber);
                    parsedOffsets.Remove(entry.ObjectNumber);
                    activeObjectNumbers?.Remove(entry.ObjectNumber);
                }

                continue;
            }

            if (entry.Offset <= 0 ||
                entry.Offset >= pdf.Length) {
                continue;
            }

            if (!parsedObjectsByOffset.TryGetValue(entry.Offset, out PdfIndirectObject? parsed)) {
                parsed = TryParseIndirectObjectAt(pdf, text, entry.Offset, map, scanBudget, out PdfIndirectObject candidate, cancellationToken)
                    ? candidate
                    : null;
                parsedObjectsByOffset[entry.Offset] = parsed;
            }

            if (parsed is not null &&
                parsed.ObjectNumber == entry.ObjectNumber &&
                parsed.Generation == entry.Generation) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = entry.Offset;
                activeObjectNumbers?.Add(entry.ObjectNumber);
            }
        }
    }

    private static List<(int Offset, (int ObjectNumber, int Offset, int Generation, bool InUse)[] Entries, int? XrefStreamOffset)> GetClassicXrefTableChain(string text, int activeXrefOffset,
        System.Threading.CancellationToken cancellationToken = default) {
        var newestToOldest = new List<(int Offset, (int ObjectNumber, int Offset, int Generation, bool InUse)[] Entries, int? XrefStreamOffset)>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (visited.Add(currentOffset) &&
            newestToOldest.Count < 64 &&
            TryParseClassicXrefTable(text, currentOffset, out var entries, out int? previousOffset, out _, out int? xrefStreamOffset, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            newestToOldest.Add((currentOffset, entries, xrefStreamOffset));
            if (!previousOffset.HasValue) {
                break;
            }

            currentOffset = previousOffset.Value;
        }

        newestToOldest.Reverse();
        return newestToOldest;
    }

    private static bool TryParseClassicXrefTable(string text, int offset, out (int ObjectNumber, int Offset, int Generation, bool InUse)[] entries, out int? previousOffset, out string trailerRaw, out int? xrefStreamOffset,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        entries = Array.Empty<(int ObjectNumber, int Offset, int Generation, bool InUse)>();
        previousOffset = null;
        trailerRaw = string.Empty;
        xrefStreamOffset = null;
        if (offset < 0 ||
            offset + 4 > text.Length ||
            !string.Equals(text.Substring(offset, 4), "xref", StringComparison.Ordinal) ||
            !HasKeywordBoundary(text, offset - 1, 0, text.Length) ||
            !HasKeywordBoundary(text, offset + 4, 0, text.Length)) {
            return false;
        }

        int trailerIndex = IndexOfKeywordCancellable(text, "trailer", offset + 4, text.Length, cancellationToken);
        if (trailerIndex < 0) {
            return false;
        }

        int position = offset + 4;
        int sectionEnd = (int)Math.Min((long)trailerIndex, (long)position + 2_000_000L);
        // The structural section estimate is deliberately capped. Growth uses pooled
        // arrays, and only the observed entries become a retained managed array.
        using var entryBuilder = new PdfPooledValueBuilder<(int ObjectNumber, int Offset, int Generation, bool InUse)>(
            PdfCollectionSizing.BoundedInitialCapacity((sectionEnd - position) / 20, 256));
        while (TryReadXrefLine(text, ref position, sectionEnd, out int lineStart, out int lineEnd, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            int tokenPosition = lineStart;
            if (!TryReadXrefToken(text, ref tokenPosition, lineEnd, out int firstStart, out int firstLength) ||
                !TryReadXrefToken(text, ref tokenPosition, lineEnd, out int countStart, out int countLength) ||
                !TryParseXrefInteger(text, firstStart, firstLength, out int firstObjectNumber) ||
                !TryParseXrefInteger(text, countStart, countLength, out int count) ||
                firstObjectNumber < 0 ||
                count <= 0 ||
                count > 1_000_000) {
                continue;
            }

            for (int i = 0; i < count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!TryReadXrefLine(text, ref position, sectionEnd, out lineStart, out lineEnd, cancellationToken)) {
                    entries = entryBuilder.ToArray();
                    return entries.Length > 0;
                }

                tokenPosition = lineStart;
                if (!TryReadXrefToken(text, ref tokenPosition, lineEnd, out int offsetStart, out int offsetLength) ||
                    !TryReadXrefToken(text, ref tokenPosition, lineEnd, out int generationStart, out int generationLength) ||
                    !TryReadXrefToken(text, ref tokenPosition, lineEnd, out int statusStart, out int statusLength) ||
                    !TryParseXrefInteger(text, offsetStart, offsetLength, out int objectOffset) ||
                    !TryParseXrefInteger(text, generationStart, generationLength, out int generation)) {
                    continue;
                }

                if (statusLength == 1 && text[statusStart] == 'n') {
                    entryBuilder.Add((firstObjectNumber + i, objectOffset, generation, true));
                } else if (statusLength == 1 && text[statusStart] == 'f') {
                    entryBuilder.Add((firstObjectNumber + i, objectOffset, generation, false));
                }
            }
        }

        entries = entryBuilder.ToArray();
        if (entries.Length == 0) {
            return false;
        }

        int dictStart = -1;
        int trailerSearchLimit = text.Length;
        const int searchWindow = 65536;
        for (long search = trailerIndex; search < trailerSearchLimit; search += searchWindow) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = (int)Math.Min(trailerSearchLimit - search, searchWindow + 1L);
            dictStart = text.IndexOf("<<", (int)search, count, StringComparison.Ordinal);
            if (dictStart >= 0) break;
        }
        if (dictStart >= 0) {
            int dictionaryLimit = (int)Math.Min((long)text.Length, (long)dictStart + 1_000_002L);
            int dictEnd = FindDictEnd(text, dictStart, dictionaryLimit, cancellationToken);
            if (dictEnd > dictStart) {
                trailerRaw = SafeSliceCancellable(text, trailerIndex, dictEnd - trailerIndex, 1_000_000, cancellationToken);
                string dictText = SafeSliceCancellable(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000, cancellationToken);
                try {
                    PdfDictionary trailer = ParseDictionary(dictText, cancellationToken: cancellationToken);
                    if (trailer.Get<PdfNumber>("Prev") is PdfNumber previous &&
                        previous.Value >= 0 &&
                        previous.Value <= int.MaxValue) {
                        previousOffset = (int)Math.Floor(previous.Value);
                    }

                    if (trailer.Get<PdfNumber>("XRefStm") is PdfNumber xrefStream &&
                        xrefStream.Value >= 0 &&
                        xrefStream.Value <= int.MaxValue) {
                        xrefStreamOffset = (int)Math.Floor(xrefStream.Value);
                    }
                } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not OperationCanceledException) {
                    previousOffset = null;
                    xrefStreamOffset = null;
                }
            }
        }

        return true;
    }

    private static bool TryReadXrefLine(string text, ref int position, int end, out int start, out int lineEnd,
        System.Threading.CancellationToken cancellationToken) {
        start = position;
        if (position >= end) {
            lineEnd = end;
            return false;
        }

        while (position < end && text[position] != '\r' && text[position] != '\n') {
            if ((position & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            position++;
        }
        lineEnd = position;
        if (position < end && text[position++] == '\r' && position < end && text[position] == '\n') position++;
        return true;
    }

    private static bool TryReadXrefToken(string text, ref int position, int end, out int start, out int length) {
        while (position < end && char.IsWhiteSpace(text[position])) position++;
        start = position;
        while (position < end && !char.IsWhiteSpace(text[position])) position++;
        length = position - start;
        return length > 0;
    }

    private static bool TryParseXrefInteger(string text, int start, int length, out int value) {
#if NET8_0_OR_GREATER
        return int.TryParse(text.AsSpan(start, length), System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out value);
#else
        return int.TryParse(text.Substring(start, length), System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out value);
#endif
    }

    private static bool ApplyXrefStreamEntries(Dictionary<int, PdfIndirectObject> map, byte[] pdf, Dictionary<int, int> parsedOffsets, PdfReadLimits limits, XrefObjectScanBudget scanBudget, PdfDecodedStreamBudget decodedStreamBudget, Action reportIncompleteXref,
        System.Threading.CancellationToken cancellationToken) {
        var xrefStreams = new List<(int ObjectNumber, int Offset, PdfStream Stream)>();
        foreach (var entry in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                int offset = parsedOffsets.TryGetValue(entry.ObjectNumber, out int parsedOffset) ? parsedOffset : int.MaxValue;
                xrefStreams.Add((entry.ObjectNumber, offset, stream));
            }
        }

        if (xrefStreams.Count == 0) {
            return false;
        }

        string text = PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset, cancellationToken)) {
            return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        SortXrefStreams(xrefStreams, cancellationToken);
        var activeChainOffsets = GetXrefStreamChainOffsets(xrefStreams, activeXrefOffset, cancellationToken);
        if (activeChainOffsets.Count == 0) {
            return false;
        }

        var classicPredecessors = GetClassicPredecessorTablesForXrefStreamChain(text, xrefStreams, activeXrefOffset, cancellationToken);
        var parsedObjectsByOffset = new Dictionary<int, PdfIndirectObject?>();
        foreach (var table in classicPredecessors) {
            cancellationToken.ThrowIfCancellationRequested();
            ApplyClassicXrefTableEntries(map, pdf, parsedOffsets, text, table.Entries, scanBudget, parsedObjectsByOffset: parsedObjectsByOffset, cancellationToken: cancellationToken);
            if (table.XrefStreamOffset.HasValue) {
                ApplyXrefStreamAtOffset(map, pdf, parsedOffsets, text, table.XrefStreamOffset.Value, limits, scanBudget, decodedStreamBudget, reportIncompleteXref, cancellationToken);
            }
        }

        foreach (int chainOffset in activeChainOffsets) {
            cancellationToken.ThrowIfCancellationRequested();
            var xrefStream = xrefStreams.First(item => item.Offset == chainOffset);
            ApplyXrefStreamObjectEntries(map, pdf, parsedOffsets, text, xrefStream.Stream, limits, scanBudget, decodedStreamBudget, reportIncompleteXref, cancellationToken);
        }

        return true;
    }

    private static bool ApplyCompressedXrefStreamEntries(Dictionary<int, PdfIndirectObject> map, byte[] pdf, Dictionary<int, int> parsedOffsets, PdfReadLimits limits, PdfDecodedStreamBudget decodedStreamBudget, Action<int> reportUnreadable,
        System.Threading.CancellationToken cancellationToken) {
        var xrefStreams = new List<(int ObjectNumber, int Offset, PdfStream Stream)>();
        foreach (var entry in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                int offset = parsedOffsets.TryGetValue(entry.ObjectNumber, out int parsedOffset) ? parsedOffset : int.MaxValue;
                xrefStreams.Add((entry.ObjectNumber, offset, stream));
            }
        }

        if (xrefStreams.Count == 0) {
            return false;
        }

        string text = PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset, cancellationToken)) {
            return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        SortXrefStreams(xrefStreams, cancellationToken);
        var activeChainOffsets = GetXrefStreamChainOffsets(xrefStreams, activeXrefOffset, cancellationToken);
        var activeEntries = new Dictionary<int, XrefStreamEntry>();
        var classicTables = activeChainOffsets.Count == 0
            ? GetClassicXrefTableChain(text, activeXrefOffset, cancellationToken)
            : GetClassicPredecessorTablesForXrefStreamChain(text, xrefStreams, activeXrefOffset, cancellationToken);
        if (classicTables.Count == 0 && activeChainOffsets.Count == 0) return false;
        foreach (var table in classicTables) {
            cancellationToken.ThrowIfCancellationRequested();
            // A hybrid stream belongs to its classic section. Apply its compressed
            // entries before that section's direct entries, then let newer sections win.
            if (table.XrefStreamOffset.HasValue) {
                var xrefStream = xrefStreams.FirstOrDefault(item => item.Offset == table.XrefStreamOffset.Value);
                if (xrefStream.Stream is not null)
                    UpdateActiveCompressedEntries(activeEntries, xrefStream.Stream, map, decodedStreamBudget,
                        () => reportUnreadable(xrefStream.ObjectNumber), cancellationToken);
            }
            for (int i = 0; i < table.Entries.Length; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                activeEntries.Remove(table.Entries[i].ObjectNumber);
            }
        }
        if (activeChainOffsets.Count != 0) {
            foreach (int chainOffset in activeChainOffsets) {
                cancellationToken.ThrowIfCancellationRequested();
                var xrefStream = xrefStreams.First(item => item.Offset == chainOffset);
                UpdateActiveCompressedEntries(activeEntries, xrefStream.Stream, map, decodedStreamBudget, () => reportUnreadable(xrefStream.ObjectNumber), cancellationToken);
            }
        }

        bool applied = false;
        foreach (XrefStreamEntry entry in activeEntries.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Field1 < 0 ||
                entry.Field1 > int.MaxValue ||
                entry.Field2 < 0 ||
                entry.Field2 > int.MaxValue) {
                reportUnreadable(entry.ObjectNumber);
                continue;
            }

            int objectStreamNumber = (int)entry.Field1;
            int objectStreamIndex = (int)entry.Field2;
            if (TryParseObjectFromObjectStream(map, parsedOffsets, objectStreamNumber, objectStreamIndex, entry.ObjectNumber, limits, decodedStreamBudget, out PdfIndirectObject parsed, out int objectStreamOffset, cancellationToken)) {
                if (parsed.Value.HasIncompleteSyntax) reportUnreadable(entry.ObjectNumber);
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = objectStreamOffset;
                applied = true;
            } else reportUnreadable(entry.ObjectNumber);
        }

        return applied;
    }

    private static void UpdateActiveCompressedEntries(
        Dictionary<int, XrefStreamEntry> activeEntries,
        PdfStream xrefStream,
        Dictionary<int, PdfIndirectObject> map,
        PdfDecodedStreamBudget decodedStreamBudget,
        Action reportIncompleteXref,
        System.Threading.CancellationToken cancellationToken) {
        byte[] data = decodedStreamBudget.Decode(xrefStream, map, int.MaxValue, cancellationToken);
        foreach (XrefStreamEntry entry in ReadXrefStreamEntries(xrefStream.Dictionary, data, reportIncompleteXref, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Type == 2) {
                activeEntries[entry.ObjectNumber] = entry;
            } else {
                activeEntries.Remove(entry.ObjectNumber);
            }
        }
    }

    private static bool ApplyXrefStreamAtOffset(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        int xrefStreamOffset,
        PdfReadLimits limits,
        XrefObjectScanBudget scanBudget,
        PdfDecodedStreamBudget decodedStreamBudget,
        Action reportIncompleteXref,
        System.Threading.CancellationToken cancellationToken) {
        PdfStream? targetStream = null;
        foreach (var entry in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!parsedOffsets.TryGetValue(entry.ObjectNumber, out int offset) ||
                offset != xrefStreamOffset ||
                entry.Value is not PdfStream stream ||
                stream.Dictionary.Get<PdfName>("Type")?.Name != "XRef") {
                continue;
            }

            targetStream = stream;
            break;
        }

        if (targetStream is null) {
            return false;
        }

        ApplyXrefStreamObjectEntries(map, pdf, parsedOffsets, text, targetStream, limits, scanBudget, decodedStreamBudget, reportIncompleteXref, cancellationToken);
        return true;
    }

    private static void ApplyXrefStreamObjectEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        PdfStream xrefStream,
        PdfReadLimits limits,
        XrefObjectScanBudget scanBudget,
        PdfDecodedStreamBudget decodedStreamBudget,
        Action reportIncompleteXref,
        System.Threading.CancellationToken cancellationToken) {
        byte[] data = decodedStreamBudget.Decode(xrefStream, map, int.MaxValue, cancellationToken);
        var entries = ReadXrefStreamEntries(xrefStream.Dictionary, data, reportIncompleteXref, cancellationToken).ToList();
        var parsedObjectsByOffset = new Dictionary<int, PdfIndirectObject?>();
        foreach (var entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Type == 0 &&
                entry.ObjectNumber != 0) {
                map.Remove(entry.ObjectNumber);
                parsedOffsets.Remove(entry.ObjectNumber);
            }
        }

        foreach (var entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Type != 1 ||
                entry.Field1 < 0 ||
                entry.Field1 > int.MaxValue ||
                entry.Field2 < 0 ||
                entry.Field2 > int.MaxValue) {
                continue;
            }

            int offset = (int)entry.Field1;
            int generation = (int)entry.Field2;
            if (map.TryGetValue(entry.ObjectNumber, out PdfIndirectObject? existing) &&
                existing.Value is PdfStream existingStream &&
                existingStream.Dictionary.Get<PdfName>("Type")?.Name == "XRef" &&
                parsedOffsets.TryGetValue(entry.ObjectNumber, out int existingOffset) &&
                existingOffset == offset) {
                // Preserve the already-decoded active xref stream instance. Re-parsing its
                // self-entry would replace the budget cache key and charge the same stream again.
                continue;
            }
            if (!parsedObjectsByOffset.TryGetValue(offset, out PdfIndirectObject? parsed)) {
                parsed = TryParseIndirectObjectAt(pdf, text, offset, map, scanBudget, out PdfIndirectObject candidate, cancellationToken)
                    ? candidate
                    : null;
                parsedObjectsByOffset[offset] = parsed;
            }

            if (parsed is not null &&
                parsed.ObjectNumber == entry.ObjectNumber &&
                parsed.Generation == generation) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = offset;
            }
        }

        // Type 2 entries are applied in one dedicated pass after any Standard Security
        // decryption has replaced encrypted object streams in the active object map.
    }

    private static List<(int Offset, (int ObjectNumber, int Offset, int Generation, bool InUse)[] Entries, int? XrefStreamOffset)> GetClassicPredecessorTablesForXrefStreamChain(
        string text,
        List<(int ObjectNumber, int Offset, PdfStream Stream)> xrefStreams,
        int activeXrefOffset,
        System.Threading.CancellationToken cancellationToken) {
        var byOffset = new Dictionary<int, PdfStream>();
        foreach (var xrefStream in xrefStreams) {
            cancellationToken.ThrowIfCancellationRequested();
            byOffset[xrefStream.Offset] = xrefStream.Stream;
        }

        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfStream? stream) &&
            visited.Add(currentOffset) &&
            visited.Count < 64) {
            cancellationToken.ThrowIfCancellationRequested();
            if (stream.Dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                return new List<(int Offset, (int ObjectNumber, int Offset, int Generation, bool InUse)[] Entries, int? XrefStreamOffset)>();
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        return GetClassicXrefTableChain(text, currentOffset, cancellationToken);
    }

    private static List<int> GetXrefStreamChainOffsets(List<(int ObjectNumber, int Offset, PdfStream Stream)> xrefStreams, int activeXrefOffset,
        System.Threading.CancellationToken cancellationToken) {
        var byOffset = new Dictionary<int, PdfStream>();
        foreach (var xrefStream in xrefStreams) {
            cancellationToken.ThrowIfCancellationRequested();
            byOffset[xrefStream.Offset] = xrefStream.Stream;
        }

        var newestToOldest = new List<int>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfStream? stream) &&
            visited.Add(currentOffset) &&
            newestToOldest.Count < 64) {
            cancellationToken.ThrowIfCancellationRequested();
            newestToOldest.Add(currentOffset);
            if (stream.Dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                break;
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        newestToOldest.Reverse();
        return newestToOldest;
    }

    private static bool TryGetLatestStartXrefOffset(string text, out int offset,
        System.Threading.CancellationToken cancellationToken = default) {
        offset = 0;
        int startXrefIndex = LastIndexOfTrailerMarker(text, "startxref", StringComparison.Ordinal, cancellationToken);
        if (startXrefIndex < 0) {
            return false;
        }

        int index = startXrefIndex + "startxref".Length;
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            index++;
        }

        long value = 0;
        int firstDigit = index;
        while (index < text.Length && char.IsDigit(text[index])) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            value = (value * 10) + (text[index] - '0');
            if (value > int.MaxValue) {
                return false;
            }

            index++;
        }

        if (index == firstDigit) {
            return false;
        }

        offset = (int)value;
        return true;
    }

    private static bool TryParseIndirectObjectAt(byte[] pdf, string text, int offset, Dictionary<int, PdfIndirectObject> map, XrefObjectScanBudget scanBudget, out PdfIndirectObject parsed,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        parsed = null!;
        if (offset < 0 || offset >= text.Length) {
            return false;
        }

        int scanLimit = (int)Math.Min(
            (long)text.Length,
            (long)offset + scanBudget.MaximumPerObject);
        int chargedThrough = offset;
        void ChargeThrough(int scannedThrough) {
            int bounded = Math.Max(offset, Math.Min(scanLimit, scannedThrough));
            scanBudget.Charge(bounded - chargedThrough);
            chargedThrough = Math.Max(chargedThrough, bounded);
        }
        if (!TryReadIndirectObjectHeaderAt(text, offset, scanLimit, out IndirectObjectHeader header, cancellationToken: cancellationToken)) {
            ChargeThrough(Math.Min(scanLimit, offset + 128));
            return false;
        }

        int id = header.ObjectNumber;
        int gen = header.Generation;
        int start = header.Index;
        int bodyStart = header.Index + header.Length;
        int valueStart = bodyStart;
        while (valueStart < scanLimit && char.IsWhiteSpace(text[valueStart])) {
            if ((valueStart & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            valueStart++;
        }

        if (valueStart + 1 < scanLimit && text[valueStart] == '<' && text[valueStart + 1] == '<') {
            int dictStart = valueStart;
            int dictEnd = FindDictEnd(text, dictStart, scanLimit, cancellationToken);
            ChargeThrough(dictEnd > dictStart ? dictEnd + 2 : scanLimit);
            if (dictEnd > dictStart) {
                string dictText = SafeSliceCancellable(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000, cancellationToken);
                PdfDictionary? dict;
                try { dict = ParseDictionary(dictText, cancellationToken: cancellationToken); }
                catch (Exception ex) when (ex is not OutOfMemoryException and not OperationCanceledException) { dict = null; }
                if (dict is null) {
                    return false;
                }

                int streamKw = dictEnd;
                while (streamKw < scanLimit && char.IsWhiteSpace(text[streamKw])) {
                    if ((streamKw & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    streamKw++;
                }

                bool hasStream = streamKw + 6 <= scanLimit &&
                    string.CompareOrdinal(text, streamKw, "stream", 0, 6) == 0 &&
                    HasKeywordBoundary(text, streamKw + 6, bodyStart, scanLimit);
                if (hasStream) {
                    int dataStart = SkipEOL(text, streamKw + 6, scanLimit);
                    int byteLen = -1;
                    TryGetResolvedLength(dict, map, out byteLen);
                    if (byteLen < 0) {
                        int fallbackEnd = FindObjectEnd(text, start, maximumIndex: scanLimit, cancellationToken: cancellationToken);
                        ChargeThrough(fallbackEnd > start ? fallbackEnd : scanLimit);
                        if (fallbackEnd < 0) {
                            return false;
                        }

                        int endStream = IndexOfKeywordCancellable(text, "endstream", dataStart, fallbackEnd, cancellationToken);
                        if (endStream > dataStart) byteLen = endStream - dataStart;
                    }

                    if (byteLen >= 0 && dataStart >= 0 && dataStart + byteLen <= pdf.Length) {
                        byte[] data = CopyBytes(pdf, dataStart, byteLen, cancellationToken);
                        parsed = new PdfIndirectObject(id, gen, new PdfStream(dict, data));
                        return true;
                    }
                }

                parsed = new PdfIndirectObject(id, gen, dict);
                return true;
            }
        }

        int end = FindObjectEnd(text, start, maximumIndex: scanLimit, cancellationToken: cancellationToken);
        ChargeThrough(end > start ? end : scanLimit);
        if (end < 0) {
            return false;
        }

        int bodyEnd = end;
        if (bodyEnd - 6 >= bodyStart && string.Equals(text.Substring(bodyEnd - 6, 6), "endobj", StringComparison.Ordinal)) {
            bodyEnd -= 6;
        }

        string body = SafeTrimmedSliceCancellable(text, bodyStart, bodyEnd - bodyStart, 1_000_000, cancellationToken);
        var topLevelObject = ParseTopLevelObject(body, cancellationToken: cancellationToken);
        if (topLevelObject is null) {
            return false;
        }

        parsed = new PdfIndirectObject(id, gen, topLevelObject);
        return true;
    }

    private static bool TryParseObjectFromObjectStream(
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int objectStreamNumber,
        int objectStreamIndex,
        int expectedObjectNumber,
        PdfReadLimits limits,
        PdfDecodedStreamBudget decodedStreamBudget,
        out PdfIndirectObject parsed,
        out int objectStreamOffset,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        parsed = null!;
        objectStreamOffset = int.MaxValue;
        if (!map.TryGetValue(objectStreamNumber, out var objectStreamIndirect) ||
            objectStreamIndirect.Value is not PdfStream objectStream ||
            objectStream.Dictionary.Get<PdfName>("Type")?.Name != "ObjStm") {
            return false;
        }

        byte[] data = decodedStreamBudget.Decode(objectStream, map, int.MaxValue, cancellationToken);
        if (!TryReadObjectStreamLayout(objectStream.Dictionary, data.Length, limits, out int n, out int first)
            || objectStreamIndex < 0 || objectStreamIndex >= n) {
            return false;
        }

        byte[] headerBytes = CopyBytes(data, 0, first, cancellationToken);
        string header = PdfEncoding.Latin1GetStringCancellable(headerBytes, cancellationToken);
        var pairs = ParsePairs(header, n, out bool completeHeader, cancellationToken);
        if (!completeHeader ||
            pairs[objectStreamIndex].Obj != expectedObjectNumber) {
            return false;
        }

        int start = first + pairs[objectStreamIndex].Off;
        int end = (objectStreamIndex + 1 < n) ? first + pairs[objectStreamIndex + 1].Off : data.Length;
        if (start < first || end > data.Length || end <= start) {
            return false;
        }

        int len = end - start;
        byte[] sliceBytes = CopyBytes(data, start, len, cancellationToken);
        var slice = PdfEncoding.Latin1GetStringCancellable(sliceBytes, cancellationToken);
        var parsedObject = ParseTopLevelObject(
            slice,
            limits,
            trackEncodedStringSourceSpans: false,
            cancellationToken: cancellationToken);
        if (parsedObject is null) {
            return false;
        }

        parsed = new PdfIndirectObject(expectedObjectNumber, 0, parsedObject);
        objectStreamOffset = parsedOffsets.TryGetValue(objectStreamNumber, out int offset) ? offset : int.MaxValue;
        return true;
    }

    private readonly struct XrefStreamEntry {
        public XrefStreamEntry(int objectNumber, long type, long field1, long field2) {
            ObjectNumber = objectNumber;
            Type = type;
            Field1 = field1;
            Field2 = field2;
        }

        public int ObjectNumber { get; }
        public long Type { get; }
        public long Field1 { get; }
        public long Field2 { get; }
    }

    private static void SortXrefStreams(
        List<(int ObjectNumber, int Offset, PdfStream Stream)> streams,
        System.Threading.CancellationToken cancellationToken) {
        try {
            streams.Sort((left, right) => {
                cancellationToken.ThrowIfCancellationRequested();
                return left.Offset.CompareTo(right.Offset);
            });
        } catch (InvalidOperationException error) when (error.InnerException is OperationCanceledException) {
            throw error.InnerException!;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }
}
