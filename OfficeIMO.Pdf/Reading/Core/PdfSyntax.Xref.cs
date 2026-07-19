namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static void ResolveIndirectStreamLengths(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        List<(int Id, int Generation, int DataStart)> streamLocations,
        PdfReadLimits limits) {
        foreach (var streamLocation in streamLocations) {
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

            if (stream.Data.Length == byteLen) {
                continue;
            }

            if (byteLen > limits.MaxRawStreamBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.RawStreamBytes, limits.MaxRawStreamBytes, byteLen);
            }

            var data = new byte[byteLen];
            Buffer.BlockCopy(pdf, byteStart, data, 0, byteLen);
            map[streamLocation.Id] = new PdfIndirectObject(streamLocation.Id, streamLocation.Generation, new PdfStream(stream.Dictionary, data, stream.DecodingFailed, stream.DecodingError));
        }
    }

    private static bool ApplyClassicXrefEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        HashSet<int> activeObjectNumbers,
        PdfReadLimits limits,
        out bool appliedClassicEntries) {
        appliedClassicEntries = false;
        string text = PdfEncoding.Latin1GetString(pdf);
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset)) {
            return false;
        }

        var tables = GetClassicXrefTableChain(text, activeXrefOffset);
        if (tables.Count == 0) {
            return false;
        }

        appliedClassicEntries = true;
        bool appliedXrefStream = false;
        foreach (var table in tables) {
            ApplyClassicXrefTableEntries(map, pdf, parsedOffsets, text, table.Entries, activeObjectNumbers);
            if (table.XrefStreamOffset.HasValue) {
                appliedXrefStream = ApplyXrefStreamAtOffset(map, pdf, parsedOffsets, text, table.XrefStreamOffset.Value, limits) || appliedXrefStream;
            }
        }

        return appliedXrefStream;
    }

    private static void ApplyClassicXrefTableEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        List<(int ObjectNumber, int Offset, int Generation, bool InUse)> entries,
        HashSet<int>? activeObjectNumbers = null) {
        foreach (var entry in entries) {
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

            if (TryParseIndirectObjectAt(pdf, text, entry.Offset, map, out var parsed) &&
                parsed.ObjectNumber == entry.ObjectNumber &&
                parsed.Generation == entry.Generation) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = entry.Offset;
                activeObjectNumbers?.Add(entry.ObjectNumber);
            }
        }
    }

    private static List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)> GetClassicXrefTableChain(string text, int activeXrefOffset) {
        var newestToOldest = new List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (visited.Add(currentOffset) &&
            newestToOldest.Count < 64 &&
            TryParseClassicXrefTable(text, currentOffset, out var entries, out int? previousOffset, out _, out int? xrefStreamOffset)) {
            newestToOldest.Add((currentOffset, entries, xrefStreamOffset));
            if (!previousOffset.HasValue) {
                break;
            }

            currentOffset = previousOffset.Value;
        }

        newestToOldest.Reverse();
        return newestToOldest;
    }

    private static bool TryParseClassicXrefTable(string text, int offset, out List<(int ObjectNumber, int Offset, int Generation, bool InUse)> entries, out int? previousOffset, out string trailerRaw, out int? xrefStreamOffset) {
        entries = new List<(int ObjectNumber, int Offset, int Generation, bool InUse)>();
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

        int trailerIndex = IndexOfKeyword(text, "trailer", offset + 4, text.Length);
        if (trailerIndex < 0) {
            return false;
        }

        string section = SafeSlice(text, offset + 4, trailerIndex - (offset + 4), 2_000_000);
        using (var reader = new StringReader(section)) {
            string? line;
            while ((line = reader.ReadLine()) is not null) {
                string[] headerParts = SplitWhitespace(line);
                if (headerParts.Length < 2 ||
                    !int.TryParse(headerParts[0], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int firstObjectNumber) ||
                    !int.TryParse(headerParts[1], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int count) ||
                    firstObjectNumber < 0 ||
                    count <= 0 ||
                    count > 1_000_000) {
                    continue;
                }

                for (int i = 0; i < count; i++) {
                    string? entryLine = reader.ReadLine();
                    if (entryLine is null) {
                        return entries.Count > 0;
                    }

                    string[] entryParts = SplitWhitespace(entryLine);
                    if (entryParts.Length < 3 ||
                        !int.TryParse(entryParts[0], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int objectOffset) ||
                        !int.TryParse(entryParts[1], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int generation)) {
                        continue;
                    }

                    if (string.Equals(entryParts[2], "n", StringComparison.Ordinal)) {
                        entries.Add((firstObjectNumber + i, objectOffset, generation, true));
                    } else if (string.Equals(entryParts[2], "f", StringComparison.Ordinal)) {
                        entries.Add((firstObjectNumber + i, objectOffset, generation, false));
                    }
                }
            }
        }

        if (entries.Count == 0) {
            return false;
        }

        int dictStart = text.IndexOf("<<", trailerIndex, StringComparison.Ordinal);
        if (dictStart >= 0) {
            int dictEnd = FindDictEnd(text, dictStart, text.Length);
            if (dictEnd > dictStart) {
                trailerRaw = SafeSlice(text, trailerIndex, dictEnd - trailerIndex, 1_000_000);
                string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                try {
                    PdfDictionary trailer = ParseDictionary(dictText);
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
                } catch (Exception ex) when (ex is not OutOfMemoryException) {
                    previousOffset = null;
                    xrefStreamOffset = null;
                }
            }
        }

        return true;
    }

    private static string[] SplitWhitespace(string value) {
        return value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
    }

    private static bool ApplyXrefStreamEntries(Dictionary<int, PdfIndirectObject> map, byte[] pdf, Dictionary<int, int> parsedOffsets, PdfReadLimits limits) {
        var xrefStreams = new List<(int ObjectNumber, int Offset, PdfStream Stream)>();
        foreach (var entry in map.Values) {
            if (entry.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                int offset = parsedOffsets.TryGetValue(entry.ObjectNumber, out int parsedOffset) ? parsedOffset : int.MaxValue;
                xrefStreams.Add((entry.ObjectNumber, offset, stream));
            }
        }

        if (xrefStreams.Count == 0) {
            return false;
        }

        string text = PdfEncoding.Latin1GetString(pdf);
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset)) {
            return false;
        }

        xrefStreams.Sort(static (left, right) => left.Offset.CompareTo(right.Offset));
        var activeChainOffsets = GetXrefStreamChainOffsets(xrefStreams, activeXrefOffset);
        if (activeChainOffsets.Count == 0) {
            return false;
        }

        var classicPredecessors = GetClassicPredecessorTablesForXrefStreamChain(text, xrefStreams, activeXrefOffset);
        foreach (var table in classicPredecessors) {
            ApplyClassicXrefTableEntries(map, pdf, parsedOffsets, text, table.Entries);
            if (table.XrefStreamOffset.HasValue) {
                ApplyXrefStreamAtOffset(map, pdf, parsedOffsets, text, table.XrefStreamOffset.Value, limits);
            }
        }

        foreach (int chainOffset in activeChainOffsets) {
            var xrefStream = xrefStreams.First(item => item.Offset == chainOffset);
            ApplyXrefStreamObjectEntries(map, pdf, parsedOffsets, text, xrefStream.Stream, limits);
        }

        return true;
    }

    private static bool ApplyCompressedXrefStreamEntries(Dictionary<int, PdfIndirectObject> map, byte[] pdf, Dictionary<int, int> parsedOffsets, PdfReadLimits limits) {
        var xrefStreams = new List<(int ObjectNumber, int Offset, PdfStream Stream)>();
        foreach (var entry in map.Values) {
            if (entry.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                int offset = parsedOffsets.TryGetValue(entry.ObjectNumber, out int parsedOffset) ? parsedOffset : int.MaxValue;
                xrefStreams.Add((entry.ObjectNumber, offset, stream));
            }
        }

        if (xrefStreams.Count == 0) {
            return false;
        }

        string text = PdfEncoding.Latin1GetString(pdf);
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset)) {
            return false;
        }

        xrefStreams.Sort(static (left, right) => left.Offset.CompareTo(right.Offset));
        var activeChainOffsets = GetXrefStreamChainOffsets(xrefStreams, activeXrefOffset);
        var activeEntries = new Dictionary<int, XrefStreamEntry>();
        if (activeChainOffsets.Count == 0) {
            var classicTables = GetClassicXrefTableChain(text, activeXrefOffset);
            if (classicTables.Count == 0) {
                return false;
            }

            foreach (var table in classicTables) {
                if (table.XrefStreamOffset.HasValue) {
                    var xrefStream = xrefStreams.FirstOrDefault(item => item.Offset == table.XrefStreamOffset.Value);
                    if (xrefStream.Stream is not null) {
                        UpdateActiveCompressedEntries(activeEntries, xrefStream.Stream, map, limits);
                    }
                }

                for (int i = 0; i < table.Entries.Count; i++) {
                    activeEntries.Remove(table.Entries[i].ObjectNumber);
                }
            }
        } else {
            foreach (int chainOffset in activeChainOffsets) {
                var xrefStream = xrefStreams.First(item => item.Offset == chainOffset);
                UpdateActiveCompressedEntries(activeEntries, xrefStream.Stream, map, limits);
            }
        }

        bool applied = false;
        foreach (XrefStreamEntry entry in activeEntries.Values) {
            if (entry.Field1 < 0 ||
                entry.Field1 > int.MaxValue ||
                entry.Field2 < 0 ||
                entry.Field2 > int.MaxValue) {
                continue;
            }

            int objectStreamNumber = (int)entry.Field1;
            int objectStreamIndex = (int)entry.Field2;
            if (TryParseObjectFromObjectStream(map, parsedOffsets, objectStreamNumber, objectStreamIndex, entry.ObjectNumber, limits, out PdfIndirectObject parsed, out int objectStreamOffset)) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = objectStreamOffset;
                applied = true;
            }
        }

        return applied;
    }

    private static void UpdateActiveCompressedEntries(
        Dictionary<int, XrefStreamEntry> activeEntries,
        PdfStream xrefStream,
        Dictionary<int, PdfIndirectObject> map,
        PdfReadLimits limits) {
        byte[] data = Filters.StreamDecoder.Decode(xrefStream.Dictionary, xrefStream.Data, map, limits.MaxDecodedStreamBytes);
        foreach (XrefStreamEntry entry in ReadXrefStreamEntries(xrefStream.Dictionary, data)) {
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
        PdfReadLimits limits) {
        PdfStream? targetStream = null;
        foreach (var entry in map.Values) {
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

        ApplyXrefStreamObjectEntries(map, pdf, parsedOffsets, text, targetStream, limits);
        return true;
    }

    private static void ApplyXrefStreamObjectEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        PdfStream xrefStream,
        PdfReadLimits limits) {
        byte[] data = Filters.StreamDecoder.Decode(xrefStream.Dictionary, xrefStream.Data, map, limits.MaxDecodedStreamBytes);
        var entries = ReadXrefStreamEntries(xrefStream.Dictionary, data).ToList();
        foreach (var entry in entries) {
            if (entry.Type == 0 &&
                entry.ObjectNumber != 0) {
                map.Remove(entry.ObjectNumber);
                parsedOffsets.Remove(entry.ObjectNumber);
            }
        }

        foreach (var entry in entries) {
            if (entry.Type != 1 ||
                entry.Field1 < 0 ||
                entry.Field1 > int.MaxValue ||
                entry.Field2 < 0 ||
                entry.Field2 > int.MaxValue) {
                continue;
            }

            int offset = (int)entry.Field1;
            int generation = (int)entry.Field2;
            if (TryParseIndirectObjectAt(pdf, text, offset, map, out var parsed) &&
                parsed.ObjectNumber == entry.ObjectNumber &&
                parsed.Generation == generation) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = offset;
            }
        }

        foreach (var entry in entries) {
            if (entry.Type != 2 ||
                entry.Field1 < 0 ||
                entry.Field1 > int.MaxValue ||
                entry.Field2 < 0 ||
                entry.Field2 > int.MaxValue) {
                continue;
            }

            int objectStreamNumber = (int)entry.Field1;
            int objectStreamIndex = (int)entry.Field2;
            if (TryParseObjectFromObjectStream(map, parsedOffsets, objectStreamNumber, objectStreamIndex, entry.ObjectNumber, limits, out var parsed, out int objectStreamOffset)) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = objectStreamOffset;
            }
        }
    }

    private static List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)> GetClassicPredecessorTablesForXrefStreamChain(
        string text,
        List<(int ObjectNumber, int Offset, PdfStream Stream)> xrefStreams,
        int activeXrefOffset) {
        var byOffset = new Dictionary<int, PdfStream>();
        foreach (var xrefStream in xrefStreams) {
            byOffset[xrefStream.Offset] = xrefStream.Stream;
        }

        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfStream? stream) &&
            visited.Add(currentOffset) &&
            visited.Count < 64) {
            if (stream.Dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                return new List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)>();
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        return GetClassicXrefTableChain(text, currentOffset);
    }

    private static List<int> GetXrefStreamChainOffsets(List<(int ObjectNumber, int Offset, PdfStream Stream)> xrefStreams, int activeXrefOffset) {
        var byOffset = new Dictionary<int, PdfStream>();
        foreach (var xrefStream in xrefStreams) {
            byOffset[xrefStream.Offset] = xrefStream.Stream;
        }

        var newestToOldest = new List<int>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfStream? stream) &&
            visited.Add(currentOffset) &&
            newestToOldest.Count < 64) {
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

    private static bool TryGetLatestStartXrefOffset(string text, out int offset) {
        offset = 0;
        int startXrefIndex = text.LastIndexOf("startxref", StringComparison.Ordinal);
        if (startXrefIndex < 0) {
            return false;
        }

        int index = startXrefIndex + "startxref".Length;
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }

        long value = 0;
        int firstDigit = index;
        while (index < text.Length && char.IsDigit(text[index])) {
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

    private static IEnumerable<XrefStreamEntry> ReadXrefStreamEntries(PdfDictionary dictionary, byte[] data) {
        if (data.Length == 0 ||
            dictionary.Get<PdfArray>("W") is not PdfArray widthsArray ||
            widthsArray.Items.Count < 3) {
            yield break;
        }

        int w0 = GetNonNegativeInt(widthsArray.Items[0]);
        int w1 = GetNonNegativeInt(widthsArray.Items[1]);
        int w2 = GetNonNegativeInt(widthsArray.Items[2]);
        int entryWidth = w0 + w1 + w2;
        if (entryWidth <= 0) {
            yield break;
        }

        var ranges = GetXrefIndexRanges(dictionary);
        int dataOffset = 0;
        foreach (var range in ranges) {
            for (int i = 0; i < range.Count; i++) {
                if (dataOffset + entryWidth > data.Length) {
                    yield break;
                }

                long type = w0 == 0 ? 1 : ReadBigEndian(data, dataOffset, w0);
                dataOffset += w0;
                long field1 = ReadBigEndian(data, dataOffset, w1);
                dataOffset += w1;
                long field2 = ReadBigEndian(data, dataOffset, w2);
                dataOffset += w2;

                yield return new XrefStreamEntry(range.FirstObjectNumber + i, type, field1, field2);
            }
        }
    }

    private static List<(int FirstObjectNumber, int Count)> GetXrefIndexRanges(PdfDictionary dictionary) {
        var ranges = new List<(int, int)>();
        if (dictionary.Get<PdfArray>("Index") is PdfArray indexArray && indexArray.Items.Count >= 2) {
            for (int i = 0; i + 1 < indexArray.Items.Count; i += 2) {
                int first = GetNonNegativeInt(indexArray.Items[i]);
                int count = GetNonNegativeInt(indexArray.Items[i + 1]);
                if (count > 0) {
                    ranges.Add((first, count));
                }
            }
        }

        if (ranges.Count == 0) {
            int size = GetNonNegativeInt(dictionary.Get<PdfNumber>("Size"));
            if (size > 0) {
                ranges.Add((0, size));
            }
        }

        return ranges;
    }

    private static int GetNonNegativeInt(PdfObject? value) {
        if (value is not PdfNumber number || number.Value <= 0) {
            return 0;
        }

        return (int)Math.Min(int.MaxValue, Math.Floor(number.Value));
    }

    private static long ReadBigEndian(byte[] data, int offset, int length) {
        long value = 0;
        for (int i = 0; i < length; i++) {
            value = (value << 8) | data[offset + i];
        }

        return value;
    }

    private static bool TryParseIndirectObjectAt(byte[] pdf, string text, int offset, Dictionary<int, PdfIndirectObject> map, out PdfIndirectObject parsed) {
        parsed = null!;
        if (offset < 0 || offset >= text.Length) {
            return false;
        }

        if (!TryReadIndirectObjectHeaderAt(text, offset, text.Length, out IndirectObjectHeader header)) {
            return false;
        }

        int id = header.ObjectNumber;
        int gen = header.Generation;
        int start = header.Index;
        int bodyStart = header.Index + header.Length;
        int valueStart = bodyStart;
        while (valueStart < text.Length && char.IsWhiteSpace(text[valueStart])) {
            valueStart++;
        }

        if (valueStart + 1 < text.Length && text[valueStart] == '<' && text[valueStart + 1] == '<') {
            int dictStart = valueStart;
            int dictEnd = FindDictEnd(text, dictStart, text.Length);
            if (dictEnd > dictStart) {
                string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                PdfDictionary? dict;
                try { dict = ParseDictionary(dictText); }
                catch (Exception ex) when (ex is not OutOfMemoryException) { dict = null; }
                if (dict is null) {
                    return false;
                }

                int streamKw = dictEnd;
                while (streamKw < text.Length && char.IsWhiteSpace(text[streamKw])) {
                    streamKw++;
                }

                bool hasStream = streamKw + 6 <= text.Length &&
                    string.CompareOrdinal(text, streamKw, "stream", 0, 6) == 0 &&
                    HasKeywordBoundary(text, streamKw + 6, bodyStart, text.Length);
                if (hasStream) {
                    int dataStart = SkipEOL(text, streamKw + 6, text.Length);
                    int byteLen = -1;
                    TryGetResolvedLength(dict, map, out byteLen);
                    if (byteLen < 0) {
                        int fallbackEnd = FindObjectEnd(text, start);
                        if (fallbackEnd < 0) {
                            return false;
                        }

                        int endStream = IndexOfKeyword(text, "endstream", dataStart, fallbackEnd);
                        if (endStream > dataStart) byteLen = endStream - dataStart;
                    }

                    if (byteLen >= 0 && dataStart >= 0 && dataStart + byteLen <= pdf.Length) {
                        var data = new byte[byteLen];
                        Buffer.BlockCopy(pdf, dataStart, data, 0, byteLen);
                        parsed = new PdfIndirectObject(id, gen, new PdfStream(dict, data));
                        return true;
                    }
                }

                parsed = new PdfIndirectObject(id, gen, dict);
                return true;
            }
        }

        int end = FindObjectEnd(text, start);
        if (end < 0) {
            return false;
        }

        int bodyEnd = end;
        if (bodyEnd - 6 >= bodyStart && string.Equals(text.Substring(bodyEnd - 6, 6), "endobj", StringComparison.Ordinal)) {
            bodyEnd -= 6;
        }

        string body = SafeSlice(text, bodyStart, bodyEnd - bodyStart, 1_000_000).Trim();
        var topLevelObject = ParseTopLevelObject(body);
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
        out PdfIndirectObject parsed,
        out int objectStreamOffset) {
        parsed = null!;
        objectStreamOffset = int.MaxValue;
        if (!map.TryGetValue(objectStreamNumber, out var objectStreamIndirect) ||
            objectStreamIndirect.Value is not PdfStream objectStream ||
            objectStream.Dictionary.Get<PdfName>("Type")?.Name != "ObjStm") {
            return false;
        }

        byte[] data = Filters.StreamDecoder.Decode(objectStream.Dictionary, objectStream.Data, map, limits.MaxDecodedStreamBytes);
        int n = (int)(objectStream.Dictionary.Get<PdfNumber>("N")?.Value ?? 0);
        int first = (int)(objectStream.Dictionary.Get<PdfNumber>("First")?.Value ?? 0);
        if (objectStreamIndex < 0 || objectStreamIndex >= n || n <= 0 || first <= 0 || first > data.Length) {
            return false;
        }

        var headerBytes = new byte[first];
        Buffer.BlockCopy(data, 0, headerBytes, 0, first);
        string header = PdfEncoding.Latin1GetString(headerBytes);
        var pairs = ParsePairs(header, n);
        if (pairs.Count != n ||
            pairs[objectStreamIndex].Obj != expectedObjectNumber) {
            return false;
        }

        int start = first + pairs[objectStreamIndex].Off;
        int end = (objectStreamIndex + 1 < n) ? first + pairs[objectStreamIndex + 1].Off : data.Length;
        if (start < 0 || end > data.Length || end <= start) {
            return false;
        }

        int len = end - start;
        var sliceBytes = new byte[len];
        Buffer.BlockCopy(data, start, sliceBytes, 0, len);
        var slice = PdfEncoding.Latin1GetString(sliceBytes);
        var parsedObject = ParseTopLevelObject(
            slice,
            limits,
            trackEncodedStringSourceSpans: false);
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
}
