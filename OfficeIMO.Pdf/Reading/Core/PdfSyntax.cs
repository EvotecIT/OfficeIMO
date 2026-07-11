using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"<<(.*?)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"<<(.*?)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline, RegexTimeout);
#endif
    private static readonly Regex TrailerRootRegex = new Regex(@"/Root\s+(\d+)\s+(\d+)\s+R", RegexOptions.Compiled, RegexTimeout);

    internal static (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) ParseObjects(byte[] pdf) {
        return ParseObjects(pdf, null);
    }

    internal static (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) ParseObjects(byte[] pdf, PdfReadOptions? options) {
        PdfReadLimits limits = options?.Limits ?? new PdfReadLimits();
        limits.Validate();
        if (pdf.LongLength > limits.MaxInputBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limits.MaxInputBytes, pdf.LongLength);
        }

        var parseTimer = System.Diagnostics.Stopwatch.StartNew();
        string text = PdfEncoding.Latin1GetString(pdf);
        var map = new Dictionary<int, PdfIndirectObject>();
        var parsedOffsets = new Dictionary<int, int>();
        var streamLocations = new List<(int Id, int Generation, int DataStart)>();
        var matches = ObjRegex.Matches(text);
        if (matches.Count > limits.MaxIndirectObjects) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.IndirectObjects, limits.MaxIndirectObjects, matches.Count);
        }

        for (int i = 0; i < matches.Count; i++) {
            if ((i & 127) == 0) {
                ThrowIfParsingTimeExceeded(parseTimer, limits);
            }

            int id = int.Parse(matches[i].Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            int gen = int.Parse(matches[i].Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
            int start = matches[i].Index;
            int bodyStart = matches[i].Index + matches[i].Length;
            int end = FindObjectEnd(text, start);
            if (end < 0) end = (i + 1 < matches.Count) ? matches[i + 1].Index : text.Length;

            int preliminaryBodyEnd = end;
            if (preliminaryBodyEnd - 6 >= bodyStart && string.Equals(text.Substring(preliminaryBodyEnd - 6, 6), "endobj", StringComparison.Ordinal)) {
                preliminaryBodyEnd -= 6;
            }

            int preliminaryBodyCharacters = preliminaryBodyEnd - bodyStart;
            int firstBodyCharacter = bodyStart;
            while (firstBodyCharacter < preliminaryBodyEnd && char.IsWhiteSpace(text[firstBodyCharacter])) {
                firstBodyCharacter++;
            }

            if (firstBodyCharacter < preliminaryBodyEnd && text[firstBodyCharacter] == '[') {
                if (preliminaryBodyCharacters > limits.MaxObjectCharacters) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, limits.MaxObjectCharacters, preliminaryBodyCharacters);
                }

                string preliminaryArrayBody = SafeSlice(text, bodyStart, preliminaryBodyCharacters, limits.MaxObjectCharacters).Trim();
                var parsedArray = ParseTopLevelObject(preliminaryArrayBody, limits);
                if (parsedArray is not null) {
                    map[id] = new PdfIndirectObject(id, gen, parsedArray);
                    parsedOffsets[id] = start;
                    continue;
                }
            }

            // Extract dictionary (balanced << >>) within object bounds
            int dictStart = text.IndexOf("<<", start, end - start, System.StringComparison.Ordinal);
            if (dictStart >= 0) {
                int dictEnd = FindDictEnd(text, dictStart, end);
                if (dictEnd > dictStart) {
                    int dictionaryCharacters = dictEnd - (dictStart + 2);
                    if (dictionaryCharacters > limits.MaxObjectCharacters) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, limits.MaxObjectCharacters, dictionaryCharacters);
                    }

                    string dictText = SafeSlice(text, dictStart + 2, dictionaryCharacters, limits.MaxObjectCharacters);
                    PdfDictionary? dict;
                    try { dict = ParseDictionary(dictText, limits); }
                    catch (Exception ex) when (ex is not OutOfMemoryException && ex is not PdfReadLimitException) { dict = null; }
                    if (dict is null) {
                        continue;
                    }

                    // Check for stream section; prefer dictionary /Length when available
                    int streamKw = IndexOfKeyword(text, "stream", dictEnd, end);
                    if (streamKw >= 0) {
                        int dataStart = SkipEOL(text, streamKw + 6, end);
                        streamLocations.Add((id, gen, dataStart));
                        // Try /Length first (inline number only)
                        int byteStart = dataStart;
                        int byteLen = -1;
                        TryGetResolvedLength(dict, map, out byteLen);
                        if (byteLen < 0) {
                            int endStream = IndexOfKeyword(text, "endstream", dataStart, end);
                            if (endStream > dataStart) byteLen = endStream - dataStart;
                        }
                        if (byteLen >= 0) {
                            if (byteLen > limits.MaxRawStreamBytes) {
                                throw PdfReadLimitException.Create(PdfReadLimitKind.RawStreamBytes, limits.MaxRawStreamBytes, byteLen);
                            }

                            if (byteStart >= 0 && byteLen >= 0 && byteStart + byteLen <= pdf.Length) {
                                var data = new byte[byteLen];
                                Buffer.BlockCopy(pdf, byteStart, data, 0, byteLen);
                                map[id] = new PdfIndirectObject(id, gen, new PdfStream(dict, data));
                                parsedOffsets[id] = start;
                                continue;
                            }
                        }
                    }
                    // No stream; store dictionary-only object
                    map[id] = new PdfIndirectObject(id, gen, dict);
                    parsedOffsets[id] = start;
                }
            }

            if (!map.ContainsKey(id)) {
                if (preliminaryBodyCharacters > limits.MaxObjectCharacters) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, limits.MaxObjectCharacters, preliminaryBodyCharacters);
                }

                string preliminaryBody = SafeSlice(text, bodyStart, preliminaryBodyCharacters, limits.MaxObjectCharacters).Trim();
                var parsed = ParseTopLevelObject(preliminaryBody, limits);
                if (parsed is not null) {
                    map[id] = new PdfIndirectObject(id, gen, parsed);
                    parsedOffsets[id] = start;
                }
            }
        }
        ResolveIndirectStreamLengths(map, pdf, streamLocations);
        var activeClassicObjectNumbers = new HashSet<int>();
        bool appliedXrefStreamEntries = ApplyClassicXrefEntries(map, pdf, parsedOffsets, activeClassicObjectNumbers, out bool appliedClassicEntries);
        appliedXrefStreamEntries = ApplyXrefStreamEntries(map, pdf, parsedOffsets) || appliedXrefStreamEntries;
        string trailerRaw = GetActiveTrailerRaw(text, map, parsedOffsets);
        PdfStandardSecurityHandler? decryptor = null;
        int? encryptObjectNumber = TryReadLastReferenceObjectNumber(trailerRaw, "Encrypt");
        if (encryptObjectNumber.HasValue) {
            TryCreateDecryptor(map, trailerRaw, options, out decryptor);
            if (decryptor is not null) {
                DecryptObjects(map, decryptor, encryptObjectNumber.Value);
                if (appliedXrefStreamEntries) {
                    ApplyCompressedXrefStreamEntries(map, pdf, parsedOffsets);
                }
            }
        }

        if (!appliedXrefStreamEntries) {
            // Compatibility fallback for simple parser-supported files whose compressed objects are only discoverable by scanning.
            ExpandObjectStreams(map, pdf, parsedOffsets, appliedClassicEntries ? activeClassicObjectNumbers : null);
        }

        if (decryptor is null) {
            ThrowIfEncryptedXrefStream(map);
        }

        if (map.Count > limits.MaxIndirectObjects) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.IndirectObjects, limits.MaxIndirectObjects, map.Count);
        }

        ThrowIfParsingTimeExceeded(parseTimer, limits);

        return (map, trailerRaw);
    }

    private static void ThrowIfParsingTimeExceeded(System.Diagnostics.Stopwatch timer, PdfReadLimits limits) {
        if (timer.Elapsed > limits.MaxObjectParsingTime) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.ObjectParsingTime,
                (long)limits.MaxObjectParsingTime.TotalMilliseconds,
                (long)timer.Elapsed.TotalMilliseconds);
        }
    }

}
