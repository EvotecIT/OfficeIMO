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
        string text = PdfEncoding.Latin1GetString(pdf);
        var map = new Dictionary<int, PdfIndirectObject>();
        var parsedOffsets = new Dictionary<int, int>();
        var streamLocations = new List<(int Id, int Generation, int DataStart)>();
        var matches = ObjRegex.Matches(text);
        for (int i = 0; i < matches.Count; i++) {
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

            string preliminaryBody = SafeSlice(text, bodyStart, preliminaryBodyEnd - bodyStart, 1_000_000).Trim();
            if (preliminaryBody.Length > 0 && preliminaryBody[0] == '[') {
                var parsedArray = ParseTopLevelObject(preliminaryBody);
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
                    string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000); // cap to 1 MB
                    PdfDictionary? dict;
                    try { dict = ParseDictionary(dictText); }
                    catch (Exception ex) when (ex is not OutOfMemoryException) { dict = null; }
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
                var parsed = ParseTopLevelObject(preliminaryBody);
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
                ExpandObjectStreams(map, pdf, parsedOffsets, null);
            }
        }

        if (!appliedXrefStreamEntries) {
            // Compatibility fallback for simple parser-supported files whose compressed objects are only discoverable by scanning.
            ExpandObjectStreams(map, pdf, parsedOffsets, appliedClassicEntries ? activeClassicObjectNumbers : null);
        }

        if (decryptor is null) {
            ThrowIfEncryptedXrefStream(map);
        }

        return (map, trailerRaw);
    }
}
