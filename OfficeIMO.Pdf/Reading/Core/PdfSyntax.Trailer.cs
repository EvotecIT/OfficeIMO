using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    internal static PdfReference? ReadTrailerReference(
        string? trailerRaw,
        string key,
        PdfReadLimits? limits = null,
        CancellationToken cancellationToken = default) =>
        TryGetTrailerReference(trailerRaw, key, limits, out PdfReference reference, cancellationToken)
            ? reference
            : null;

    internal static bool TryGetTrailerReference(
        string? trailerRaw,
        string key,
        PdfReadLimits? limits,
        out PdfReference reference,
        CancellationToken cancellationToken = default) {
        reference = null!;
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(trailerRaw) || string.IsNullOrWhiteSpace(key)) return false;
        int searchIndex = 0;
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        while (TryReadNextTrailerDictionary(trailerRaw!, ref searchIndex, effectiveLimits, out PdfDictionary dictionary, cancellationToken)) {
            if (dictionary.Items.TryGetValue(key, out PdfObject? value)) {
                if (value is PdfReference found) {
                    reference = found;
                    return true;
                }
                // A later trailer entry overrides the same key in every earlier revision.
                return false;
            }
        }
        return false;
    }

    /// <summary>Reads up to three trailer references in one chain scan, retaining explicit non-reference overrides.</summary>
    internal static (PdfReference? First, PdfReference? Second, PdfReference? Third) ReadTrailerReferences(
        string? trailerRaw,
        string firstKey,
        string? secondKey,
        string? thirdKey,
        PdfReadLimits? limits,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(trailerRaw) || string.IsNullOrWhiteSpace(firstKey)) return default;
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        PdfReference? first = null;
        PdfReference? second = null;
        PdfReference? third = null;
        bool foundFirst = false;
        bool foundSecond = string.IsNullOrWhiteSpace(secondKey);
        bool foundThird = string.IsNullOrWhiteSpace(thirdKey);
        int searchIndex = 0;
        while (TryReadNextTrailerDictionary(trailerRaw!, ref searchIndex, effectiveLimits, out PdfDictionary dictionary, cancellationToken)) {
            if (!foundFirst && dictionary.Items.TryGetValue(firstKey, out PdfObject? firstValue)) {
                foundFirst = true;
                first = firstValue as PdfReference;
            }
            if (!foundSecond && dictionary.Items.TryGetValue(secondKey!, out PdfObject? secondValue)) {
                foundSecond = true;
                second = secondValue as PdfReference;
            }
            if (!foundThird && dictionary.Items.TryGetValue(thirdKey!, out PdfObject? thirdValue)) {
                foundThird = true;
                third = thirdValue as PdfReference;
            }
            if (foundFirst && foundSecond && foundThird) break;
        }
        return (first, second, third);
    }

    private static bool TryReadNextTrailerDictionary(
        string raw,
        ref int searchIndex,
        PdfReadLimits limits,
        out PdfDictionary dictionary,
        CancellationToken cancellationToken) {
        dictionary = null!;
        cancellationToken.ThrowIfCancellationRequested();
        if (searchIndex >= raw.Length) return false;
        int trailerIndex = IndexOfTrailerMarker(raw, searchIndex, cancellationToken);
        if (trailerIndex < 0) return false;
        int dictionaryStart = SkipWhitespaceAndComments(raw, trailerIndex + 7, raw.Length, cancellationToken);
        if (dictionaryStart > raw.Length - 2 ||
            raw[dictionaryStart] != '<' ||
            raw[dictionaryStart + 1] != '<') return false;
        int dictionaryEnd = FindDictEnd(raw, dictionaryStart, raw.Length, cancellationToken);
        if (dictionaryEnd <= dictionaryStart ||
            dictionaryEnd - dictionaryStart - 2 > limits.MaxObjectCharacters) return false;
        try {
            dictionary = ParseDictionary(
                raw.Substring(dictionaryStart + 2, dictionaryEnd - dictionaryStart - 2),
                limits,
                cancellationToken: cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            searchIndex = dictionaryEnd;
            return true;
        } catch (Exception exception) when (exception is not OutOfMemoryException && exception is not OperationCanceledException) {
            return false;
        }
    }

    private static int IndexOfTrailerMarker(string raw, int searchIndex, CancellationToken cancellationToken) {
        const string marker = "trailer";
        const int window = 65536;
        for (int index = searchIndex; index < raw.Length; index += window) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(raw.Length - index, window + marker.Length - 1);
            int found = raw.IndexOf(marker, index, count, StringComparison.OrdinalIgnoreCase);
            if (found >= 0) return found;
        }
        return -1;
    }

    private static int LastIndexOfTrailerMarker(string raw, string marker, StringComparison comparison,
        CancellationToken cancellationToken) {
        const int window = 65536;
        for (int end = raw.Length; end > 0; end -= window) {
            cancellationToken.ThrowIfCancellationRequested();
            int start = Math.Max(0, end - window - marker.Length + 1);
            int found = raw.LastIndexOf(marker, end - 1, end - start, comparison);
            if (found >= 0) return found;
        }
        return -1;
    }

    internal static byte[]? ReadPermanentTrailerIdentifier(string trailerRaw) {
        string entry = PdfIncrementalObjectWriter.ReadTrailerIdEntry(trailerRaw);
        if (string.IsNullOrWhiteSpace(entry)) return null;
        try {
            PdfDictionary dictionary = ParseDictionary("<<" + entry + ">>");
            if (dictionary.Get<PdfArray>("ID") is PdfArray identifiers &&
                identifiers.Items.Count > 0 && identifiers.Items[0] is PdfStringObj permanent) {
                return (byte[])permanent.RawBytes.Clone();
            }
        } catch {
            // A malformed optional identifier is not allowed to block an otherwise safe rewrite.
        }
        return null;
    }

    private static string GetActiveTrailerRaw(
        string text,
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int maximumTrailerCharacters,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (TryGetLatestStartXrefOffset(text, out int activeXrefOffset, cancellationToken)) {
            if (TryGetClassicTrailerChainRaw(text, map, parsedOffsets, activeXrefOffset, out string trailerRaw, cancellationToken)) {
                return trailerRaw;
            }

            if (TryGetXrefStreamTrailerChainRaw(text, map, parsedOffsets, activeXrefOffset, out trailerRaw, cancellationToken)) {
                return trailerRaw;
            }
        }

        int trailerIdx = LastIndexOfTrailerMarker(text, "trailer", StringComparison.OrdinalIgnoreCase, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return trailerIdx >= 0
            ? SafeSlice(text, trailerIdx, text.Length - trailerIdx, maximumTrailerCharacters)
            : string.Empty;
    }

    private static bool TryGetXrefStreamTrailerChainRaw(
        string text,
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int activeXrefOffset,
        out string trailerRaw,
        CancellationToken cancellationToken) {
        trailerRaw = string.Empty;
        var byOffset = new Dictionary<int, PdfDictionary>();
        foreach (var entry in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!parsedOffsets.TryGetValue(entry.ObjectNumber, out int offset)) {
                continue;
            }

            PdfDictionary? dictionary = entry.Value is PdfStream stream ? stream.Dictionary : entry.Value as PdfDictionary;
            if (dictionary?.Get<PdfName>("Type")?.Name == "XRef") {
                byOffset[offset] = dictionary;
            }
        }

        var trailers = new List<string>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfDictionary? dictionary) &&
            visited.Add(currentOffset) &&
            trailers.Count < 64) {
            cancellationToken.ThrowIfCancellationRequested();
            trailers.Add(BuildXrefStreamTrailerRaw(dictionary, cancellationToken));
            if (dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                break;
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        if (trailers.Count > 0 &&
            TryGetClassicTrailerChainRaw(text, map, parsedOffsets, currentOffset, out string classicTrailerRaw, cancellationToken)) {
            trailers.Add(classicTrailerRaw);
        }

        if (trailers.Count == 0) {
            return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        trailerRaw = JoinTrailerParts(trailers, "\n", cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }

    private static string BuildXrefStreamTrailerRaw(PdfDictionary dictionary, CancellationToken cancellationToken) {
        var parts = new List<string>();
        AppendTrailerEntry(parts, dictionary, "Size", cancellationToken);
        AppendTrailerEntry(parts, dictionary, "Root", cancellationToken);
        AppendTrailerEntry(parts, dictionary, "Info", cancellationToken);
        AppendTrailerEntry(parts, dictionary, "ID", cancellationToken);
        AppendTrailerEntry(parts, dictionary, "Encrypt", cancellationToken);
        AppendTrailerEntry(parts, dictionary, "Prev", cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return "trailer\n<< " + JoinTrailerParts(parts, " ", cancellationToken) + " >>";
    }

    private static void AppendTrailerEntry(List<string> parts, PdfDictionary dictionary, string key, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            TryFormatTrailerValue(value, out string? formatted, cancellationToken)) {
            parts.Add("/" + key + " " + formatted);
        }
    }

    private static bool TryFormatTrailerValue(PdfObject value, out string? formatted, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfReference reference:
                formatted = reference.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " +
                    reference.Generation.ToString(System.Globalization.CultureInfo.InvariantCulture) + " R";
                return true;
            case PdfNumber number:
                formatted = number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                return true;
            case PdfName name:
                formatted = "/" + name.Name;
                return true;
            case PdfStringObj text:
                var hex = new System.Text.StringBuilder(text.RawBytes.Length * 2 + 2);
                hex.Append('<');
                foreach (byte valueByte in text.RawBytes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    hex.Append(valueByte.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
                }
                hex.Append('>');
                formatted = hex.ToString();
                return true;
            case PdfArray array:
                var items = new List<string>();
                foreach (PdfObject item in array.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!TryFormatTrailerValue(item, out string? itemText, cancellationToken)) {
                        formatted = null;
                        return false;
                    }

                    if (itemText is null) {
                        formatted = null;
                        return false;
                    }

                    items.Add(itemText);
                }

                formatted = "[" + JoinTrailerParts(items, " ", cancellationToken) + "]";
                return true;
            case PdfNull:
                formatted = "null";
                return true;
            default:
                formatted = null;
                return false;
        }
    }

    private static bool TryGetClassicTrailerChainRaw(
        string text,
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int activeXrefOffset,
        out string trailerRaw,
        CancellationToken cancellationToken) {
        trailerRaw = string.Empty;
        var trailers = new List<string>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (visited.Add(currentOffset) &&
            trailers.Count < 64 &&
            TryParseClassicXrefTable(text, currentOffset, out _, out int? previousOffset, out string currentTrailerRaw, out int? xrefStreamOffset, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.IsNullOrWhiteSpace(currentTrailerRaw)) {
                trailers.Add(currentTrailerRaw);
            }

            if (xrefStreamOffset.HasValue &&
                TryGetXrefStreamTrailerRawAtOffset(map, parsedOffsets, xrefStreamOffset.Value, out string xrefStreamTrailerRaw, cancellationToken)) {
                trailers.Add(xrefStreamTrailerRaw);
            }

            if (!previousOffset.HasValue) {
                break;
            }

            currentOffset = previousOffset.Value;
        }

        if (trailers.Count == 0) {
            return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        trailerRaw = JoinTrailerParts(trailers, "\n", cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }

    private static bool TryGetXrefStreamTrailerRawAtOffset(
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int xrefStreamOffset,
        out string trailerRaw,
        CancellationToken cancellationToken) {
        trailerRaw = string.Empty;
        foreach (var entry in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!parsedOffsets.TryGetValue(entry.ObjectNumber, out int offset) ||
                offset != xrefStreamOffset ||
                entry.Value is not PdfStream stream ||
                stream.Dictionary.Get<PdfName>("Type")?.Name != "XRef") {
                continue;
            }

            trailerRaw = BuildXrefStreamTrailerRaw(stream.Dictionary, cancellationToken);
            return true;
        }

        return false;
    }

    private static string JoinTrailerParts(List<string> parts, string separator,
        CancellationToken cancellationToken) {
        var result = new System.Text.StringBuilder();
        for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (partIndex > 0) result.Append(separator);
            string part = parts[partIndex];
            for (int offset = 0; offset < part.Length; offset += 65536) {
                cancellationToken.ThrowIfCancellationRequested();
                result.Append(part, offset, Math.Min(65536, part.Length - offset));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return result.ToString();
    }

}
