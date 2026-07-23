namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static string GetActiveTrailerRaw(
        string text,
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int maximumTrailerCharacters) {
        if (TryGetLatestStartXrefOffset(text, out int activeXrefOffset)) {
            if (TryGetClassicTrailerChainRaw(text, map, parsedOffsets, activeXrefOffset, out string trailerRaw)) {
                return trailerRaw;
            }

            if (TryGetXrefStreamTrailerChainRaw(text, map, parsedOffsets, activeXrefOffset, out trailerRaw)) {
                return trailerRaw;
            }
        }

        int trailerIdx = text.LastIndexOf("trailer", StringComparison.OrdinalIgnoreCase);
        return trailerIdx >= 0
            ? SafeSlice(text, trailerIdx, text.Length - trailerIdx, maximumTrailerCharacters)
            : string.Empty;
    }

    private static bool TryGetXrefStreamTrailerChainRaw(
        string text,
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int activeXrefOffset,
        out string trailerRaw) {
        trailerRaw = string.Empty;
        var byOffset = new Dictionary<int, PdfDictionary>();
        foreach (var entry in map.Values) {
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
            trailers.Add(BuildXrefStreamTrailerRaw(dictionary));
            if (dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                break;
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        if (trailers.Count > 0 &&
            TryGetClassicTrailerChainRaw(text, map, parsedOffsets, currentOffset, out string classicTrailerRaw)) {
            trailers.Add(classicTrailerRaw);
        }

        if (trailers.Count == 0) {
            return false;
        }

        trailerRaw = string.Join("\n", trailers);
        return true;
    }

    private static string BuildXrefStreamTrailerRaw(PdfDictionary dictionary) {
        var parts = new List<string>();
        AppendTrailerEntry(parts, dictionary, "Size");
        AppendTrailerEntry(parts, dictionary, "Root");
        AppendTrailerEntry(parts, dictionary, "Info");
        AppendTrailerEntry(parts, dictionary, "ID");
        AppendTrailerEntry(parts, dictionary, "Encrypt");
        AppendTrailerEntry(parts, dictionary, "Prev");
        return "trailer\n<< " + string.Join(" ", parts) + " >>";
    }

    private static void AppendTrailerEntry(List<string> parts, PdfDictionary dictionary, string key) {
        if (dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            TryFormatTrailerValue(value, out string? formatted)) {
            parts.Add("/" + key + " " + formatted);
        }
    }

    private static bool TryFormatTrailerValue(PdfObject value, out string? formatted) {
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
                formatted = "(" + text.Value.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)") + ")";
                return true;
            case PdfArray array:
                var items = new List<string>();
                foreach (PdfObject item in array.Items) {
                    if (!TryFormatTrailerValue(item, out string? itemText)) {
                        formatted = null;
                        return false;
                    }

                    if (itemText is null) {
                        formatted = null;
                        return false;
                    }

                    items.Add(itemText);
                }

                formatted = "[" + string.Join(" ", items) + "]";
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
        out string trailerRaw) {
        trailerRaw = string.Empty;
        var trailers = new List<string>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (visited.Add(currentOffset) &&
            trailers.Count < 64 &&
            TryParseClassicXrefTable(text, currentOffset, out _, out int? previousOffset, out string currentTrailerRaw, out int? xrefStreamOffset)) {
            if (!string.IsNullOrWhiteSpace(currentTrailerRaw)) {
                trailers.Add(currentTrailerRaw);
            }

            if (xrefStreamOffset.HasValue &&
                TryGetXrefStreamTrailerRawAtOffset(map, parsedOffsets, xrefStreamOffset.Value, out string xrefStreamTrailerRaw)) {
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

        trailerRaw = string.Join("\n", trailers);
        return true;
    }

    private static bool TryGetXrefStreamTrailerRawAtOffset(
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int xrefStreamOffset,
        out string trailerRaw) {
        trailerRaw = string.Empty;
        foreach (var entry in map.Values) {
            if (!parsedOffsets.TryGetValue(entry.ObjectNumber, out int offset) ||
                offset != xrefStreamOffset ||
                entry.Value is not PdfStream stream ||
                stream.Dictionary.Get<PdfName>("Type")?.Name != "XRef") {
                continue;
            }

            trailerRaw = BuildXrefStreamTrailerRaw(stream.Dictionary);
            return true;
        }

        return false;
    }

}
