using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private static readonly HashSet<string> StandardCidOrderings = new(StringComparer.Ordinal) { "GB1", "CNS1", "Japan1", "Korea1" };

    /// <summary>PDF/A-u requires a Unicode value for every drawn glyph; adds or completes the ToUnicode map from the font's own encoding.</summary>
    private void EnsureUnicode(FontUse use, string location) {
        PdfDictionary font = use.Font;
        bool composite = GetName(font, "Subtype") == "Type0";
        PdfFontResource resource;
        try {
            resource = ResourceResolver.CreateFontResource(use.ResourceName, font, _objects);
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException or IndexOutOfRangeException or PdfReadLimitException) {
            AddViolation(PdfAConversionRefusalReason.UnicodeMappingUnavailable, location, "The font cannot be read to determine Unicode values: " + exception.Message, font);
            return;
        }
        ToUnicodeCMap? cmap = resource.CMap;

        if (composite && use.CompositeStrings.Count > 0) {
            if (cmap is not null && use.CompositeStrings.All(text => cmap.TryMapBytes(text, 4096, out string mapped) && IsValidUnicode(mapped))) return;
            PdfDictionary? descendant = Resolve(font.Items.TryGetValue("DescendantFonts", out PdfObject? descendants) ? descendants : null) is PdfArray array &&
                array.Items.Count > 0 ? ResolveDictionary(array.Items[0]) : null;
            PdfDictionary? systemInfo = descendant is null ? null : ResolveDictionary(descendant.Items.TryGetValue("CIDSystemInfo", out PdfObject? info) ? info : null);
            string? registry = systemInfo is null ? null : TextOf(systemInfo, "Registry");
            string? ordering = systemInfo is null ? null : TextOf(systemInfo, "Ordering");
            if (registry == "Adobe" && ordering is not null && StandardCidOrderings.Contains(ordering)) return;
            AddViolation(PdfAConversionRefusalReason.UnicodeMappingUnavailable, location,
                "The composite font uses a CMap whose codes have no Unicode mapping, and its character collection is not a standard Adobe one.", font);
            return;
        }

        Func<byte[], int, string> fallback = ResourceResolver.CreateBudgetedDecoder(resource, preferToUnicode: false, useWinAnsiFallback: false);
        var mapping = new SortedDictionary<int, string>();
        var missing = new List<int>();
        bool needsMap = cmap is null;
        foreach (int code in use.Codes) {
            byte[] bytes = composite ? new[] { (byte)(code >> 8), (byte)code } : new[] { (byte)code };
            string? value = null;
            if (cmap is not null && cmap.TryMapBytes(bytes, 64, out string mapped) && IsValidUnicode(mapped)) {
                value = mapped;
            } else {
                needsMap = true;
                try {
                    string decoded = fallback(bytes, 64);
                    if (IsValidUnicode(decoded)) value = decoded;
                } catch (Exception exception) when (exception is PdfUnsupportedTextMappingException or InvalidDataException or ArgumentException or PdfReadLimitException) {
                }
            }
            if (value is null) missing.Add(code);
            else mapping[code] = value;
        }
        if (missing.Count > 0) {
            AddViolation(PdfAConversionRefusalReason.UnicodeMappingUnavailable, location,
                "Character code(s) " + FormatCodes(missing) + " have no Unicode value in the font's ToUnicode map or encoding, which PDF/A-u requires.", font);
            return;
        }
        if (!needsMap || mapping.Count == 0) return;
        if (Change(PdfAConversionChangeKind.ToUnicodeAdded, location,
                (cmap is null ? "Added a ToUnicode map" : "Completed the ToUnicode map") + " for " + mapping.Count.ToString(CultureInfo.InvariantCulture) +
                " drawn character code(s), derived from the font's encoding, as PDF/A-u requires.", font)) {
            font.Items["ToUnicode"] = AddObject(BuildToUnicodeStream(mapping, composite ? 2 : 1, cmap));
        }
    }

    /// <summary>
    /// PDF/A-4 (6.2.10.7) forbids ToUnicode values of U+0000, U+FEFF and U+FFFE. PDF/A-4 does not require Unicode values, so
    /// the map is rewritten without those entries; the glyphs drawn are unchanged.
    /// </summary>
    private void RemoveInvalidUnicodeValues(FontUse use, string location) {
        PdfDictionary font = use.Font;
        if (Resolve(font.Items.TryGetValue("ToUnicode", out PdfObject? value) ? value : null) is not PdfStream) return;
        bool composite = GetName(font, "Subtype") == "Type0";
        if (composite && GetName(font, "Encoding") is not ("Identity-H" or "Identity-V")) return;
        ToUnicodeCMap? cmap;
        try {
            cmap = ResourceResolver.CreateFontResource(use.ResourceName, font, _objects).CMap;
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException or IndexOutOfRangeException or PdfReadLimitException) {
            return;
        }
        if (cmap is null) return;
        int codeBytes = composite ? 2 : 1;
        int invalid = 0;
        foreach (KeyValuePair<string, string> entry in cmap.Mappings) {
            if (entry.Key.Length == codeBytes * 2 && HasForbiddenUnicodeValue(entry.Value)) invalid++;
        }
        if (invalid == 0) return;
        if (Change(PdfAConversionChangeKind.ToUnicodeAdded, location,
                "Removed " + invalid.ToString(CultureInfo.InvariantCulture) + " ToUnicode value(s) of U+0000, U+FEFF or U+FFFE, which PDF/A-4 forbids; those codes have no text value now.",
                font, OfficeConversionLossKind.Omission)) {
            // Keep every other value, including U+FFFD, which PDF/A-4 allows.
            font.Items["ToUnicode"] = AddObject(BuildToUnicodeStream(new SortedDictionary<int, string>(), codeBytes, cmap,
                static mapped => !string.IsNullOrEmpty(mapped) && !HasForbiddenUnicodeValue(mapped!)));
        }
    }

    private static bool HasForbiddenUnicodeValue(string value) {
        foreach (char character in value) {
            if (character is '\0' or '\uFEFF' or '\uFFFE') return true;
        }
        return false;
    }

    private static bool IsValidUnicode(string? value) {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (char character in value!) {
            if (character is '\0' or '﻿' or '￾' or '�') return false;
        }
        return true;
    }

    private PdfStream BuildToUnicodeStream(SortedDictionary<int, string> mapping, int codeBytes, ToUnicodeCMap? existing,
        Func<string?, bool>? keepExisting = null) {
        keepExisting ??= IsValidUnicode;
        var all = new SortedDictionary<int, string>(mapping);
        if (existing is not null) {
            foreach (KeyValuePair<string, string> entry in existing.Mappings) {
                if (entry.Key.Length != codeBytes * 2 ||
                    !int.TryParse(entry.Key, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code) ||
                    all.ContainsKey(code) || !keepExisting(entry.Value)) continue;
                all[code] = entry.Value;
            }
        }
        string codeFormat = codeBytes == 1 ? "X2" : "X4";
        var text = new StringBuilder();
        text.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n")
            .Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n")
            .Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n")
            .Append("1 begincodespacerange\n<").Append(0.ToString(codeFormat, CultureInfo.InvariantCulture)).Append("> <")
            .Append((codeBytes == 1 ? 0xFF : 0xFFFF).ToString(codeFormat, CultureInfo.InvariantCulture)).Append(">\nendcodespacerange\n");
        var entries = all.ToList();
        for (int start = 0; start < entries.Count; start += 100) {
            int count = Math.Min(100, entries.Count - start);
            text.Append(count.ToString(CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
            for (int index = start; index < start + count; index++) {
                text.Append('<').Append(entries[index].Key.ToString(codeFormat, CultureInfo.InvariantCulture)).Append("> <");
                foreach (char character in entries[index].Value) text.Append(((int)character).ToString("X4", CultureInfo.InvariantCulture));
                text.Append(">\n");
            }
            text.Append("endbfchar\n");
        }
        text.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        return new PdfStream(dictionary, PdfFlateEncoder.Compress(Encoding.ASCII.GetBytes(text.ToString()), _cancellationToken));
    }
}
