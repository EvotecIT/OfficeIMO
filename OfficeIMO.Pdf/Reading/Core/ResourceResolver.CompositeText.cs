namespace OfficeIMO.Pdf;

// Composite (Type0) fonts select CIDs. When ToUnicode is absent or does not cover a code, the CID's text comes from
// the descendant's Adobe character collection and then from the embedded font program, the way PDF viewers recover it.
internal static partial class ResourceResolver {
    private static PdfCidTextFallback? CreateCidTextFallback(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects,
        byte[]? trueTypeProgram) {
        if (!string.Equals(font.Get<PdfName>("Subtype")?.Name, "Type0", StringComparison.Ordinal) ||
            !font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) ||
            ResolveArray(descendantsObject, objects) is not { Items.Count: > 0 } descendants ||
            ResolveDict(descendants.Items[0], objects) is not PdfDictionary descendant) return null;
        string? collection = null;
        if (descendant.Items.TryGetValue("CIDSystemInfo", out PdfObject? infoObject) && ResolveDict(infoObject, objects) is PdfDictionary info) {
            string registry = ReadInfoString(info, "Registry");
            string ordering = ReadInfoString(info, "Ordering");
            if (PdfPredefinedCMap.HasUnicodeCollection(registry, ordering)) collection = ordering;
        }

        Func<Func<int, string?>?>? program = null;
        byte[]? programBytes = null;
        string? descendantSubtype = descendant.Get<PdfName>("Subtype")?.Name;
        if (string.Equals(descendantSubtype, "CIDFontType2", StringComparison.Ordinal) && trueTypeProgram != null) {
            byte[]? cidToGlyph = TryReadCidToGlyphMap(font, objects);
            if (cidToGlyph != null) {
                program = () => CreateTrueTypeCidText(trueTypeProgram, cidToGlyph);
                programBytes = trueTypeProgram;
            }
        } else if (string.Equals(descendantSubtype, "CIDFontType0", StringComparison.Ordinal) &&
            ResolveEmbeddedCffStream(font, objects, out _) != null) {
            program = () => CreateCffCidText(font, objects);
        }
        return collection == null && program == null ? null : new PdfCidTextFallback(collection, program, programBytes);

        string ReadInfoString(PdfDictionary dictionary, string key) =>
            dictionary.Items.TryGetValue(key, out PdfObject? value) && ResolveObject(value, objects) is PdfStringObj text
                ? text.Value
                : string.Empty;
    }

    private static bool HasIdentityToUnicode(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects) =>
        string.Equals(font.Get<PdfName>("Subtype")?.Name, "Type0", StringComparison.Ordinal) &&
        font.Items.TryGetValue("ToUnicode", out PdfObject? toUnicode) &&
        ResolveObject(toUnicode, objects) is PdfName { Name: "Identity-H" or "Identity-V" };

    // A CIDFontType2 selects glyphs through CIDToGIDMap (empty for Identity); the program's Unicode cmap names each glyph.
    private static Func<int, string?>? CreateTrueTypeCidText(byte[] program, byte[] cidToGlyph) {
        Dictionary<int, int>? scalars;
        try {
            scalars = PdfTrueTypeUnicodeCmap.TryReadGlyphScalars(program);
        } catch (Exception exception) when (exception is IndexOutOfRangeException or ArgumentException or OverflowException) {
            return null;
        }
        if (scalars == null) return null;
        return cid => {
            int glyph = cidToGlyph.Length == 0
                ? cid
                : 2 * cid + 1 < cidToGlyph.Length ? (cidToGlyph[2 * cid] << 8) | cidToGlyph[2 * cid + 1] : 0;
            return glyph > 0 && scalars.TryGetValue(glyph, out int scalar) ? char.ConvertFromUtf32(scalar) : null;
        };
    }

    // A CID-keyed CFF program selects glyphs through its charset; a name-keyed one uses the CID as the glyph index.
    // OpenType wrappers can name glyphs through their cmap; bare name-keyed programs through their glyph names.
    private static Func<int, string?>? CreateCffCidText(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects) {
        byte[]? data = TryReadEmbeddedCffProgram(font, objects, out string? subtype);
        if (data == null || string.Equals(subtype, "Type1", StringComparison.Ordinal)) return null;
        bool openType = string.Equals(subtype, "OpenType", StringComparison.Ordinal);
        PdfCffFontProgram? cff;
        Dictionary<int, int>? scalars;
        try {
            byte[]? cffData = openType ? PdfCffFontProgram.TryExtractOpenTypeCff(data) : data;
            cff = cffData == null ? null : PdfCffFontProgram.TryParse(cffData);
            scalars = openType ? PdfTrueTypeUnicodeCmap.TryReadGlyphScalars(data) : null;
        } catch (Exception exception) when (exception is IndexOutOfRangeException or ArgumentException or OverflowException) {
            return null;
        }
        if (cff == null || scalars == null && (cff.IsCidKeyed || cff.GlyphNames == null)) return null;
        return cid => {
            if (!cff.TryGetGlyphForCid(cid, out int glyph)) return null;
            if (scalars != null) return scalars.TryGetValue(glyph, out int scalar) ? char.ConvertFromUtf32(scalar) : null;
            string?[]? names = cff.GlyphNames;
            return names != null && glyph < names.Length && names[glyph] is string name &&
                TryDecodeGlyphName(name, out string? value) && !string.IsNullOrEmpty(value) ? value : null;
        };
    }

    private static bool HasCompositeTextSources(PdfFontResource font) =>
        string.Equals(font.FontSubtype, "Type0", StringComparison.Ordinal) &&
        (font.CidTextFallback != null || font.HasIdentityToUnicode || IsIdentityComposite(font));

    private static bool IsIdentityComposite(PdfFontResource font) => font.Encoding is "Identity-H" or "Identity-V";

    /// <summary>
    /// Decodes a composite run code by code: ToUnicode, then the CID's collection or font-program text. A code no
    /// source maps refuses the run, except that an Identity encoding under the fallback policy keeps raw text: the
    /// ToUnicode map's partial match within the code, or otherwise the code's own value as one UTF-16 character.
    /// </summary>
    private static string DecodeCompositeWithTextSources(PdfFontResource font, byte[] bytes, int maximum,
        bool preferToUnicode, bool useWinAnsiFallback) {
        if (bytes.Length == 0) return string.Empty;
        bool named = UsesNamedCompositeEncoding(font);
        bool useToUnicode = preferToUnicode && font.HasToUnicode;
        ToUnicodeCMap? toUnicode = useToUnicode ? font.CMap : null;
        if (toUnicode != null && toUnicode.TryMapBytes(bytes, maximum, out string mapped)) return mapped;
        PdfPredefinedCMap? codeMap = named ? font.PredefinedCMap?.Value : null;
        if (named && !useToUnicode && codeMap != null && codeMap.TryDecode(bytes, maximum, out string decoded)) return decoded;
        bool keepRawText = !named && useWinAnsiFallback;

        var text = new StringBuilder(bytes.Length);
        for (int index = 0; index < bytes.Length;) {
            int length = named ? codeMap?.GetCodeLength(bytes, index) ?? 2 : 2;
            if (index + length > bytes.Length) {
                // The string ends inside a code: a width probe's prefix or a truncated operand.
                if (!keepRawText) throw new PdfUnsupportedTextMappingException(font);
                AppendRawCompositeText(toUnicode, bytes, index, bytes.Length - index, text, maximum);
                break;
            }
            if (TryAppendCompositeCode(font, bytes, index, length, named, codeMap, toUnicode, useToUnicode, text) == null) {
                if (!keepRawText) throw new PdfUnsupportedTextMappingException(font);
                AppendRawCompositeText(toUnicode, bytes, index, length, text, maximum);
            }
            if (text.Length > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maximum, text.Length);
            index += length;
        }
        return text.ToString();
    }

    // Returns the source that supplied the code's text, or null when no source maps it.
    private static PdfTextDecodingSource? TryAppendCompositeCode(PdfFontResource font, byte[] bytes, int index, int length, bool named,
        PdfPredefinedCMap? codeMap, ToUnicodeCMap? toUnicode, bool useToUnicode, StringBuilder text) {
        if (toUnicode != null && toUnicode.TryMapCode(bytes, index, length, out string mapped)) {
            text.Append(mapped);
            return PdfTextDecodingSource.ToUnicode;
        }
        if (useToUnicode && font.HasIdentityToUnicode && TryAppendIdentityUnicode(bytes, index, length, text)) return PdfTextDecodingSource.ToUnicode;
        int cid;
        if (named) {
            if (codeMap == null || !codeMap.TryGetPaintedCid(bytes, index, length, out cid)) return null;
        } else {
            cid = (bytes[index] << 8) | bytes[index + 1];
        }
        return font.CidTextFallback?.TryAppend(cid, text);
    }

    // ToUnicode /Identity-H names no map: each code is read as the UTF-16 or scalar value of its text.
    private static bool TryAppendIdentityUnicode(byte[] bytes, int index, int length, StringBuilder text) {
        int value = 0;
        for (int offset = 0; offset < length; offset++) value = (value << 8) | bytes[index + offset];
        if (value < 0x20 || value > 0x10FFFF || value is >= 0xD800 and <= 0xDFFF || value == 0xFFFD) return false;
        text.Append(char.ConvertFromUtf32(value));
        return true;
    }

    private static void AppendRawCompositeText(ToUnicodeCMap? toUnicode, byte[] bytes, int index, int length, StringBuilder text, int maximum) {
        int remaining = maximum - text.Length;
        if (remaining <= 0) throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maximum, (long)text.Length + length);
        if (length == 2 && (toUnicode == null || toUnicode.CountUnmappedCodeBytes(SliceCode(bytes, index, length), null) == length)) {
            // One glyph keeps one character. A control or surrogate value is no text: U+0000 marks the glyph as undecoded.
            char value = (char)((bytes[index] << 8) | bytes[index + 1]);
            text.Append(char.IsControl(value) || char.IsSurrogate(value) ? '\0' : value);
            return;
        }
        byte[] code = SliceCode(bytes, index, length);
        text.Append(toUnicode != null ? toUnicode.MapBytes(code, remaining) : PdfWinAnsiEncoding.Decode(code, remaining));
    }

    private static byte[] SliceCode(byte[] bytes, int index, int length) {
        var code = new byte[length];
        Buffer.BlockCopy(bytes, index, code, 0, length);
        return code;
    }

    // Mirrors the decoder: the source of every code and the bytes of codes no source maps.
    private static PdfTextDecodingSource GetCompositeTextSource(PdfFontResource font, byte[] bytes, bool preferToUnicode,
        Action? cancellationCheck, out int unmappedBytes) {
        unmappedBytes = 0;
        bool named = UsesNamedCompositeEncoding(font);
        bool useToUnicode = preferToUnicode && font.HasToUnicode;
        ToUnicodeCMap? toUnicode = useToUnicode ? font.CMap : null;
        if (toUnicode != null) {
            // Without another source the run keeps the ToUnicode map's own count of raw bytes.
            int unmappedToUnicodeBytes = toUnicode.CountUnmappedCodeBytes(bytes, cancellationCheck);
            if (unmappedToUnicodeBytes == 0 || font.CidTextFallback == null && !font.HasIdentityToUnicode) {
                unmappedBytes = unmappedToUnicodeBytes;
                return PdfTextDecodingSource.ToUnicode;
            }
        }
        PdfPredefinedCMap? codeMap = named ? font.PredefinedCMap?.Value : null;
        bool usedToUnicode = false;
        bool usedProgram = false;
        var scratch = new StringBuilder();
        for (int index = 0; index < bytes.Length;) {
            cancellationCheck?.Invoke();
            int length = named ? codeMap?.GetCodeLength(bytes, index) ?? 2 : 2;
            bool complete = index + length <= bytes.Length;
            if (!complete) length = bytes.Length - index;
            PdfTextDecodingSource? source = complete
                ? TryAppendCompositeCode(font, bytes, index, length, named, codeMap, toUnicode, useToUnicode, scratch.Clear())
                : null;
            if (source == null) unmappedBytes += length;
            else if (source == PdfTextDecodingSource.ToUnicode) usedToUnicode = true;
            else if (source == PdfTextDecodingSource.FontProgram) usedProgram = true;
            index += length;
        }
        if (unmappedBytes == 0 && !usedToUnicode) return usedProgram ? PdfTextDecodingSource.FontProgram : PdfTextDecodingSource.PredefinedCMap;
        if (usedToUnicode || toUnicode != null) return PdfTextDecodingSource.ToUnicode;
        // The source already says that no mapping supplied the run.
        unmappedBytes = 0;
        return named ? PdfTextDecodingSource.Unmappable : PdfTextDecodingSource.WinAnsiFallback;
    }
}
