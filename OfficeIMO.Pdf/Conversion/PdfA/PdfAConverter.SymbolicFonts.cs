using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private static readonly string[] SymbolFamilies = { "Symbol" };
    private static readonly string[] DingbatsFamilies = { "Zapf Dingbats", "ZapfDingbats", "ITC Zapf Dingbats" };

    /// <summary>True for the standard Symbol and ZapfDingbats fonts, whose codes select glyphs by name in their own encodings.</summary>
    private static bool IsStandardSymbolicName(string name, out bool dingbats) {
        dingbats = name is "ZapfDingbats" or "ZapfDingbatsITC" or "ITCZapfDingbats" or "Dingbats";
        return dingbats || name is "Symbol" or "SymbolMT";
    }

    /// <summary>
    /// Embeds a supplied or installed Symbol or ZapfDingbats face for the non-embedded standard font. Each code selects a glyph name
    /// through the font's built-in encoding or Differences; the face must have that glyph with the width the page positions it
    /// by (the Widths entry, or the Adobe metrics of the standard font), within one unit.
    /// </summary>
    private void EmbedSymbolicSubstitute(FontUse use, PdfDictionary? descriptor, string location, string plainName, bool dingbats) {
        PdfDictionary font = use.Font;
        string encodingName = dingbats ? "ZapfDingbatsEncoding" : "SymbolEncoding";
        string?[] builtin = PdfFontEncodingNames.Get(encodingName)!;
        string?[] glyphNames = BuildSymbolicEncoding(font, builtin);
        IEnumerable<int> codes = use.Codes.Count > 0 ? use.Codes : new[] { 32 };
        var unnamed = codes.Where(code => code >= glyphNames.Length || glyphNames[code] is null or ".notdef").ToList();
        if (unnamed.Count > 0) {
            AddViolation(PdfAConversionRefusalReason.FontGlyphMissing, location,
                "Code(s) " + FormatCodes(unnamed) + " select no glyph in the " + plainName + " encoding, so they draw .notdef, which PDF/A forbids.", font);
            return;
        }
        var builtinCodes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int code = 0; code < builtin.Length; code++) {
            if (builtin[code] is string name && !builtinCodes.ContainsKey(name)) builtinCodes.Add(name, code);
        }

        string? failure = null;
        PdfAConversionRefusalReason reason = PdfAConversionRefusalReason.FontNotEmbeddable;
        string[] families = dingbats ? DingbatsFamilies : SymbolFamilies;
        foreach ((string family, string origin, PdfEmbeddedFontFamily substitute) in _options.FontSubstitution.Resolve(families)) {
            byte[] face = substitute.RegularSnapshot;
            PdfTrueTypeFontProgram program;
            try {
                program = PdfTrueTypeFontProgram.Parse(face);
            } catch (Exception exception) when (exception is NotSupportedException or InvalidDataException or ArgumentException or IndexOutOfRangeException) {
                failure = "the " + origin + " " + family + " face cannot be embedded: " + exception.Message;
                reason = PdfAConversionRefusalReason.FontNotEmbeddable;
                continue;
            }
            byte[] faceData = program.FontDataForInspection;
            var glyphs = new SortedDictionary<int, int>();
            var missing = new List<int>();
            var mismatched = new List<string>();
            foreach (int code in codes) {
                string name = glyphNames[code]!;
                int glyph = builtinCodes.TryGetValue(name, out int builtinCode) ? PdfTrueTypeUnicodeCmap.ResolveSymbolicCodeGlyph(faceData, builtinCode) : 0;
                if (glyph <= 0 || glyph >= program.GlyphCount) {
                    missing.Add(code);
                    continue;
                }
                glyphs[code] = glyph;
                if (!use.VisibleCodes.Contains(code)) continue;
                double? expected = ExpectedWidth(font, descriptor, code, 0, null) ??
                    (PdfStandardFontWidths.TryGetSymbolicWidth1000(dingbats, name, out int standardWidth) ? standardWidth : null);
                double actual = program.GetExactGlyphWidth1000(glyph);
                if (expected is not double width || Math.Abs(width - actual) > WidthTolerance) {
                    if (mismatched.Count < 8) {
                        mismatched.Add(FormatCode(code) + " " + (expected?.ToString("0.#", CultureInfo.InvariantCulture) ?? "?") + " vs " +
                            actual.ToString("0.#", CultureInfo.InvariantCulture));
                    }
                }
            }
            if (missing.Count > 0) {
                failure = "the " + origin + " " + family + " face has no glyph for code(s) " + FormatCodes(missing);
                reason = PdfAConversionRefusalReason.FontGlyphMissing;
                continue;
            }
            if (mismatched.Count > 0) {
                failure = "the " + origin + " " + family + " face has different glyph widths (" + string.Join(", ", mismatched) + ")";
                reason = PdfAConversionRefusalReason.FontMetricsMismatch;
                continue;
            }
            if (Change(PdfAConversionChangeKind.FontEmbedded, location,
                    "Embedded the " + origin + " " + family + " face (" + program.FontName + ") for the non-embedded font " + plainName +
                    "; every code selects the glyph its " + plainName + " glyph name names, with the advance width the page already positions it by.",
                    font, OfficeConversionLossKind.Approximation)) {
                WriteSymbolicSubstitute(font, descriptor, plainName, program, glyphNames, glyphs, dingbats);
            }
            return;
        }
        AddViolation(reason, location, failure is null
            ? "The " + plainName + " font is not embedded and no " + SubstituteSources + " substitute (" + string.Join(", ", families) + ") was found."
            : "The " + plainName + " font is not embedded and cannot be substituted without changing the page: " + failure + ".", font);
    }

    /// <summary>The built-in encoding of Symbol or ZapfDingbats, replaced by a named encoding or completed by Differences.</summary>
    private string?[] BuildSymbolicEncoding(PdfDictionary font, string?[] builtin) {
        PdfObject? encoding = Resolve(font.Items.TryGetValue("Encoding", out PdfObject? value) ? value : null);
        if (encoding is PdfName named) return (string?[])(PdfFontEncodingNames.Get(named.Name) ?? builtin).Clone();
        if (encoding is not PdfDictionary dictionary) return (string?[])builtin.Clone();
        string?[] names = (string?[])((GetName(dictionary, "BaseEncoding") is string baseName ? PdfFontEncodingNames.Get(baseName) : null) ?? builtin).Clone();
        if (Resolve(dictionary.Items.TryGetValue("Differences", out PdfObject? differences) ? differences : null) is PdfArray array) {
            int code = 0;
            foreach (PdfObject item in array.Items) {
                PdfObject? resolved = Resolve(item);
                if (resolved is PdfNumber number) code = (int)number.Value;
                else if (resolved is PdfName name) {
                    if (code >= 0 && code < names.Length) names[code] = name.Name;
                    code++;
                }
            }
        }
        return names;
    }

    /// <summary>
    /// Writes a symbolic TrueType subset whose single (3,0) cmap maps each code to its glyph, as PDF/A requires of symbolic
    /// TrueType fonts, which carry no Encoding entry. A ToUnicode map keeps the text the glyph names stand for.
    /// </summary>
    private void WriteSymbolicSubstitute(PdfDictionary font, PdfDictionary? descriptor, string plainName, PdfTrueTypeFontProgram program,
        string?[] glyphNames, SortedDictionary<int, int> glyphs, bool dingbats) {
        byte[] subset = program.BuildGlyphRetainingSubset(glyphs.Values.Append(0), PdfTrueTypeUnicodeCmap.BuildSymbolCmap(glyphs));
        string taggedName = program.SubsettingForbidden ? SanitizeFontName(plainName)
            : PdfSubsetTag.Apply(SanitizeFontName(plainName), glyphs.Values.Distinct().OrderBy(static glyph => glyph), subset);

        var fileDictionary = new PdfDictionary();
        fileDictionary.Items["Length1"] = new PdfNumber(subset.Length);
        fileDictionary.Items["Filter"] = new PdfName("FlateDecode");
        PdfReference fontFile = AddObject(new PdfStream(fileDictionary, PdfFlateEncoder.Compress(subset, _cancellationToken)));

        bool descriptorIsNew = descriptor is null;
        descriptor ??= new PdfDictionary();
        descriptor.Items["Type"] = new PdfName("FontDescriptor");
        descriptor.Items["FontName"] = new PdfName(taggedName);
        descriptor.Items["Flags"] = new PdfNumber((program.Flags & ~32) | 4);
        var box = new PdfArray();
        foreach (int value in program.FontBBox) box.Items.Add(new PdfNumber(value));
        descriptor.Items["FontBBox"] = box;
        descriptor.Items["ItalicAngle"] = new PdfNumber(program.ItalicAngle);
        descriptor.Items["Ascent"] = new PdfNumber(program.Ascent);
        descriptor.Items["Descent"] = new PdfNumber(program.Descent);
        descriptor.Items["CapHeight"] = new PdfNumber(program.CapHeight);
        descriptor.Items["StemV"] = new PdfNumber(program.StemV);
        descriptor.Items.Remove("FontFile");
        descriptor.Items.Remove("FontFile3");
        descriptor.Items.Remove("CharSet");
        descriptor.Items["FontFile2"] = fontFile;
        if (descriptorIsNew) font.Items["FontDescriptor"] = AddObject(descriptor);

        font.Items["Subtype"] = new PdfName("TrueType");
        font.Items["BaseFont"] = new PdfName(taggedName);
        font.Items.Remove("Encoding");
        if (!font.Items.ContainsKey("Widths")) {
            int first = glyphs.Keys.First();
            int last = glyphs.Keys.Last();
            var widths = new PdfArray();
            for (int code = first; code <= last; code++) {
                double width = glyphs.TryGetValue(code, out int glyph)
                    ? PdfStandardFontWidths.TryGetSymbolicWidth1000(dingbats, glyphNames[code]!, out int standardWidth) ? standardWidth : program.GetExactGlyphWidth1000(glyph)
                    : 0D;
                widths.Items.Add(new PdfNumber(Math.Round(width, 3)));
            }
            font.Items["FirstChar"] = new PdfNumber(first);
            font.Items["LastChar"] = new PdfNumber(last);
            font.Items["Widths"] = widths;
        }
        if (!font.Items.ContainsKey("ToUnicode")) {
            var mapping = new SortedDictionary<int, string>();
            foreach (int code in glyphs.Keys) {
                string name = glyphNames[code]!;
                if (PdfFontEncodingNames.TryGetSymbolicUnicode(name, out int scalar)) mapping[code] = char.ConvertFromUtf32(scalar);
                else if (ResourceResolver.TryDecodeStandardGlyphName(name, out string? text) && IsValidUnicode(text)) mapping[code] = text!;
            }
            if (mapping.Count > 0) font.Items["ToUnicode"] = AddObject(BuildToUnicodeStream(mapping, 1, null));
        }
    }
}
