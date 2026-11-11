using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    // PDF/A validators compare dictionary and program widths with a tolerance of one unit.
    private const double WidthTolerance = 1D;

    private static readonly char[] FamilySeparators = { ',', '-' };

    private void ProcessFonts() {
        foreach (FontUse use in _fontUses.Values) {
            _cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary font = use.Font;
            string subtype = GetName(font, "Subtype") ?? string.Empty;
            string baseFont = GetName(font, "BaseFont") ?? "(unnamed)";
            string location = use.Location + ", font " + use.ResourceName + " (" + baseFont + ")";
            switch (subtype) {
                case "Type3":
                    break;
                case "Type0":
                    ProcessCompositeFont(use, location);
                    break;
                default:
                    ProcessSimpleFont(use, location, baseFont);
                    break;
            }
            NormalizeFontNames(use, location);
            if (_unicode) EnsureUnicode(use, location);
        }
    }

    /// <summary>PDF/A requires font names to be valid UTF-8; legacy names in other encodings are renamed, which no viewer uses to draw.</summary>
    private void NormalizeFontNames(FontUse use, string location) {
        var named = new List<(PdfDictionary Dictionary, string Key)> { (use.Font, "BaseFont") };
        if (Resolve(use.Font.Items.TryGetValue("DescendantFonts", out PdfObject? descendants) ? descendants : null) is PdfArray array) {
            foreach (PdfObject item in array.Items) {
                if (ResolveDictionary(item) is PdfDictionary descendant) {
                    named.Add((descendant, "BaseFont"));
                    if (ResolveDictionary(descendant.Items.TryGetValue("FontDescriptor", out PdfObject? cidDescriptor) ? cidDescriptor : null) is PdfDictionary d) named.Add((d, "FontName"));
                }
            }
        }
        if (ResolveDictionary(use.Font.Items.TryGetValue("FontDescriptor", out PdfObject? fd) ? fd : null) is PdfDictionary descriptor) named.Add((descriptor, "FontName"));
        foreach ((PdfDictionary dictionary, string key) in named) {
            if (GetName(dictionary, key) is not string name || IsPortableName(name)) continue;
            string replacement = ToPortableFontName(name);
            if (Change(PdfAConversionChangeKind.FontDescriptorUpdated, location,
                    "Renamed the font from a name that is not valid UTF-8 to " + replacement + ", as PDF/A requires. Font names do not affect drawing.", dictionary)) {
                dictionary.Items[key] = new PdfName(replacement);
            }
        }
    }

    private static bool IsPortableName(string name) {
        foreach (char character in name) {
            if (character < 0x21 || (character >= 0x7F && character <= 0x9F) || character == '\uFFFD') return false;
        }
        return true;
    }

    private static string ToPortableFontName(string name) {
        string prefix = name.Length > 7 && name[6] == '+' && name.Take(6).All(static c => c is >= 'A' and <= 'Z') ? name.Substring(0, 7) : string.Empty;
        uint hash = 2166136261;
        foreach (char character in name) hash = unchecked((hash ^ character) * 16777619);
        return prefix + "Font" + hash.ToString("X8", CultureInfo.InvariantCulture);
    }

    private void ProcessCompositeFont(FontUse use, string location) {
        PdfDictionary? descendant = Resolve(use.Font.Items.TryGetValue("DescendantFonts", out PdfObject? descendants) ? descendants : null) is PdfArray array &&
            array.Items.Count > 0 ? ResolveDictionary(array.Items[0]) : null;
        PdfDictionary? descriptor = descendant is null ? null : ResolveDictionary(descendant.Items.TryGetValue("FontDescriptor", out PdfObject? fd) ? fd : null);
        if (descriptor is null || !HasFontProgram(descriptor)) {
            if (use.Visible) {
                AddViolation(PdfAConversionRefusalReason.FontNotEmbeddable, location,
                    "The composite (CID) font is not embedded. Its glyphs are addressed by CID, so no installed font can stand in without changing the page.", use.Font);
            }
            return;
        }
        if (GetName(descendant!, "Subtype") == "CIDFontType2" && !descendant!.Items.ContainsKey("CIDToGIDMap") &&
            Change(PdfAConversionChangeKind.FontDescriptorUpdated, location, "Wrote the implied /Identity CIDToGIDMap, which PDF/A requires to be explicit.", descendant)) {
            descendant!.Items["CIDToGIDMap"] = new PdfName("Identity");
        }
        if (descriptor.Items.ContainsKey("CIDSet") &&
            Change(PdfAConversionChangeKind.FontDescriptorUpdated, location, "Removed the optional CIDSet, which PDF/A requires to list every glyph exactly.", descriptor)) {
            descriptor.Items.Remove("CIDSet");
        }
        CheckEmbeddedProgram(use, location, composite: true);
    }

    private void ProcessSimpleFont(FontUse use, string location, string baseFont) {
        PdfDictionary? descriptor = ResolveDictionary(use.Font.Items.TryGetValue("FontDescriptor", out PdfObject? fd) ? fd : null);
        if (descriptor is not null && HasFontProgram(descriptor)) {
            if (descriptor.Items.ContainsKey("CharSet") &&
                Change(PdfAConversionChangeKind.FontDescriptorUpdated, location, "Removed the optional CharSet, which PDF/A requires to list every glyph exactly.", descriptor)) {
                descriptor.Items.Remove("CharSet");
            }
            FixTrueTypeEncoding(use, descriptor, location);
            CheckEmbeddedProgram(use, location, composite: false);
            return;
        }
        int flags = descriptor is not null && Resolve(descriptor.Items.TryGetValue("Flags", out PdfObject? flagValue) ? flagValue : null) is PdfNumber flagNumber
            ? (int)flagNumber.Value : 32;
        if (((flags & 4) == 0 && !IsSymbolicName(StripSubsetTag(baseFont))) || PdfWriter.TryGetStandardFontByBaseFontName(StripSubsetTag(baseFont), out _)) {
            string?[] names = BuildSimpleEncoding(use.Font);
            var notdef = use.Codes.Where(code => code >= names.Length || names[code] is null || names[code] == ".notdef").ToList();
            if (notdef.Count > 0) {
                AddViolation(PdfAConversionRefusalReason.FontGlyphMissing, location,
                    "Code(s) " + FormatCodes(notdef) + " select no glyph in the font's encoding, so they draw .notdef, which PDF/A forbids even for invisible text.", use.Font);
                return;
            }
        }
        if (!use.Visible) return;
        if (!_options.EmbedSubstituteFonts) {
            AddViolation(PdfAConversionRefusalReason.FontNotEmbeddable, location,
                "The font is not embedded, and EmbedSubstituteFonts is false.", use.Font);
            return;
        }
        EmbedSubstitute(use, descriptor, location, baseFont);
    }

    private bool HasFontProgram(PdfDictionary descriptor) =>
        Resolve(descriptor.Items.TryGetValue("FontFile", out PdfObject? f1) ? f1 : null) is PdfStream ||
        Resolve(descriptor.Items.TryGetValue("FontFile2", out PdfObject? f2) ? f2 : null) is PdfStream ||
        Resolve(descriptor.Items.TryGetValue("FontFile3", out PdfObject? f3) ? f3 : null) is PdfStream;

    private void EmbedSubstitute(FontUse use, PdfDictionary? descriptor, string location, string baseFont) {
        PdfDictionary font = use.Font;
        string plainName = StripSubsetTag(baseFont);
        bool standard = PdfWriter.TryGetStandardFontByBaseFontName(plainName, out PdfStandardFont standardFont);
        int? descriptorFlags = descriptor is not null && Resolve(descriptor.Items.TryGetValue("Flags", out PdfObject? flagValue) ? flagValue : null) is PdfNumber flagNumber
            ? (int)flagNumber.Value : null;
        bool symbolicName = IsSymbolicName(plainName);
        if (symbolicName || (!standard && descriptorFlags is int flags && (flags & 4) != 0 && (flags & 32) == 0)) {
            AddViolation(PdfAConversionRefusalReason.FontNotEmbeddable, location,
                "The symbolic font is not embedded. Its glyphs have no standard names to find them in a substitute.", font);
            return;
        }

        string?[] glyphNames = BuildSimpleEncoding(font);
        var unicodeByCode = new Dictionary<int, int>();
        var unnamed = new List<int>();
        // A font shown only with empty strings still needs a program; the space glyph keeps the subset valid.
        IEnumerable<int> codes = use.Codes.Count > 0 ? use.Codes : new[] { 32 };
        foreach (int code in codes) {
            string? glyphName = code < glyphNames.Length ? glyphNames[code] : null;
            if (glyphName is null || !ResourceResolver.TryDecodeStandardGlyphName(glyphName, out string? text) || text is null ||
                text.Length == 0 || char.ConvertToUtf32(text, 0) is int scalar && text.Length != char.ConvertFromUtf32(scalar).Length) {
                if (use.VisibleCodes.Contains(code)) unnamed.Add(code);
                continue;
            }
            unicodeByCode[code] = char.ConvertToUtf32(text, 0);
        }
        if (unnamed.Count > 0) {
            AddViolation(PdfAConversionRefusalReason.FontGlyphMissing, location,
                "The font is not embedded and character code(s) " + FormatCodes(unnamed) +
                " have no standard glyph name, so they cannot be found in a substitute font.", font);
            return;
        }

        bool bold = PdfFontStyleEvidence.IsBold(plainName, descriptor is not null &&
            Resolve(descriptor.Items.TryGetValue("FontWeight", out PdfObject? weight) ? weight : null) is PdfNumber weightNumber ? (int)weightNumber.Value : null);
        bool italic = PdfFontStyleEvidence.IsItalic(plainName, descriptorFlags);
        IEnumerable<string> families = standard ? StandardSubstitutes(standardFont) : FamilyCandidates(plainName);

        string? lastFailure = null;
        PdfAConversionRefusalReason failureReason = PdfAConversionRefusalReason.FontNotEmbeddable;
        foreach (string family in families) {
            if (!PdfEmbeddedFontFamily.TryFromSystem(family, out PdfEmbeddedFontFamily? installed) || installed is null) continue;
            byte[]? face = bold && italic ? installed.BoldItalicSnapshot : bold ? installed.BoldSnapshot : italic ? installed.ItalicSnapshot : installed.RegularSnapshot;
            if (face is null) {
                lastFailure ??= "the installed " + family + " family has no " + (bold ? italic ? "bold italic" : "bold" : "italic") + " face";
                continue;
            }
            PdfTrueTypeFontProgram program;
            try {
                program = PdfTrueTypeFontProgram.Parse(face);
            } catch (Exception exception) when (exception is NotSupportedException or InvalidDataException or ArgumentException or IndexOutOfRangeException) {
                lastFailure = "the installed " + family + " face cannot be embedded: " + exception.Message;
                continue;
            }

            var glyphs = new Dictionary<int, int>();
            var missing = new List<int>();
            var mismatched = new List<string>();
            foreach (KeyValuePair<int, int> mapping in unicodeByCode) {
                if (!program.TryGetGlyphId(mapping.Value, out int glyph) || glyph <= 0) {
                    if (use.VisibleCodes.Contains(mapping.Key)) missing.Add(mapping.Key);
                    continue;
                }
                glyphs[mapping.Key] = glyph;
                if (!use.VisibleCodes.Contains(mapping.Key)) continue;
                double? expected = ExpectedWidth(font, descriptor, mapping.Key, mapping.Value, standard ? standardFont : null);
                double actual = program.GetExactGlyphWidth1000(glyph);
                if (expected is double width && Math.Abs(width - actual) > WidthTolerance && mismatched.Count < 8) {
                    mismatched.Add(FormatCode(mapping.Key) + " " + width.ToString("0.#", CultureInfo.InvariantCulture) + " vs " + actual.ToString("0.#", CultureInfo.InvariantCulture));
                }
            }
            if (missing.Count > 0) {
                lastFailure = "the installed " + family + " face has no glyph for code(s) " + FormatCodes(missing);
                failureReason = PdfAConversionRefusalReason.FontGlyphMissing;
                continue;
            }
            if (mismatched.Count > 0) {
                lastFailure = "the installed " + family + " face has different glyph widths (" + string.Join(", ", mismatched) + ")";
                failureReason = PdfAConversionRefusalReason.FontMetricsMismatch;
                continue;
            }
            if (glyphs.Count == 0) {
                lastFailure = "the installed " + family + " face has none of the font's glyphs";
                failureReason = PdfAConversionRefusalReason.FontGlyphMissing;
                continue;
            }

            if (Change(PdfAConversionChangeKind.FontEmbedded, location,
                    "Embedded the installed " + family + " face (" + program.FontName + ") for the non-embedded font " + plainName +
                    "; every drawn glyph has the advance width the page already positions it by.", font, OfficeConversionLossKind.Approximation)) {
                WriteSubstitute(font, descriptor, plainName, program, glyphNames, glyphs, standard ? standardFont : null);
            }
            return;
        }

        if (lastFailure is null) {
            AddViolation(PdfAConversionRefusalReason.FontNotEmbeddable, location,
                "The font is not embedded and no installed substitute (" + string.Join(", ", families) + ") was found.", font);
        } else {
            AddViolation(failureReason, location, "The font is not embedded and cannot be substituted without changing the page: " + lastFailure + ".", font);
        }
    }

    private void WriteSubstitute(PdfDictionary font, PdfDictionary? descriptor, string plainName, PdfTrueTypeFontProgram program,
        string?[] glyphNames, Dictionary<int, int> glyphs, PdfStandardFont? standardFont) {
        byte[] subset = program.BuildGlyphRetainingSubset(glyphs.Values.Append(0));
        string taggedName = PdfSubsetTag.Apply(SanitizeFontName(plainName), glyphs.Values.OrderBy(static glyph => glyph), subset);

        var fileDictionary = new PdfDictionary();
        fileDictionary.Items["Length1"] = new PdfNumber(subset.Length);
        fileDictionary.Items["Filter"] = new PdfName("FlateDecode");
        PdfReference fontFile = AddObject(new PdfStream(fileDictionary, PdfFlateEncoder.Compress(subset, _cancellationToken)));

        bool descriptorIsNew = descriptor is null;
        descriptor ??= new PdfDictionary();
        descriptor.Items["Type"] = new PdfName("FontDescriptor");
        descriptor.Items["FontName"] = new PdfName(taggedName);
        descriptor.Items["Flags"] = new PdfNumber((program.Flags & ~4) | 32);
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
        string?[] winAnsi = PdfFontEncodingNames.Get("WinAnsiEncoding")!;
        var differences = new PdfArray();
        int previous = -2;
        foreach (int code in glyphs.Keys.OrderBy(static code => code)) {
            string name = glyphNames[code]!;
            if (string.Equals(winAnsi[code], name, StringComparison.Ordinal)) continue;
            if (code != previous + 1) differences.Items.Add(new PdfNumber(code));
            differences.Items.Add(new PdfName(name));
            previous = code;
        }
        if (differences.Items.Count == 0) {
            font.Items["Encoding"] = new PdfName("WinAnsiEncoding");
        } else {
            var encoding = new PdfDictionary();
            encoding.Items["Type"] = new PdfName("Encoding");
            encoding.Items["BaseEncoding"] = new PdfName("WinAnsiEncoding");
            encoding.Items["Differences"] = differences;
            font.Items["Encoding"] = encoding;
        }

        if (!font.Items.ContainsKey("Widths")) {
            int first = glyphs.Keys.Min();
            int last = glyphs.Keys.Max();
            var widths = new PdfArray();
            for (int code = first; code <= last; code++) {
                double width = 0D;
                if (glyphs.TryGetValue(code, out int glyph)) {
                    int unicode = ResourceResolver.TryDecodeStandardGlyphName(glyphNames[code]!, out string? text) && text is { Length: > 0 } ? char.ConvertToUtf32(text, 0) : 0;
                    width = ExpectedWidth(font, null, code, unicode, standardFont) ?? program.GetExactGlyphWidth1000(glyph);
                }
                widths.Items.Add(new PdfNumber(Math.Round(width, 3)));
            }
            font.Items["FirstChar"] = new PdfNumber(first);
            font.Items["LastChar"] = new PdfNumber(last);
            font.Items["Widths"] = widths;
        }
    }

    /// <summary>Returns the width a viewer positions the glyph by: the Widths entry, or the standard font metrics.</summary>
    private double? ExpectedWidth(PdfDictionary font, PdfDictionary? descriptor, int code, int unicode, PdfStandardFont? standardFont) {
        if (Resolve(font.Items.TryGetValue("Widths", out PdfObject? widthsValue) ? widthsValue : null) is PdfArray widths) {
            int first = Resolve(font.Items.TryGetValue("FirstChar", out PdfObject? firstValue) ? firstValue : null) is PdfNumber firstNumber ? (int)firstNumber.Value : 0;
            int index = code - first;
            if (index >= 0 && index < widths.Items.Count && Resolve(widths.Items[index]) is PdfNumber width) return width.Value;
            return descriptor is not null && Resolve(descriptor.Items.TryGetValue("MissingWidth", out PdfObject? missing) ? missing : null) is PdfNumber missingWidth
                ? missingWidth.Value : 0D;
        }
        if (standardFont is PdfStandardFont standard) {
            if (standard is PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique) return 600D;
            if (unicode > 0 && unicode <= 0xFFFF && PdfStandardFontWidths.TryGetWidth1000(standard, (char)unicode, out int standardWidth)) return standardWidth;
            return -1D;
        }
        return null;
    }

    /// <summary>The named family first, as viewers show it when installed, then the library's metric-compatible substitutes.</summary>
    private static List<string> StandardSubstitutes(PdfStandardFont font) {
        string family = font switch {
            PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => "Times",
            PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => "Courier",
            _ => "Helvetica"
        };
        var candidates = new List<string> { family };
        foreach (string substitute in OfficeIMO.Drawing.OfficeDocumentFontSubstitutions.EnumerateGeneric(family)) {
            if (!candidates.Contains(substitute, StringComparer.OrdinalIgnoreCase)) candidates.Add(substitute);
        }
        return candidates;
    }

    private static List<string> FamilyCandidates(string plainName) {
        var candidates = new List<string>();
        string? family = PdfFontFamilyNames.ToFamilyName(plainName);
        if (!string.IsNullOrWhiteSpace(family)) candidates.Add(family!);
        int separator = plainName.IndexOfAny(FamilySeparators);
        string stem = separator > 0 ? plainName.Substring(0, separator) : plainName;
        if (!candidates.Contains(stem, StringComparer.OrdinalIgnoreCase)) candidates.Add(stem);
        foreach (string substitute in OfficeIMO.Drawing.OfficeDocumentFontSubstitutions.Enumerate(candidates[0])) {
            if (!candidates.Contains(substitute, StringComparer.OrdinalIgnoreCase)) candidates.Add(substitute);
        }
        return candidates;
    }

    /// <summary>Returns the glyph name each simple-font code selects, from the base encoding and Differences.</summary>
    private string?[] BuildSimpleEncoding(PdfDictionary font) {
        string plainName = StripSubsetTag(GetName(font, "BaseFont") ?? string.Empty);
        string?[] baseNames = PdfFontEncodingNames.Get("StandardEncoding")!;
        PdfObject? encoding = Resolve(font.Items.TryGetValue("Encoding", out PdfObject? encodingValue) ? encodingValue : null);
        PdfArray? differences = null;
        if (encoding is PdfName encodingName) {
            baseNames = PdfFontEncodingNames.Get(encodingName.Name) ?? baseNames;
        } else if (encoding is PdfDictionary encodingDictionary) {
            if (GetName(encodingDictionary, "BaseEncoding") is string baseEncoding) baseNames = PdfFontEncodingNames.Get(baseEncoding) ?? baseNames;
            differences = Resolve(encodingDictionary.Items.TryGetValue("Differences", out PdfObject? differencesValue) ? differencesValue : null) as PdfArray;
        }
        var names = (string?[])baseNames.Clone();
        if (differences is not null) {
            int code = 0;
            foreach (PdfObject item in differences.Items) {
                PdfObject? resolved = Resolve(item);
                if (resolved is PdfNumber number) {
                    code = (int)number.Value;
                } else if (resolved is PdfName name) {
                    if (code >= 0 && code < names.Length) names[code] = name.Name;
                    code++;
                }
            }
        }
        _ = plainName;
        return names;
    }

    private static bool IsSymbolicName(string name) =>
        name.StartsWith("Symbol", StringComparison.OrdinalIgnoreCase) || name.StartsWith("ZapfDingbats", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase) || name.StartsWith("Webdings", StringComparison.OrdinalIgnoreCase);

    private static string StripSubsetTag(string name) =>
        name.Length > 7 && name[6] == '+' && name.Take(6).All(static character => character is >= 'A' and <= 'Z') ? name.Substring(7) : name;

    private static string SanitizeFontName(string name) {
        var builder = new StringBuilder(name.Length);
        foreach (char character in name) {
            if (character > ' ' && character < 127 && character is not ('(' or ')' or '<' or '>' or '[' or ']' or '{' or '}' or '/' or '%' or '#')) builder.Append(character);
        }
        return builder.Length == 0 ? "Font" : builder.ToString();
    }

    private static string FormatCode(int code) => code.ToString(CultureInfo.InvariantCulture);

    private static string FormatCodes(List<int> codes) {
        string text = string.Join(", ", codes.Take(12).Select(FormatCode));
        return codes.Count > 12 ? text + " and " + (codes.Count - 12).ToString(CultureInfo.InvariantCulture) + " more" : text;
    }
}
