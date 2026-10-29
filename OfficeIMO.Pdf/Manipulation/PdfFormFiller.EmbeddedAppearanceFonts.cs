namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static bool TryCreateEmbeddedTextAppearanceFontPlan(
        PdfFormFillerOptions? options,
        string displayValue,
        string diagnosticSource,
        ref int nextObjectNumber,
        out TextAppearanceFontPlan? fontPlan,
        out string? failureMessage) {
        fontPlan = null;
        failureMessage = null;
        if (string.IsNullOrEmpty(displayValue) || options?.AppearanceFontFamilySnapshot is not PdfEmbeddedFontFamily family) {
            return false;
        }

        return TryCreateEmbeddedTextAppearanceFontPlan(family, options, displayValue, diagnosticSource, ref nextObjectNumber, out fontPlan, out failureMessage);
    }

    /// <summary>
    /// Uses the installed font behind an embedded subset in the form's default resources when the subset lacks glyphs
    /// for a new value. A form authored with an embedded body font then keeps that font for filled values instead of
    /// falling back to a non-embedded standard font.
    /// </summary>
    private static bool TryCreateInstalledInheritedTextAppearanceFontPlan(
        Dictionary<int, PdfIndirectObject> objects,
        PdfFormFillerOptions? options,
        string displayValue,
        string diagnosticSource,
        ref int nextObjectNumber,
        out TextAppearanceFontPlan? fontPlan,
        params PdfDictionary?[] resourceCandidates) {
        fontPlan = null;
        if (string.IsNullOrEmpty(displayValue) || options?.HasAppearanceFontFamily == true || options?.HasAppearanceFontFallbacks == true) {
            return false;
        }

        foreach (PdfDictionary resources in EnumerateCandidateTextAppearanceResources(resourceCandidates)) {
            if (!resources.Items.TryGetValue("Font", out PdfObject? fontsObject) ||
                ResolveDictionary(objects, fontsObject) is not PdfDictionary fonts) {
                continue;
            }

            IEnumerable<string> names = fonts.Items.Keys
                .OrderBy(name => string.Equals(name, DefaultAppearanceFontName, StringComparison.Ordinal) ? 0 : 1)
                .ThenBy(name => name, StringComparer.Ordinal);
            foreach (string name in names) {
                string? familyName = TryReadEmbeddedFontFamilyName(objects, fonts.Items[name]);
                if (familyName == null ||
                    !PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfEmbeddedFontFamily? family) || family == null) {
                    continue;
                }

                if (TryCreateEmbeddedTextAppearanceFontPlan(family, options, displayValue, diagnosticSource, ref nextObjectNumber, out fontPlan, out _, shareWith: objects)) {
                    return true;
                }
            }
        }

        return false;
    }

    // Reads the family of a Type0 or TrueType font that carries an embedded font file, without its subset tag or style suffix.
    private static string? TryReadEmbeddedFontFamilyName(Dictionary<int, PdfIndirectObject> objects, PdfObject fontObject) {
        if (ResolveDictionary(objects, fontObject) is not PdfDictionary font) return null;
        PdfDictionary? descriptorOwner = font;
        if (font.Get<PdfName>("Subtype")?.Name == "Type0") {
            descriptorOwner = ResolveObject(objects, font.Items.TryGetValue("DescendantFonts", out PdfObject? descendants) ? descendants : null) is PdfArray array &&
                array.Items.Count > 0
                    ? ResolveDictionary(objects, array.Items[0])
                    : null;
        }

        if (descriptorOwner == null ||
            ResolveDictionary(objects, descriptorOwner.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject) ? descriptorObject : null) is not PdfDictionary descriptor ||
            !(descriptor.Items.ContainsKey("FontFile2") || descriptor.Items.ContainsKey("FontFile3") || descriptor.Items.ContainsKey("FontFile"))) {
            return null;
        }

        string? baseFont = (ResolveObject(objects, descriptor.Items.TryGetValue("FontFamily", out PdfObject? familyObject) ? familyObject : null) as PdfStringObj)?.Value
            ?? font.Get<PdfName>("BaseFont")?.Name;
        if (string.IsNullOrWhiteSpace(baseFont)) return null;
        string name = baseFont!;
        int plus = name.IndexOf('+');
        if (plus == 6) name = name.Substring(plus + 1);
        for (int i = 1; i < name.Length; i++) {
            if (name[i] == '-' || name[i] == ',') {
                name = name.Substring(0, i);
                break;
            }
        }
        return string.IsNullOrWhiteSpace(name) ? null : name;
    }

    private static bool TryCreateEmbeddedTextAppearanceFontPlan(
        PdfEmbeddedFontFamily family,
        PdfFormFillerOptions? options,
        string displayValue,
        string diagnosticSource,
        ref int nextObjectNumber,
        out TextAppearanceFontPlan? fontPlan,
        out string? failureMessage,
        Dictionary<int, PdfIndirectObject>? shareWith = null) {
        fontPlan = null;
        failureMessage = null;

        // Fields filled in one operation share one embedded font: each field adds its glyphs to the same
        // program and rewrites the same objects, so the last write covers every value.
        SharedEmbeddedAppearanceFont? shared = null;
        if (shareWith != null) {
            Dictionary<string, SharedEmbeddedAppearanceFont> sharedFonts = SharedEmbeddedAppearanceFonts.GetOrCreateValue(shareWith);
            sharedFonts.TryGetValue(family.FamilyName, out shared);
        }

        EmbeddedTextAppearanceFontProgram fontProgram;
        try {
            fontProgram = shared?.Program ?? EmbeddedTextAppearanceFontProgram.Parse(family.RegularSnapshot, family.FamilyName);
        } catch (Exception exception) when (
            exception is NotSupportedException ||
            exception is ArgumentException ||
            exception is ArithmeticException ||
            exception is FormatException ||
            exception is IndexOutOfRangeException ||
            exception is InvalidOperationException) {
            failureMessage = $"The configured appearance font '{family.FamilyName}' could not be parsed as a supported embedded font.";
            return false;
        }

        options?.AddTextShapingDiagnostics(PdfTextDiagnostics.AnalyzeAdvancedTextLayout(displayValue, family.RegularSnapshot, diagnosticSource, family.FamilyName));

        if (!TryCreateEmbeddedTextSegmentEncoder(fontProgram, displayValue, out string? mappedHex, out Func<string, string?>? segmentEncoder)) {
            options?.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeEmbeddedFontText(displayValue, family.RegularSnapshot, diagnosticSource, family.FamilyName));
            failureMessage = $"The configured appearance font '{family.FamilyName}' cannot encode every character required by the form field appearance.";
            return false;
        }
        options?.AddFontDiagnostics(fontProgram.AnalyzeFullFontEmbedding(diagnosticSource));

        if (shared == null) {
            shared = new SharedEmbeddedAppearanceFont(fontProgram, nextObjectNumber);
            nextObjectNumber += 5;
            if (shareWith != null) SharedEmbeddedAppearanceFonts.GetOrCreateValue(shareWith)[family.FamilyName] = shared;
        }
        int fontFileObjectNumber = shared.FirstObjectNumber;
        int descriptorObjectNumber = fontFileObjectNumber + 1;
        int descendantFontObjectNumber = fontFileObjectNumber + 2;
        int toUnicodeObjectNumber = fontFileObjectNumber + 3;
        int type0FontObjectNumber = fontFileObjectNumber + 4;

        var appearanceFonts = new PdfDictionary();
        appearanceFonts.Items[DefaultAppearanceFontName] = new PdfReference(type0FontObjectNumber, 0);

        var resources = new PdfDictionary();
        resources.Items["Font"] = appearanceFonts;

        fontPlan = new TextAppearanceFontPlan(
            DefaultAppearanceFontName,
            resources,
            mappedHex,
            segmentEncoder,
            (text, fontSize) => fontProgram.MeasureTextWidth(text, fontSize),
            encodeTextSegments: null,
            (objects, _) => fontProgram.Materialize(
                objects,
                fontFileObjectNumber,
                descriptorObjectNumber,
                descendantFontObjectNumber,
                toUnicodeObjectNumber,
                type0FontObjectNumber));
        return true;
    }

    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<Dictionary<int, PdfIndirectObject>, Dictionary<string, SharedEmbeddedAppearanceFont>> SharedEmbeddedAppearanceFonts =
        new System.Runtime.CompilerServices.ConditionalWeakTable<Dictionary<int, PdfIndirectObject>, Dictionary<string, SharedEmbeddedAppearanceFont>>();

    private sealed class SharedEmbeddedAppearanceFont {
        internal SharedEmbeddedAppearanceFont(EmbeddedTextAppearanceFontProgram program, int firstObjectNumber) {
            Program = program;
            FirstObjectNumber = firstObjectNumber;
        }

        internal EmbeddedTextAppearanceFontProgram Program { get; }
        internal int FirstObjectNumber { get; }
    }

    private static bool TryCreateEmbeddedTextSegmentEncoder(EmbeddedTextAppearanceFontProgram fontProgram, string displayValue, out string? encodedTextHex, out Func<string, string?>? encodeTextSegmentHex) {
        encodedTextHex = null;
        encodeTextSegmentHex = segment => TryEncodeEmbeddedText(fontProgram, segment, out string segmentHex) ? segmentHex : null;
        if (TryEncodeEmbeddedText(fontProgram, displayValue, out string mappedHex)) {
            encodedTextHex = mappedHex;
            return true;
        }

        string normalized = displayValue.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        if (lines.Length == 1) {
            encodeTextSegmentHex = null;
            return false;
        }

        for (int i = 0; i < lines.Length; i++) {
            if (lines[i].Length > 0 && !TryEncodeEmbeddedText(fontProgram, lines[i], out _)) {
                encodeTextSegmentHex = null;
                return false;
            }
        }

        return true;
    }

    private static bool TryEncodeEmbeddedText(EmbeddedTextAppearanceFontProgram fontProgram, string text, out string hex) {
        try {
            hex = fontProgram.EncodeTextAsGlyphHex(text);
            return true;
        } catch (Exception exception) when (
            exception is ArgumentException ||
            exception is InvalidOperationException ||
            exception is NotSupportedException) {
            hex = string.Empty;
            return false;
        }
    }

    private sealed class EmbeddedTextAppearanceFontProgram {
        private readonly PdfTrueTypeFontProgram? _trueTypeFont;
        private readonly PdfOpenTypeCffFontProgram? _cffFont;

        private EmbeddedTextAppearanceFontProgram(PdfTrueTypeFontProgram font) {
            _trueTypeFont = font;
            _cffFont = null;
        }

        private EmbeddedTextAppearanceFontProgram(PdfOpenTypeCffFontProgram font) {
            _trueTypeFont = null;
            _cffFont = font;
        }

        public static EmbeddedTextAppearanceFontProgram Parse(byte[] fontData, string fontName) =>
            IsOpenTypeCffFontData(fontData)
                ? new EmbeddedTextAppearanceFontProgram(PdfFontProgramCache.GetOpenTypeCff(fontData, fontName))
                : new EmbeddedTextAppearanceFontProgram(PdfFontProgramCache.GetTrueType(fontData, fontName));

        public string EncodeTextAsGlyphHex(string text) {
            if (_trueTypeFont != null) {
                return _trueTypeFont.EncodeTextAsGlyphHex(text);
            }

            if (_cffFont != null) {
                return _cffFont.EncodeTextAsGlyphHex(text);
            }

            throw new InvalidOperationException("The configured appearance font was not parsed.");
        }

        public double MeasureTextWidth(string text, double fontSize) {
            if (_trueTypeFont != null) {
                return _trueTypeFont.MeasureTextWidth(text, fontSize);
            }

            if (_cffFont != null) {
                return _cffFont.MeasureTextWidth(text, fontSize);
            }

            throw new InvalidOperationException("The configured appearance font was not parsed.");
        }

        public IReadOnlyList<PdfFontEmbeddingDiagnostic> AnalyzeFullFontEmbedding(string source) {
            if (_cffFont != null) {
                return PdfFontDiagnostics.AnalyzeOpenTypeCffCompactEmbedding(_cffFont, source);
            }

            return Array.Empty<PdfFontEmbeddingDiagnostic>();
        }

        public void Materialize(
            Dictionary<int, PdfIndirectObject> objects,
            int fontFileObjectNumber,
            int descriptorObjectNumber,
            int descendantFontObjectNumber,
            int toUnicodeObjectNumber,
            int type0FontObjectNumber) {
            if (_trueTypeFont != null) {
                MaterializeEmbeddedTextAppearanceFont(
                    objects,
                    _trueTypeFont,
                    fontFileObjectNumber,
                    descriptorObjectNumber,
                    descendantFontObjectNumber,
                    toUnicodeObjectNumber,
                    type0FontObjectNumber);
                return;
            }

            if (_cffFont != null) {
                MaterializeEmbeddedTextAppearanceFont(
                    objects,
                    _cffFont,
                    fontFileObjectNumber,
                    descriptorObjectNumber,
                    descendantFontObjectNumber,
                    toUnicodeObjectNumber,
                    type0FontObjectNumber);
                return;
            }

            throw new InvalidOperationException("The configured appearance font was not parsed.");
        }
    }

    /// <summary>
    /// Embeds a subset of <paramref name="fontData"/> for <paramref name="text"/>: the font file, descriptor, descendant
    /// and ToUnicode objects are added to <paramref name="objects"/>, and the Type0 font dictionary is returned for the
    /// caller to place. Returns false when the font cannot be parsed or lacks a character of the text.
    /// </summary>
    internal static bool TryCreateEmbeddedTextFont(byte[] fontData, string fontName, string text, double fontSize,
        Dictionary<int, PdfIndirectObject> objects, ref int nextObjectNumber,
        out PdfObject type0Font, out string glyphHex, out double textWidth) {
        type0Font = null!;
        glyphHex = string.Empty;
        textWidth = 0D;
        EmbeddedTextAppearanceFontProgram program;
        try {
            program = EmbeddedTextAppearanceFontProgram.Parse(fontData, fontName);
        } catch (Exception exception) when (
            exception is NotSupportedException ||
            exception is ArgumentException ||
            exception is ArithmeticException ||
            exception is FormatException ||
            exception is IndexOutOfRangeException ||
            exception is InvalidOperationException) {
            return false;
        }

        if (!TryEncodeEmbeddedText(program, text, out glyphHex)) return false;
        textWidth = program.MeasureTextWidth(text, fontSize);
        int fontFile = nextObjectNumber++;
        int descriptor = nextObjectNumber++;
        int descendant = nextObjectNumber++;
        int toUnicode = nextObjectNumber++;
        int type0 = nextObjectNumber++;
        program.Materialize(objects, fontFile, descriptor, descendant, toUnicode, type0);
        type0Font = objects[type0].Value;
        objects.Remove(type0);
        return true;
    }

    internal static void MaterializeEmbeddedTextAppearanceFont(
        Dictionary<int, PdfIndirectObject> objects,
        PdfTrueTypeFontProgram fontProgram,
        int fontFileObjectNumber,
        int descriptorObjectNumber,
        int descendantFontObjectNumber,
        int toUnicodeObjectNumber,
        int type0FontObjectNumber) {
        byte[] fontData = fontProgram.BuildIdentitySubsetFontFile();
        objects[fontFileObjectNumber] = new PdfIndirectObject(fontFileObjectNumber, 0, CreateEmbeddedFontFileStream(fontData));
        objects[descriptorObjectNumber] = new PdfIndirectObject(descriptorObjectNumber, 0, CreateTrueTypeFontDescriptor(fontProgram, fontFileObjectNumber));
        objects[descendantFontObjectNumber] = new PdfIndirectObject(descendantFontObjectNumber, 0, CreateCidFontType2Descendant(fontProgram, descriptorObjectNumber));
        objects[toUnicodeObjectNumber] = new PdfIndirectObject(toUnicodeObjectNumber, 0, CreateToUnicodeStream(fontProgram));
        objects[type0FontObjectNumber] = new PdfIndirectObject(type0FontObjectNumber, 0, CreateType0Font(fontProgram, descendantFontObjectNumber, toUnicodeObjectNumber));
        ApplySubsetFontName(objects, PdfSubsetTag.Apply(fontProgram.FontName, fontProgram.GetUsedGlyphIds(), fontData),
            descriptorObjectNumber, descendantFontObjectNumber, type0FontObjectNumber);
    }

    // An embedded subset is named with a subset tag, as PDF requires and archival validators check.
    private static void ApplySubsetFontName(Dictionary<int, PdfIndirectObject> objects, string subsetName,
        int descriptorObjectNumber, int descendantFontObjectNumber, int type0FontObjectNumber) {
        if (objects[descriptorObjectNumber].Value is PdfDictionary descriptor) descriptor.Items["FontName"] = new PdfName(subsetName);
        if (objects[descendantFontObjectNumber].Value is PdfDictionary descendant) descendant.Items["BaseFont"] = new PdfName(subsetName);
        if (objects[type0FontObjectNumber].Value is PdfDictionary type0) type0.Items["BaseFont"] = new PdfName(subsetName);
    }

    private static void MaterializeEmbeddedTextAppearanceFont(
        Dictionary<int, PdfIndirectObject> objects,
        PdfOpenTypeCffFontProgram fontProgram,
        int fontFileObjectNumber,
        int descriptorObjectNumber,
        int descendantFontObjectNumber,
        int toUnicodeObjectNumber,
        int type0FontObjectNumber) {
        byte[] fontData = fontProgram.BuildCompactOpenTypeFontFile();
        objects[fontFileObjectNumber] = new PdfIndirectObject(fontFileObjectNumber, 0, CreateOpenTypeCffFontFileStream(fontData));
        objects[descriptorObjectNumber] = new PdfIndirectObject(descriptorObjectNumber, 0, CreateOpenTypeCffFontDescriptor(fontProgram, fontFileObjectNumber));
        objects[descendantFontObjectNumber] = new PdfIndirectObject(descendantFontObjectNumber, 0, CreateCidFontType0Descendant(fontProgram, descriptorObjectNumber));
        objects[toUnicodeObjectNumber] = new PdfIndirectObject(toUnicodeObjectNumber, 0, CreateToUnicodeStream(fontProgram));
        objects[type0FontObjectNumber] = new PdfIndirectObject(type0FontObjectNumber, 0, CreateType0Font(fontProgram, descendantFontObjectNumber, toUnicodeObjectNumber));
        ApplySubsetFontName(objects, PdfSubsetTag.Apply(fontProgram.FontName, System.Array.Empty<int>(), fontData),
            descriptorObjectNumber, descendantFontObjectNumber, type0FontObjectNumber);
    }

    private static PdfStream CreateEmbeddedFontFileStream(byte[] fontData) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Length1"] = new PdfNumber(fontData.Length);
        return new PdfStream(dictionary, fontData);
    }

    private static PdfStream CreateOpenTypeCffFontFileStream(byte[] fontData) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Subtype"] = new PdfName("OpenType");
        dictionary.Items["Length1"] = new PdfNumber(fontData.Length);
        return new PdfStream(dictionary, fontData);
    }

    private static PdfDictionary CreateTrueTypeFontDescriptor(PdfTrueTypeFontProgram fontProgram, int fontFileObjectNumber) {
        var descriptor = new PdfDictionary();
        descriptor.Items["Type"] = new PdfName("FontDescriptor");
        descriptor.Items["FontName"] = new PdfName(fontProgram.FontName);
        descriptor.Items["Flags"] = new PdfNumber(fontProgram.Flags);
        descriptor.Items["FontBBox"] = CreateNumberArray(fontProgram.FontBBox.Select(value => (double)value));
        descriptor.Items["ItalicAngle"] = new PdfNumber(fontProgram.ItalicAngle);
        descriptor.Items["Ascent"] = new PdfNumber(fontProgram.Ascent);
        descriptor.Items["Descent"] = new PdfNumber(fontProgram.Descent);
        descriptor.Items["CapHeight"] = new PdfNumber(fontProgram.CapHeight);
        descriptor.Items["StemV"] = new PdfNumber(fontProgram.StemV);
        descriptor.Items["FontFile2"] = new PdfReference(fontFileObjectNumber, 0);
        return descriptor;
    }

    private static PdfDictionary CreateOpenTypeCffFontDescriptor(PdfOpenTypeCffFontProgram fontProgram, int fontFileObjectNumber) {
        var descriptor = new PdfDictionary();
        descriptor.Items["Type"] = new PdfName("FontDescriptor");
        descriptor.Items["FontName"] = new PdfName(fontProgram.FontName);
        descriptor.Items["Flags"] = new PdfNumber(fontProgram.Flags);
        descriptor.Items["FontBBox"] = CreateNumberArray(fontProgram.FontBBox.Select(value => (double)value));
        descriptor.Items["ItalicAngle"] = new PdfNumber(fontProgram.ItalicAngle);
        descriptor.Items["Ascent"] = new PdfNumber(fontProgram.Ascent);
        descriptor.Items["Descent"] = new PdfNumber(fontProgram.Descent);
        descriptor.Items["CapHeight"] = new PdfNumber(fontProgram.CapHeight);
        descriptor.Items["StemV"] = new PdfNumber(fontProgram.StemV);
        descriptor.Items["FontFile3"] = new PdfReference(fontFileObjectNumber, 0);
        return descriptor;
    }

    private static PdfDictionary CreateCidFontType2Descendant(PdfTrueTypeFontProgram fontProgram, int descriptorObjectNumber) {
        var cidSystemInfo = new PdfDictionary();
        cidSystemInfo.Items["Registry"] = new PdfStringObj("Adobe");
        cidSystemInfo.Items["Ordering"] = new PdfStringObj("Identity");
        cidSystemInfo.Items["Supplement"] = new PdfNumber(0);

        var descendant = new PdfDictionary();
        descendant.Items["Type"] = new PdfName("Font");
        descendant.Items["Subtype"] = new PdfName("CIDFontType2");
        descendant.Items["BaseFont"] = new PdfName(fontProgram.FontName);
        descendant.Items["CIDSystemInfo"] = cidSystemInfo;
        descendant.Items["FontDescriptor"] = new PdfReference(descriptorObjectNumber, 0);
        descendant.Items["DW"] = new PdfNumber(500);
        descendant.Items["W"] = CreateUsedGlyphWidthArray(fontProgram);
        descendant.Items["CIDToGIDMap"] = new PdfName("Identity");
        return descendant;
    }

    private static PdfDictionary CreateCidFontType0Descendant(PdfOpenTypeCffFontProgram fontProgram, int descriptorObjectNumber) {
        var cidSystemInfo = new PdfDictionary();
        cidSystemInfo.Items["Registry"] = new PdfStringObj("Adobe");
        cidSystemInfo.Items["Ordering"] = new PdfStringObj("Identity");
        cidSystemInfo.Items["Supplement"] = new PdfNumber(0);

        var descendant = new PdfDictionary();
        descendant.Items["Type"] = new PdfName("Font");
        descendant.Items["Subtype"] = new PdfName("CIDFontType0");
        descendant.Items["BaseFont"] = new PdfName(fontProgram.FontName);
        descendant.Items["CIDSystemInfo"] = cidSystemInfo;
        descendant.Items["FontDescriptor"] = new PdfReference(descriptorObjectNumber, 0);
        descendant.Items["DW"] = new PdfNumber(500);
        descendant.Items["W"] = CreateUsedGlyphWidthArray(fontProgram);
        return descendant;
    }

    private static PdfStream CreateToUnicodeStream(PdfTrueTypeFontProgram fontProgram) {
        var dictionary = new PdfDictionary();
        return new PdfStream(dictionary, PdfToUnicodeCMapBuilder.BuildToUnicodeCMap(fontProgram));
    }

    private static PdfStream CreateToUnicodeStream(PdfOpenTypeCffFontProgram fontProgram) {
        var dictionary = new PdfDictionary();
        return new PdfStream(dictionary, PdfToUnicodeCMapBuilder.BuildToUnicodeCMap(fontProgram));
    }

    private static PdfDictionary CreateType0Font(PdfTrueTypeFontProgram fontProgram, int descendantFontObjectNumber, int toUnicodeObjectNumber) {
        var descendants = new PdfArray();
        descendants.Items.Add(new PdfReference(descendantFontObjectNumber, 0));

        var font = new PdfDictionary();
        font.Items["Type"] = new PdfName("Font");
        font.Items["Subtype"] = new PdfName("Type0");
        font.Items["BaseFont"] = new PdfName(fontProgram.FontName);
        font.Items["Encoding"] = new PdfName("Identity-H");
        font.Items["DescendantFonts"] = descendants;
        font.Items["ToUnicode"] = new PdfReference(toUnicodeObjectNumber, 0);
        return font;
    }

    private static PdfDictionary CreateType0Font(PdfOpenTypeCffFontProgram fontProgram, int descendantFontObjectNumber, int toUnicodeObjectNumber) {
        var descendants = new PdfArray();
        descendants.Items.Add(new PdfReference(descendantFontObjectNumber, 0));

        var font = new PdfDictionary();
        font.Items["Type"] = new PdfName("Font");
        font.Items["Subtype"] = new PdfName("Type0");
        font.Items["BaseFont"] = new PdfName(fontProgram.FontName);
        font.Items["Encoding"] = new PdfName("Identity-H");
        font.Items["DescendantFonts"] = descendants;
        font.Items["ToUnicode"] = new PdfReference(toUnicodeObjectNumber, 0);
        return font;
    }

    private static PdfArray CreateUsedGlyphWidthArray(PdfTrueTypeFontProgram fontProgram) {
        IReadOnlyList<int> usedGlyphIds = fontProgram.GetUsedGlyphIds();
        var array = new PdfArray();
        if (usedGlyphIds.Count == 0) {
            return array;
        }

        int index = 0;
        while (index < usedGlyphIds.Count) {
            int rangeStart = usedGlyphIds[index];
            int rangeEnd = rangeStart;
            int rangeIndex = index + 1;
            while (rangeIndex < usedGlyphIds.Count && usedGlyphIds[rangeIndex] == rangeEnd + 1) {
                rangeEnd = usedGlyphIds[rangeIndex];
                rangeIndex++;
            }

            var widths = new PdfArray();
            for (int glyphId = rangeStart; glyphId <= rangeEnd; glyphId++) {
                widths.Items.Add(new PdfNumber(fontProgram.GetGlyphWidth1000(glyphId)));
            }

            array.Items.Add(new PdfNumber(rangeStart));
            array.Items.Add(widths);
            index = rangeIndex;
        }

        return array;
    }

    private static PdfArray CreateUsedGlyphWidthArray(PdfOpenTypeCffFontProgram fontProgram) {
        IReadOnlyList<int> usedGlyphIds = fontProgram.GetUsedGlyphIds();
        var array = new PdfArray();
        if (usedGlyphIds.Count == 0) {
            return array;
        }

        int index = 0;
        while (index < usedGlyphIds.Count) {
            int rangeStart = usedGlyphIds[index];
            int rangeEnd = rangeStart;
            int rangeIndex = index + 1;
            while (rangeIndex < usedGlyphIds.Count && usedGlyphIds[rangeIndex] == rangeEnd + 1) {
                rangeEnd = usedGlyphIds[rangeIndex];
                rangeIndex++;
            }

            var widths = new PdfArray();
            for (int glyphId = rangeStart; glyphId <= rangeEnd; glyphId++) {
                widths.Items.Add(new PdfNumber(fontProgram.GetGlyphWidth1000(glyphId)));
            }

            array.Items.Add(new PdfNumber(rangeStart));
            array.Items.Add(widths);
            index = rangeIndex;
        }

        return array;
    }

    private static PdfArray CreateNumberArray(IEnumerable<double> values) {
        var array = new PdfArray();
        foreach (double value in values) {
            array.Items.Add(new PdfNumber(value));
        }

        return array;
    }
}
