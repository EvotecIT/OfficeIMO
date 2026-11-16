namespace OfficeIMO.Pdf;

internal sealed partial class PdfAccessibilityInspector {
    // ISO 32000-1:2008, 9.7.5.2, Table 118 predefined CMaps and the Adobe character collection each one uses.
    private static readonly Dictionary<string, string> PredefinedCMapOrderings = BuildPredefinedCMaps();

    private static readonly HashSet<string> StandardFourteen = new(StringComparer.Ordinal) {
        "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic", "Helvetica", "Helvetica-Bold", "Helvetica-Oblique",
        "Helvetica-BoldOblique", "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique", "Symbol", "ZapfDingbats"
    };

    private static Dictionary<string, string> BuildPredefinedCMaps() {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        void Add(string ordering, params string[] names) { foreach (string name in names) map[name] = ordering; }
        Add("GB1", "GB-EUC-H", "GB-EUC-V", "GBpc-EUC-H", "GBpc-EUC-V", "GBK-EUC-H", "GBK-EUC-V", "GBKp-EUC-H", "GBKp-EUC-V",
            "GBK2K-H", "GBK2K-V", "UniGB-UCS2-H", "UniGB-UCS2-V", "UniGB-UTF16-H", "UniGB-UTF16-V");
        Add("CNS1", "B5pc-H", "B5pc-V", "HKscs-B5-H", "HKscs-B5-V", "ETen-B5-H", "ETen-B5-V", "ETenms-B5-H", "ETenms-B5-V",
            "CNS-EUC-H", "CNS-EUC-V", "UniCNS-UCS2-H", "UniCNS-UCS2-V", "UniCNS-UTF16-H", "UniCNS-UTF16-V");
        Add("Japan1", "83pv-RKSJ-H", "90ms-RKSJ-H", "90ms-RKSJ-V", "90msp-RKSJ-H", "90msp-RKSJ-V", "90pv-RKSJ-H", "Add-RKSJ-H",
            "Add-RKSJ-V", "EUC-H", "EUC-V", "Ext-RKSJ-H", "Ext-RKSJ-V", "H", "V", "UniJIS-UCS2-H", "UniJIS-UCS2-V",
            "UniJIS-UCS2-HW-H", "UniJIS-UCS2-HW-V", "UniJIS-UTF16-H", "UniJIS-UTF16-V");
        Add("Korea1", "KSC-EUC-H", "KSC-EUC-V", "KSCms-UHC-H", "KSCms-UHC-V", "KSCms-UHC-HW-H", "KSCms-UHC-HW-V", "KSCpc-EUC-H",
            "UniKS-UCS2-H", "UniKS-UCS2-V", "UniKS-UTF16-H", "UniKS-UTF16-V");
        Add("Identity", "Identity-H", "Identity-V");
        return map;
    }

    private void CheckFonts(Dictionary<PdfDictionary, (bool Rendered, int? ObjectNumber, int Page)> fonts) {
        foreach (KeyValuePair<PdfDictionary, (bool Rendered, int? ObjectNumber, int Page)> entry in fonts) {
            _cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary font = entry.Key;
            (bool rendered, int? objectNumber, int page) = entry.Value;
            string subtype = ReadName(font, "Subtype") ?? string.Empty;
            string fontName = ReadName(font, "BaseFont") ?? "(unnamed)";
            try {
                if (subtype == "Type0") CheckCompositeFont(font, fontName, rendered, objectNumber, page);
                else CheckSimpleFont(font, subtype, fontName, rendered, objectNumber, page);
            } catch (Exception exception) when (IsReadFailure(exception)) {
                Unknown("31-009", null, "The font " + fontName + " could not be inspected: " + exception.Message, page, objectNumber);
            }
        }
    }

    private void CheckSimpleFont(PdfDictionary font, string subtype, string fontName, bool rendered, int? objectNumber, int page) {
        PdfDictionary? descriptor = ResolveDict(font.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject) ? descriptorObject : null);
        if (subtype != "Type3" && rendered) {
            Applies("31-009");
            if (!HasFontFile(descriptor)) {
                Fail("31-009", "7.21.4.1-1", "The font " + fontName + " is used to render text but is not embedded.", page, objectNumber);
            }
        }
        PdfObject? encoding = font.Items.TryGetValue("Encoding", out PdfObject? encodingObject) ? Resolve(encodingObject) : null;
        bool symbolic = ((ReadInteger(descriptor ?? new PdfDictionary(), "Flags") ?? 0) & 4) == 4;
        if (subtype == "TrueType") {
            if (symbolic) {
                Applies("31-024");
                if (encoding is not null) Fail("31-024", "7.21.6-3", "The symbolic TrueType font " + fontName + " has an Encoding entry.", page, objectNumber);
            } else {
                Applies("31-019");
                if (encoding is null) {
                    Fail("31-019", "7.21.6-2", "The non-symbolic TrueType font " + fontName + " has no Encoding; use WinAnsiEncoding or MacRomanEncoding.", page, objectNumber);
                } else if (encoding is PdfDictionary encodingDictionary && ReadName(encodingDictionary, "BaseEncoding") is null) {
                    Applies("31-020");
                    Fail("31-020", "7.21.6-2", "The non-symbolic TrueType font " + fontName + " has an Encoding dictionary without BaseEncoding.", page, objectNumber);
                } else {
                    Applies("31-021");
                    string? name = encoding is PdfName encodingName ? encodingName.Name
                        : encoding is PdfDictionary dictionary ? ReadName(dictionary, "BaseEncoding") : null;
                    if (name is not ("WinAnsiEncoding" or "MacRomanEncoding")) {
                        Fail("31-021", "7.21.6-2", "The non-symbolic TrueType font " + fontName + " uses " + (name ?? "an unreadable encoding") +
                            "; only WinAnsiEncoding and MacRomanEncoding are allowed.", page, objectNumber);
                    }
                }
            }
        }

        if (font.Items.ContainsKey("ToUnicode")) { Applies("31-027"); return; }
        Applies("31-027");
        string? baseEncoding = encoding is PdfName simpleEncoding ? simpleEncoding.Name
            : encoding is PdfDictionary differencesEncoding ? ReadName(differencesEncoding, "BaseEncoding") : null;
        bool hasDifferences = encoding is PdfDictionary withDifferences && withDifferences.Items.ContainsKey("Differences");
        if (subtype == "TrueType" && !symbolic) return;
        if (baseEncoding is "WinAnsiEncoding" or "MacRomanEncoding" or "MacExpertEncoding" && !hasDifferences) return;
        if (subtype is "Type1" or "MMType1" && !hasDifferences && StandardFourteen.Contains(StripSubsetPrefix(fontName)) && !HasFontFile(descriptor)) return;
        if (subtype is "Type1" or "MMType1" or "Type3") {
            Unknown("31-027", "7.21.7-1", "The font " + fontName + " has no ToUnicode map; its glyph names must all be in the Adobe Glyph List, which this check cannot confirm yet.", page, objectNumber);
            return;
        }
        Fail("31-027", "7.21.7-1", "The font " + fontName + " has no ToUnicode map and its character codes cannot be mapped to Unicode otherwise.", page, objectNumber);
    }

    private void CheckCompositeFont(PdfDictionary font, string fontName, bool rendered, int? objectNumber, int page) {
        PdfArray? descendants = ResolveArray(font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) ? descendantsObject : null);
        PdfObject? descendantObject = descendants is not null && descendants.Items.Count > 0 ? descendants.Items[0] : null;
        PdfDictionary? cidFont = ResolveDict(descendantObject);
        PdfDictionary? descriptor = cidFont is null ? null : ResolveDict(cidFont.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject) ? descriptorObject : null);
        int? cidObject = ObjectNumberOf(descendantObject) ?? objectNumber;
        if (rendered) {
            Applies("31-009");
            if (!HasFontFile(descriptor)) {
                Fail("31-009", "7.21.4.1-1", "The font " + fontName + " is used to render text but is not embedded.", page, cidObject);
            }
        }
        if (cidFont is not null && ReadName(cidFont, "Subtype") == "CIDFontType2" && HasFontFile(descriptor)) {
            Applies("31-005");
            if (!cidFont.Items.TryGetValue("CIDToGIDMap", out PdfObject? map)) {
                Fail("31-005", "7.21.3.2-1", "The embedded Type 2 CIDFont " + fontName + " has no CIDToGIDMap.", page, cidObject);
            } else {
                Applies("31-004");
                PdfObject? mapValue = Resolve(map);
                if (mapValue is not PdfStream && !(mapValue is PdfName { Name: "Identity" })) {
                    Fail("31-004", "7.21.3.2-1", "The CIDToGIDMap of " + fontName + " is neither a stream nor Identity.", page, cidObject);
                }
            }
        }

        PdfObject? encodingObject = font.Items.TryGetValue("Encoding", out PdfObject? encodingEntry) ? encodingEntry : null;
        PdfObject? encoding = Resolve(encodingObject);
        string? cmapName = null;
        PdfDictionary? cmapSystemInfo = null;
        Applies("31-006");
        if (encoding is PdfName encodingName) {
            cmapName = encodingName.Name;
            if (!PredefinedCMapOrderings.ContainsKey(cmapName)) {
                Fail("31-006", "7.21.3.3-1", "The CMap " + cmapName + " is not predefined and is not embedded.", page, objectNumber);
            }
        } else if (encoding is PdfStream cmapStream) {
            cmapName = ReadName(cmapStream.Dictionary, "CMapName");
            cmapSystemInfo = ResolveDict(cmapStream.Dictionary.Items.TryGetValue("CIDSystemInfo", out PdfObject? infoObject) ? infoObject : null);
            CheckEmbeddedCMap(cmapStream, fontName, page, ObjectNumberOf(encodingObject));
        } else {
            Fail("31-006", "7.21.3.3-1", "The Type 0 font " + fontName + " has no usable Encoding CMap.", page, objectNumber);
        }

        PdfDictionary? fontSystemInfo = cidFont is null ? null : ResolveDict(cidFont.Items.TryGetValue("CIDSystemInfo", out PdfObject? fontInfoObject) ? fontInfoObject : null);
        if (cmapName is not ("Identity-H" or "Identity-V")) {
            Applies("31-001");
            Applies("31-002");
            Applies("31-003");
            string? fontRegistry = fontSystemInfo is null ? null : ReadText(fontSystemInfo, "Registry");
            string? fontOrdering = fontSystemInfo is null ? null : ReadText(fontSystemInfo, "Ordering");
            string? cmapRegistry, cmapOrdering;
            int? cmapSupplement;
            if (cmapSystemInfo is not null) {
                cmapRegistry = ReadText(cmapSystemInfo, "Registry");
                cmapOrdering = ReadText(cmapSystemInfo, "Ordering");
                cmapSupplement = ReadInteger(cmapSystemInfo, "Supplement");
            } else if (cmapName is not null && PredefinedCMapOrderings.TryGetValue(cmapName, out string? predefinedOrdering)) {
                cmapRegistry = "Adobe";
                cmapOrdering = predefinedOrdering;
                cmapSupplement = null;
            } else {
                cmapRegistry = cmapOrdering = null;
                cmapSupplement = null;
            }
            if (fontRegistry is null || !string.Equals(fontRegistry, cmapRegistry, StringComparison.Ordinal)) {
                Fail("31-001", "7.21.3.1-1", "The CIDSystemInfo Registry of " + fontName + " (" + (fontRegistry ?? "none") + ") differs from its CMap's (" + (cmapRegistry ?? "none") + ").", page, cidObject);
            }
            if (fontOrdering is null || !string.Equals(fontOrdering, cmapOrdering, StringComparison.Ordinal)) {
                Fail("31-002", "7.21.3.1-1", "The CIDSystemInfo Ordering of " + fontName + " (" + (fontOrdering ?? "none") + ") differs from its CMap's (" + (cmapOrdering ?? "none") + ").", page, cidObject);
            }
            int? fontSupplement = fontSystemInfo is null ? null : ReadInteger(fontSystemInfo, "Supplement");
            if (cmapSystemInfo is not null && (fontSupplement is null || cmapSupplement is null || fontSupplement > cmapSupplement)) {
                Fail("31-003", "7.21.3.1-1", "The CIDSystemInfo Supplement of " + fontName + " is missing or greater than its CMap's.", page, cidObject);
            } else if (cmapSystemInfo is null && cmapOrdering is not null && fontSupplement is > 0) {
                // A predefined CMap's Supplement lives in its resource file, which this check does not carry.
                Unknown("31-003", "7.21.3.1-1", "The CIDSystemInfo Supplement of " + fontName + " is " +
                    fontSupplement.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + "; the Supplement of the predefined CMap " +
                    cmapName + " is not known to this check.", page, cidObject);
            }
        }

        Applies("31-027");
        if (font.Items.ContainsKey("ToUnicode")) return;
        string? registry = fontSystemInfo is null ? null : ReadText(fontSystemInfo, "Registry");
        string? ordering = fontSystemInfo is null ? null : ReadText(fontSystemInfo, "Ordering");
        if (registry == "Adobe" && ordering is "GB1" or "CNS1" or "Japan1" or "Korea1") return;
        Fail("31-027", "7.21.7-1", "The composite font " + fontName + " has no ToUnicode map and does not use an Adobe GB1, CNS1, Japan1 or Korea1 collection.", page, objectNumber);
    }

    private void CheckEmbeddedCMap(PdfStream cmap, string fontName, int page, int? cmapObject) {
        if (cmap.Dictionary.Items.TryGetValue("UseCMap", out PdfObject? useCMap)) {
            Applies("31-008");
            PdfObject? used = Resolve(useCMap);
            if (used is not PdfName usedName || !PredefinedCMapOrderings.ContainsKey(usedName.Name)) {
                Fail("31-008", "7.21.3.3-3", "The CMap of " + fontName + " uses another CMap that is not predefined.", page, cmapObject);
            }
        }
        Applies("31-007");
        int dictionaryMode = ReadInteger(cmap.Dictionary, "WMode") ?? 0;
        string text;
        try {
            text = PdfEncoding.Latin1GetString(_document.DecodedStreamBudget.DecodeRequired(cmap, _objects, _document.ReadOptions.Limits.MaxDecodedStreamBytes, _cancellationToken));
        } catch (Exception exception) when (IsReadFailure(exception)) {
            Unknown("31-007", "7.21.3.3-2", "The embedded CMap of " + fontName + " could not be decoded: " + exception.Message, page, cmapObject);
            return;
        }
        string streamMode = ReadCMapProgramWMode(text) ?? "0";
        if (!int.TryParse(streamMode, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int mode) || mode != dictionaryMode) {
            Fail("31-007", "7.21.3.3-2", "The CMap of " + fontName + " declares WMode " + streamMode +
                " in its program but " + dictionaryMode.ToString(System.Globalization.CultureInfo.InvariantCulture) + " in its dictionary.", page, cmapObject);
        }
    }

    // Returns the token after the first /WMode key of a CMap program, skipping comments and strings.
    private static string? ReadCMapProgramWMode(string text) {
        int i = 0;
        while (i < text.Length) {
            char c = text[i];
            if (c == '%') {
                while (i < text.Length && text[i] != '\n' && text[i] != '\r') i++;
            } else if (c == '(') {
                int depth = 0;
                for (; i < text.Length; i++) {
                    if (text[i] == '\\') { i++; continue; }
                    if (text[i] == '(') depth++;
                    else if (text[i] == ')' && --depth == 0) { i++; break; }
                }
            } else if (c == '<' && i + 1 < text.Length && text[i + 1] != '<') {
                int end = text.IndexOf('>', i + 1);
                i = end < 0 ? text.Length : end + 1;
            } else if (c == '/' && string.CompareOrdinal(text, i, "/WMode", 0, 6) == 0 &&
                (i + 6 == text.Length || IsCMapDelimiter(text[i + 6]))) {
                int start = i + 6;
                while (start < text.Length) {
                    if (IsCMapWhiteSpace(text[start])) start++;
                    else if (text[start] == '%') { while (start < text.Length && text[start] != '\n' && text[start] != '\r') start++; }
                    else break;
                }
                int end = start;
                while (end < text.Length && !IsCMapDelimiter(text[end])) end++;
                return end > start ? text.Substring(start, end - start) : null;
            } else {
                i++;
            }
        }
        return null;
    }

    private static bool IsCMapWhiteSpace(char c) => c is ' ' or '\t' or '\r' or '\n' or '\f' or '\0';

    private static bool IsCMapDelimiter(char c) => IsCMapWhiteSpace(c) || c is '/' or '(' or ')' or '<' or '>' or '[' or ']' or '{' or '}' or '%';

    private bool HasFontFile(PdfDictionary? descriptor) => descriptor is not null &&
        (Resolve(descriptor.Items.TryGetValue("FontFile", out PdfObject? f1) ? f1 : null) is PdfStream ||
         Resolve(descriptor.Items.TryGetValue("FontFile2", out PdfObject? f2) ? f2 : null) is PdfStream ||
         Resolve(descriptor.Items.TryGetValue("FontFile3", out PdfObject? f3) ? f3 : null) is PdfStream);

    private static string StripSubsetPrefix(string name) =>
        name.Length > 7 && name[6] == '+' && name.Take(6).All(static c => c >= 'A' && c <= 'Z') ? name.Substring(7) : name;
}
