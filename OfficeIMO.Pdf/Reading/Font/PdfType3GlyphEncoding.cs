namespace OfficeIMO.Pdf;

internal static class PdfType3GlyphEncoding {
    internal static bool TryCreate(string encodingName, out Dictionary<int, string> glyphNames) {
        glyphNames = new Dictionary<int, string>();
        AddAscii(glyphNames);
        if (string.Equals(encodingName, "StandardEncoding", StringComparison.Ordinal)) {
            AddStandard(glyphNames);
            return true;
        }
        if (string.Equals(encodingName, "WinAnsiEncoding", StringComparison.Ordinal)) {
            AddWinAnsi(glyphNames);
            return true;
        }
        glyphNames.Clear();
        return false;
    }

    private static void AddAscii(Dictionary<int, string> map) {
        string[] punctuation = {
            "space", "exclam", "quotedbl", "numbersign", "dollar", "percent", "ampersand", "quotesingle",
            "parenleft", "parenright", "asterisk", "plus", "comma", "hyphen", "period", "slash"
        };
        for (int index = 0; index < punctuation.Length; index++) map[32 + index] = punctuation[index];
        for (int code = 48; code <= 57; code++) map[code] = ((char)code).ToString();
        map[58] = "colon"; map[59] = "semicolon"; map[60] = "less"; map[61] = "equal";
        map[62] = "greater"; map[63] = "question"; map[64] = "at";
        for (int code = 65; code <= 90; code++) map[code] = ((char)code).ToString();
        map[91] = "bracketleft"; map[92] = "backslash"; map[93] = "bracketright";
        map[94] = "asciicircum"; map[95] = "underscore"; map[96] = "grave";
        for (int code = 97; code <= 122; code++) map[code] = ((char)code).ToString();
        map[123] = "braceleft"; map[124] = "bar"; map[125] = "braceright"; map[126] = "asciitilde";
    }

    private static void AddStandard(Dictionary<int, string> map) {
        map[39] = "quoteright";
        map[96] = "quoteleft";
        Add(map, 161,
            "exclamdown", "cent", "sterling", "fraction", "yen", "florin", "section", "currency",
            "quotesingle", "quotedblleft", "guillemotleft", "guilsinglleft", "guilsinglright", "fi", "fl");
        Add(map, 177,
            "endash", "dagger", "daggerdbl", "periodcentered", null, "paragraph", "bullet", "quotesinglbase",
            "quotedblbase", "quotedblright", "guillemotright", "ellipsis", "perthousand", null, "questiondown");
        Add(map, 193,
            "grave", "acute", "circumflex", "tilde", "macron", "breve", "dotaccent", "dieresis",
            null, "ring", "cedilla", null, "hungarumlaut", "ogonek", "caron", "emdash");
        map[225] = "AE"; map[227] = "ordfeminine"; map[232] = "Lslash"; map[233] = "Oslash";
        map[234] = "OE"; map[235] = "ordmasculine"; map[241] = "ae"; map[245] = "dotlessi";
        map[248] = "lslash"; map[249] = "oslash"; map[250] = "oe"; map[251] = "germandbls";
    }

    private static void AddWinAnsi(Dictionary<int, string> map) {
        Add(map, 128,
            "Euro", null, "quotesinglbase", "florin", "quotedblbase", "ellipsis", "dagger", "daggerdbl",
            "circumflex", "perthousand", "Scaron", "guilsinglleft", "OE", null, "Zcaron", null,
            null, "quoteleft", "quoteright", "quotedblleft", "quotedblright", "bullet", "endash", "emdash",
            "tilde", "trademark", "scaron", "guilsinglright", "oe", null, "zcaron", "Ydieresis");
        Add(map, 160,
            "space", "exclamdown", "cent", "sterling", "currency", "yen", "brokenbar", "section",
            "dieresis", "copyright", "ordfeminine", "guillemotleft", "logicalnot", "hyphen", "registered", "macron",
            "degree", "plusminus", "twosuperior", "threesuperior", "acute", "mu", "paragraph", "periodcentered",
            "cedilla", "onesuperior", "ordmasculine", "guillemotright", "onequarter", "onehalf", "threequarters", "questiondown",
            "Agrave", "Aacute", "Acircumflex", "Atilde", "Adieresis", "Aring", "AE", "Ccedilla",
            "Egrave", "Eacute", "Ecircumflex", "Edieresis", "Igrave", "Iacute", "Icircumflex", "Idieresis",
            "Eth", "Ntilde", "Ograve", "Oacute", "Ocircumflex", "Otilde", "Odieresis", "multiply",
            "Oslash", "Ugrave", "Uacute", "Ucircumflex", "Udieresis", "Yacute", "Thorn", "germandbls",
            "agrave", "aacute", "acircumflex", "atilde", "adieresis", "aring", "ae", "ccedilla",
            "egrave", "eacute", "ecircumflex", "edieresis", "igrave", "iacute", "icircumflex", "idieresis",
            "eth", "ntilde", "ograve", "oacute", "ocircumflex", "otilde", "odieresis", "divide",
            "oslash", "ugrave", "uacute", "ucircumflex", "udieresis", "yacute", "thorn", "ydieresis");
    }

    private static void Add(Dictionary<int, string> map, int firstCode, params string?[] names) {
        for (int index = 0; index < names.Length; index++) {
            if (names[index] is string name) map[firstCode + index] = name;
        }
    }
}
