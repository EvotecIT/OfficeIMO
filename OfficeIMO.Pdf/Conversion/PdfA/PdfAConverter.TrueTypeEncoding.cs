using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private static readonly string[] DeclarableEncodings = { "WinAnsiEncoding", "MacRomanEncoding" };

    /// <summary>
    /// PDF/A requires a non-symbolic embedded TrueType font to declare WinAnsi or MacRoman encoding. A font that relies on the
    /// implied standard encoding gets an explicit encoding selecting the same glyph names, so every code draws the same glyph.
    /// </summary>
    private void FixTrueTypeEncoding(FontUse use, PdfDictionary descriptor, string location) {
        PdfDictionary font = use.Font;
        if (GetName(font, "Subtype") != "TrueType" || Resolve(descriptor.Items.TryGetValue("FontFile2", out PdfObject? file) ? file : null) is not PdfStream program) return;
        int flags = Resolve(descriptor.Items.TryGetValue("Flags", out PdfObject? flagValue) ? flagValue : null) is PdfNumber number ? (int)number.Value : 0;
        if ((flags & 4) != 0) return;
        PdfObject? encoding = Resolve(font.Items.TryGetValue("Encoding", out PdfObject? encodingValue) ? encodingValue : null);
        if (encoding is PdfName { Name: "WinAnsiEncoding" or "MacRomanEncoding" }) return;
        if (encoding is PdfDictionary existing && GetName(existing, "BaseEncoding") is string declared && declared is "WinAnsiEncoding" or "MacRomanEncoding") {
            if (!existing.Items.ContainsKey("Differences") || HasWindowsUnicodeCmap(program)) return;
            // Differences need a (3,1) cmap. Entries for codes the pages never draw can go without changing any glyph.
            string?[] declaredNames = BuildSimpleEncoding(font);
            string?[] baseNames = PdfFontEncodingNames.Get(declared)!;
            if (use.Codes.Any(code => !string.Equals(declaredNames[code], baseNames[code], StringComparison.Ordinal))) {
                AddViolation(PdfAConversionRefusalReason.FontEncodingNotCompliant, location,
                    "The embedded TrueType font changes drawn codes with Differences but has no Windows Unicode cmap, which PDF/A requires.", font);
                return;
            }
            if (Change(PdfAConversionChangeKind.FontDescriptorUpdated, location,
                    "Removed encoding Differences for codes the pages never draw, which PDF/A forbids without a Windows Unicode cmap.", font)) {
                font.Items["Encoding"] = new PdfName(declared);
            }
            return;
        }

        string?[] names = BuildSimpleEncoding(font);
        foreach (string baseEncoding in DeclarableEncodings) {
            string?[] baseNames = PdfFontEncodingNames.Get(baseEncoding)!;
            if (use.Codes.All(code => names[code] is null || string.Equals(names[code], baseNames[code], StringComparison.Ordinal))) {
                if (Change(PdfAConversionChangeKind.FontDescriptorUpdated, location,
                        "Declared /" + baseEncoding + ", which selects the same glyphs for every drawn code, as PDF/A requires for TrueType fonts.", font)) {
                    font.Items["Encoding"] = new PdfName(baseEncoding);
                }
                return;
            }
        }

        bool namesAreStandard = use.Codes.All(code => names[code] is null || ResourceResolver.TryDecodeStandardGlyphName(names[code]!, out _));
        if (namesAreStandard && HasWindowsUnicodeCmap(program)) {
            if (Change(PdfAConversionChangeKind.FontDescriptorUpdated, location,
                    "Declared WinAnsi encoding with differences that select the same glyph names for every drawn code, as PDF/A requires for TrueType fonts.", font)) {
                string?[] winAnsi = PdfFontEncodingNames.Get("WinAnsiEncoding")!;
                var differences = new PdfArray();
                int previous = -2;
                foreach (int code in use.Codes) {
                    if (names[code] is not string name || string.Equals(winAnsi[code], name, StringComparison.Ordinal)) continue;
                    if (code != previous + 1) differences.Items.Add(new PdfNumber(code));
                    differences.Items.Add(new PdfName(name));
                    previous = code;
                }
                var dictionary = new PdfDictionary();
                dictionary.Items["Type"] = new PdfName("Encoding");
                dictionary.Items["BaseEncoding"] = new PdfName("WinAnsiEncoding");
                dictionary.Items["Differences"] = differences;
                font.Items["Encoding"] = dictionary;
            }
            return;
        }
        AddViolation(PdfAConversionRefusalReason.FontEncodingNotCompliant, location,
            "The embedded TrueType font has no WinAnsi or MacRoman encoding, and its glyphs cannot be selected the same way with one.", font);
    }

    private bool HasWindowsUnicodeCmap(PdfStream program) =>
        StreamDecoder.TryDecode(program, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] data, _objects) && HasWindowsUnicodeCmap(data);

    /// <summary>True when an sfnt program has a (3,1) Windows Unicode cmap, which PDF/A requires for encodings with Differences.</summary>
    internal static bool HasWindowsUnicodeCmap(byte[] data) {
        if (data.Length < 12) return false;
        int tables = (data[4] << 8) | data[5];
        for (int table = 0; table < tables; table++) {
            int record = 12 + table * 16;
            if (record + 16 > data.Length) return false;
            if (data[record] != 'c' || data[record + 1] != 'm' || data[record + 2] != 'a' || data[record + 3] != 'p') continue;
            int offset = (data[record + 8] << 24) | (data[record + 9] << 16) | (data[record + 10] << 8) | data[record + 11];
            if (offset < 0 || offset + 4 > data.Length) return false;
            int count = (data[offset + 2] << 8) | data[offset + 3];
            for (int entry = 0; entry < count; entry++) {
                int at = offset + 4 + entry * 8;
                if (at + 4 > data.Length) return false;
                int platform = (data[at] << 8) | data[at + 1];
                int encodingId = (data[at + 2] << 8) | data[at + 3];
                if (platform == 3 && encodingId == 1) return true;
            }
        }
        return false;
    }
}
