using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    /// <summary>
    /// PDF/A requires the Widths of an embedded font to agree with its program and forbids drawing .notdef.
    /// Correcting either would move or change glyphs, so a disagreement refuses the source.
    /// </summary>
    private void CheckEmbeddedProgram(FontUse use, string location, bool composite) {
        if (use.Codes.Count == 0) return;
        PdfDictionary? descendant = composite && Resolve(use.Font.Items.TryGetValue("DescendantFonts", out PdfObject? descendants) ? descendants : null) is PdfArray array &&
            array.Items.Count > 0 ? ResolveDictionary(array.Items[0]) : null;
        if (composite && (descendant is null || GetName(use.Font, "Encoding") is not ("Identity-H" or "Identity-V"))) return;
        PdfDictionary? descriptor = ResolveDictionary((composite ? descendant! : use.Font).Items.TryGetValue("FontDescriptor", out PdfObject? fd) ? fd : null);
        if (descriptor is null) return;

        Func<int, int>? glyphForCode;
        Func<int, double?> advance;
        int glyphCount;
        if (Resolve(descriptor.Items.TryGetValue("FontFile2", out PdfObject? trueType) ? trueType : null) is PdfStream) {
            PdfDrawingFontProgram? drawing;
            try {
                drawing = ResourceResolver.CreateFontResource(use.ResourceName, use.Font, _objects).ForDrawing(_cancellationToken).DrawingProgram;
            } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException or IndexOutOfRangeException or PdfReadLimitException) {
                AddViolation(PdfAConversionRefusalReason.FontNotEmbeddable, location, "The embedded font program cannot be read: " + exception.Message, use.Font);
                return;
            }
            if (drawing is null || !SfntMetrics.TryRead(drawing.Program, out SfntMetrics metrics)) return;
            glyphForCode = drawing.GlyphForCode;
            advance = glyph => metrics.GetAdvance1000(glyph);
            glyphCount = metrics.GlyphCount;
        } else if (TryReadCompactProgram(descriptor, out PdfCffFontProgram? program, out string?[]? type1Encoding)) {
            double scale = program!.FontMatrix[0] * 1000D;
            glyphCount = program.GlyphCount;
            advance = glyph => program.GetAdvance(glyph) is double units ? units * scale : null;
            if (composite) {
                glyphForCode = code => program.TryGetGlyphForCid(code, out int glyph) ? glyph : 0;
            } else {
                string?[] builtin = type1Encoding ?? BuiltinGlyphNames(program);
                string?[] names = BuildEmbeddedSimpleEncoding(use.Font, builtin);
                glyphForCode = code => code < names.Length && names[code] is string name ? program.TryGetGlyph(name, out int glyph) ? glyph : 0 : 0;
            }
        } else {
            return;
        }

        Dictionary<int, double>? cidWidths = composite ? ReadCidWidths(descendant!) : null;
        double defaultWidth = composite && Resolve(descendant!.Items.TryGetValue("DW", out PdfObject? dw) ? dw : null) is PdfNumber dwNumber ? dwNumber.Value : 1000D;
        bool symbolic = Resolve(descriptor.Items.TryGetValue("Flags", out PdfObject? flags) ? flags : null) is PdfNumber flagNumber &&
            ((int)flagNumber.Value & 4) != 0;

        var notdef = new List<int>();
        var mismatched = new List<string>();
        int mismatchCount = 0;
        foreach (int code in use.Codes) {
            if (code < 0) continue;
            int glyph = glyphForCode(code);
            if (glyph < 0) continue;
            if (glyph == 0 || glyph >= glyphCount) {
                if (!symbolic || composite) notdef.Add(code);
                continue;
            }
            if (!use.VisibleCodes.Contains(code) || advance(glyph) is not double actual) continue;
            double expected = composite
                ? cidWidths!.TryGetValue(code, out double cidWidth) ? cidWidth : defaultWidth
                : ExpectedWidth(use.Font, descriptor, code, 0, null) ?? 0D;
            if (Math.Abs(expected - actual) > WidthTolerance) {
                mismatchCount++;
                if (mismatched.Count < 6) {
                    mismatched.Add(FormatCode(code) + ": " + expected.ToString("0.#", CultureInfo.InvariantCulture) + " vs " + actual.ToString("0.#", CultureInfo.InvariantCulture));
                }
            }
        }
        if (notdef.Count > 0) {
            AddViolation(PdfAConversionRefusalReason.FontGlyphMissing, location,
                "Code(s) " + FormatCodes(notdef) + " draw the .notdef glyph of the embedded font, which PDF/A forbids.", use.Font);
        }
        if (mismatchCount > 0) {
            AddViolation(PdfAConversionRefusalReason.FontMetricsMismatch, location,
                mismatchCount.ToString(CultureInfo.InvariantCulture) + " drawn glyph(s) have widths in the font dictionary that differ from the embedded program (" +
                string.Join(", ", mismatched) + "). Correcting them would move text.", use.Font);
        }
    }

    /// <summary>Reads a FontFile3 CFF or OpenType program, or a FontFile Type 1 program converted to CFF.</summary>
    private bool TryReadCompactProgram(PdfDictionary descriptor, out PdfCffFontProgram? program, out string?[]? type1Encoding) {
        program = null;
        type1Encoding = null;
        byte[]? cff = null;
        int limit = _document.ReadOptions.Limits.MaxDecodedStreamBytes;
        if (Resolve(descriptor.Items.TryGetValue("FontFile3", out PdfObject? compact) ? compact : null) is PdfStream compactStream) {
            if (!Filters.StreamDecoder.TryDecode(compactStream, limit, out byte[] data, _objects)) return false;
            cff = GetName(compactStream.Dictionary, "Subtype") == "OpenType" ? PdfCffFontProgram.TryExtractOpenTypeCff(data) : data;
        } else if (Resolve(descriptor.Items.TryGetValue("FontFile", out PdfObject? type1) ? type1 : null) is PdfStream type1Stream) {
            if (!Filters.StreamDecoder.TryDecode(type1Stream, limit, out byte[] data, _objects)) return false;
            PdfType1FontConverter.Result? converted = PdfType1FontConverter.TryConvert(data);
            cff = converted?.Cff;
            type1Encoding = converted?.BuiltinEncoding;
        }
        program = cff is null ? null : PdfCffFontProgram.TryParse(cff);
        return program is not null && program.HasUniformFontMatrix && program.FontMatrix.Length >= 1 && program.FontMatrix[0] > 0D;
    }

    private static string?[] BuiltinGlyphNames(PdfCffFontProgram program) {
        var names = new string?[256];
        if (program.GlyphNames is not string[] glyphNames) return names;
        for (int code = 0; code < 256 && code < program.BuiltinEncoding.Length; code++) {
            int glyph = program.BuiltinEncoding[code];
            if (glyph > 0 && glyph < glyphNames.Length) names[code] = glyphNames[glyph];
        }
        return names;
    }

    /// <summary>An embedded simple font's Differences apply to its BaseEncoding or, without one, to the program's built-in encoding.</summary>
    private string?[] BuildEmbeddedSimpleEncoding(PdfDictionary font, string?[] builtin) {
        PdfObject? encoding = Resolve(font.Items.TryGetValue("Encoding", out PdfObject? encodingValue) ? encodingValue : null);
        string?[] names = (string?[])builtin.Clone();
        if (names.Length < 256) Array.Resize(ref names, 256);
        if (encoding is PdfName encodingName && PdfFontEncodingNames.Get(encodingName.Name) is string?[] named) return (string?[])named.Clone();
        if (encoding is not PdfDictionary dictionary) return names;
        if (GetName(dictionary, "BaseEncoding") is string baseName && PdfFontEncodingNames.Get(baseName) is string?[] baseNames) names = (string?[])baseNames.Clone();
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

    private Dictionary<int, double> ReadCidWidths(PdfDictionary descendant) {
        var widths = new Dictionary<int, double>();
        if (Resolve(descendant.Items.TryGetValue("W", out PdfObject? w) ? w : null) is not PdfArray array) return widths;
        int index = 0;
        while (index < array.Items.Count) {
            if (Resolve(array.Items[index]) is not PdfNumber first) break;
            PdfObject? next = index + 1 < array.Items.Count ? Resolve(array.Items[index + 1]) : null;
            if (next is PdfArray list) {
                for (int item = 0; item < list.Items.Count && item < 65536; item++) {
                    if (Resolve(list.Items[item]) is PdfNumber width) widths[(int)first.Value + item] = width.Value;
                }
                index += 2;
            } else if (next is PdfNumber last && index + 2 < array.Items.Count && Resolve(array.Items[index + 2]) is PdfNumber rangeWidth) {
                for (int cid = (int)first.Value; cid <= (int)last.Value && cid - (int)first.Value < 65536; cid++) widths[cid] = rangeWidth.Value;
                index += 3;
            } else {
                break;
            }
        }
        return widths;
    }

    /// <summary>Reads advance widths from an sfnt program.</summary>
    private readonly struct SfntMetrics {
        private readonly byte[] _data;
        private readonly int _hmtx;
        private readonly int _metricCount;
        private readonly int _unitsPerEm;

        private SfntMetrics(byte[] data, int hmtx, int metricCount, int unitsPerEm, int glyphCount) {
            _data = data;
            _hmtx = hmtx;
            _metricCount = metricCount;
            _unitsPerEm = unitsPerEm;
            GlyphCount = glyphCount;
        }

        internal int GlyphCount { get; }

        internal double GetAdvance1000(int glyph) {
            int metric = Math.Min(glyph, _metricCount - 1);
            int offset = _hmtx + metric * 4;
            return ((_data[offset] << 8) | _data[offset + 1]) * 1000D / _unitsPerEm;
        }

        internal static bool TryRead(byte[] data, out SfntMetrics metrics) {
            metrics = default;
            if (data.Length < 12) return false;
            int tableCount = (data[4] << 8) | data[5];
            int head = -1, hhea = -1, hmtx = -1, maxp = -1, hmtxLength = 0;
            for (int table = 0; table < tableCount; table++) {
                int record = 12 + table * 16;
                if (record + 16 > data.Length) return false;
                string tag = Encoding.ASCII.GetString(data, record, 4);
                int offset = (data[record + 8] << 24) | (data[record + 9] << 16) | (data[record + 10] << 8) | data[record + 11];
                int length = (data[record + 12] << 24) | (data[record + 13] << 16) | (data[record + 14] << 8) | data[record + 15];
                if (offset < 0 || length < 0 || offset > data.Length - length) return false;
                switch (tag) {
                    case "head": head = offset; break;
                    case "hhea": hhea = offset; break;
                    case "hmtx": hmtx = offset; hmtxLength = length; break;
                    case "maxp": maxp = offset; break;
                }
            }
            if (head < 0 || hhea < 0 || hmtx < 0 || maxp < 0 || head + 20 > data.Length || hhea + 36 > data.Length || maxp + 6 > data.Length) return false;
            int unitsPerEm = (data[head + 18] << 8) | data[head + 19];
            int metricCount = (data[hhea + 34] << 8) | data[hhea + 35];
            int glyphCount = (data[maxp + 4] << 8) | data[maxp + 5];
            if (unitsPerEm == 0 || metricCount == 0 || metricCount * 4 > hmtxLength) return false;
            metrics = new SfntMetrics(data, hmtx, metricCount, unitsPerEm, glyphCount);
            return true;
        }
    }
}
