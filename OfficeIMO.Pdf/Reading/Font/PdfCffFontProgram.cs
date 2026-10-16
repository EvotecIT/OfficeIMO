namespace OfficeIMO.Pdf;

/// <summary>
/// Bounded reader for a bare CFF font program (FontFile3 Type1C or CIDFontType0C, or the CFF table of an
/// OpenType program). It reads the glyph naming, encoding, CID charset, metrics and font matrix that a PDF
/// renderer needs to select glyphs; outlines are executed by the shared OpenType CFF engine.
/// </summary>
/// <remarks>Adobe Technical Note 5176 (The Compact Font Format Specification).</remarks>
internal sealed class PdfCffFontProgram {
    internal const int MaximumGlyphs = 65535;
    private const int MaximumCharStringScan = 4096;
    private const int MaximumFontDicts = 256;

    private readonly CffIndex _charStrings;
    private readonly byte[]? _fdSelect;
    private readonly double[] _defaultWidths;
    private readonly double[] _nominalWidths;
    private readonly Dictionary<string, int>? _glyphByName;
    private readonly Dictionary<int, int>? _glyphByCid;

    private PdfCffFontProgram(byte[] data, CffIndex charStrings, bool isCidKeyed, string[]? glyphNames, int[]? cids,
        int[] builtinEncoding, double[] fontMatrix, double[] fontBBox, byte[]? fdSelect, double[] defaultWidths, double[] nominalWidths) {
        Data = data;
        _charStrings = charStrings;
        IsCidKeyed = isCidKeyed;
        GlyphNames = glyphNames;
        Cids = cids;
        BuiltinEncoding = builtinEncoding;
        FontMatrix = fontMatrix;
        FontBBox = fontBBox;
        _fdSelect = fdSelect;
        _defaultWidths = defaultWidths;
        _nominalWidths = nominalWidths;
        if (glyphNames != null) {
            _glyphByName = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int glyph = glyphNames.Length - 1; glyph >= 0; glyph--) _glyphByName[glyphNames[glyph]] = glyph;
        }
        if (cids != null) {
            _glyphByCid = new Dictionary<int, int>();
            for (int glyph = cids.Length - 1; glyph >= 0; glyph--) _glyphByCid[cids[glyph]] = glyph;
        }
    }

    /// <summary>The bare CFF bytes.</summary>
    internal byte[] Data { get; }

    internal int GlyphCount => _charStrings.Count;

    /// <summary>True when the Top DICT declares ROS and the charset names CIDs.</summary>
    internal bool IsCidKeyed { get; }

    /// <summary>Glyph names by glyph id for name-keyed fonts.</summary>
    internal string[]? GlyphNames { get; }

    /// <summary>CIDs by glyph id for CID-keyed fonts.</summary>
    internal int[]? Cids { get; }

    /// <summary>Glyph id for each of the 256 codes of the program's built-in encoding, or 0.</summary>
    internal int[] BuiltinEncoding { get; }

    /// <summary>Effective glyph-space to text-space matrix.</summary>
    internal double[] FontMatrix { get; }

    internal double[] FontBBox { get; }

    internal bool TryGetGlyph(string name, out int glyph) {
        glyph = 0;
        return _glyphByName != null && _glyphByName.TryGetValue(name, out glyph);
    }

    internal bool TryGetGlyphForCid(int cid, out int glyph) {
        glyph = 0;
        if (_glyphByCid == null) {
            if (cid <= 0 || cid >= GlyphCount) return false;
            glyph = cid;
            return true;
        }
        return _glyphByCid.TryGetValue(cid, out glyph) && glyph > 0;
    }

    internal static PdfCffFontProgram? TryParse(byte[] data) {
        try {
            return Parse(data);
        } catch (Exception exception) when (exception is InvalidDataException or IndexOutOfRangeException or ArgumentException or OverflowException) {
            return null;
        }
    }

    /// <summary>Returns the CFF table of an OpenType ('OTTO') program, or null.</summary>
    internal static byte[]? TryExtractOpenTypeCff(byte[] data) {
        if (data.Length < 12 || ReadUInt32(data, 0) != 0x4F54544FU) return null;
        int count = ReadUInt16(data, 4);
        if (count == 0 || count > 512 || 12 + count * 16 > data.Length) return null;
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (ReadUInt32(data, record) != 0x43464620U) continue;
            uint offset = ReadUInt32(data, record + 8), length = ReadUInt32(data, record + 12);
            if (offset > (uint)data.Length || length > (uint)data.Length - offset || length < 4) return null;
            var cff = new byte[length];
            Buffer.BlockCopy(data, (int)offset, cff, 0, (int)length);
            return cff;
        }
        return null;
    }

    /// <summary>Returns the design-unit advance encoded by a charstring, or null when it cannot be read.</summary>
    internal double? GetAdvance(int glyph) {
        if (glyph < 0 || glyph >= GlyphCount) return null;
        int fd = SelectFontDict(glyph);
        CffSlice charString = _charStrings[glyph];
        return ScanCharString(charString, out double? width, out _)
            ? width.HasValue ? _nominalWidths[fd] + width.Value : _defaultWidths[fd]
            : null;
    }

    /// <summary>True when the charstring certainly paints no outline.</summary>
    internal bool IsEmptyGlyph(int glyph) {
        if (glyph < 0 || glyph >= GlyphCount) return true;
        return ScanCharString(_charStrings[glyph], out _, out bool paints) && !paints;
    }

    private int SelectFontDict(int glyph) => _fdSelect == null || glyph >= _fdSelect.Length || _fdSelect[glyph] >= _defaultWidths.Length ? 0 : _fdSelect[glyph];

    // Reads the optional width operand and whether any outline can be painted. Subroutine calls stop the scan
    // and are reported as painting, so callers never mistake a subroutine-built glyph for an empty one.
    private static bool ScanCharString(CffSlice charString, out double? width, out bool paints) {
        try {
            return ScanCharStringCore(charString, out width, out paints);
        } catch (InvalidDataException) {
            width = null;
            paints = true;
            return false;
        }
    }

    private static bool ScanCharStringCore(CffSlice charString, out double? width, out bool paints) {
        width = null;
        paints = false;
        byte[] data = charString.Data;
        int offset = charString.Offset;
        int end = Math.Min(charString.Offset + charString.Length, charString.Offset + MaximumCharStringScan);
        int stack = 0;
        double first = 0D;
        bool widthResolved = false;
        int stems = 0;
        while (offset < end) {
            byte b0 = data[offset];
            if (b0 >= 32 || b0 == 28) {
                double value = ReadCharStringNumber(data, ref offset, end);
                if (stack == 0) first = value;
                stack++;
                continue;
            }
            offset++;
            int op = b0;
            if (b0 == 12) {
                if (offset >= end) return false;
                op = 1200 + data[offset++];
            }
            switch (op) {
                case 1: case 3: case 18: case 23: // stems
                    if (!widthResolved) { if ((stack & 1) == 1) width = first; widthResolved = true; }
                    stems += stack / 2;
                    stack = 0;
                    break;
                case 19: case 20: // hintmask, cntrmask
                    if (!widthResolved) { if ((stack & 1) == 1) width = first; widthResolved = true; }
                    stems += stack / 2;
                    stack = 0;
                    offset += (stems + 7) / 8;
                    break;
                case 21: // rmoveto
                    if (!widthResolved) { if (stack > 2) width = first; widthResolved = true; }
                    paints = true;
                    return true;
                case 4: case 22: // vmoveto, hmoveto
                    if (!widthResolved) { if (stack > 1) width = first; widthResolved = true; }
                    paints = true;
                    return true;
                case 14: // endchar
                    if (!widthResolved && (stack == 1 || stack == 5)) width = first;
                    paints = stack >= 4;
                    return true;
                case 10: case 29: // callsubr, callgsubr
                    paints = true;
                    return widthResolved;
                default:
                    paints = true;
                    return widthResolved;
            }
        }
        return widthResolved;
    }

    private static double ReadCharStringNumber(byte[] data, ref int offset, int end) {
        byte b0 = data[offset];
        if (b0 == 28) {
            if (offset + 3 > end) throw new InvalidDataException("Truncated CFF charstring number.");
            short value = (short)((data[offset + 1] << 8) | data[offset + 2]);
            offset += 3;
            return value;
        }
        if (b0 <= 246) { offset++; return b0 - 139; }
        if (b0 <= 250) {
            if (offset + 2 > end) throw new InvalidDataException("Truncated CFF charstring number.");
            int value = ((b0 - 247) * 256) + data[offset + 1] + 108;
            offset += 2;
            return value;
        }
        if (b0 <= 254) {
            if (offset + 2 > end) throw new InvalidDataException("Truncated CFF charstring number.");
            int value = -((b0 - 251) * 256) - data[offset + 1] - 108;
            offset += 2;
            return value;
        }
        if (offset + 5 > end) throw new InvalidDataException("Truncated CFF charstring number.");
        int fixedValue = (data[offset + 1] << 24) | (data[offset + 2] << 16) | (data[offset + 3] << 8) | data[offset + 4];
        offset += 5;
        return fixedValue / 65536D;
    }

    private static PdfCffFontProgram Parse(byte[] data) {
        if (data.Length < 4 || data[0] != 1) throw new InvalidDataException("Not a CFF1 program.");
        int cursor = data[2];
        if (cursor < 4 || cursor > data.Length) throw new InvalidDataException("Invalid CFF header size.");
        CffIndex names = CffIndex.Read(data, ref cursor);
        CffIndex topDicts = CffIndex.Read(data, ref cursor);
        CffIndex strings = CffIndex.Read(data, ref cursor);
        CffIndex.Read(data, ref cursor); // Global subroutines are executed by the shared engine.
        if (names.Count < 1 || topDicts.Count < 1) throw new InvalidDataException("The CFF program has no font.");
        Dictionary<int, double[]> top = ParseDict(topDicts[0]);
        if (top.TryGetValue(1206, out double[]? type) && type.Length > 0 && type[0] != 2) throw new InvalidDataException("Unsupported charstring type.");
        int charStringsOffset = RequireOffset(top, 17, data.Length);
        CffIndex charStrings = CffIndex.ReadAt(data, charStringsOffset);
        int glyphCount = charStrings.Count;
        if (glyphCount < 1 || glyphCount > MaximumGlyphs) throw new InvalidDataException("Invalid CFF glyph count.");

        bool isCidKeyed = top.ContainsKey(1230);
        double[] topMatrix = top.TryGetValue(1207, out double[]? matrix) && matrix.Length == 6 ? matrix : new[] { 0.001D, 0D, 0D, 0.001D, 0D, 0D };
        double[] bbox = top.TryGetValue(5, out double[]? box) && box.Length == 4 ? box : new[] { 0D, -200D, 1000D, 800D };
        int charsetOffset = top.TryGetValue(15, out double[]? charsetValue) && charsetValue.Length > 0 ? CheckedInt(charsetValue[0]) : 0;
        int[] charset = ReadCharset(data, charsetOffset, glyphCount);

        double[] defaultWidths;
        double[] nominalWidths;
        byte[]? fdSelect = null;
        double[] effectiveMatrix = topMatrix;
        if (isCidKeyed) {
            int fdArrayOffset = RequireOffset(top, 1236, data.Length);
            CffIndex fdArray = CffIndex.ReadAt(data, fdArrayOffset);
            if (fdArray.Count < 1 || fdArray.Count > MaximumFontDicts) throw new InvalidDataException("Invalid CFF FDArray.");
            defaultWidths = new double[fdArray.Count];
            nominalWidths = new double[fdArray.Count];
            for (int index = 0; index < fdArray.Count; index++) {
                Dictionary<int, double[]> fontDict = ParseDict(fdArray[index]);
                ReadPrivateWidths(data, fontDict, out defaultWidths[index], out nominalWidths[index]);
                if (index == 0 && fontDict.TryGetValue(1207, out double[]? fdMatrix) && fdMatrix.Length == 6) {
                    effectiveMatrix = top.ContainsKey(1207) ? Multiply(fdMatrix, topMatrix) : fdMatrix;
                }
            }
            if (top.TryGetValue(1237, out double[]? fdSelectValue) && fdSelectValue.Length > 0) {
                fdSelect = ReadFdSelect(data, CheckedInt(fdSelectValue[0]), glyphCount);
            }
        } else {
            defaultWidths = new double[1];
            nominalWidths = new double[1];
            ReadPrivateWidths(data, top, out defaultWidths[0], out nominalWidths[0]);
        }

        string[]? glyphNames = null;
        int[]? cids = null;
        if (isCidKeyed) {
            cids = charset;
        } else {
            glyphNames = new string[glyphCount];
            var decodedNames = new Dictionary<int, string>();
            long decodedCharacters = 0;
            for (int glyph = 0; glyph < glyphCount; glyph++) {
                int sid = charset[glyph];
                if (!decodedNames.TryGetValue(sid, out string? name)) {
                    name = ResolveString(strings, sid);
                    decodedCharacters += name.Length;
                    if (decodedCharacters > 1024 * 1024) throw new InvalidDataException("CFF glyph names exceed the decoded-name budget.");
                    decodedNames.Add(sid, name);
                }
                glyphNames[glyph] = name;
            }
        }
        int[] builtinEncoding = new int[256];
        if (!isCidKeyed) ReadEncoding(data, top, charset, glyphNames!, builtinEncoding);
        if (!IsUsableMatrix(effectiveMatrix)) throw new InvalidDataException("The CFF font matrix is singular.");
        return new PdfCffFontProgram(data, charStrings, isCidKeyed, glyphNames, cids, builtinEncoding, effectiveMatrix, bbox,
            fdSelect, defaultWidths, nominalWidths);
    }

    private static bool IsUsableMatrix(double[] matrix) {
        foreach (double value in matrix) if (double.IsNaN(value) || double.IsInfinity(value)) return false;
        return Math.Abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]) > 1e-12D;
    }

    private static double[] Multiply(double[] first, double[] second) => new[] {
        first[0] * second[0] + first[1] * second[2],
        first[0] * second[1] + first[1] * second[3],
        first[2] * second[0] + first[3] * second[2],
        first[2] * second[1] + first[3] * second[3],
        first[4] * second[0] + first[5] * second[2] + second[4],
        first[4] * second[1] + first[5] * second[3] + second[5]
    };

    private static void ReadPrivateWidths(byte[] data, Dictionary<int, double[]> dict, out double defaultWidth, out double nominalWidth) {
        defaultWidth = 0D;
        nominalWidth = 0D;
        if (!dict.TryGetValue(18, out double[]? privateValue) || privateValue.Length != 2) return;
        int size = CheckedInt(privateValue[0]);
        int offset = CheckedInt(privateValue[1]);
        if (size <= 0 || offset < 0 || offset > data.Length - size) return;
        Dictionary<int, double[]> privateDict = ParseDict(new CffSlice(data, offset, size));
        if (privateDict.TryGetValue(20, out double[]? defaultValue) && defaultValue.Length > 0) defaultWidth = defaultValue[0];
        if (privateDict.TryGetValue(21, out double[]? nominalValue) && nominalValue.Length > 0) nominalWidth = nominalValue[0];
    }

    private static int RequireOffset(Dictionary<int, double[]> dict, int op, int length) {
        if (!dict.TryGetValue(op, out double[]? value) || value.Length < 1) throw new InvalidDataException("A required CFF offset is missing.");
        int offset = CheckedInt(value[0]);
        if (offset <= 0 || offset >= length) throw new InvalidDataException("A CFF offset is out of range.");
        return offset;
    }

    private static int CheckedInt(double value) {
        if (double.IsNaN(value) || value < int.MinValue || value > int.MaxValue || value != Math.Truncate(value)) throw new InvalidDataException("Invalid CFF integer.");
        return (int)value;
    }

    private static string ResolveString(CffIndex strings, int sid) {
        if (sid < PdfCffStandardStrings.Count) return PdfCffStandardStrings.Get(sid);
        int index = sid - PdfCffStandardStrings.Count;
        if (index >= strings.Count) return "sid" + sid.ToString(System.Globalization.CultureInfo.InvariantCulture);
        CffSlice slice = strings[index];
        if (slice.Length > 1024 * 1024) throw new InvalidDataException("CFF glyph name exceeds the decoded-name budget.");
        return System.Text.Encoding.ASCII.GetString(slice.Data, slice.Offset, slice.Length);
    }

    private static int[] ReadCharset(byte[] data, int offset, int glyphCount) {
        var result = new int[glyphCount];
        if (offset == 0) {
            // ISOAdobe: SIDs 1..228 are glyphs 1..228 in order.
            for (int glyph = 1; glyph < glyphCount; glyph++) result[glyph] = glyph <= 228 ? glyph : 0;
            return result;
        }
        if (offset == 1 || offset == 2) return result; // Expert charsets name no PDF-encoded glyphs.
        if (offset < 0 || offset >= data.Length) throw new InvalidDataException("Invalid CFF charset offset.");
        int cursor = offset;
        int format = data[cursor++];
        int gid = 1;
        if (format == 0) {
            for (; gid < glyphCount; gid++) {
                if (cursor + 2 > data.Length) throw new InvalidDataException("Truncated CFF charset.");
                result[gid] = ReadUInt16(data, cursor);
                cursor += 2;
            }
        } else if (format == 1 || format == 2) {
            while (gid < glyphCount) {
                int rangeSize = format == 1 ? 3 : 4;
                if (cursor + rangeSize > data.Length) throw new InvalidDataException("Truncated CFF charset.");
                int first = ReadUInt16(data, cursor);
                int left = format == 1 ? data[cursor + 2] : ReadUInt16(data, cursor + 2);
                cursor += rangeSize;
                for (int index = 0; index <= left && gid < glyphCount; index++) result[gid++] = first + index;
            }
        } else {
            throw new InvalidDataException("Unsupported CFF charset format.");
        }
        return result;
    }

    private static byte[] ReadFdSelect(byte[] data, int offset, int glyphCount) {
        if (offset <= 0 || offset >= data.Length) throw new InvalidDataException("Invalid CFF FDSelect offset.");
        var result = new byte[glyphCount];
        int format = data[offset];
        if (format == 0) {
            if (offset + 1 + glyphCount > data.Length) throw new InvalidDataException("Truncated CFF FDSelect.");
            Buffer.BlockCopy(data, offset + 1, result, 0, glyphCount);
            return result;
        }
        if (format != 3) throw new InvalidDataException("Unsupported CFF FDSelect format.");
        if (offset + 3 > data.Length) throw new InvalidDataException("Truncated CFF FDSelect.");
        int ranges = ReadUInt16(data, offset + 1);
        int cursor = offset + 3;
        if (cursor + ranges * 3 + 2 > data.Length) throw new InvalidDataException("Truncated CFF FDSelect.");
        for (int index = 0; index < ranges; index++) {
            int first = ReadUInt16(data, cursor);
            byte fd = data[cursor + 2];
            int next = ReadUInt16(data, cursor + 3);
            for (int glyph = first; glyph < next && glyph < glyphCount; glyph++) result[glyph] = fd;
            cursor += 3;
        }
        return result;
    }

    private static void ReadEncoding(byte[] data, Dictionary<int, double[]> top, int[] charset, string[] glyphNames, int[] encoding) {
        int offset = top.TryGetValue(16, out double[]? value) && value.Length > 0 ? CheckedInt(value[0]) : 0;
        if (offset == 0 || offset == 1) {
            string?[] names = PdfFontEncodingNames.Get("StandardEncoding")!;
            var glyphByName = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int glyph = glyphNames.Length - 1; glyph > 0; glyph--) glyphByName[glyphNames[glyph]] = glyph;
            for (int code = 0; code < 256; code++) {
                if (names[code] is string name && glyphByName.TryGetValue(name, out int glyph)) encoding[code] = glyph;
            }
            return;
        }
        if (offset < 0 || offset >= data.Length) return;
        int cursor = offset;
        int format = data[cursor++];
        int glyphId = 1;
        if ((format & 0x7F) == 0) {
            if (cursor >= data.Length) return;
            int codes = data[cursor++];
            for (int index = 0; index < codes && cursor < data.Length; index++, glyphId++) {
                int code = data[cursor++];
                if (glyphId < charset.Length) encoding[code] = glyphId;
            }
        } else if ((format & 0x7F) == 1) {
            if (cursor >= data.Length) return;
            int ranges = data[cursor++];
            for (int index = 0; index < ranges && cursor + 2 <= data.Length; index++) {
                int first = data[cursor], left = data[cursor + 1];
                cursor += 2;
                for (int code = first; code <= first + left && code < 256; code++, glyphId++) {
                    if (glyphId < charset.Length) encoding[code] = glyphId;
                }
            }
        } else {
            return;
        }
        if ((format & 0x80) != 0 && cursor < data.Length) {
            int supplements = data[cursor++];
            for (int index = 0; index < supplements && cursor + 3 <= data.Length; index++) {
                int code = data[cursor];
                int sid = ReadUInt16(data, cursor + 1);
                cursor += 3;
                int glyph = Array.IndexOf(charset, sid, 1);
                if (glyph > 0) encoding[code] = glyph;
            }
        }
    }

    private static Dictionary<int, double[]> ParseDict(CffSlice slice) {
        var result = new Dictionary<int, double[]>();
        var operands = new List<double>(8);
        byte[] data = slice.Data;
        int offset = slice.Offset;
        int end = slice.Offset + slice.Length;
        while (offset < end) {
            byte b0 = data[offset];
            if (b0 <= 21) {
                offset++;
                int op = b0;
                if (b0 == 12) {
                    if (offset >= end) break;
                    op = 1200 + data[offset++];
                }
                result[op] = operands.ToArray();
                operands.Clear();
                continue;
            }
            if (operands.Count >= 48) throw new InvalidDataException("CFF DICT operand stack overflow.");
            operands.Add(ReadDictNumber(data, ref offset, end));
        }
        return result;
    }

    private static double ReadDictNumber(byte[] data, ref int offset, int end) {
        byte b0 = data[offset];
        if (b0 == 28) {
            if (offset + 3 > end) throw new InvalidDataException("Truncated CFF DICT number.");
            short value = (short)((data[offset + 1] << 8) | data[offset + 2]);
            offset += 3;
            return value;
        }
        if (b0 == 29) {
            if (offset + 5 > end) throw new InvalidDataException("Truncated CFF DICT number.");
            int value = (data[offset + 1] << 24) | (data[offset + 2] << 16) | (data[offset + 3] << 8) | data[offset + 4];
            offset += 5;
            return value;
        }
        if (b0 == 30) {
            offset++;
            var text = new System.Text.StringBuilder();
            while (offset < end) {
                byte packed = data[offset++];
                for (int shift = 4; shift >= 0; shift -= 4) {
                    int nibble = (packed >> shift) & 0xF;
                    if (nibble == 0xF) {
                        return double.TryParse(text.ToString(), System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double real) ? real : 0D;
                    }
                    if (text.Length > 64) throw new InvalidDataException("CFF real number is too long.");
                    text.Append(nibble switch {
                        <= 9 => ((char)('0' + nibble)).ToString(),
                        0xA => ".",
                        0xB => "E",
                        0xC => "E-",
                        0xE => "-",
                        _ => string.Empty
                    });
                }
            }
            throw new InvalidDataException("Unterminated CFF real number.");
        }
        if (b0 >= 32 && b0 <= 246) { offset++; return b0 - 139; }
        if (b0 >= 247 && b0 <= 250) {
            if (offset + 2 > end) throw new InvalidDataException("Truncated CFF DICT number.");
            int value = ((b0 - 247) * 256) + data[offset + 1] + 108;
            offset += 2;
            return value;
        }
        if (b0 >= 251 && b0 <= 254) {
            if (offset + 2 > end) throw new InvalidDataException("Truncated CFF DICT number.");
            int value = -((b0 - 251) * 256) - data[offset + 1] - 108;
            offset += 2;
            return value;
        }
        throw new InvalidDataException("Invalid CFF DICT byte.");
    }

    private static int ReadUInt16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];

    internal readonly struct CffSlice {
        internal CffSlice(byte[] data, int offset, int length) { Data = data; Offset = offset; Length = length; }
        internal byte[] Data { get; }
        internal int Offset { get; }
        internal int Length { get; }
    }

    internal sealed class CffIndex {
        private readonly byte[] _data;
        private readonly int[] _offsets;

        private CffIndex(byte[] data, int[] offsets) { _data = data; _offsets = offsets; }

        internal int Count => _offsets.Length == 0 ? 0 : _offsets.Length - 1;

        internal CffSlice this[int index] => new(_data, _offsets[index], _offsets[index + 1] - _offsets[index]);

        internal static CffIndex ReadAt(byte[] data, int offset) {
            int cursor = offset;
            return Read(data, ref cursor);
        }

        internal static CffIndex Read(byte[] data, ref int cursor) {
            if (cursor < 0 || cursor + 2 > data.Length) throw new InvalidDataException("Truncated CFF INDEX.");
            int count = ReadUInt16(data, cursor);
            cursor += 2;
            if (count == 0) return new CffIndex(data, Array.Empty<int>());
            if (cursor >= data.Length) throw new InvalidDataException("Truncated CFF INDEX.");
            int offSize = data[cursor++];
            if (offSize < 1 || offSize > 4) throw new InvalidDataException("Invalid CFF INDEX offset size.");
            long offsetsLength = (long)(count + 1) * offSize;
            if (cursor + offsetsLength > data.Length) throw new InvalidDataException("Truncated CFF INDEX offsets.");
            int dataStart = cursor + (int)offsetsLength - 1;
            var offsets = new int[count + 1];
            for (int index = 0; index <= count; index++) {
                uint value = 0;
                for (int b = 0; b < offSize; b++) value = (value << 8) | data[cursor++];
                long absolute = dataStart + (long)value;
                if (value < 1 || absolute > data.Length || index > 0 && absolute < offsets[index - 1]) throw new InvalidDataException("Invalid CFF INDEX offset.");
                offsets[index] = (int)absolute;
            }
            cursor = offsets[count];
            return new CffIndex(data, offsets);
        }
    }
}
