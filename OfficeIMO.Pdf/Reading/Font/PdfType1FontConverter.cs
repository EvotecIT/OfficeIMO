using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Converts an embedded Type 1 font program (FontFile) into an equivalent bare CFF program. The cleartext
/// part supplies the font matrix and built-in encoding; the eexec-encrypted private part supplies the
/// charstrings, which are interpreted (subroutines, flex, hint replacement and seac accents) and re-emitted
/// as hint-free Type 2 outlines.
/// </summary>
/// <remarks>Adobe Type 1 Font Format and Technical Note 5176. All work is bounded for untrusted input.</remarks>
internal static partial class PdfType1FontConverter {
    private const int MaximumGlyphs = 4096;
    private const int MaximumSubroutines = 65535;
    private const int MaximumOperations = 20_000;
    private const int MaximumCallDepth = 10;
    private const int MaximumOutputPoints = 20_000;

    private static readonly char[] ArrayOpeners = { '[', '{' };
    private static readonly char[] ArrayClosers = { ']', '}' };
    private static readonly char[] Separators = { ' ', '\t', '\r', '\n' };

    internal sealed class Result {
        internal Result(byte[] cff, string?[] builtinEncoding) {
            Cff = cff;
            BuiltinEncoding = builtinEncoding;
        }

        internal byte[] Cff { get; }

        /// <summary>Glyph names of the program's built-in encoding by code.</summary>
        internal string?[] BuiltinEncoding { get; }
    }

    internal static Result? TryConvert(byte[] program) {
        try {
            return Convert(program);
        } catch (Exception exception) when (exception is InvalidDataException or IndexOutOfRangeException or ArgumentException or OverflowException) {
            return null;
        }
    }

    private static Result? Convert(byte[] program) {
        byte[] data = StripPfbHeaders(program);
        int eexec = IndexOf(data, "eexec", 0);
        if (eexec < 0) return null;
        string clearText = Encoding.ASCII.GetString(data, 0, eexec);
        int encryptedStart = eexec + 5;
        while (encryptedStart < data.Length && IsWhitespace(data[encryptedStart])) encryptedStart++;
        byte[] encrypted = ExtractEncrypted(data, encryptedStart);
        byte[] privatePart = Decrypt(encrypted, 55665, 4);

        double[] fontMatrix = ReadNumberArray(clearText, "/FontMatrix", 6) ?? new[] { 0.001D, 0D, 0D, 0.001D, 0D, 0D };
        double[] fontBBox = ReadNumberArray(clearText, "/FontBBox", 4) ?? new[] { 0D, -200D, 1000D, 800D };
        string fontName = ReadName(clearText, "/FontName") ?? "Type1";
        string?[] builtinEncoding = ReadEncoding(clearText);

        int lenIv = 4;
        int lenIvAt = IndexOf(privatePart, "/lenIV", 0);
        int charStringsAt = IndexOf(privatePart, "/CharStrings", 0);
        if (charStringsAt < 0) return null;
        if (lenIvAt >= 0 && lenIvAt < charStringsAt) {
            int cursor = lenIvAt + 6;
            if (TryReadInteger(privatePart, ref cursor, out int value) && value >= -1 && value <= 255) lenIv = value;
        }
        Dictionary<int, byte[]> subrs = ReadSubrs(privatePart, charStringsAt, lenIv);
        List<(string Name, byte[] CharString)> charStrings = ReadCharStrings(privatePart, charStringsAt, lenIv);
        if (charStrings.Count == 0) return null;

        var byName = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in charStrings) byName[entry.Name] = entry.CharString;
        var names = new List<string> { ".notdef" };
        var seen = new HashSet<string>(StringComparer.Ordinal) { ".notdef" };
        foreach (var entry in charStrings) {
            if (seen.Add(entry.Name)) names.Add(entry.Name);
            if (names.Count >= MaximumGlyphs) break;
        }
        var outlines = new List<byte[]>(names.Count);
        foreach (string name in names) {
            var path = new GlyphPath();
            if (byName.TryGetValue(name, out byte[]? charString)) {
                var interpreter = new Interpreter(subrs, byName, path);
                try {
                    interpreter.Run(charString);
                } catch (InvalidDataException) {
                    path = new GlyphPath { Width = path.Width };
                }
            }
            outlines.Add(path.ToType2CharString());
        }
        return new Result(BuildCff(fontName, names, outlines, fontMatrix, fontBBox), builtinEncoding);
    }

    private static byte[] StripPfbHeaders(byte[] data) {
        if (data.Length < 6 || data[0] != 0x80) return data;
        using var output = new MemoryStream(data.Length);
        int cursor = 0;
        while (cursor + 2 <= data.Length && data[cursor] == 0x80) {
            int type = data[cursor + 1];
            if (type == 3) break;
            if (cursor + 6 > data.Length) break;
            int length = data[cursor + 2] | (data[cursor + 3] << 8) | (data[cursor + 4] << 16) | (data[cursor + 5] << 24);
            cursor += 6;
            if (length < 0 || length > data.Length - cursor) length = data.Length - cursor;
            output.Write(data, cursor, length);
            cursor += length;
        }
        return output.ToArray();
    }

    private static byte[] ExtractEncrypted(byte[] data, int start) {
        // Binary eexec sections are the norm in PDF; a hexadecimal section starts with four hex digits.
        bool hex = start + 4 <= data.Length;
        for (int index = start; hex && index < start + 4; index++) hex = IsHexDigit(data[index]);
        if (!hex) {
            var binary = new byte[data.Length - start];
            Buffer.BlockCopy(data, start, binary, 0, binary.Length);
            return binary;
        }
        var decoded = new List<byte>((data.Length - start) / 2);
        int high = -1;
        for (int index = start; index < data.Length; index++) {
            byte value = data[index];
            if (IsWhitespace(value)) continue;
            if (!IsHexDigit(value)) break;
            int nibble = value <= '9' ? value - '0' : (value | 0x20) - 'a' + 10;
            if (high < 0) high = nibble;
            else {
                decoded.Add((byte)((high << 4) | nibble));
                high = -1;
            }
        }
        return decoded.ToArray();
    }

    private static byte[] Decrypt(byte[] encrypted, ushort key, int skip) {
        if (skip < 0) skip = 0;
        if (encrypted.Length <= skip) return Array.Empty<byte>();
        var result = new byte[encrypted.Length - skip];
        ushort state = key;
        for (int index = 0; index < encrypted.Length; index++) {
            byte cipher = encrypted[index];
            byte plain = (byte)(cipher ^ (state >> 8));
            state = unchecked((ushort)((cipher + state) * 52845 + 22719));
            if (index >= skip) result[index - skip] = plain;
        }
        return result;
    }

    private static double[]? ReadNumberArray(string text, string key, int count) {
        int at = text.IndexOf(key, StringComparison.Ordinal);
        if (at < 0) return null;
        int start = text.IndexOfAny(ArrayOpeners, at + key.Length);
        if (start < 0 || start - at > key.Length + 8) return null;
        int end = text.IndexOfAny(ArrayClosers, start + 1);
        if (end < 0) return null;
        string[] parts = text.Substring(start + 1, end - start - 1).Split(Separators, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != count) return null;
        var values = new double[count];
        for (int index = 0; index < count; index++) {
            if (!double.TryParse(parts[index], NumberStyles.Float, CultureInfo.InvariantCulture, out values[index]) ||
                double.IsNaN(values[index]) || double.IsInfinity(values[index])) return null;
        }
        return values;
    }

    private static string? ReadName(string text, string key) {
        int at = text.IndexOf(key, StringComparison.Ordinal);
        if (at < 0) return null;
        int start = text.IndexOf('/', at + key.Length);
        if (start < 0) return null;
        int end = start + 1;
        while (end < text.Length && !char.IsWhiteSpace(text[end]) && text[end] is not ('/' or '[' or '{' or '(' or '<')) end++;
        return end > start + 1 ? text.Substring(start + 1, end - start - 1) : null;
    }

    private static Dictionary<int, byte[]> ReadSubrs(byte[] data, int limit, int lenIv) {
        var subrs = new Dictionary<int, byte[]>();
        int at = IndexOf(data, "/Subrs", 0);
        if (at < 0 || at > limit) return subrs;
        int cursor = at + 6;
        if (!TryReadInteger(data, ref cursor, out int declared) || declared < 0 || declared > MaximumSubroutines) return subrs;
        while (cursor < limit) {
            int dup = IndexOf(data, "dup", cursor);
            if (dup < 0 || dup >= limit) break;
            cursor = dup + 3;
            if (!TryReadInteger(data, ref cursor, out int index) || !TryReadInteger(data, ref cursor, out int length) ||
                index < 0 || index >= Math.Max(declared, 1) || length < 0 || !TryReadBinaryStart(data, ref cursor) ||
                length > data.Length - cursor) continue;
            subrs[index] = DecryptCharString(data, cursor, length, lenIv);
            cursor += length;
        }
        return subrs;
    }

    private static List<(string Name, byte[] CharString)> ReadCharStrings(byte[] data, int start, int lenIv) {
        var result = new List<(string, byte[])>();
        int cursor = start + 12;
        if (!TryReadInteger(data, ref cursor, out _)) return result;
        while (cursor < data.Length && result.Count < MaximumGlyphs) {
            while (cursor < data.Length && data[cursor] != '/') {
                if (data[cursor] == 'e' && cursor + 3 <= data.Length && data[cursor + 1] == 'n' && data[cursor + 2] == 'd' &&
                    (cursor + 3 == data.Length || IsWhitespace(data[cursor + 3]))) return result;
                cursor++;
            }
            if (cursor >= data.Length) break;
            int nameStart = ++cursor;
            while (cursor < data.Length && !IsWhitespace(data[cursor]) && data[cursor] != '/') cursor++;
            if (cursor == nameStart) continue;
            string name = Encoding.ASCII.GetString(data, nameStart, cursor - nameStart);
            int afterName = cursor;
            if (!TryReadInteger(data, ref cursor, out int length) || length < 0 || !TryReadBinaryStart(data, ref cursor) ||
                length > data.Length - cursor) {
                cursor = afterName;
                continue;
            }
            result.Add((name, DecryptCharString(data, cursor, length, lenIv)));
            cursor += length;
        }
        return result;
    }

    private static byte[] DecryptCharString(byte[] data, int offset, int length, int lenIv) {
        var slice = new byte[length];
        Buffer.BlockCopy(data, offset, slice, 0, length);
        return lenIv < 0 ? slice : Decrypt(slice, 4330, lenIv);
    }

    // Reads the RD/-| token after a length and positions the cursor at the first binary byte.
    private static bool TryReadBinaryStart(byte[] data, ref int cursor) {
        while (cursor < data.Length && IsWhitespace(data[cursor])) cursor++;
        int start = cursor;
        while (cursor < data.Length && !IsWhitespace(data[cursor])) cursor++;
        int length = cursor - start;
        if (length is < 1 or > 8) return false;
        string token = Encoding.ASCII.GetString(data, start, length);
        if (token is not ("RD" or "-|")) return false;
        if (cursor >= data.Length) return false;
        cursor++; // exactly one separator byte precedes the binary data
        return true;
    }

    private static bool TryReadInteger(byte[] data, ref int cursor, out int value) {
        value = 0;
        while (cursor < data.Length && IsWhitespace(data[cursor])) cursor++;
        int start = cursor;
        if (cursor < data.Length && data[cursor] == '-') cursor++;
        while (cursor < data.Length && data[cursor] >= '0' && data[cursor] <= '9') cursor++;
        if (cursor == start || cursor - start > 9) return false;
        return int.TryParse(Encoding.ASCII.GetString(data, start, cursor - start), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value);
    }

    private static int IndexOf(byte[] data, string token, int start) {
        if (token.Length == 0) return -1;
        byte first = (byte)token[0];
        for (int index = Math.Max(0, start); index <= data.Length - token.Length; index++) {
            if (data[index] != first) continue;
            int matched = 1;
            while (matched < token.Length && data[index + matched] == token[matched]) matched++;
            if (matched == token.Length) return index;
        }
        return -1;
    }

    private static bool IsWhitespace(byte value) => value is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'\f' or 0;

    private static bool IsHexDigit(byte value) => value is >= (byte)'0' and <= (byte)'9' or >= (byte)'a' and <= (byte)'f' or >= (byte)'A' and <= (byte)'F';

    private sealed class GlyphPath {
        private readonly List<(byte Kind, double[] Points)> _segments = new();
        private int _points;

        internal double Width { get; set; }

        internal void MoveTo(double x, double y) => Add(0, x, y);

        internal void LineTo(double x, double y) => Add(1, x, y);

        internal void CurveTo(double x1, double y1, double x2, double y2, double x3, double y3) => Add(2, x1, y1, x2, y2, x3, y3);

        private void Add(byte kind, params double[] points) {
            _points += points.Length / 2;
            if (_points > MaximumOutputPoints) throw new InvalidDataException("The Type 1 glyph exceeds the outline budget.");
            _segments.Add((kind, points));
        }

        internal byte[] ToType2CharString() {
            using var output = new MemoryStream();
            double width = Width < 0D || double.IsNaN(Width) || double.IsInfinity(Width) ? 0D : Width;
            bool widthWritten = false;
            double currentX = 0D, currentY = 0D;
            foreach (var segment in _segments) {
                double[] points = segment.Points;
                if (segment.Kind == 0) {
                    if (!widthWritten) { WriteNumber(output, width); widthWritten = true; }
                    WriteDelta(output, points[0], points[1], ref currentX, ref currentY);
                    output.WriteByte(21);
                } else if (segment.Kind == 1) {
                    WriteDelta(output, points[0], points[1], ref currentX, ref currentY);
                    output.WriteByte(5);
                } else {
                    WriteDelta(output, points[0], points[1], ref currentX, ref currentY);
                    WriteDelta(output, points[2], points[3], ref currentX, ref currentY);
                    WriteDelta(output, points[4], points[5], ref currentX, ref currentY);
                    output.WriteByte(8);
                }
            }
            if (!widthWritten) WriteNumber(output, width);
            output.WriteByte(14);
            return output.ToArray();
        }

        private static void WriteDelta(MemoryStream output, double x, double y, ref double currentX, ref double currentY) {
            double qx = Quantize(x), qy = Quantize(y);
            WriteNumber(output, qx - currentX);
            WriteNumber(output, qy - currentY);
            currentX = qx;
            currentY = qy;
        }

        private static double Quantize(double value) => Math.Round(value * 65536D) / 65536D;
    }

    private static void WriteNumber(MemoryStream output, double value) {
        if (value == Math.Truncate(value) && value >= -32768D && value <= 32767D) {
            int integer = (int)value;
            if (integer >= -107 && integer <= 107) {
                output.WriteByte((byte)(integer + 139));
            } else if (integer >= 108 && integer <= 1131) {
                integer -= 108;
                output.WriteByte((byte)((integer >> 8) + 247));
                output.WriteByte((byte)integer);
            } else if (integer >= -1131 && integer <= -108) {
                integer = -integer - 108;
                output.WriteByte((byte)((integer >> 8) + 251));
                output.WriteByte((byte)integer);
            } else {
                output.WriteByte(28);
                output.WriteByte((byte)(integer >> 8));
                output.WriteByte((byte)integer);
            }
            return;
        }
        double clamped = Math.Max(-32768D, Math.Min(32767.99D, value));
        int fixedValue = (int)Math.Round(clamped * 65536D);
        output.WriteByte(255);
        output.WriteByte((byte)(fixedValue >> 24));
        output.WriteByte((byte)(fixedValue >> 16));
        output.WriteByte((byte)(fixedValue >> 8));
        output.WriteByte((byte)fixedValue);
    }

    private sealed class Interpreter {
        private readonly Dictionary<int, byte[]> _subrs;
        private readonly Dictionary<string, byte[]> _glyphs;
        private readonly GlyphPath _path;
        private readonly List<double> _stack = new(32);
        private readonly Stack<double> _postScriptStack = new();
        private readonly List<(double X, double Y)> _flexPoints = new();
        private double _x, _y, _offsetX, _offsetY;
        private bool _flexing, _open, _widthSet;
        private int _operations;

        internal Interpreter(Dictionary<int, byte[]> subrs, Dictionary<string, byte[]> glyphs, GlyphPath path) {
            _subrs = subrs;
            _glyphs = glyphs;
            _path = path;
        }

        internal void Run(byte[] charString) => Execute(charString, 0);

        // Returns true when endchar finished the glyph.
        private bool Execute(byte[] program, int depth) {
            if (depth > MaximumCallDepth) throw new InvalidDataException("Type 1 subroutine nesting is too deep.");
            int offset = 0;
            while (offset < program.Length) {
                if (++_operations > MaximumOperations) throw new InvalidDataException("The Type 1 charstring exceeds its operation budget.");
                int b0 = program[offset++];
                if (b0 >= 32) {
                    Push(ReadNumber(program, ref offset, b0));
                    continue;
                }
                switch (b0) {
                    case 1: case 3: _stack.Clear(); break; // hstem, vstem
                    case 4: RMoveTo(0D, Arg(0)); break; // vmoveto
                    case 5: RLineTo(Arg(0), Arg(1)); break;
                    case 6: RLineTo(Arg(0), 0D); break;
                    case 7: RLineTo(0D, Arg(0)); break;
                    case 8: RCurveTo(Arg(0), Arg(1), Arg(2), Arg(3), Arg(4), Arg(5)); break;
                    case 9: _open = false; _stack.Clear(); break; // closepath: CFF contours close implicitly
                    case 10: {
                        int index = (int)Pop();
                        if (!_subrs.TryGetValue(index, out byte[]? subr)) throw new InvalidDataException("Missing Type 1 subroutine.");
                        if (Execute(subr, depth + 1)) return true;
                        break;
                    }
                    case 11: return false; // return
                    case 13: { // hsbw
                        double sbx = Arg(0), wx = Arg(1);
                        SetSideBearing(sbx, 0D, wx);
                        break;
                    }
                    case 14: _stack.Clear(); return true; // endchar
                    case 21: RMoveTo(Arg(0), Arg(1)); break;
                    case 22: RMoveTo(Arg(0), 0D); break;
                    case 30: RCurveTo(0D, Arg(0), Arg(1), Arg(2), Arg(3), 0D); break; // vhcurveto
                    case 31: RCurveTo(Arg(0), 0D, Arg(1), Arg(2), 0D, Arg(3)); break; // hvcurveto
                    case 12: {
                        if (offset >= program.Length) throw new InvalidDataException("Truncated Type 1 escape.");
                        if (ExecuteEscape(program[offset++], depth)) return true;
                        break;
                    }
                    default:
                        _stack.Clear();
                        break;
                }
            }
            return false;
        }

        private bool ExecuteEscape(int op, int depth) {
            switch (op) {
                case 0: case 1: case 2: _stack.Clear(); return false; // dotsection, vstem3, hstem3
                case 6: Seac(Arg(0), Arg(1), Arg(2), (int)Arg(3), (int)Arg(4), depth); return true;
                case 7: SetSideBearing(Arg(0), Arg(1), Arg(2)); return false; // sbw
                case 12: { // div
                    if (_stack.Count < 2) throw new InvalidDataException("Type 1 div underflow.");
                    double divisor = _stack[_stack.Count - 1], dividend = _stack[_stack.Count - 2];
                    _stack.RemoveRange(_stack.Count - 2, 2);
                    Push(divisor == 0D ? 0D : dividend / divisor);
                    return false;
                }
                case 16: CallOtherSubr(); return false;
                case 17: Push(_postScriptStack.Count > 0 ? _postScriptStack.Pop() : 0D); return false; // pop
                case 33: { // setcurrentpoint
                    double x = Arg(0), y = Arg(1);
                    _stack.Clear();
                    _x = x; _y = y;
                    return false;
                }
                default: _stack.Clear(); return false;
            }
        }

        private void CallOtherSubr() {
            int other = (int)Pop();
            int count = (int)Pop();
            if (count < 0 || count > _stack.Count) throw new InvalidDataException("Invalid Type 1 othersubr arguments.");
            var arguments = new double[count];
            for (int index = count - 1; index >= 0; index--) arguments[index] = Pop();
            switch (other) {
                case 0: { // end flex: the first rmoveto names the reference point; the next six are two curves
                    if (_flexPoints.Count >= 8) {
                        var p = _flexPoints;
                        CurveAbsolute(p[2].X, p[2].Y, p[3].X, p[3].Y, p[4].X, p[4].Y);
                        CurveAbsolute(p[5].X, p[5].Y, p[6].X, p[6].Y, p[7].X, p[7].Y);
                        _x = p[7].X; _y = p[7].Y;
                    }
                    _flexing = false;
                    _flexPoints.Clear();
                    // The program pops the end point and calls setcurrentpoint.
                    _postScriptStack.Clear();
                    _postScriptStack.Push(_y);
                    _postScriptStack.Push(_x);
                    break;
                }
                case 1: // start flex
                    _flexing = true;
                    _flexPoints.Clear();
                    _flexPoints.Add((_x, _y));
                    break;
                case 2: break; // flex point marker
                default:
                    // Hint replacement (3) and unknown othersubrs return their arguments to pop.
                    _postScriptStack.Clear();
                    for (int index = arguments.Length - 1; index >= 0; index--) _postScriptStack.Push(arguments[index]);
                    break;
            }
        }

        private void Seac(double asb, double adx, double ady, int baseCode, int accentCode, int depth) {
            string?[] standard = PdfFontEncodingNames.Get("StandardEncoding")!;
            if (baseCode is < 0 or > 255 || accentCode is < 0 or > 255 ||
                standard[baseCode] is not string baseName || standard[accentCode] is not string accentName ||
                !_glyphs.TryGetValue(baseName, out byte[]? baseProgram) || !_glyphs.TryGetValue(accentName, out byte[]? accentProgram) ||
                depth > 0 && _offsetX != 0D) {
                _stack.Clear();
                return;
            }
            double width = _path.Width;
            bool widthSet = _widthSet;
            _stack.Clear();
            _widthSet = false;
            _x = _y = 0D;
            _offsetX = _offsetY = 0D;
            Execute(baseProgram, depth + 1);
            _stack.Clear();
            _widthSet = false;
            _x = _y = 0D;
            _offsetX = adx - asb;
            _offsetY = ady;
            Execute(accentProgram, depth + 1);
            _offsetX = _offsetY = 0D;
            _path.Width = width;
            _widthSet = widthSet;
        }

        private void SetSideBearing(double sbx, double sby, double wx) {
            _stack.Clear();
            _x = sbx;
            _y = sby;
            if (!_widthSet) {
                _path.Width = wx;
                _widthSet = true;
            }
        }

        private void RMoveTo(double dx, double dy) {
            _stack.Clear();
            _x += dx;
            _y += dy;
            if (_flexing) {
                _flexPoints.Add((_x, _y));
                return;
            }
            _path.MoveTo(_x + _offsetX, _y + _offsetY);
            _open = true;
        }

        private void RLineTo(double dx, double dy) {
            _stack.Clear();
            EnsureOpen();
            _x += dx;
            _y += dy;
            _path.LineTo(_x + _offsetX, _y + _offsetY);
        }

        private void RCurveTo(double dx1, double dy1, double dx2, double dy2, double dx3, double dy3) {
            _stack.Clear();
            EnsureOpen();
            double x1 = _x + dx1, y1 = _y + dy1;
            double x2 = x1 + dx2, y2 = y1 + dy2;
            double x3 = x2 + dx3, y3 = y2 + dy3;
            _path.CurveTo(x1 + _offsetX, y1 + _offsetY, x2 + _offsetX, y2 + _offsetY, x3 + _offsetX, y3 + _offsetY);
            _x = x3;
            _y = y3;
        }

        private void CurveAbsolute(double x1, double y1, double x2, double y2, double x3, double y3) {
            EnsureOpen();
            _path.CurveTo(x1 + _offsetX, y1 + _offsetY, x2 + _offsetX, y2 + _offsetY, x3 + _offsetX, y3 + _offsetY);
        }

        // A drawing operator after closepath starts a new contour at the current point.
        private void EnsureOpen() {
            if (_open) return;
            _path.MoveTo(_x + _offsetX, _y + _offsetY);
            _open = true;
        }

        private double Arg(int index) {
            if (index >= _stack.Count) throw new InvalidDataException("Type 1 operand stack underflow.");
            return _stack[index];
        }

        private double Pop() {
            if (_stack.Count == 0) throw new InvalidDataException("Type 1 operand stack underflow.");
            double value = _stack[_stack.Count - 1];
            _stack.RemoveAt(_stack.Count - 1);
            return value;
        }

        private void Push(double value) {
            if (_stack.Count >= 64) throw new InvalidDataException("Type 1 operand stack overflow.");
            _stack.Add(value);
        }

        private static double ReadNumber(byte[] program, ref int offset, int b0) {
            if (b0 <= 246) return b0 - 139;
            if (b0 <= 250) {
                if (offset >= program.Length) throw new InvalidDataException("Truncated Type 1 number.");
                return ((b0 - 247) * 256) + program[offset++] + 108;
            }
            if (b0 <= 254) {
                if (offset >= program.Length) throw new InvalidDataException("Truncated Type 1 number.");
                return -((b0 - 251) * 256) - program[offset++] - 108;
            }
            if (offset + 4 > program.Length) throw new InvalidDataException("Truncated Type 1 number.");
            int value = (program[offset] << 24) | (program[offset + 1] << 16) | (program[offset + 2] << 8) | program[offset + 3];
            offset += 4;
            return value;
        }
    }

    private static byte[] BuildCff(string fontName, List<string> names, List<byte[]> charStrings, double[] fontMatrix, double[] fontBBox) {
        var customStrings = new List<string>();
        var sids = new int[names.Count];
        for (int glyph = 1; glyph < names.Count; glyph++) {
            int standard = PdfCffStandardStrings.IndexOf(names[glyph]);
            if (standard >= 0) {
                sids[glyph] = standard;
            } else {
                sids[glyph] = PdfCffStandardStrings.Count + customStrings.Count;
                customStrings.Add(names[glyph]);
            }
        }
        byte[] nameIndex = BuildIndex(new[] { Encoding.ASCII.GetBytes(SanitizeName(fontName)) });
        var stringItems = new List<byte[]>(customStrings.Count);
        foreach (string value in customStrings) stringItems.Add(Encoding.ASCII.GetBytes(value));
        byte[] stringIndex = BuildIndex(stringItems);
        byte[] globalSubrs = BuildIndex(Array.Empty<byte[]>());
        using var charset = new MemoryStream();
        charset.WriteByte(0);
        for (int glyph = 1; glyph < sids.Length; glyph++) {
            charset.WriteByte((byte)(sids[glyph] >> 8));
            charset.WriteByte((byte)sids[glyph]);
        }
        byte[] charStringIndex = BuildIndex(charStrings);
        byte[] privateDict = { 139, 20 }; // defaultWidthX 0; widths are explicit in every charstring

        // The Top DICT uses fixed five-byte offsets, so its size is known before the layout.
        byte[] topDict = BuildTopDict(0, 0, 0, fontMatrix, fontBBox);
        byte[] topIndex = BuildIndex(new[] { topDict });
        int charsetOffset = 4 + nameIndex.Length + topIndex.Length + stringIndex.Length + globalSubrs.Length;
        int charStringsOffset = charsetOffset + (int)charset.Length;
        int privateOffset = charStringsOffset + charStringIndex.Length;
        topDict = BuildTopDict(charsetOffset, charStringsOffset, privateOffset, fontMatrix, fontBBox);
        topIndex = BuildIndex(new[] { topDict });

        using var output = new MemoryStream();
        output.Write(new byte[] { 1, 0, 4, 4 }, 0, 4);
        output.Write(nameIndex, 0, nameIndex.Length);
        output.Write(topIndex, 0, topIndex.Length);
        output.Write(stringIndex, 0, stringIndex.Length);
        output.Write(globalSubrs, 0, globalSubrs.Length);
        charset.Position = 0;
        charset.CopyTo(output);
        output.Write(charStringIndex, 0, charStringIndex.Length);
        output.Write(privateDict, 0, privateDict.Length);
        return output.ToArray();
    }

    private static string SanitizeName(string name) {
        var builder = new StringBuilder(Math.Min(name.Length, 63));
        foreach (char value in name) {
            if (builder.Length >= 63) break;
            if (value > ' ' && value < 127 && value is not ('[' or ']' or '(' or ')' or '{' or '}' or '<' or '>' or '/' or '%')) builder.Append(value);
        }
        return builder.Length == 0 ? "Type1" : builder.ToString();
    }

    private static byte[] BuildTopDict(int charsetOffset, int charStringsOffset, int privateOffset, double[] fontMatrix, double[] fontBBox) {
        using var output = new MemoryStream();
        bool defaultMatrix = fontMatrix[0] == 0.001D && fontMatrix[1] == 0D && fontMatrix[2] == 0D && fontMatrix[3] == 0.001D &&
            fontMatrix[4] == 0D && fontMatrix[5] == 0D;
        if (!defaultMatrix) {
            foreach (double value in fontMatrix) WriteDictReal(output, value);
            output.WriteByte(12);
            output.WriteByte(7);
        }
        foreach (double value in fontBBox) WriteDictInteger(output, (int)Math.Round(Math.Max(-32768D, Math.Min(32767D, value))));
        output.WriteByte(5);
        WriteDictOffset(output, charsetOffset);
        output.WriteByte(15);
        WriteDictOffset(output, charStringsOffset);
        output.WriteByte(17);
        WriteDictOffset(output, 2);
        WriteDictOffset(output, privateOffset);
        output.WriteByte(18);
        return output.ToArray();
    }

    private static void WriteDictOffset(MemoryStream output, int value) {
        output.WriteByte(29);
        output.WriteByte((byte)(value >> 24));
        output.WriteByte((byte)(value >> 16));
        output.WriteByte((byte)(value >> 8));
        output.WriteByte((byte)value);
    }

    private static void WriteDictInteger(MemoryStream output, int value) {
        if (value >= -107 && value <= 107) {
            output.WriteByte((byte)(value + 139));
        } else {
            WriteDictOffset(output, value);
        }
    }

    private static void WriteDictReal(MemoryStream output, double value) {
        string text = value.ToString("0.##########E+0", CultureInfo.InvariantCulture).Replace("E+0", string.Empty);
        var nibbles = new List<int>(text.Length + 2);
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (current >= '0' && current <= '9') nibbles.Add(current - '0');
            else if (current == '.') nibbles.Add(0xA);
            else if (current == '-') nibbles.Add(0xE);
            else if (current == 'E') {
                if (index + 1 < text.Length && text[index + 1] == '-') { nibbles.Add(0xC); index++; }
                else { nibbles.Add(0xB); if (index + 1 < text.Length && text[index + 1] == '+') index++; }
            }
        }
        nibbles.Add(0xF);
        if ((nibbles.Count & 1) == 1) nibbles.Add(0xF);
        output.WriteByte(30);
        for (int index = 0; index < nibbles.Count; index += 2) output.WriteByte((byte)((nibbles[index] << 4) | nibbles[index + 1]));
    }

    internal static byte[] BuildIndex(IReadOnlyList<byte[]> items) {
        using var output = new MemoryStream();
        output.WriteByte((byte)(items.Count >> 8));
        output.WriteByte((byte)items.Count);
        if (items.Count == 0) return output.ToArray();
        long total = 1;
        foreach (byte[] item in items) total += item.Length;
        int offSize = total <= 0xFF ? 1 : total <= 0xFFFF ? 2 : total <= 0xFFFFFF ? 3 : 4;
        output.WriteByte((byte)offSize);
        long offset = 1;
        WriteOffset(output, offset, offSize);
        foreach (byte[] item in items) {
            offset += item.Length;
            WriteOffset(output, offset, offSize);
        }
        foreach (byte[] item in items) output.Write(item, 0, item.Length);
        return output.ToArray();
    }

    private static void WriteOffset(MemoryStream output, long value, int size) {
        for (int shift = (size - 1) * 8; shift >= 0; shift -= 8) output.WriteByte((byte)(value >> shift));
    }
}
