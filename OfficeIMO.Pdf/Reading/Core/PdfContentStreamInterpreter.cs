using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Bounded lexical interpreter shared by text, visual, and XObject content-stream visitors.
/// It owns PDF operand parsing, operator boundaries, comments, and inline-image framing once.
/// </summary>
internal static class PdfContentStreamInterpreter {
    internal static void Interpret(
        PdfContentSource content,
        int maxOperations,
        Action<PdfContentOperation> visit,
        Func<string, int>? inlineImageComponentCount = null,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands,
        bool dispatchInvalidOperations = false,
        Func<PdfArray, int>? inlineImageArrayComponentCount = null,
        PdfContentOperands? operandBuffer = null) {
        content.ThrowIfNull(nameof(content));
        Guard.NotNull(visit, nameof(visit));
        Func<PdfContentOperation, bool> visitAll = operation => {
            visit(operation);
            return true;
        };
        if (content.Bytes is byte[] bytes) {
            new Reader<PdfByteContentText>(new PdfByteContentText(bytes), maxOperations, maxOperands, maxNestingDepth, inlineImageComponentCount, inlineImageArrayComponentCount, dispatchInvalidOperations, operandBuffer)
                .InterpretUntil(visitAll);
        } else {
            new Reader<PdfStringContentText>(new PdfStringContentText(content.Text!), maxOperations, maxOperands, maxNestingDepth, inlineImageComponentCount, inlineImageArrayComponentCount, dispatchInvalidOperations, operandBuffer)
                .InterpretUntil(visitAll);
        }
    }

    internal static bool InterpretUntil(
        PdfContentSource content,
        int maxOperations,
        Func<PdfContentOperation, bool> visit,
        Func<string, int>? inlineImageComponentCount = null,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands,
        bool dispatchInvalidOperations = false,
        Func<PdfArray, int>? inlineImageArrayComponentCount = null) {
        content.ThrowIfNull(nameof(content));
        Guard.NotNull(visit, nameof(visit));
        return content.Bytes is byte[] bytes
            ? new Reader<PdfByteContentText>(new PdfByteContentText(bytes), maxOperations, maxOperands, maxNestingDepth, inlineImageComponentCount, inlineImageArrayComponentCount, dispatchInvalidOperations)
                .InterpretUntil(visit)
            : new Reader<PdfStringContentText>(new PdfStringContentText(content.Text!), maxOperations, maxOperands, maxNestingDepth, inlineImageComponentCount, inlineImageArrayComponentCount, dispatchInvalidOperations)
                .InterpretUntil(visit);
    }

    // Operators are short Latin-1 tokens; reusing one instance per operator avoids a string per operation.
    private static readonly Dictionary<int, string> KnownOperators = CreateKnownOperators();

    private static Dictionary<int, string> CreateKnownOperators() {
        string[] operators = {
            "b", "B", "b*", "B*", "BDC", "BI", "BMC", "BT", "BX", "c", "cm", "CS", "cs", "d", "d0", "d1", "Do", "DP",
            "EI", "EMC", "ET", "EX", "f", "F", "f*", "G", "g", "gs", "h", "i", "ID", "j", "J", "K", "k", "l", "m", "M",
            "MP", "n", "q", "Q", "re", "RG", "rg", "ri", "s", "S", "SC", "sc", "SCN", "scn", "sh", "T*", "Tc", "Td",
            "TD", "Tf", "Tj", "TJ", "TL", "Tm", "Tr", "Ts", "Tw", "Tz", "v", "w", "W", "W*", "y"
        };
        var known = new Dictionary<int, string>(operators.Length);
        foreach (string op in operators) {
            if (TryGetOperatorKey(op, 0, op.Length, out int key)) known.Add(key, op);
        }
        return known;
    }

    // Packs up to three Latin-1 characters with the length. Longer tokens are not interned.
    private static bool TryGetOperatorKey(string source, int start, int length, out int key) =>
        TryGetOperatorKey(new PdfStringContentText(source), start, length, out key);

    private static bool TryGetOperatorKey<TText>(TText source, int start, int length, out int key)
        where TText : struct, IPdfContentText {
        key = 0;
        if (length is < 1 or > 3) return false;
        int packed = length << 24;
        for (int index = 0; index < length; index++) {
            char character = source[start + index];
            if (character > 0xFF) return false;
            packed |= character << (index * 8);
        }
        key = packed;
        return true;
    }

    // Generic over the content representation so decoded stream bytes are read directly. The byte
    // view behaves exactly like the Latin-1 string the reader used before: one character per byte.
    private sealed class Reader<TText> where TText : struct, IPdfContentText {
        private const int CacheSlots = 64;
        private readonly TText _content;
        private string? _contentString;
        private readonly int _maxOperations;
        private readonly int _maxOperands;
        private readonly int _maxNestingDepth;
        private readonly Func<string, int>? _inlineImageComponentCount;
        private readonly Func<PdfArray, int>? _inlineImageArrayComponentCount;
        private readonly bool _dispatchInvalidOperations;
        private readonly PdfContentOperands _operands;
        private readonly bool _borrowOperands;
        private int _index;
        private int _operationCount;
        private int _operandCount;
        private int _markedContentDepth;
        private bool _hasInvalidOperand;
        private string?[]? _rawNames;
        private string?[]? _decodedNames;
        private PdfContentOperands?[]? _arrayScratch;
        // Borrowed operands may also borrow their mixed arrays, such as TJ arrays. A synchronous
        // visitor reads them only during its callback, so the lists are reused by later operations.
        private List<PdfContentOperands>? _borrowedArrays;
        private int _borrowedArrayCount;

        internal Reader(
            TText content,
            int maxOperations,
            int maxOperands,
            int maxNestingDepth,
            Func<string, int>? inlineImageComponentCount,
            Func<PdfArray, int>? inlineImageArrayComponentCount,
            bool dispatchInvalidOperations,
            PdfContentOperands? operandBuffer = null) {
            _content = content;
            _maxOperations = maxOperations;
            _maxOperands = maxOperands;
            _maxNestingDepth = maxNestingDepth;
            _inlineImageComponentCount = inlineImageComponentCount;
            _inlineImageArrayComponentCount = inlineImageArrayComponentCount;
            _dispatchInvalidOperations = dispatchInvalidOperations;
            // An explicitly supplied buffer belongs to a synchronous visitor. Its contents
            // are valid only during that callback; ordinary visitors retain stable snapshots.
            _borrowOperands = operandBuffer != null;
            _operands = operandBuffer ?? new PdfContentOperands();
            _operands.Clear();
        }

        internal bool InterpretUntil(Func<PdfContentOperation, bool> visit) {
            while (_index < _content.Length) {
                SkipWhitespaceAndComments();
                if (_index >= _content.Length) {
                    break;
                }

                char current = _content[_index];
                if (current == ']') {
                    _index++;
                    continue;
                }

                if (TryReadOperand(0, _operands)) {
                    continue;
                }

                int operatorOffset = _index;
                string name = ReadOperator();
                if (name.Length == 0) {
                    _index++;
                    continue;
                }

                if (++_operationCount > _maxOperations) {
                    throw PdfReadLimitException.Create(
                        PdfReadLimitKind.ContentOperations,
                        _maxOperations,
                        _operationCount);
                }

                bool hasInvalidOperand = _hasInvalidOperand;
                _hasInvalidOperand = false;
                PdfContentInlineImage? inlineImage = string.Equals(name, "BI", StringComparison.Ordinal)
                    ? ReadInlineImage()
                    : null;
                bool hasInvalidInlineImageOperand = _hasInvalidOperand;
                bool invalidOperation = hasInvalidOperand || hasInvalidInlineImageOperand;
                bool skipOperation = hasInvalidInlineImageOperand && !_dispatchInvalidOperations ||
                    hasInvalidOperand && !_dispatchInvalidOperations && !CanDispatchWithInvalidOperands(name);
                if (invalidOperation) ClearOperands();
                var operation = new PdfContentOperation(
                    name,
                    _borrowOperands ? _operands : _operands.ToObjectArray(),
                    operatorOffset,
                    inlineImage,
                    invalidOperation);
                if (!_borrowOperands || skipOperation) ClearOperands();
                _hasInvalidOperand = false;
                if (skipOperation) {
                    continue;
                }

                if (name is "BDC" or "BMC") {
                    if (++_markedContentDepth > _maxNestingDepth)
                        throw PdfReadLimitException.Create(PdfReadLimitKind.ContentNestingDepth, _maxNestingDepth, _markedContentDepth);
                } else if (name == "EMC" && _markedContentDepth > 0) {
                    _markedContentDepth--;
                }

                bool keepReading;
                try {
                    keepReading = visit(operation);
                } finally {
                    if (_borrowOperands) ClearOperands();
                }
                if (!keepReading) {
                    return false;
                }
            }

            if (_dispatchInvalidOperations && (_hasInvalidOperand || _operands.Count > 0)) {
                if (++_operationCount > _maxOperations) {
                    throw PdfReadLimitException.Create(
                        PdfReadLimitKind.ContentOperations,
                        _maxOperations,
                        _operationCount);
                }
                ClearOperands();
                _hasInvalidOperand = false;
                return visit(new PdfContentOperation(
                    string.Empty,
                    Array.Empty<object>(),
                    _content.Length,
                    null,
                    hasInvalidOperands: true));
            }

            return true;
        }

        // Reads one operand into a typed list: numbers stay unboxed and strings stay in its buffer.
        private bool TryReadOperand(int nestingDepth, PdfContentOperands target) {
            if (_index >= _content.Length) {
                return false;
            }

            char current = _content[_index];
            if (current == '/') {
                target.AddObject(ReadName());
                CountOperand();
                return true;
            }

            if (current == '(') {
                ReadLiteralString(target);
                CountOperand();
                return true;
            }

            if (current == '<') {
                if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                    int childDepth = EnsureNestingBudget(nestingDepth);
                    target.AddObject(ReadDictionary(childDepth));
                } else {
                    ReadHexString(target);
                }
                CountOperand();
                return true;
            }

            if (current == '[') {
                target.AddObject(ReadArray(EnsureNestingBudget(nestingDepth), borrowMixedResult: _borrowOperands && nestingDepth == 0));
                CountOperand();
                return true;
            }

            if (IsNumberStart(current)) {
                double? number = ReadNumber();
                if (number.HasValue) {
                    target.AddNumber(number.Value);
                } else {
                    _hasInvalidOperand = true;
                }

                CountOperand();
                return true;
            }

            return false;
        }

        private bool TryReadValue(int nestingDepth, out object? value) {
            value = null;
            if (_index >= _content.Length) {
                return false;
            }

            char current = _content[_index];
            if (current == '/') {
                value = ReadName();
                CountOperand();
                return true;
            }

            if (current == '(') {
                value = ReadLiteralStringBytes();
                CountOperand();
                return true;
            }

            if (current == '<') {
                if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                    int childDepth = EnsureNestingBudget(nestingDepth);
                    value = ReadDictionary(childDepth);
                } else {
                    value = ReadHexStringBytes();
                }
                CountOperand();
                return true;
            }

            if (current == '[') {
                value = ReadArray(EnsureNestingBudget(nestingDepth), borrowMixedResult: _borrowOperands && nestingDepth == 0);
                CountOperand();
                return true;
            }

            if (IsNumberStart(current)) {
                double? number = ReadNumber();
                if (number.HasValue) {
                    value = PdfContentNumberBoxes.Box(number.Value);
                } else {
                    _hasInvalidOperand = true;
                }

                CountOperand();
                return true;
            }

            return false;
        }

        private string ReadName() {
            _index++;
            int start = _index;
            uint hash = 2166136261;
            while (_index < _content.Length && !IsDelimiter(_content[_index])) {
                hash = unchecked((hash ^ _content[_index]) * 16777619);
                _index++;
            }

            // Resource and tag names repeat throughout a stream; decode each distinct spelling once.
            int length = _index - start;
            int slot = (int)(hash % CacheSlots);
            _rawNames ??= new string?[CacheSlots];
            _decodedNames ??= new string?[CacheSlots];
            string? raw = _rawNames[slot];
            if (raw != null && raw.Length == length && EqualsAt(start, raw)) {
                return _decodedNames[slot]!;
            }

            raw = _content.Substring(start, length);
            string decoded = PdfSyntax.DecodeName(raw);
            _rawNames[slot] = raw;
            _decodedNames[slot] = decoded;
            return decoded;
        }


        private double? ReadNumber() {
            int start = _index++;
            while (_index < _content.Length) {
                char current = _content[_index];
                if (!(char.IsDigit(current) ||
                      current == '.' ||
                      current == '-' ||
                      current == '+' ||
                      current == 'e' ||
                      current == 'E')) {
                    break;
                }

                _index++;
            }

#if NET8_0_OR_GREATER
            if (TryParseSimpleDecimal(start, _index - start, out double simple)) return simple;
#endif
            bool parsed = _content.TryParseDouble(start, _index - start, out double value);
            return parsed && !double.IsNaN(value) && !double.IsInfinity(value)
                ? value
                : (double?)null;
        }

#if NET8_0_OR_GREATER
        // Plain decimals such as "12", "-0.5" or "612.000" with at most 15 significant digits and
        // 22 fractional digits. Both the digits and the power of ten are exact doubles, so one IEEE
        // division yields the correctly rounded value that double.TryParse returns on this runtime.
        private bool TryParseSimpleDecimal(int start, int length, out double value) {
            value = 0D;
            int index = start;
            int end = start + length;
            bool negative = false;
            if (index < end && (_content[index] == '-' || _content[index] == '+')) {
                negative = _content[index] == '-';
                index++;
            }
            long mantissa = 0;
            int digits = 0;
            int fractionDigits = 0;
            bool seenPoint = false;
            bool seenDigit = false;
            for (; index < end; index++) {
                char character = _content[index];
                if (character >= '0' && character <= '9') {
                    seenDigit = true;
                    if (mantissa == 0 && character == '0') {
                        if (seenPoint) fractionDigits++;
                        continue;
                    }
                    if (++digits > 15) return false;
                    mantissa = (mantissa * 10) + (character - '0');
                    if (seenPoint) fractionDigits++;
                } else if (character == '.' && !seenPoint) {
                    seenPoint = true;
                } else {
                    return false;
                }
            }
            if (!seenDigit || fractionDigits > 22) return false;
            double magnitude = fractionDigits == 0 ? mantissa : mantissa / ExactPowersOfTen[fractionDigits];
            value = negative ? -magnitude : magnitude;
            return true;
        }

        private static readonly double[] ExactPowersOfTen = {
            1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11,
            1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22
        };
#endif

        private byte[] ReadLiteralStringBytes() {
            int start = ScanLiteralString(out int length);
            return ParseLiteralBytes(start, length);
        }

        private void ReadLiteralString(PdfContentOperands target) {
            int start = ScanLiteralString(out int length);
            byte[] buffer = target.BeginString(length, out int offset);
            target.CommitString(offset, DecodeLiteralInto(start, length, buffer, offset));
        }

        private int ScanLiteralString(out int length) {
            int start = ++_index;
            int depth = 1;
            bool escaped = false;
            while (_index < _content.Length && depth > 0) {
                char current = _content[_index++];
                if (escaped) {
                    escaped = false;
                } else if (current == '\\') {
                    escaped = true;
                } else if (current == '(') {
                    depth++;
                } else if (current == ')') {
                    depth--;
                }
            }

            if (depth != 0 || escaped) _hasInvalidOperand = true;

            int end = _index - (depth == 0 || escaped ? 1 : 0);
            length = end - start;
            return start;
        }

        private byte[] ReadHexStringBytes() {
            int start = ScanHexString(out int length);
            return DecodeHexBytes(start, length);
        }

        private void ReadHexString(PdfContentOperands target) {
            int start = ScanHexString(out int length);
            byte[] buffer = target.BeginString((length + 1) / 2, out int offset);
            target.CommitString(offset, DecodeHexInto(start, length, buffer, offset));
        }

        private int ScanHexString(out int length) {
            _index++;
            int start = _index;
            while (_index < _content.Length && _content[_index] != '>') {
                _index++;
            }

            length = _index - start;
            if (_index < _content.Length) {
                _index++;
            } else {
                _hasInvalidOperand = true;
            }

            return start;
        }

        // Same decoding as PdfStringParser.ParseLiteralToBytes over the reader's content.
        private byte[] ParseLiteralBytes(int start, int length) {
            if (length == 0) return Array.Empty<byte>();
            var bytes = new byte[length];
            int count = DecodeLiteralInto(start, length, bytes, 0);
            if (count == bytes.Length) return bytes;
            var result = new byte[count];
            Buffer.BlockCopy(bytes, 0, result, 0, count);
            return result;
        }

        // Escapes only remove characters, so the decoded bytes never exceed the source length.
        private int DecodeLiteralInto(int start, int length, byte[] bytes, int offset) {
            int count = offset;
            int end = start + length;
            for (int i = start; i < end; i++) {
                char c = _content[i];
                if (c == '\\') {
                    if (i + 1 >= end) break;
                    char n = _content[++i];
                    switch (n) {
                        case 'n': bytes[count++] = (byte)'\n'; break;
                        case 'r': bytes[count++] = (byte)'\r'; break;
                        case 't': bytes[count++] = (byte)'\t'; break;
                        case 'b': bytes[count++] = (byte)'\b'; break;
                        case 'f': bytes[count++] = (byte)'\f'; break;
                        case '\\': bytes[count++] = (byte)'\\'; break;
                        case '(': bytes[count++] = (byte)'('; break;
                        case ')': bytes[count++] = (byte)')'; break;
                        case '\n': break;
                        case '\r':
                            if (i + 1 < end && _content[i + 1] == '\n') i++;
                            break;
                        default:
                            if (IsOctalDigit(n)) {
                                int v = n - '0';
                                for (int k = 0; k < 2 && i + 1 < end && IsOctalDigit(_content[i + 1]); k++) {
                                    v = (v << 3) + (_content[++i] - '0');
                                }
                                bytes[count++] = (byte)(v & 0xFF);
                            } else {
                                bytes[count++] = (byte)(n & 0xFF);
                            }
                            break;
                    }
                } else {
                    bytes[count++] = (byte)(c & 0xFF);
                }
            }

            return count - offset;
        }

        private static bool IsOctalDigit(char c) => c >= '0' && c <= '7';

        // Same decoding as PdfTextString.DecodeHexBytes over the reader's content.
        private byte[] DecodeHexBytes(int start, int length) {
            int end = start + length;
            int nibbleCount = 0;
            for (int i = start; i < end; i++) {
                if (!char.IsWhiteSpace(_content[i])) nibbleCount++;
            }

            if (nibbleCount == 0) return Array.Empty<byte>();
            var bytes = new byte[(nibbleCount + 1) / 2];
            int highNibble = -1;
            int byteIndex = 0;
            for (int i = start; i < end; i++) {
                char ch = _content[i];
                if (char.IsWhiteSpace(ch)) continue;
                int nibble = HexNibble(ch);
                if (highNibble < 0) {
                    highNibble = nibble;
                } else {
                    bytes[byteIndex++] = (byte)((highNibble << 4) | nibble);
                    highNibble = -1;
                }
            }

            if (highNibble >= 0) bytes[byteIndex] = (byte)(highNibble << 4);
            return bytes;
        }

        // One pass of DecodeHexBytes: the first invalid character throws the same exception.
        private int DecodeHexInto(int start, int length, byte[] bytes, int offset) {
            int end = start + length;
            int highNibble = -1;
            int byteIndex = offset;
            for (int i = start; i < end; i++) {
                char ch = _content[i];
                if (char.IsWhiteSpace(ch)) continue;
                int nibble = HexNibble(ch);
                if (highNibble < 0) {
                    highNibble = nibble;
                } else {
                    bytes[byteIndex++] = (byte)((highNibble << 4) | nibble);
                    highNibble = -1;
                }
            }

            if (highNibble >= 0) bytes[byteIndex++] = (byte)(highNibble << 4);
            return byteIndex - offset;
        }

        private static int HexNibble(char c) {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
            if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
            throw new FormatException($"Invalid hex character '{c}'.");
        }

        private bool EqualsAt(int start, string value) {
            for (int index = 0; index < value.Length; index++) {
                if (_content[start + index] != value[index]) return false;
            }
            return true;
        }

        // Rare paths (inline image data scanning) still read a Latin-1 string; byte content builds it once.
        private string ContentString => _contentString ??= _content.Substring(0, _content.Length);

        private void ClearOperands() {
            _operands.Clear();
            for (int index = 0; index < _borrowedArrayCount; index++) {
                _borrowedArrays![index].Clear();
            }
            _borrowedArrayCount = 0;
        }

        private object ReadArray(int nestingDepth, bool borrowMixedResult = false) {
            _index++;
            // Each nesting level collects into its own reusable typed list; the result is sized exactly.
            if (_arrayScratch == null || _arrayScratch.Length <= nestingDepth) {
                Array.Resize(ref _arrayScratch, Math.Max(4, nestingDepth + 1));
            }
            PdfContentOperands values = _arrayScratch[nestingDepth] ??= new PdfContentOperands();
            values.Clear();
            bool terminated = false;
            while (_index < _content.Length) {
                SkipWhitespaceAndComments();
                if (_index >= _content.Length) {
                    break;
                }

                if (_content[_index] == ']') {
                    _index++;
                    terminated = true;
                    break;
                }

                if (!TryReadOperand(nestingDepth, values)) {
                    string token = ReadOperator();
                    if (string.Equals(token, "true", StringComparison.Ordinal)) {
                        values.AddObject(true);
                        CountOperand();
                    } else if (string.Equals(token, "false", StringComparison.Ordinal)) {
                        values.AddObject(false);
                        CountOperand();
                    } else if (string.Equals(token, "R", StringComparison.Ordinal) &&
                        values.Count >= 2 &&
                        TryReadReferencePart(values, values.Count - 2, out int objectNumber) &&
                        TryReadReferencePart(values, values.Count - 1, out int generation)) {
                        values.RemoveLast();
                        values.RemoveLast();
                        values.AddObject(new PdfReference(objectNumber, generation));
                    } else if (token.Length == 0) {
                        _index++;
                    }
                }
            }
            if (!terminated) _hasInvalidOperand = true;

            object result;
            if (values.AllNumbers()) {
                result = values.ToNumberArray();
            } else if (borrowMixedResult) {
                _borrowedArrays ??= new List<PdfContentOperands>();
                if (_borrowedArrayCount == _borrowedArrays.Count) _borrowedArrays.Add(new PdfContentOperands(values.Count));
                PdfContentOperands borrowed = _borrowedArrays[_borrowedArrayCount++];
                borrowed.CopyFrom(values);
                result = borrowed;
            } else {
                result = values.ToList();
            }
            values.Clear();
            return result;
        }

        private static bool TryReadReferencePart(PdfContentOperands values, int index, out int result) {
            result = 0;
            if (!values.TryGetNumber(index, out double number) ||
                number < 0D ||
                number > int.MaxValue ||
                number != Math.Truncate(number)) {
                return false;
            }

            result = (int)number;
            return true;
        }

        private PdfContentDictionary ReadDictionary(int nestingDepth) {
            int dictionaryStart = _index;
            _index += 2;
            var dictionary = new PdfContentDictionary();
            bool terminated = false;
            while (_index < _content.Length) {
                SkipWhitespaceAndComments();
                if (IsAt(">>")) {
                    _index += 2;
                    terminated = true;
                    break;
                }

                if (_index >= _content.Length) {
                    break;
                }

                if (_content[_index] != '/') {
                    SkipOneValue(nestingDepth);
                    continue;
                }

                string key = ReadName();
                CountOperand();
                SkipWhitespaceAndComments();
                if (TryReadValue(nestingDepth, out object? value)) {
                    if (value is not null) {
                        dictionary.Items[key] = value;
                    }
                } else {
                    string token = ReadOperator();
                    if (string.Equals(token, "true", StringComparison.Ordinal)) {
                        dictionary.Items[key] = true;
                        CountOperand();
                    } else if (string.Equals(token, "false", StringComparison.Ordinal)) {
                        dictionary.Items[key] = false;
                        CountOperand();
                    }
                }
            }
            if (!terminated) _hasInvalidOperand = true;

            dictionary.SetOptionalContentSource(_content.Source, dictionaryStart, Math.Max(0, _index - dictionaryStart));
            return dictionary;
        }

        private void SkipOneValue(int nestingDepth = 0) {
            if (TryReadValue(nestingDepth, out _)) {
                return;
            }

            string token = ReadOperator();
            if (token.Length == 0 && _index < _content.Length) {
                _index++;
            }
        }

        private string ReadOperator() {
            if (_index >= _content.Length) {
                return string.Empty;
            }

            int start = _index++;
            char first = _content[start];
            if (first == '\'' || first == '"') {
                return first == '\'' ? "'" : "\"";
            }

            while (_index < _content.Length && !IsDelimiter(_content[_index])) {
                _index++;
            }

            int length = _index - start;
            return TryGetOperatorKey(_content, start, length, out int key) && KnownOperators.TryGetValue(key, out string? known)
                ? known
                : _content.Substring(start, length);
        }

        private PdfContentInlineImage? ReadInlineImage() {
            var dictionary = new PdfDictionary();
            while (_index < _content.Length) {
                SkipWhitespaceAndComments();
                if (_index >= _content.Length) {
                    return null;
                }

                if (IsOperatorAt("ID")) {
                    _index += 2;
                    break;
                }

                if (_content[_index] != '/') {
                    SkipOneValue();
                    continue;
                }

                string key = NormalizeInlineImageKey(ReadName());
                CountOperand();
                SkipWhitespaceAndComments();
                if (TryReadInlineImageValue(key, out PdfObject? value) && value is not null) {
                    dictionary.Items[key] = value;
                }
            }

            if (_index < _content.Length && _content[_index] == '\r') {
                _index += _index + 1 < _content.Length && _content[_index + 1] == '\n' ? 2 : 1;
            } else if (_index < _content.Length && char.IsWhiteSpace(_content[_index])) {
                _index++;
            }

            int dataStart = _index;
            int dataLength = TryGetRawInlineImageLength(dictionary, out int rawLength)
                ? rawLength
                : PdfInlineImageDataScanner.FindLength(ContentString, dataStart);
            if (dataLength < 0 || dataStart + dataLength > _content.Length) {
                _index = _content.Length;
                return null;
            }

            byte[] data = ReadBytes(dataStart, dataLength);
            _index = dataStart + dataLength;
            SkipWhitespaceAndComments();
            if (IsOperatorAt("EI")) {
                _index += 2;
            } else {
                _hasInvalidOperand = true;
            }

            return new PdfContentInlineImage(dictionary, data);
        }

        private bool TryReadInlineImageValue(string key, out PdfObject? value) {
            value = null;
            if (_index >= _content.Length) {
                return false;
            }

            char current = _content[_index];
            if (current == '/') {
                value = new PdfName(NormalizeInlineImageName(ReadName(), key));
                CountOperand();
                return true;
            }

            if (IsNumberStart(current)) {
                double? number = ReadNumber();
                if (number.HasValue) {
                    value = new PdfNumber(number.Value);
                } else {
                    _hasInvalidOperand = true;
                }

                CountOperand();
                return true;
            }

            if (current == '[') {
                object arrayValue = ReadArray(EnsureNestingBudget(0));
                var array = new PdfArray();
                IEnumerable<object> items = arrayValue is double[] numbers
                    ? numbers.Cast<object>()
                    : (IEnumerable<object>)arrayValue;
                foreach (object item in items) {
                    PdfObject? converted = ConvertToPdfObject(item, key);
                    if (converted is not null) {
                        array.Items.Add(converted);
                    }
                }

                value = array;
                CountOperand();
                return true;
            }

            if (current == '<') {
                if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                    PdfContentDictionary contentDictionary = ReadDictionary(EnsureNestingBudget(0));
                    value = ConvertDictionary(contentDictionary);
                } else {
                    value = new PdfStringObj(ReadHexStringBytes());
                }

                CountOperand();
                return true;
            }

            string token = ReadOperator();
            if (string.Equals(token, "true", StringComparison.Ordinal)) {
                value = new PdfBoolean(true);
                CountOperand();
                return true;
            }

            if (string.Equals(token, "false", StringComparison.Ordinal)) {
                value = new PdfBoolean(false);
                CountOperand();
                return true;
            }

            return false;
        }

        private void CountOperand() {
            int observedCount = ++_operandCount;
            if (observedCount > _maxOperands) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.ContentOperands,
                    _maxOperands,
                    observedCount);
            }
        }

        private int EnsureNestingBudget(int currentDepth) {
            int observedDepth = currentDepth + 1;
            if (observedDepth > _maxNestingDepth) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.ContentNestingDepth,
                    _maxNestingDepth,
                    observedDepth);
            }

            return observedDepth;
        }

        private static PdfDictionary ConvertDictionary(PdfContentDictionary contentDictionary) {
            var dictionary = new PdfDictionary();
            foreach (KeyValuePair<string, object> item in contentDictionary.Items) {
                PdfObject? value = ConvertToPdfObject(item.Value);
                if (value is not null) {
                    dictionary.Items[item.Key] = value;
                }
            }

            return dictionary;
        }

        private static PdfObject? ConvertToPdfObject(object value, string? ownerKey = null) {
            if (value is PdfObject pdfObject) {
                return pdfObject;
            }

            if (value is double number) {
                return new PdfNumber(number);
            }

            if (value is string name) {
                return new PdfName(NormalizeInlineImageName(name, ownerKey));
            }

            if (value is bool boolean) {
                return new PdfBoolean(boolean);
            }

            if (value is byte[] bytes) {
                return new PdfStringObj(bytes);
            }

            if (value is PdfContentDictionary dictionary) {
                return ConvertDictionary(dictionary);
            }

            if (value is double[] numbers) {
                var array = new PdfArray();
                foreach (double item in numbers) {
                    array.Items.Add(new PdfNumber(item));
                }

                return array;
            }

            if (value is List<object> values) {
                var array = new PdfArray();
                foreach (object item in values) {
                    PdfObject? converted = ConvertToPdfObject(item, ownerKey);
                    if (converted is not null) {
                        array.Items.Add(converted);
                    }
                }

                return array;
            }

            return null;
        }

        private bool TryGetRawInlineImageLength(PdfDictionary dictionary, out int length) {
            length = 0;
            if (dictionary.Items.ContainsKey("Filter")) {
                return false;
            }

            int width = ReadPositiveInteger(dictionary, "Width");
            int height = ReadPositiveInteger(dictionary, "Height");
            if (width <= 0 || height <= 0) {
                return false;
            }

            int bitsPerComponent = ReadPositiveInteger(dictionary, "BitsPerComponent");
            bool imageMask = dictionary.Items.TryGetValue("ImageMask", out PdfObject? maskObject) &&
                maskObject is PdfBoolean mask &&
                mask.Value;
            if (imageMask && bitsPerComponent == 0) {
                bitsPerComponent = 1;
            }

            int components = imageMask ? 1 : GetInlineImageComponentCount(dictionary);
            if (bitsPerComponent <= 0 || components <= 0) {
                return false;
            }

            long rowBytes = (((long)width * components * bitsPerComponent) + 7L) / 8L;
            long byteCount = rowBytes * height;
            if (byteCount <= 0L || byteCount > int.MaxValue) {
                return false;
            }

            length = (int)byteCount;
            return true;
        }

        private int GetInlineImageComponentCount(PdfDictionary dictionary) {
            if (!dictionary.Items.TryGetValue("ColorSpace", out PdfObject? value)) {
                return 1;
            }
            if (value is PdfArray array) {
                return GetInlineImageArrayComponentCount(array);
            }
            if (value is not PdfName name) {
                return 0;
            }
            string colorSpace = name.Name;
            switch (colorSpace) {
                case "DeviceRGB":
                    return 3;
                case "DeviceCMYK":
                    return 4;
                default:
                    return _inlineImageComponentCount == null
                        ? 1
                        : Math.Max(0, _inlineImageComponentCount(colorSpace));
            }
        }

        private int GetInlineImageArrayComponentCount(PdfArray array) {
            if (array.Items.Count == 0 || array.Items[0] is not PdfName kind) {
                return 0;
            }

            switch (kind.Name) {
                case "DeviceGray":
                case "CalGray":
                case "Indexed":
                case "Separation":
                    return 1;
                case "DeviceRGB":
                case "CalRGB":
                case "Lab":
                    return 3;
                case "DeviceCMYK":
                    return 4;
                case "DeviceN":
                case "NChannel":
                    return array.Items.Count > 1 && array.Items[1] is PdfArray colorants && colorants.Items.Count > 0
                        ? colorants.Items.Count
                        : 0;
                default:
                    return Math.Max(0, _inlineImageArrayComponentCount?.Invoke(array) ?? 0);
            }
        }
        private static int ReadPositiveInteger(PdfDictionary dictionary, string key) =>
            dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            value is PdfNumber number &&
            number.Value > 0D &&
            number.Value <= int.MaxValue
                ? (int)number.Value
                : 0;

        private byte[] ReadBytes(int start, int length) {
            var bytes = new byte[length];
            for (int i = 0; i < length; i++) {
                bytes[i] = (byte)_content[start + i];
            }

            return bytes;
        }

        private void SkipWhitespaceAndComments() {
            while (_index < _content.Length) {
                if (char.IsWhiteSpace(_content[_index])) {
                    _index++;
                    continue;
                }

                if (_content[_index] != '%') {
                    return;
                }

                while (_index < _content.Length &&
                       _content[_index] != '\r' &&
                       _content[_index] != '\n') {
                    _index++;
                }
            }
        }

        private bool IsOperatorAt(string value) =>
            IsAt(value) &&
            (_index == 0 || IsDelimiter(_content[_index - 1])) &&
            (_index + value.Length >= _content.Length || IsDelimiter(_content[_index + value.Length]));

        private bool IsAt(string value) =>
            _index + value.Length <= _content.Length &&
            EqualsAt(_index, value);

        private static bool IsNumberStart(char value) =>
            value == '-' || value == '+' || value == '.' || char.IsDigit(value);

        private static bool CanDispatchWithInvalidOperands(string value) {
            switch (value) {
                case "q":
                case "Q":
                case "h":
                case "W":
                case "W*":
                case "S":
                case "s":
                case "f":
                case "F":
                case "f*":
                case "B":
                case "B*":
                case "b":
                case "b*":
                case "n":
                case "BT":
                case "ET":
                case "w":
                case "J":
                case "j":
                case "Tc":
                case "Tw":
                case "Tz":
                case "TL":
                case "Tf":
                case "Tr":
                case "Ts":
                case "Td":
                case "TD":
                case "Tm":
                case "T*":
                case "d0":
                case "d1":
                case "BMC":
                case "BDC":
                case "EMC":
                case "BX":
                case "EX":
                case "BI":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsDelimiter(char value) =>
            char.IsWhiteSpace(value) ||
            value == '/' ||
            value == '[' ||
            value == ']' ||
            value == '(' ||
            value == ')' ||
            value == '<' ||
            value == '>' ||
            value == '%';

        private static string NormalizeInlineImageKey(string key) {
            switch (key) {
                case "W": return "Width";
                case "H": return "Height";
                case "BPC": return "BitsPerComponent";
                case "CS": return "ColorSpace";
                case "F": return "Filter";
                case "D": return "Decode";
                case "DP": return "DecodeParms";
                case "IM": return "ImageMask";
                default: return key;
            }
        }

        private static string NormalizeInlineImageName(string name, string? ownerKey) {
            if (string.Equals(ownerKey, "ColorSpace", StringComparison.Ordinal)) {
                switch (name) {
                    case "G": return "DeviceGray";
                    case "RGB": return "DeviceRGB";
                    case "CMYK": return "DeviceCMYK";
                    case "I": return "Indexed";
                }
            }

            if (string.Equals(ownerKey, "Filter", StringComparison.Ordinal)) {
                switch (name) {
                    case "Fl": return "FlateDecode";
                    case "AHx": return "ASCIIHexDecode";
                    case "A85": return "ASCII85Decode";
                    case "RL": return "RunLengthDecode";
                    case "DCT": return "DCTDecode";
                }
            }

            return name;
        }
    }
}

internal readonly struct PdfContentOperation {
    internal PdfContentOperation(
        string name,
        IReadOnlyList<object> operands,
        int operatorOffset,
        PdfContentInlineImage? inlineImage,
        bool hasInvalidOperands) {
        Name = name;
        Operands = operands;
        OperatorOffset = operatorOffset;
        InlineImage = inlineImage;
        HasInvalidOperands = hasInvalidOperands;
    }

    internal string Name { get; }
    internal IReadOnlyList<object> Operands { get; }
    internal int OperatorOffset { get; }
    internal PdfContentInlineImage? InlineImage { get; }
    internal bool HasInvalidOperands { get; }
}

internal sealed class PdfContentDictionary {
    private object? _optionalContentSource;
    private int _optionalContentStart;
    private int _optionalContentLength;
    private PdfInlineOptionalContentReferences? _optionalContentReferences;

    internal Dictionary<string, object> Items { get; } = new Dictionary<string, object>(StringComparer.Ordinal);

    // Only /OC marked content reads these references, so the dictionary text is scanned on first use.
    internal PdfInlineOptionalContentReferences? OptionalContentReferences {
        get {
            if (_optionalContentReferences is null && _optionalContentSource is not null) {
                _optionalContentReferences = _optionalContentSource is string text
                    ? PdfInlineOptionalContentReferenceParser.Parse(text, _optionalContentStart, _optionalContentLength)
                    // The parser reads only inside the dictionary window, so byte content passes that slice.
                    : PdfInlineOptionalContentReferenceParser.Parse(
                        PdfEncoding.Latin1GetString((byte[])_optionalContentSource, _optionalContentStart, _optionalContentLength),
                        0,
                        _optionalContentLength);
            }
            return _optionalContentReferences;
        }
    }

    internal void SetOptionalContentSource(object source, int start, int length) {
        _optionalContentSource = source;
        _optionalContentStart = start;
        _optionalContentLength = length;
        _optionalContentReferences = null;
    }
}

internal sealed class PdfContentInlineImage {
    internal PdfContentInlineImage(PdfDictionary dictionary, byte[] data) {
        Dictionary = dictionary;
        Data = data;
    }

    internal PdfDictionary Dictionary { get; }
    internal byte[] Data { get; }
}
