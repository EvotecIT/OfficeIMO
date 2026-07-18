using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Bounded lexical interpreter shared by text, visual, and XObject content-stream visitors.
/// It owns PDF operand parsing, operator boundaries, comments, and inline-image framing once.
/// </summary>
internal static class PdfContentStreamInterpreter {
    internal static void Interpret(
        string content,
        int maxOperations,
        Action<PdfContentOperation> visit,
        Func<string, int>? inlineImageComponentCount = null,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNull(visit, nameof(visit));
        var reader = new Reader(content, maxOperations, maxOperands, maxNestingDepth, inlineImageComponentCount);
        reader.InterpretUntil(operation => {
            visit(operation);
            return true;
        });
    }

    internal static bool InterpretUntil(
        string content,
        int maxOperations,
        Func<PdfContentOperation, bool> visit,
        Func<string, int>? inlineImageComponentCount = null,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNull(visit, nameof(visit));
        var reader = new Reader(content, maxOperations, maxOperands, maxNestingDepth, inlineImageComponentCount);
        return reader.InterpretUntil(visit);
    }

    private sealed class Reader {
        private readonly string _content;
        private readonly int _maxOperations;
        private readonly int _maxOperands;
        private readonly int _maxNestingDepth;
        private readonly Func<string, int>? _inlineImageComponentCount;
        private readonly List<object> _operands = new List<object>(8);
        private int _index;
        private int _operationCount;
        private int _operandCount;

        internal Reader(
            string content,
            int maxOperations,
            int maxOperands,
            int maxNestingDepth,
            Func<string, int>? inlineImageComponentCount) {
            _content = content;
            _maxOperations = maxOperations;
            _maxOperands = maxOperands;
            _maxNestingDepth = maxNestingDepth;
            _inlineImageComponentCount = inlineImageComponentCount;
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

                if (TryReadValue(0, out object? value)) {
                    if (value is not null) {
                        _operands.Add(value);
                    }

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

                PdfContentInlineImage? inlineImage = string.Equals(name, "BI", StringComparison.Ordinal)
                    ? ReadInlineImage()
                    : null;
                var operation = new PdfContentOperation(
                    name,
                    _operands.Count == 0 ? Array.Empty<object>() : _operands.ToArray(),
                    operatorOffset,
                    inlineImage);
                _operands.Clear();
                if (!visit(operation)) {
                    return false;
                }
            }

            return true;
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
                value = ReadArray(EnsureNestingBudget(nestingDepth));
                CountOperand();
                return true;
            }

            if (IsNumberStart(current)) {
                value = ReadNumber();
                CountOperand();
                return true;
            }

            return false;
        }

        private string ReadName() {
            _index++;
            int start = _index;
            while (_index < _content.Length && !IsDelimiter(_content[_index])) {
                _index++;
            }

            return PdfSyntax.DecodeName(_content.Substring(start, _index - start));
        }

        private double ReadNumber() {
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

#pragma warning disable CA1846 // Keep netstandard2.0-safe parsing.
            return double.TryParse(
                _content.Substring(start, _index - start),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out double value)
#pragma warning restore CA1846
                ? value
                : 0D;
        }

        private byte[] ReadLiteralStringBytes() {
            _index++;
            int depth = 1;
            bool escaped = false;
            var value = new StringBuilder();
            while (_index < _content.Length && depth > 0) {
                char current = _content[_index++];
                if (escaped) {
                    value.Append('\\');
                    value.Append(current);
                    escaped = false;
                } else if (current == '\\') {
                    escaped = true;
                } else if (current == '(') {
                    depth++;
                    value.Append(current);
                } else if (current == ')') {
                    depth--;
                    if (depth > 0) {
                        value.Append(current);
                    }
                } else {
                    value.Append(current);
                }
            }

            return PdfStringParser.ParseLiteralToBytes(value.ToString());
        }

        private byte[] ReadHexStringBytes() {
            _index++;
            int start = _index;
            while (_index < _content.Length && _content[_index] != '>') {
                _index++;
            }

            string hex = _content.Substring(start, _index - start);
            if (_index < _content.Length) {
                _index++;
            }

            return PdfTextString.DecodeHexBytes(hex);
        }

        private object ReadArray(int nestingDepth) {
            _index++;
            var values = new List<object>();
            while (_index < _content.Length) {
                SkipWhitespaceAndComments();
                if (_index >= _content.Length) {
                    break;
                }

                if (_content[_index] == ']') {
                    _index++;
                    break;
                }

                if (TryReadValue(nestingDepth, out object? value)) {
                    if (value is not null) {
                        values.Add(value);
                    }
                } else {
                    string token = ReadOperator();
                    if (string.Equals(token, "true", StringComparison.Ordinal)) {
                        values.Add(true);
                        CountOperand();
                    } else if (string.Equals(token, "false", StringComparison.Ordinal)) {
                        values.Add(false);
                        CountOperand();
                    } else if (token.Length == 0) {
                        _index++;
                    }
                }
            }

            if (values.All(value => value is double)) {
                return values.Cast<double>().ToArray();
            }

            return values;
        }

        private PdfContentDictionary ReadDictionary(int nestingDepth) {
            int dictionaryStart = _index;
            _index += 2;
            var dictionary = new PdfContentDictionary();
            while (_index < _content.Length) {
                SkipWhitespaceAndComments();
                if (IsAt(">>")) {
                    _index += 2;
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
                if (TryReadValue(nestingDepth, out object? value) && value is not null) {
                    dictionary.Items[key] = value;
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

            dictionary.OptionalContentReferences = PdfInlineOptionalContentReferenceParser.Parse(
                _content,
                dictionaryStart,
                Math.Max(0, _index - dictionaryStart));
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
                return first.ToString();
            }

            while (_index < _content.Length && !IsDelimiter(_content[_index])) {
                _index++;
            }

            return _content.Substring(start, _index - start);
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
                if (TryReadInlineImageValue(out PdfObject? value) && value is not null) {
                    dictionary.Items[key] = value;
                }
            }

            if (_index < _content.Length && char.IsWhiteSpace(_content[_index])) {
                _index++;
            }

            int dataStart = _index;
            int dataLength = TryGetRawInlineImageLength(dictionary, out int rawLength)
                ? rawLength
                : PdfInlineImageDataScanner.FindLength(_content, dataStart);
            if (dataLength < 0 || dataStart + dataLength > _content.Length) {
                _index = _content.Length;
                return null;
            }

            byte[] data = ReadBytes(dataStart, dataLength);
            _index = dataStart + dataLength;
            SkipWhitespaceAndComments();
            if (IsOperatorAt("EI")) {
                _index += 2;
            }

            return new PdfContentInlineImage(dictionary, data);
        }

        private bool TryReadInlineImageValue(out PdfObject? value) {
            value = null;
            if (_index >= _content.Length) {
                return false;
            }

            char current = _content[_index];
            if (current == '/') {
                value = new PdfName(NormalizeInlineImageName(ReadName()));
                CountOperand();
                return true;
            }

            if (IsNumberStart(current)) {
                value = new PdfNumber(ReadNumber());
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
                    PdfObject? converted = ConvertToPdfObject(item);
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

        private static PdfObject? ConvertToPdfObject(object value) {
            if (value is double number) {
                return new PdfNumber(number);
            }

            if (value is string name) {
                return new PdfName(NormalizeInlineImageName(name));
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
                    PdfObject? converted = ConvertToPdfObject(item);
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
            string colorSpace = dictionary.Items.TryGetValue("ColorSpace", out PdfObject? value) &&
                                value is PdfName name
                ? name.Name
                : "DeviceGray";
            switch (colorSpace) {
                case "DeviceRGB":
                    return 3;
                case "DeviceCMYK":
                    return 4;
                default:
                    return Math.Max(1, _inlineImageComponentCount?.Invoke(colorSpace) ?? 1);
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
            string.CompareOrdinal(_content, _index, value, 0, value.Length) == 0;

        private static bool IsNumberStart(char value) =>
            value == '-' || value == '+' || value == '.' || char.IsDigit(value);

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

        private static string NormalizeInlineImageName(string name) {
            switch (name) {
                case "G": return "DeviceGray";
                case "RGB": return "DeviceRGB";
                case "CMYK": return "DeviceCMYK";
                case "I": return "Indexed";
                case "Fl": return "FlateDecode";
                case "AHx": return "ASCIIHexDecode";
                case "A85": return "ASCII85Decode";
                case "RL": return "RunLengthDecode";
                default: return name;
            }
        }
    }
}

internal readonly struct PdfContentOperation {
    internal PdfContentOperation(
        string name,
        IReadOnlyList<object> operands,
        int operatorOffset,
        PdfContentInlineImage? inlineImage) {
        Name = name;
        Operands = operands;
        OperatorOffset = operatorOffset;
        InlineImage = inlineImage;
    }

    internal string Name { get; }
    internal IReadOnlyList<object> Operands { get; }
    internal int OperatorOffset { get; }
    internal PdfContentInlineImage? InlineImage { get; }
}

internal sealed class PdfContentDictionary {
    internal Dictionary<string, object> Items { get; } = new Dictionary<string, object>(StringComparer.Ordinal);
    internal PdfInlineOptionalContentReferences? OptionalContentReferences { get; set; }
}

internal sealed class PdfContentInlineImage {
    internal PdfContentInlineImage(PdfDictionary dictionary, byte[] data) {
        Dictionary = dictionary;
        Data = data;
    }

    internal PdfDictionary Dictionary { get; }
    internal byte[] Data { get; }
}
