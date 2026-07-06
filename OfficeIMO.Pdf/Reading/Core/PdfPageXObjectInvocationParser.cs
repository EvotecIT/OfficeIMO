using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfPageXObjectInvocationParser {
    public static IReadOnlyList<PdfPageXObjectInvocation> Parse(string content, Matrix2D baseTransform, double pageHeight) {
        return Parse(content, baseTransform, pageHeight, null);
    }

    public static IReadOnlyList<PdfPageXObjectInvocation> Parse(string content, Matrix2D baseTransform, double pageHeight, IReadOnlyDictionary<string, PdfPageColorSpaceKind>? colorSpaces) {
        return Parse(content, baseTransform, pageHeight, null, colorSpaces);
    }

    public static IReadOnlyList<PdfPageXObjectInvocation> Parse(
        string content,
        Matrix2D baseTransform,
        double pageHeight,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        IReadOnlyDictionary<string, PdfPageColorSpaceKind>? colorSpaces,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpaceKind initialFillColorSpace = PdfPageColorSpaceKind.DeviceGray,
        double? initialFillOpacity = null,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null) {
        if (string.IsNullOrEmpty(content)) {
            return Array.Empty<PdfPageXObjectInvocation>();
        }

        var parser = new Parser(content, baseTransform, pageHeight, graphicsStates, colorSpaces, optionalContentVisibility, initialFillColor, initialFillColorSpace, initialFillOpacity, paintOrderBase, paintOrderScale, paintOrderOffset, initialClipPath);
        return parser.Parse();
    }

    private sealed class Parser {
        private readonly string _content;
        private readonly double _pageHeight;
        private readonly Matrix2D _baseTransform;
        private readonly IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? _graphicsStates;
        private readonly IReadOnlyDictionary<string, PdfPageColorSpaceKind>? _colorSpaces;
        private readonly PdfPageOptionalContentVisibility? _optionalContentVisibility;
        private readonly double _paintOrderBase;
        private readonly double _paintOrderScale;
        private readonly double _paintOrderOffset;
        private readonly List<PdfPageXObjectInvocation> _invocations = new List<PdfPageXObjectInvocation>();
        private readonly List<object> _args = new List<object>(8);
        private readonly Stack<GraphicsState> _stack = new Stack<GraphicsState>();
        private readonly Stack<bool> _hiddenContentStack = new Stack<bool>();
        private readonly List<(double X, double Y)> _path = new List<(double X, double Y)>();
        private readonly List<OfficePathCommand> _pathCommands = new List<OfficePathCommand>();
        private readonly GraphicsState _initialState;
        private GraphicsState _state;
        private int _currentSubpathStartIndex = -1;
        private int _index;
        private int _inlineImageIndex;

        public Parser(
            string content,
            Matrix2D baseTransform,
            double pageHeight,
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
            IReadOnlyDictionary<string, PdfPageColorSpaceKind>? colorSpaces,
            PdfPageOptionalContentVisibility? optionalContentVisibility,
            OfficeColor? initialFillColor,
            PdfPageColorSpaceKind initialFillColorSpace,
            double? initialFillOpacity,
            double paintOrderBase,
            double paintOrderScale,
            double paintOrderOffset,
            PdfPageClipPath? initialClipPath) {
            _content = content;
            _baseTransform = baseTransform;
            _graphicsStates = graphicsStates;
            _colorSpaces = colorSpaces;
            _optionalContentVisibility = optionalContentVisibility;
            _initialState = GraphicsState.Create(baseTransform, initialFillColor, initialFillColorSpace, initialFillOpacity, initialClipPath);
            _state = _initialState;
            _pageHeight = pageHeight;
            _paintOrderBase = paintOrderBase;
            _paintOrderScale = paintOrderScale;
            _paintOrderOffset = paintOrderOffset;
        }

        public IReadOnlyList<PdfPageXObjectInvocation> Parse() {
            while (_index < _content.Length) {
                SkipWhitespace();
                if (_index >= _content.Length) {
                    break;
                }

                char current = _content[_index];
                if (current == '%') {
                    SkipComment();
                } else if (current == '/') {
                    _args.Add(ReadName());
                } else if (current == '(') {
                    SkipLiteralString();
                } else if (current == '<') {
                    if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                        _args.Add(PdfInlineOptionalContentReferenceParser.Read(_content, ref _index));
                    } else {
                        SkipAngleObject();
                    }
                } else if (current == '[') {
                    SkipArray();
                } else if (IsNumberStart(current)) {
                    _args.Add(ReadNumber());
                } else {
                    double paintOrder = GetPaintOrder(_index);
                    string op = ReadOperator();
                    if (op.Length == 0) {
                        _index++;
                    } else {
                        ApplyOperator(op, paintOrder);
                    }
                }
            }

            return _invocations.Count == 0 ? Array.Empty<PdfPageXObjectInvocation>() : _invocations.AsReadOnly();
        }

        private double GetPaintOrder(int operatorIndex) => _paintOrderBase + ((operatorIndex + _paintOrderOffset) * _paintOrderScale);

        private void ApplyOperator(string op, double paintOrder) {
            switch (op) {
                case "q":
                    _stack.Push(_state);
                    break;
                case "Q":
                    _state = _stack.Count > 0 ? _stack.Pop() : _initialState;
                    break;
                case "cm":
                    if (_args.Count >= 6) {
                        Matrix2D matrix = new Matrix2D(
                            NumberAt(_args.Count - 6),
                            NumberAt(_args.Count - 5),
                            NumberAt(_args.Count - 4),
                            NumberAt(_args.Count - 3),
                            NumberAt(_args.Count - 2),
                            NumberAt(_args.Count - 1));
                        _state = _state.WithTransform(Matrix2D.Multiply(_state.Transform, matrix));
                    }

                    break;
                case "re":
                    if (_args.Count >= 4) {
                        AddRectanglePath(NumberAt(_args.Count - 4), NumberAt(_args.Count - 3), NumberAt(_args.Count - 2), NumberAt(_args.Count - 1));
                    }

                    break;
                case "m":
                    if (_args.Count >= 2) {
                        MoveTo(NumberAt(_args.Count - 2), NumberAt(_args.Count - 1));
                    }

                    break;
                case "l":
                    if (_args.Count >= 2) {
                        LineTo(NumberAt(_args.Count - 2), NumberAt(_args.Count - 1));
                    }

                    break;
                case "c":
                    if (_args.Count >= 6) {
                        CubicTo(
                            NumberAt(_args.Count - 6),
                            NumberAt(_args.Count - 5),
                            NumberAt(_args.Count - 4),
                            NumberAt(_args.Count - 3),
                            NumberAt(_args.Count - 2),
                            NumberAt(_args.Count - 1));
                    }

                    break;
                case "v":
                    if (_args.Count >= 4 && _path.Count > 0) {
                        (double X, double Y) currentPoint = _path[_path.Count - 1];
                        CubicTo(
                            currentPoint.X,
                            currentPoint.Y,
                            NumberAt(_args.Count - 4),
                            NumberAt(_args.Count - 3),
                            NumberAt(_args.Count - 2),
                            NumberAt(_args.Count - 1),
                            firstControlAlreadyTransformed: true);
                    }

                    break;
                case "y":
                    if (_args.Count >= 4) {
                        CubicTo(
                            NumberAt(_args.Count - 4),
                            NumberAt(_args.Count - 3),
                            NumberAt(_args.Count - 2),
                            NumberAt(_args.Count - 1),
                            NumberAt(_args.Count - 2),
                            NumberAt(_args.Count - 1));
                    }

                    break;
                case "h":
                    ClosePath();

                    break;
                case "W":
                    if (!HasHiddenContent()) {
                        CaptureClipPath(OfficeFillRule.NonZero);
                    }

                    break;
                case "W*":
                    if (!HasHiddenContent()) {
                        CaptureClipPath(OfficeFillRule.EvenOdd);
                    }

                    break;
                case "n":
                    ClearPath();
                    break;
                case "S":
                case "s":
                case "f":
                case "F":
                case "f*":
                case "B":
                case "B*":
                case "b":
                case "b*":
                    ClearPath();
                    break;
                case "gs":
                    if (_args.Count >= 1 && _args[_args.Count - 1] is string graphicsStateName) {
                        ApplyGraphicsStateResource(graphicsStateName);
                    }

                    break;
                case "cs":
                    if (_args.Count >= 1 &&
                        _args[_args.Count - 1] is string fillColorSpaceName &&
                        TryReadColorSpace(fillColorSpaceName, out PdfPageColorSpaceKind fillColorSpace)) {
                        _state = _state.WithFillColorSpace(fillColorSpace);
                    }

                    break;
                case "sc":
                case "scn":
                    if (TryReadColor(_state.FillColorSpace, out OfficeColor fillColor)) {
                        _state = _state.WithFillColor(fillColor);
                    }

                    break;
                case "rg":
                    if (_args.Count >= 3) {
                        _state = _state.WithFillColor(ReadRgb(_args.Count - 3), PdfPageColorSpaceKind.DeviceRgb);
                    }

                    break;
                case "g":
                    if (_args.Count >= 1) {
                        _state = _state.WithFillColor(ReadGray(_args.Count - 1), PdfPageColorSpaceKind.DeviceGray);
                    }

                    break;
                case "k":
                    if (_args.Count >= 4) {
                        _state = _state.WithFillColor(ReadCmyk(_args.Count - 4), PdfPageColorSpaceKind.DeviceCmyk);
                    }

                    break;
                case "Do":
                    if (!HasHiddenContent() &&
                        _args.Count >= 1 &&
                        _args[_args.Count - 1] is string name &&
                        !string.IsNullOrEmpty(name)) {
                        _invocations.Add(new PdfPageXObjectInvocation(name, _state.Transform, _state.ClipPath, _state.FillColor, _state.FillColorSpace, _state.FillOpacity, paintOrder));
                    }

                    break;
                case "BI":
                    if (TryReadInlineImage(out PdfPageInlineImage? inlineImage) && inlineImage != null && !HasHiddenContent()) {
                        _invocations.Add(new PdfPageXObjectInvocation(inlineImage, _state.Transform, _state.ClipPath, _state.FillColor, _state.FillColorSpace, _state.FillOpacity, paintOrder));
                    }

                    break;
                case "BDC":
                    _hiddenContentStack.Push(IsHiddenOptionalContent(_args.Count > 1 ? _args[_args.Count - 2] : null, _args.Count > 0 ? _args[_args.Count - 1] : null));
                    break;
                case "BMC":
                    _hiddenContentStack.Push(false);
                    break;
                case "EMC":
                    if (_hiddenContentStack.Count > 0) {
                        _hiddenContentStack.Pop();
                    }

                    break;
            }

            _args.Clear();
        }

        private void AddRectanglePath(double x, double y, double width, double height) {
            var p0 = TransformPoint(x, y);
            var p1 = TransformPoint(x + width, y);
            var p2 = TransformPoint(x + width, y + height);
            var p3 = TransformPoint(x, y + height);
            _currentSubpathStartIndex = _path.Count;
            _path.Add(p0);
            _path.Add(p1);
            _path.Add(p2);
            _path.Add(p3);
            _path.Add(p0);
            _pathCommands.Add(OfficePathCommand.MoveTo(ToOfficePoint(p0)));
            _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(p1)));
            _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(p2)));
            _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(p3)));
            _pathCommands.Add(OfficePathCommand.Close());
        }

        private void MoveTo(double x, double y) {
            (double X, double Y) point = TransformPoint(x, y);
            _currentSubpathStartIndex = _path.Count;
            _path.Add(point);
            _pathCommands.Add(OfficePathCommand.MoveTo(ToOfficePoint(point)));
        }

        private void LineTo(double x, double y) {
            if (_currentSubpathStartIndex < 0) {
                MoveTo(x, y);
                return;
            }

            (double X, double Y) point = TransformPoint(x, y);
            _path.Add(point);
            _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(point)));
        }

        private void CubicTo(double c1x, double c1y, double c2x, double c2y, double endX, double endY, bool firstControlAlreadyTransformed = false) {
            if (_path.Count == 0 || _currentSubpathStartIndex < 0) {
                MoveTo(endX, endY);
                return;
            }

            (double X, double Y) control1 = firstControlAlreadyTransformed ? (c1x, c1y) : TransformPoint(c1x, c1y);
            (double X, double Y) control2 = TransformPoint(c2x, c2y);
            (double X, double Y) end = TransformPoint(endX, endY);
            _path.Add(end);
            _pathCommands.Add(OfficePathCommand.CubicBezierTo(ToOfficePoint(control1), ToOfficePoint(control2), ToOfficePoint(end)));
        }

        private void CaptureClipPath(OfficeFillRule fillRule) {
            if (TryCreateAxisAlignedRectangle(out double x, out double y, out double width, out double height)) {
                _state = _state.WithClipPath(PdfPageClipPath.ResolveActiveClip(_state.ClipPath, PdfPageClipPath.Rectangle(x, y, width, height)));
                return;
            }

            if (PdfPageClipPath.TryCreatePath(_pathCommands, fillRule, out PdfPageClipPath clipPath)) {
                _state = _state.WithClipPath(PdfPageClipPath.ResolveActiveClip(_state.ClipPath, clipPath));
            }
        }

        private bool TryCreateAxisAlignedRectangle(out double x, out double y, out double width, out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (_path.Count < 4) {
                return false;
            }

            if (_path.Count != 5 ||
                _pathCommands.Count != 5 ||
                _pathCommands[0].Kind != OfficePathCommandKind.MoveTo ||
                _pathCommands[1].Kind != OfficePathCommandKind.LineTo ||
                _pathCommands[2].Kind != OfficePathCommandKind.LineTo ||
                _pathCommands[3].Kind != OfficePathCommandKind.LineTo ||
                _pathCommands[4].Kind != OfficePathCommandKind.Close ||
                !NearlyEqual(_path[0].X, _path[4].X) ||
                !NearlyEqual(ToTop(_path[0].Y), ToTop(_path[4].Y))) {
                return false;
            }

            double left = _path.Min(point => point.X);
            double right = _path.Max(point => point.X);
            double top = _path.Min(point => ToTop(point.Y));
            double bottom = _path.Max(point => ToTop(point.Y));
            width = right - left;
            height = bottom - top;
            if (width <= 0D || height <= 0D) {
                return false;
            }

            for (int i = 0; i < _path.Count; i++) {
                bool onVertical = NearlyEqual(_path[i].X, left) || NearlyEqual(_path[i].X, right);
                bool onHorizontal = NearlyEqual(ToTop(_path[i].Y), top) || NearlyEqual(ToTop(_path[i].Y), bottom);
                if (!onVertical || !onHorizontal) {
                    return false;
                }
            }

            for (int i = 0; i < 4; i++) {
                double x1 = _path[i].X;
                double y1 = ToTop(_path[i].Y);
                double x2 = _path[i + 1].X;
                double y2 = ToTop(_path[i + 1].Y);
                bool horizontal = NearlyEqual(y1, y2) && !NearlyEqual(x1, x2);
                bool vertical = NearlyEqual(x1, x2) && !NearlyEqual(y1, y2);
                if (!horizontal && !vertical) {
                    return false;
                }
            }

            x = left;
            y = top;
            return true;
        }

        private void ClosePath() {
            if (_path.Count == 0 || _currentSubpathStartIndex < 0 || _currentSubpathStartIndex >= _path.Count) {
                return;
            }

            _path.Add(_path[_currentSubpathStartIndex]);
            _pathCommands.Add(OfficePathCommand.Close());
        }

        private void ClearPath() {
            _path.Clear();
            _pathCommands.Clear();
            _currentSubpathStartIndex = -1;
        }

        private (double X, double Y) TransformPoint(double x, double y) => _state.Transform.Transform(x, y);

        private double ToTop(double pdfY) => _pageHeight - pdfY;

        private OfficePoint ToOfficePoint((double X, double Y) point) => new OfficePoint(point.X, ToTop(point.Y));

        private double NumberAt(int index) => _args[index] is double value ? value : 0D;

        private void ApplyGraphicsStateResource(string name) {
            if (_graphicsStates == null || !_graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource resource)) {
                return;
            }

            _state = _state.WithOpacity(resource.FillOpacity);
        }

        private bool HasHiddenContent() {
            foreach (bool hidden in _hiddenContentStack) {
                if (hidden) {
                    return true;
                }
            }

            return false;
        }

        private bool IsHiddenOptionalContent(object? tag, object? property) =>
            tag is string tagName &&
            string.Equals(tagName, "OC", StringComparison.Ordinal) &&
            ((property is string propertyName &&
                _optionalContentVisibility?.IsHidden(propertyName) == true) ||
             (property is PdfInlineOptionalContentReferences references &&
                _optionalContentVisibility?.IsHidden(references) == true));

        private bool TryReadInlineImage(out PdfPageInlineImage? inlineImage) {
            inlineImage = null;
            var dictionary = new PdfDictionary();
            while (_index < _content.Length) {
                SkipWhitespace();
                if (_index >= _content.Length) {
                    return false;
                }

                if (IsOperatorAt("ID")) {
                    _index += 2;
                    break;
                }

                if (_content[_index] != '/') {
                    ReadOperator();
                    continue;
                }

                string key = NormalizeInlineImageKey(ReadName());
                SkipWhitespace();
                if (_index >= _content.Length) {
                    return false;
                }

                if (!TryReadInlineImageValue(out PdfObject? value) || value == null) {
                    continue;
                }

                dictionary.Items[key] = value;
            }

            if (_index < _content.Length && char.IsWhiteSpace(_content[_index])) {
                _index++;
            }

            int dataStart = _index;
            int dataLength = TryGetRawInlineImageLength(dictionary, out int rawLength)
                ? rawLength
                : FindInlineImageDataLength(dataStart);
            if (dataLength < 0 || dataStart + dataLength > _content.Length) {
                return false;
            }

            byte[] data = ReadBytes(dataStart, dataLength);
            _index = dataStart + dataLength;
            SkipWhitespace();
            if (IsOperatorAt("EI")) {
                _index += 2;
            }

            var stream = new PdfStream(dictionary, data);
            inlineImage = new PdfPageInlineImage("__inline" + (++_inlineImageIndex).ToString(CultureInfo.InvariantCulture), stream);
            return true;
        }

        private bool TryReadInlineImageValue(out PdfObject? value) {
            value = null;
            char current = _content[_index];
            if (current == '/') {
                value = new PdfName(NormalizeInlineImageName(ReadName()));
                return true;
            }

            if (IsNumberStart(current)) {
                value = new PdfNumber(ReadNumber());
                return true;
            }

            if (current == '[') {
                value = ReadInlineImageArray();
                return true;
            }

            string token = ReadOperator();
            if (string.Equals(token, "true", StringComparison.Ordinal)) {
                value = new PdfBoolean(true);
                return true;
            }

            if (string.Equals(token, "false", StringComparison.Ordinal)) {
                value = new PdfBoolean(false);
                return true;
            }

            return false;
        }

        private PdfArray ReadInlineImageArray() {
            var array = new PdfArray();
            _index++;
            while (_index < _content.Length) {
                SkipWhitespace();
                if (_index >= _content.Length) {
                    break;
                }

                if (_content[_index] == ']') {
                    _index++;
                    break;
                }

                if (TryReadInlineImageValue(out PdfObject? value) && value != null) {
                    array.Items.Add(value);
                } else {
                    _index++;
                }
            }

            return array;
        }

        private static bool TryGetRawInlineImageLength(PdfDictionary dictionary, out int length) {
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
            bool isImageMask = dictionary.Items.TryGetValue("ImageMask", out PdfObject? imageMaskObject) &&
                imageMaskObject is PdfBoolean imageMask &&
                imageMask.Value;
            if (isImageMask && bitsPerComponent == 0) {
                bitsPerComponent = 1;
            }

            int components = isImageMask ? 1 : GetInlineImageComponentCount(dictionary);
            if (bitsPerComponent <= 0 || components <= 0) {
                return false;
            }

            long rowBitCount = (long)width * components * bitsPerComponent;
            long rowByteCount = (rowBitCount + 7L) / 8L;
            long byteCount = rowByteCount * height;
            if (byteCount <= 0L || byteCount > int.MaxValue) {
                return false;
            }

            length = (int)byteCount;
            return true;
        }

        private static int GetInlineImageComponentCount(PdfDictionary dictionary) {
            string colorSpace = dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) && colorSpaceObject is PdfName colorSpaceName
                ? colorSpaceName.Name
                : "DeviceGray";
            switch (colorSpace) {
                case "DeviceRGB":
                    return 3;
                case "DeviceCMYK":
                    return 4;
                default:
                    return 1;
            }
        }

        private static int ReadPositiveInteger(PdfDictionary dictionary, string key) =>
            dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            value is PdfNumber number &&
            number.Value > 0D &&
            number.Value <= int.MaxValue
                ? (int)number.Value
                : 0;

        private int FindInlineImageDataLength(int dataStart) {
            int index = dataStart;
            while (index + 2 < _content.Length) {
                if (char.IsWhiteSpace(_content[index]) &&
                    _content[index + 1] == 'E' &&
                    _content[index + 2] == 'I' &&
                    (index + 3 >= _content.Length || IsDelimiter(_content[index + 3]))) {
                    return index - dataStart;
                }

                index++;
            }

            return -1;
        }

        private byte[] ReadBytes(int start, int length) {
            var data = new byte[length];
            for (int i = 0; i < length; i++) {
                data[i] = (byte)_content[start + i];
            }

            return data;
        }

        private bool IsOperatorAt(string op) =>
            _index + op.Length <= _content.Length &&
            string.CompareOrdinal(_content, _index, op, 0, op.Length) == 0 &&
            (_index + op.Length >= _content.Length || IsDelimiter(_content[_index + op.Length]));

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

        private OfficeColor ReadRgb(int startIndex) =>
            OfficeColor.FromRgb(ToByte(NumberAt(startIndex)), ToByte(NumberAt(startIndex + 1)), ToByte(NumberAt(startIndex + 2)));

        private OfficeColor ReadGray(int index) {
            byte value = ToByte(NumberAt(index));
            return OfficeColor.FromRgb(value, value, value);
        }

        private OfficeColor ReadCmyk(int startIndex) {
            double cyan = Clamp01(NumberAt(startIndex));
            double magenta = Clamp01(NumberAt(startIndex + 1));
            double yellow = Clamp01(NumberAt(startIndex + 2));
            double black = Clamp01(NumberAt(startIndex + 3));
            return OfficeColor.FromRgb(
                ToByte((1D - cyan) * (1D - black)),
                ToByte((1D - magenta) * (1D - black)),
                ToByte((1D - yellow) * (1D - black)));
        }

        private bool TryReadColor(PdfPageColorSpaceKind colorSpace, out OfficeColor color) {
            color = OfficeColor.Black;
            int componentCount = GetColorComponentCount(colorSpace);
            int endIndex = _args.Count;
            while (endIndex > 0 && !(_args[endIndex - 1] is double)) {
                endIndex--;
            }

            if (endIndex < componentCount) {
                return false;
            }

            int startIndex = endIndex - componentCount;
            switch (colorSpace) {
                case PdfPageColorSpaceKind.DeviceRgb:
                    color = ReadRgb(startIndex);
                    return true;
                case PdfPageColorSpaceKind.DeviceCmyk:
                    color = ReadCmyk(startIndex);
                    return true;
                default:
                    color = ReadGray(startIndex);
                    return true;
            }
        }

        private static int GetColorComponentCount(PdfPageColorSpaceKind colorSpace) {
            switch (colorSpace) {
                case PdfPageColorSpaceKind.DeviceRgb:
                    return 3;
                case PdfPageColorSpaceKind.DeviceCmyk:
                    return 4;
                default:
                    return 1;
            }
        }

        private bool TryReadColorSpace(string name, out PdfPageColorSpaceKind colorSpace) {
            switch (name) {
                case "DeviceRGB":
                case "RGB":
                    colorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    return true;
                case "DeviceCMYK":
                case "CMYK":
                    colorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    return true;
                case "DeviceGray":
                case "G":
                    colorSpace = PdfPageColorSpaceKind.DeviceGray;
                    return true;
                default:
                    if (_colorSpaces != null && _colorSpaces.TryGetValue(name, out colorSpace)) {
                        return true;
                    }

                    colorSpace = PdfPageColorSpaceKind.DeviceGray;
                    return false;
            }
        }

        private static byte ToByte(double value) => (byte)Math.Round(Clamp01(value) * 255D);

        private static double Clamp01(double value) {
            if (value < 0D) {
                return 0D;
            }

            return value > 1D ? 1D : value;
        }

        private void SkipWhitespace() {
            while (_index < _content.Length && char.IsWhiteSpace(_content[_index])) {
                _index++;
            }
        }

        private void SkipComment() {
            while (_index < _content.Length && _content[_index] != '\r' && _content[_index] != '\n') {
                _index++;
            }
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
            int start = _index;
            _index++;
            while (_index < _content.Length) {
                char ch = _content[_index];
                if (!(char.IsDigit(ch) || ch == '.' || ch == '-' || ch == '+' || ch == 'e' || ch == 'E')) {
                    break;
                }

                _index++;
            }

#pragma warning disable CA1846 // Keep netstandard2.0-safe parsing instead of requiring span overloads.
            return double.TryParse(_content.Substring(start, _index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
#pragma warning restore CA1846
                ? value
                : 0D;
        }

        private string ReadOperator() {
            int start = _index;
            while (_index < _content.Length && !IsDelimiter(_content[_index])) {
                _index++;
            }

            return _content.Substring(start, _index - start);
        }

        private void SkipLiteralString() {
            int depth = 1;
            bool escaped = false;
            _index++;
            while (_index < _content.Length && depth > 0) {
                char ch = _content[_index++];
                if (escaped) {
                    escaped = false;
                } else if (ch == '\\') {
                    escaped = true;
                } else if (ch == '(') {
                    depth++;
                } else if (ch == ')') {
                    depth--;
                }
            }
        }

        private void SkipAngleObject() {
            if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                _index += 2;
                int depth = 1;
                while (_index < _content.Length && depth > 0) {
                    if (_index + 1 < _content.Length && _content[_index] == '<' && _content[_index + 1] == '<') {
                        depth++;
                        _index += 2;
                    } else if (_index + 1 < _content.Length && _content[_index] == '>' && _content[_index + 1] == '>') {
                        depth--;
                        _index += 2;
                    } else {
                        _index++;
                    }
                }
                return;
            }

            _index++;
            while (_index < _content.Length && _content[_index] != '>') {
                _index++;
            }

            if (_index < _content.Length) {
                _index++;
            }
        }

        private void SkipArray() {
            int depth = 1;
            _index++;
            while (_index < _content.Length && depth > 0) {
                char ch = _content[_index];
                if (ch == '(') {
                    SkipLiteralString();
                } else if (ch == '<') {
                    SkipAngleObject();
                } else {
                    if (ch == '[') {
                        depth++;
                    } else if (ch == ']') {
                        depth--;
                    }

                    _index++;
                }
            }
        }

        private static bool IsNumberStart(char ch) => ch == '-' || ch == '+' || ch == '.' || char.IsDigit(ch);

        private static bool IsDelimiter(char ch) =>
            char.IsWhiteSpace(ch) || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>' || ch == '%';

        private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
    }

    private readonly struct GraphicsState {
        private GraphicsState(Matrix2D transform, PdfPageClipPath? clipPath, OfficeColor fillColor, PdfPageColorSpaceKind fillColorSpace, double? fillOpacity) {
            Transform = transform;
            ClipPath = clipPath;
            FillColor = fillColor;
            FillColorSpace = fillColorSpace;
            FillOpacity = fillOpacity;
        }

        public Matrix2D Transform { get; }

        public PdfPageClipPath? ClipPath { get; }

        public OfficeColor FillColor { get; }

        public PdfPageColorSpaceKind FillColorSpace { get; }

        public double? FillOpacity { get; }

        public static GraphicsState Create(Matrix2D transform) =>
            Create(transform, null, PdfPageColorSpaceKind.DeviceGray, null, null);

        public static GraphicsState Create(Matrix2D transform, OfficeColor? fillColor, PdfPageColorSpaceKind fillColorSpace, double? fillOpacity, PdfPageClipPath? clipPath) =>
            new GraphicsState(transform, clipPath, fillColor ?? OfficeColor.Black, fillColorSpace, fillOpacity);

        public GraphicsState WithTransform(Matrix2D transform) => new GraphicsState(transform, ClipPath, FillColor, FillColorSpace, FillOpacity);

        public GraphicsState WithClipPath(PdfPageClipPath clipPath) => new GraphicsState(Transform, clipPath, FillColor, FillColorSpace, FillOpacity);

        public GraphicsState WithFillColor(OfficeColor color) => new GraphicsState(Transform, ClipPath, color, FillColorSpace, FillOpacity);

        public GraphicsState WithFillColor(OfficeColor color, PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, ClipPath, color, colorSpace, FillOpacity);

        public GraphicsState WithFillColorSpace(PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, ClipPath, FillColor, colorSpace, FillOpacity);

        public GraphicsState WithOpacity(double? fillOpacity) =>
            new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, fillOpacity ?? FillOpacity);
    }
}

internal readonly struct PdfPageXObjectInvocation {
    public PdfPageXObjectInvocation(string name, Matrix2D transform, PdfPageClipPath? clipPath, OfficeColor fillColor, PdfPageColorSpaceKind fillColorSpace, double? fillOpacity, double paintOrder = 0D) {
        Name = name;
        InlineImage = null;
        Transform = transform;
        ClipPath = clipPath;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillOpacity = fillOpacity;
        PaintOrder = paintOrder;
    }

    public PdfPageXObjectInvocation(PdfPageInlineImage inlineImage, Matrix2D transform, PdfPageClipPath? clipPath, OfficeColor fillColor, PdfPageColorSpaceKind fillColorSpace, double? fillOpacity, double paintOrder = 0D) {
        Name = inlineImage.ResourceName;
        InlineImage = inlineImage;
        Transform = transform;
        ClipPath = clipPath;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillOpacity = fillOpacity;
        PaintOrder = paintOrder;
    }

    public string Name { get; }

    public PdfPageInlineImage? InlineImage { get; }

    public Matrix2D Transform { get; }

    public PdfPageClipPath? ClipPath { get; }

    public OfficeColor FillColor { get; }

    public PdfPageColorSpaceKind FillColorSpace { get; }

    public double? FillOpacity { get; }

    public double PaintOrder { get; }
}
