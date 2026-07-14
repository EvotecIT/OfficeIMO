using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfPageXObjectInvocationParser {
    private const double HairlineStrokeWidth = 0.25D;

    private static double ResolveStrokeWidth(double value) {
        if (value < 0D) {
            return 0D;
        }

        return Math.Abs(value) <= 0.001D ? HairlineStrokeWidth : value;
    }

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
        PdfPageClipPath? initialClipPath = null,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpaceKind initialStrokeColorSpace = PdfPageColorSpaceKind.DeviceGray,
        double? initialStrokeOpacity = null,
        double? initialStrokeWidth = null,
        OfficeStrokeDashStyle? initialStrokeDashStyle = null,
        OfficeStrokeLineCap? initialStrokeLineCap = null,
        OfficeStrokeLineJoin? initialStrokeLineJoin = null,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations) {
        if (string.IsNullOrEmpty(content)) {
            return Array.Empty<PdfPageXObjectInvocation>();
        }

        var parser = new Parser(content, baseTransform, pageHeight, graphicsStates, colorSpaces, optionalContentVisibility, initialFillColor, initialFillColorSpace, initialFillOpacity, paintOrderBase, paintOrderScale, paintOrderOffset, initialClipPath, initialStrokeColor, initialStrokeColorSpace, initialStrokeOpacity, initialStrokeWidth, initialStrokeDashStyle, initialStrokeLineCap, initialStrokeLineJoin, maxOperations);
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
        private readonly Stack<TextState> _textStack = new Stack<TextState>();
        private readonly Stack<bool> _hiddenContentStack = new Stack<bool>();
        private readonly List<(double X, double Y)> _path = new List<(double X, double Y)>();
        private readonly List<OfficePathCommand> _pathCommands = new List<OfficePathCommand>();
        private readonly GraphicsState _initialState;
        private GraphicsState _state;
        private bool _inText;
        private double _textSize = 12D;
        private double _textLeading = 14.4D;
        private double _textCharSpacing;
        private double _textWordSpacing;
        private double _textHScale = 1D;
        private double _textRise;
        private int _textRenderingMode;
        private Matrix2D _textMatrix = Matrix2D.Identity;
        private Matrix2D _lineMatrix = Matrix2D.Identity;
        private int _currentSubpathStartIndex = -1;
        private int _index;
        private int _inlineImageIndex;
        private readonly int _maxOperations;
        private int _operationCount;

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
            PdfPageClipPath? initialClipPath,
            OfficeColor? initialStrokeColor,
            PdfPageColorSpaceKind initialStrokeColorSpace,
            double? initialStrokeOpacity,
            double? initialStrokeWidth,
            OfficeStrokeDashStyle? initialStrokeDashStyle,
            OfficeStrokeLineCap? initialStrokeLineCap,
            OfficeStrokeLineJoin? initialStrokeLineJoin,
            int maxOperations) {
            _content = content;
            _baseTransform = baseTransform;
            _graphicsStates = graphicsStates;
            _colorSpaces = colorSpaces;
            _optionalContentVisibility = optionalContentVisibility;
            _initialState = GraphicsState.Create(baseTransform, initialFillColor, initialFillColorSpace, initialFillOpacity, initialClipPath, initialStrokeColor, initialStrokeColorSpace, initialStrokeOpacity, initialStrokeWidth, initialStrokeDashStyle, initialStrokeLineCap, initialStrokeLineJoin);
            _state = _initialState;
            _pageHeight = pageHeight;
            _paintOrderBase = paintOrderBase;
            _paintOrderScale = paintOrderScale;
            _paintOrderOffset = paintOrderOffset;
            _maxOperations = maxOperations;
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
                    _args.Add(ReadLiteralStringBytes());
                } else if (current == '<') {
                    if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                        _args.Add(PdfInlineOptionalContentReferenceParser.Read(_content, ref _index));
                    } else {
                        _args.Add(ReadHexStringBytes());
                    }
                } else if (current == '[') {
                    _args.Add(ReadArray());
                } else if (IsNumberStart(current)) {
                    _args.Add(ReadNumber());
                } else {
                    double paintOrder = GetPaintOrder(_index);
                    string op = ReadOperator();
                    if (op.Length == 0) {
                        _index++;
                    } else {
                        if (++_operationCount > _maxOperations) {
                            throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperations, _maxOperations, _operationCount);
                        }
                        ApplyOperator(op, paintOrder);
                    }
                }
            }

            return _invocations.Count == 0 ? Array.Empty<PdfPageXObjectInvocation>() : _invocations.AsReadOnly();
        }

        private double GetPaintOrder(int operatorIndex) => _paintOrderBase + ((operatorIndex + _paintOrderOffset) * _paintOrderScale);

        private TextState CaptureTextState() =>
            new TextState(_inText, _textSize, _textLeading, _textCharSpacing, _textWordSpacing, _textHScale, _textRise, _textRenderingMode, _textMatrix, _lineMatrix);

        private void RestoreTextState(TextState state) {
            _inText = state.InText;
            _textSize = state.Size;
            _textLeading = state.Leading;
            _textCharSpacing = state.CharSpacing;
            _textWordSpacing = state.WordSpacing;
            _textHScale = state.HScale;
            _textRise = state.TextRise;
            _textRenderingMode = state.TextRenderingMode;
            _textMatrix = state.TextMatrix;
            _lineMatrix = state.LineMatrix;
        }

        private void SetTextMatrix(int startIndex) {
            _lineMatrix = new Matrix2D(
                NumberAt(startIndex),
                NumberAt(startIndex + 1),
                NumberAt(startIndex + 2),
                NumberAt(startIndex + 3),
                NumberAt(startIndex + 4),
                NumberAt(startIndex + 5));
            _textMatrix = _lineMatrix;
        }

        private void MoveTextLine(double tx, double ty) {
            _lineMatrix = Matrix2D.Multiply(_lineMatrix, Matrix2D.Translation(tx, ty));
            _textMatrix = _lineMatrix;
        }

        private void MoveToNextTextLine() {
            _lineMatrix = Matrix2D.Multiply(_lineMatrix, Matrix2D.Translation(0D, -_textLeading));
            _textMatrix = _lineMatrix;
        }

        private void ShowText(object textObject) {
            if (!_inText || textObject is not byte[] bytes || bytes.Length == 0) {
                return;
            }

            double advance = EstimateTextAdvance(bytes);
            ApplyTextClippingPath(advance);
            _textMatrix = Matrix2D.Multiply(_textMatrix, Matrix2D.Translation(advance, 0D));
        }

        private void ShowTextArray(object arrayObject) {
            if (arrayObject is not List<object> items) {
                ShowText(arrayObject);
                return;
            }

            for (int i = 0; i < items.Count; i++) {
                if (items[i] is byte[] bytes) {
                    ShowText(bytes);
                } else if (items[i] is double kerning) {
                    double delta = -kerning / 1000D * _textSize * _textHScale;
                    _textMatrix = Matrix2D.Multiply(_textMatrix, Matrix2D.Translation(delta, 0D));
                }
            }
        }

        private double EstimateTextAdvance(byte[] bytes) {
            double glyphAdvance = Math.Max(0.001D, _textSize * 0.5D);
            double advance = 0D;
            for (int i = 0; i < bytes.Length; i++) {
                advance += glyphAdvance + _textCharSpacing;
                if (bytes[i] == 32) {
                    advance += _textWordSpacing;
                }
            }

            return advance * _textHScale;
        }

        private void ApplyTextClippingPath(double advance) {
            if (!AddsTextToClippingPath(_textRenderingMode) || _textSize <= 0D || Math.Abs(advance) <= 0.000001D) {
                return;
            }

            double left = advance < 0D ? advance : 0D;
            double width = Math.Abs(advance);
            double descent = Math.Max(0.001D, _textSize * 0.25D);
            double height = Math.Max(0.001D, _textSize + descent);
            Matrix2D textToPage = Matrix2D.Multiply(_state.Transform, _textMatrix);
            var textClipBuilder = new PdfPageClipPathBuilder(_pageHeight);
            textClipBuilder.AddRectanglePath(textToPage, left, _textRise - descent, width, height);
            if (textClipBuilder.TryCreateClipPath(OfficeFillRule.NonZero, out PdfPageClipPath textClipPath)) {
                _state = _state.WithClipPath(PdfPageClipPath.ResolveActiveClip(_state.ClipPath, textClipPath));
            }
        }

        private void ApplyOperator(string op, double paintOrder) {
            switch (op) {
                case "q":
                    _stack.Push(_state);
                    _textStack.Push(CaptureTextState());
                    break;
                case "Q":
                    _state = _stack.Count > 0 ? _stack.Pop() : _initialState;
                    RestoreTextState(_textStack.Count > 0 ? _textStack.Pop() : TextState.Default);
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
                case "w":
                    if (_args.Count >= 1) {
                        _state = _state.WithStrokeWidth(ResolveStrokeWidth(NumberAt(_args.Count - 1)));
                    }

                    break;
                case "J":
                    if (_args.Count >= 1) {
                        _state = _state.WithStrokeLineCap(ReadLineCap(NumberAt(_args.Count - 1)));
                    }

                    break;
                case "j":
                    if (_args.Count >= 1) {
                        _state = _state.WithStrokeLineJoin(ReadLineJoin(NumberAt(_args.Count - 1)));
                    }

                    break;
                case "d":
                    if (_args.Count >= 2 && TryGetNumberArray(_args[_args.Count - 2], out double[] dashArray)) {
                        _state = _state.WithStrokeDashStyle(ReadDashStyle(dashArray));
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
                case "CS":
                    if (_args.Count >= 1 &&
                        _args[_args.Count - 1] is string strokeColorSpaceName &&
                        TryReadColorSpace(strokeColorSpaceName, out PdfPageColorSpaceKind strokeColorSpace)) {
                        _state = _state.WithStrokeColorSpace(strokeColorSpace);
                    }

                    break;
                case "sc":
                case "scn":
                    if (TryReadColor(_state.FillColorSpace, out OfficeColor fillColor)) {
                        _state = _state.WithFillColor(fillColor);
                    }

                    break;
                case "SC":
                case "SCN":
                    if (TryReadColor(_state.StrokeColorSpace, out OfficeColor strokeColor)) {
                        _state = _state.WithStrokeColor(strokeColor);
                    }

                    break;
                case "rg":
                    if (_args.Count >= 3) {
                        _state = _state.WithFillColor(ReadRgb(_args.Count - 3), PdfPageColorSpaceKind.DeviceRgb);
                    }

                    break;
                case "RG":
                    if (_args.Count >= 3) {
                        _state = _state.WithStrokeColor(ReadRgb(_args.Count - 3), PdfPageColorSpaceKind.DeviceRgb);
                    }

                    break;
                case "g":
                    if (_args.Count >= 1) {
                        _state = _state.WithFillColor(ReadGray(_args.Count - 1), PdfPageColorSpaceKind.DeviceGray);
                    }

                    break;
                case "G":
                    if (_args.Count >= 1) {
                        _state = _state.WithStrokeColor(ReadGray(_args.Count - 1), PdfPageColorSpaceKind.DeviceGray);
                    }

                    break;
                case "k":
                    if (_args.Count >= 4) {
                        _state = _state.WithFillColor(ReadCmyk(_args.Count - 4), PdfPageColorSpaceKind.DeviceCmyk);
                    }

                    break;
                case "K":
                    if (_args.Count >= 4) {
                        _state = _state.WithStrokeColor(ReadCmyk(_args.Count - 4), PdfPageColorSpaceKind.DeviceCmyk);
                    }

                    break;
                case "BT":
                    _inText = true;
                    _textMatrix = Matrix2D.Identity;
                    _lineMatrix = Matrix2D.Identity;
                    break;
                case "ET":
                    _inText = false;
                    break;
                case "Tf":
                    if (_args.Count >= 2) {
                        _textSize = NumberAt(_args.Count - 1);
                    }

                    break;
                case "Tm":
                    if (_args.Count >= 6) {
                        SetTextMatrix(_args.Count - 6);
                    }

                    break;
                case "Td":
                    if (_args.Count >= 2) {
                        MoveTextLine(NumberAt(_args.Count - 2), NumberAt(_args.Count - 1));
                    }

                    break;
                case "TD":
                    if (_args.Count >= 2) {
                        double tx = NumberAt(_args.Count - 2);
                        double ty = NumberAt(_args.Count - 1);
                        _textLeading = -ty;
                        MoveTextLine(tx, ty);
                    }

                    break;
                case "TL":
                    if (_args.Count >= 1) {
                        _textLeading = NumberAt(_args.Count - 1);
                    }

                    break;
                case "T*":
                    MoveToNextTextLine();
                    break;
                case "Tc":
                    if (_args.Count >= 1) {
                        _textCharSpacing = NumberAt(_args.Count - 1);
                    }

                    break;
                case "Tw":
                    if (_args.Count >= 1) {
                        _textWordSpacing = NumberAt(_args.Count - 1);
                    }

                    break;
                case "Tz":
                    if (_args.Count >= 1) {
                        _textHScale = NumberAt(_args.Count - 1) / 100D;
                    }

                    break;
                case "Ts":
                    if (_args.Count >= 1) {
                        _textRise = NumberAt(_args.Count - 1);
                    }

                    break;
                case "Tr":
                    if (_args.Count >= 1) {
                        _textRenderingMode = ReadTextRenderingMode(NumberAt(_args.Count - 1));
                    }

                    break;
                case "'":
                    if (_args.Count >= 1) {
                        MoveToNextTextLine();
                        ShowText(_args[_args.Count - 1]);
                    }

                    break;
                case "\"":
                    if (_args.Count >= 3) {
                        _textWordSpacing = NumberAt(_args.Count - 3);
                        _textCharSpacing = NumberAt(_args.Count - 2);
                        MoveToNextTextLine();
                        ShowText(_args[_args.Count - 1]);
                    }

                    break;
                case "Tj":
                    if (_args.Count >= 1) {
                        ShowText(_args[_args.Count - 1]);
                    }

                    break;
                case "TJ":
                    if (_args.Count >= 1) {
                        ShowTextArray(_args[_args.Count - 1]);
                    }

                    break;
                case "Do":
                    if (!HasHiddenContent() &&
                        _args.Count >= 1 &&
                        _args[_args.Count - 1] is string name &&
                        !string.IsNullOrEmpty(name)) {
                        _invocations.Add(new PdfPageXObjectInvocation(name, _state.Transform, _state.ClipPath, _state.FillColor, _state.FillColorSpace, _state.FillOpacity, _state.StrokeColor, _state.StrokeColorSpace, _state.StrokeOpacity, _state.StrokeWidth, _state.StrokeDashStyle, _state.StrokeLineCap, _state.StrokeLineJoin, paintOrder));
                    }

                    break;
                case "BI":
                    if (TryReadInlineImage(out PdfPageInlineImage? inlineImage) && inlineImage != null && !HasHiddenContent()) {
                        _invocations.Add(new PdfPageXObjectInvocation(inlineImage, _state.Transform, _state.ClipPath, _state.FillColor, _state.FillColorSpace, _state.FillOpacity, _state.StrokeColor, _state.StrokeColorSpace, _state.StrokeOpacity, _state.StrokeWidth, _state.StrokeDashStyle, _state.StrokeLineCap, _state.StrokeLineJoin, paintOrder));
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

            _state = _state.WithGraphicsStateResource(resource);
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
                : PdfInlineImageDataScanner.FindLength(_content, dataStart);
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

            if (current == '<') {
                value = ReadInlineImageAngleObject();
                return true;
            }

            int start = _index;
            string token = ReadOperator();
            if (string.Equals(token, "true", StringComparison.Ordinal)) {
                value = new PdfBoolean(true);
                return true;
            }

            if (string.Equals(token, "false", StringComparison.Ordinal)) {
                value = new PdfBoolean(false);
                return true;
            }

            if (_index == start && _index < _content.Length) {
                _index++;
            }

            return false;
        }

        private PdfObject ReadInlineImageAngleObject() {
            if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                return ReadInlineImageDictionary();
            }

            _index++;
            int start = _index;
            while (_index < _content.Length && _content[_index] != '>') {
                _index++;
            }

            string hex = _content.Substring(start, _index - start);
            if (_index < _content.Length) {
                _index++;
            }

            return new PdfStringObj(PdfTextString.DecodeHexBytes(hex));
        }

        private PdfDictionary ReadInlineImageDictionary() {
            var dictionary = new PdfDictionary();
            _index += 2;
            while (_index < _content.Length) {
                SkipWhitespace();
                if (_index + 1 < _content.Length && _content[_index] == '>' && _content[_index + 1] == '>') {
                    _index += 2;
                    break;
                }

                if (_index >= _content.Length) {
                    break;
                }

                if (_content[_index] != '/') {
                    if (!TryReadInlineImageValue(out _)) {
                        _index++;
                    }

                    continue;
                }

                string key = ReadName();
                SkipWhitespace();
                if (_index >= _content.Length) {
                    break;
                }

                if (TryReadInlineImageValue(out PdfObject? value) && value != null) {
                    dictionary.Items[key] = value;
                }
            }

            return dictionary;
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

        private int GetInlineImageComponentCount(PdfDictionary dictionary) {
            string colorSpace = dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) && colorSpaceObject is PdfName colorSpaceName
                ? colorSpaceName.Name
                : "DeviceGray";
            switch (colorSpace) {
                case "DeviceRGB":
                    return 3;
                case "DeviceCMYK":
                    return 4;
                default:
                    return _colorSpaces != null && _colorSpaces.TryGetValue(colorSpace, out PdfPageColorSpaceKind resolved)
                        ? GetComponentCount(resolved)
                        : 1;
            }
        }

        private static int GetComponentCount(PdfPageColorSpaceKind colorSpace) {
            switch (colorSpace) {
                case PdfPageColorSpaceKind.DeviceRgb:
                case PdfPageColorSpaceKind.CalRgb:
                    return 3;
                case PdfPageColorSpaceKind.DeviceCmyk:
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
            return OfficeColorSpaceConverter.FromCmyk(
                NumberAt(startIndex),
                NumberAt(startIndex + 1),
                NumberAt(startIndex + 2),
                NumberAt(startIndex + 3));
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
                case PdfPageColorSpaceKind.CalGray:
                    color = PdfPageColorConverter.FromCalGray(NumberAt(startIndex));
                    return true;
                case PdfPageColorSpaceKind.CalRgb:
                    color = PdfPageColorConverter.FromCalRgb(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2));
                    return true;
                case PdfPageColorSpaceKind.Lab:
                    color = PdfPageColorConverter.FromLab(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2));
                    return true;
                default:
                    color = ReadGray(startIndex);
                    return true;
            }
        }

        private static int GetColorComponentCount(PdfPageColorSpaceKind colorSpace) {
            switch (colorSpace) {
                case PdfPageColorSpaceKind.DeviceRgb:
                case PdfPageColorSpaceKind.CalRgb:
                case PdfPageColorSpaceKind.Lab:
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
                case "CalGray":
                    colorSpace = PdfPageColorSpaceKind.CalGray;
                    return true;
                case "CalRGB":
                    colorSpace = PdfPageColorSpaceKind.CalRgb;
                    return true;
                case "Lab":
                    colorSpace = PdfPageColorSpaceKind.Lab;
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

        private static OfficeStrokeLineCap? ReadLineCap(double value) {
            int mode = (int)Math.Round(value);
            return mode switch {
                1 => OfficeStrokeLineCap.Round,
                2 => OfficeStrokeLineCap.Square,
                _ => OfficeStrokeLineCap.Butt
            };
        }

        private static OfficeStrokeLineJoin? ReadLineJoin(double value) {
            int mode = (int)Math.Round(value);
            return mode switch {
                1 => OfficeStrokeLineJoin.Round,
                2 => OfficeStrokeLineJoin.Bevel,
                _ => OfficeStrokeLineJoin.Miter
            };
        }

        private static OfficeStrokeDashStyle ReadDashStyle(double[] dashArray) {
            if (dashArray.Length == 0) {
                return OfficeStrokeDashStyle.Solid;
            }

            if (dashArray.Length >= 6) {
                return OfficeStrokeDashStyle.DashDotDot;
            }

            if (dashArray.Length >= 4) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (dashArray.Length >= 2) {
                return dashArray[0] <= dashArray[1] ? OfficeStrokeDashStyle.Dot : OfficeStrokeDashStyle.Dash;
            }

            return OfficeStrokeDashStyle.Solid;
        }

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

        private byte[] ReadLiteralStringBytes() {
            _index++;
            int depth = 1;
            bool escaped = false;
            var builder = new System.Text.StringBuilder();
            while (_index < _content.Length && depth > 0) {
                char ch = _content[_index++];
                if (escaped) {
                    builder.Append('\\');
                    builder.Append(ch);
                    escaped = false;
                } else if (ch == '\\') {
                    escaped = true;
                } else if (ch == '(') {
                    depth++;
                    builder.Append(ch);
                } else if (ch == ')') {
                    depth--;
                    if (depth > 0) {
                        builder.Append(ch);
                    }
                } else {
                    builder.Append(ch);
                }
            }

            return PdfStringParser.ParseLiteralToBytes(builder.ToString());
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

            var builder = new System.Text.StringBuilder(hex.Length);
            for (int i = 0; i < hex.Length; i++) {
                if (!char.IsWhiteSpace(hex[i])) {
                    builder.Append(hex[i]);
                }
            }

            hex = builder.ToString();
            if ((hex.Length & 1) == 1) {
                hex += "0";
            }

            var bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++) {
                bytes[i] = (byte)((HexNibble(hex[i * 2]) << 4) | HexNibble(hex[(i * 2) + 1]));
            }

            return bytes;

            static int HexNibble(char ch) {
                if (ch >= '0' && ch <= '9') return ch - '0';
                if (ch >= 'a' && ch <= 'f') return 10 + (ch - 'a');
                if (ch >= 'A' && ch <= 'F') return 10 + (ch - 'A');
                return 0;
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

        private List<object> ReadArray() {
            var items = new List<object>();
            _index++;
            while (_index < _content.Length) {
                SkipWhitespace();
                if (_index >= _content.Length) {
                    break;
                }

                char current = _content[_index];
                if (current == ']') {
                    _index++;
                    break;
                }

                if (current == '(') {
                    items.Add(ReadLiteralStringBytes());
                } else if (current == '<') {
                    if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                        SkipAngleObject();
                    } else {
                        items.Add(ReadHexStringBytes());
                    }
                } else if (IsNumberStart(current)) {
                    items.Add(ReadNumber());
                } else if (current == '/') {
                    items.Add(ReadName());
                } else if (current == '[') {
                    SkipArray();
                } else {
                    ReadOperator();
                }
            }

            return items;
        }

        private void SkipArray() {
            int depth = 1;
            _index++;
            while (_index < _content.Length && depth > 0) {
                char current = _content[_index];
                if (current == '(') {
                    SkipLiteralString();
                } else if (current == '<') {
                    SkipAngleObject();
                } else if (current == '[') {
                    depth++;
                    _index++;
                } else if (current == ']') {
                    depth--;
                    _index++;
                } else {
                    _index++;
                }
            }
        }

        private static bool TryGetNumberArray(object value, out double[] numbers) {
            if (value is double[] direct) {
                numbers = direct;
                return true;
            }

            if (value is List<object> items) {
                var collected = new List<double>(items.Count);
                for (int i = 0; i < items.Count; i++) {
                    if (items[i] is not double number) {
                        numbers = Array.Empty<double>();
                        return false;
                    }

                    collected.Add(number);
                }

                numbers = collected.ToArray();
                return true;
            }

            numbers = Array.Empty<double>();
            return false;
        }

        private static bool IsNumberStart(char ch) => ch == '-' || ch == '+' || ch == '.' || char.IsDigit(ch);

        private static bool IsDelimiter(char ch) =>
            char.IsWhiteSpace(ch) || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>' || ch == '%';

        private static int ReadTextRenderingMode(double value) {
            int mode = (int)Math.Round(value);
            return mode < 0 || mode > 7 ? 0 : mode;
        }

        private static bool AddsTextToClippingPath(int renderingMode) =>
            renderingMode >= 4 && renderingMode <= 7;

        private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
    }

    private readonly struct TextState {
        public TextState(bool inText, double size, double leading, double charSpacing, double wordSpacing, double hScale, double textRise, int textRenderingMode, Matrix2D textMatrix, Matrix2D lineMatrix) {
            InText = inText;
            Size = size;
            Leading = leading;
            CharSpacing = charSpacing;
            WordSpacing = wordSpacing;
            HScale = hScale;
            TextRise = textRise;
            TextRenderingMode = textRenderingMode;
            TextMatrix = textMatrix;
            LineMatrix = lineMatrix;
        }

        public static TextState Default { get; } = new TextState(false, 12D, 14.4D, 0D, 0D, 1D, 0D, 0, Matrix2D.Identity, Matrix2D.Identity);

        public bool InText { get; }

        public double Size { get; }

        public double Leading { get; }

        public double CharSpacing { get; }

        public double WordSpacing { get; }

        public double HScale { get; }

        public double TextRise { get; }

        public int TextRenderingMode { get; }

        public Matrix2D TextMatrix { get; }

        public Matrix2D LineMatrix { get; }
    }

    private readonly struct GraphicsState {
        private GraphicsState(
            Matrix2D transform,
            PdfPageClipPath? clipPath,
            OfficeColor fillColor,
            PdfPageColorSpaceKind fillColorSpace,
            double? fillOpacity,
            OfficeColor strokeColor,
            PdfPageColorSpaceKind strokeColorSpace,
            double? strokeOpacity,
            double strokeWidth,
            OfficeStrokeDashStyle? strokeDashStyle,
            OfficeStrokeLineCap? strokeLineCap,
            OfficeStrokeLineJoin? strokeLineJoin) {
            Transform = transform;
            ClipPath = clipPath;
            FillColor = fillColor;
            FillColorSpace = fillColorSpace;
            FillOpacity = fillOpacity;
            StrokeColor = strokeColor;
            StrokeColorSpace = strokeColorSpace;
            StrokeOpacity = strokeOpacity;
            StrokeWidth = strokeWidth;
            StrokeDashStyle = strokeDashStyle;
            StrokeLineCap = strokeLineCap;
            StrokeLineJoin = strokeLineJoin;
        }

        public Matrix2D Transform { get; }

        public PdfPageClipPath? ClipPath { get; }

        public OfficeColor FillColor { get; }

        public PdfPageColorSpaceKind FillColorSpace { get; }

        public double? FillOpacity { get; }

        public OfficeColor StrokeColor { get; }

        public PdfPageColorSpaceKind StrokeColorSpace { get; }

        public double? StrokeOpacity { get; }

        public double StrokeWidth { get; }

        public OfficeStrokeDashStyle? StrokeDashStyle { get; }

        public OfficeStrokeLineCap? StrokeLineCap { get; }

        public OfficeStrokeLineJoin? StrokeLineJoin { get; }

        public static GraphicsState Create(Matrix2D transform) =>
            Create(transform, null, PdfPageColorSpaceKind.DeviceGray, null, null, null, PdfPageColorSpaceKind.DeviceGray, null, null, null, null, null);

        public static GraphicsState Create(
            Matrix2D transform,
            OfficeColor? fillColor,
            PdfPageColorSpaceKind fillColorSpace,
            double? fillOpacity,
            PdfPageClipPath? clipPath,
            OfficeColor? strokeColor,
            PdfPageColorSpaceKind strokeColorSpace,
            double? strokeOpacity,
            double? strokeWidth,
            OfficeStrokeDashStyle? strokeDashStyle,
            OfficeStrokeLineCap? strokeLineCap,
            OfficeStrokeLineJoin? strokeLineJoin) =>
            new GraphicsState(
                transform,
                clipPath,
                fillColor ?? OfficeColor.Black,
                fillColorSpace,
                fillOpacity,
                strokeColor ?? OfficeColor.Black,
                strokeColorSpace,
                strokeOpacity,
                strokeWidth.HasValue ? ResolveStrokeWidth(strokeWidth.Value) : 1D,
                strokeDashStyle,
                strokeLineCap,
                strokeLineJoin);

        public GraphicsState WithTransform(Matrix2D transform) => new GraphicsState(transform, ClipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithClipPath(PdfPageClipPath clipPath) => new GraphicsState(Transform, clipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithFillColor(OfficeColor color) => new GraphicsState(Transform, ClipPath, color, FillColorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithFillColor(OfficeColor color, PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, ClipPath, color, colorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithFillColorSpace(PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, ClipPath, FillColor, colorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeColor(OfficeColor color) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, color, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeColor(OfficeColor color, PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, color, colorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeColorSpace(PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, colorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeWidth(double strokeWidth) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, strokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeDashStyle(OfficeStrokeDashStyle? strokeDashStyle) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, strokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeLineCap(OfficeStrokeLineCap? strokeLineCap) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, strokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeLineJoin(OfficeStrokeLineJoin? strokeLineJoin) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, strokeLineJoin);

        public GraphicsState WithGraphicsStateResource(PdfPageGraphicsStateResource resource) =>
            new GraphicsState(
                Transform,
                ClipPath,
                FillColor,
                FillColorSpace,
                resource.FillOpacity ?? FillOpacity,
                StrokeColor,
                StrokeColorSpace,
                resource.StrokeOpacity ?? StrokeOpacity,
                resource.StrokeWidth.HasValue ? ResolveStrokeWidth(resource.StrokeWidth.Value) : StrokeWidth,
                resource.StrokeDashStyle ?? StrokeDashStyle,
                resource.StrokeLineCap ?? StrokeLineCap,
                resource.StrokeLineJoin ?? StrokeLineJoin);
    }
}

internal readonly struct PdfPageXObjectInvocation {
    public PdfPageXObjectInvocation(
        string name,
        Matrix2D transform,
        PdfPageClipPath? clipPath,
        OfficeColor fillColor,
        PdfPageColorSpaceKind fillColorSpace,
        double? fillOpacity,
        OfficeColor strokeColor,
        PdfPageColorSpaceKind strokeColorSpace,
        double? strokeOpacity,
        double strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin,
        double paintOrder = 0D) {
        Name = name;
        InlineImage = null;
        Transform = transform;
        ClipPath = clipPath;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillOpacity = fillOpacity;
        StrokeColor = strokeColor;
        StrokeColorSpace = strokeColorSpace;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        PaintOrder = paintOrder;
    }

    public PdfPageXObjectInvocation(
        PdfPageInlineImage inlineImage,
        Matrix2D transform,
        PdfPageClipPath? clipPath,
        OfficeColor fillColor,
        PdfPageColorSpaceKind fillColorSpace,
        double? fillOpacity,
        OfficeColor strokeColor,
        PdfPageColorSpaceKind strokeColorSpace,
        double? strokeOpacity,
        double strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin,
        double paintOrder = 0D) {
        Name = inlineImage.ResourceName;
        InlineImage = inlineImage;
        Transform = transform;
        ClipPath = clipPath;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillOpacity = fillOpacity;
        StrokeColor = strokeColor;
        StrokeColorSpace = strokeColorSpace;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        PaintOrder = paintOrder;
    }

    public string Name { get; }

    public PdfPageInlineImage? InlineImage { get; }

    public Matrix2D Transform { get; }

    public PdfPageClipPath? ClipPath { get; }

    public OfficeColor FillColor { get; }

    public PdfPageColorSpaceKind FillColorSpace { get; }

    public double? FillOpacity { get; }

    public OfficeColor StrokeColor { get; }

    public PdfPageColorSpaceKind StrokeColorSpace { get; }

    public double? StrokeOpacity { get; }

    public double StrokeWidth { get; }

    public OfficeStrokeDashStyle? StrokeDashStyle { get; }

    public OfficeStrokeLineCap? StrokeLineCap { get; }

    public OfficeStrokeLineJoin? StrokeLineJoin { get; }

    public double PaintOrder { get; }
}
