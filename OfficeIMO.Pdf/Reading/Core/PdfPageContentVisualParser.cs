using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfPageContentVisualParser {
    public static IReadOnlyList<PdfPageVisualPrimitive> Parse(string content, double pageHeight) {
        return Parse(content, pageHeight, null);
    }

    public static IReadOnlyList<PdfPageVisualPrimitive> Parse(string content, double pageHeight, IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates) {
        return Parse(content, pageHeight, graphicsStates, null);
    }

    public static IReadOnlyList<PdfPageVisualPrimitive> Parse(
        string content,
        double pageHeight,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        IReadOnlyDictionary<string, PdfPageColorSpaceKind>? colorSpaces,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null) {
        return Parse(content, 0D, pageHeight, graphicsStates, colorSpaces, null, null, optionalContentVisibility);
    }

    public static IReadOnlyList<PdfPageVisualPrimitive> Parse(
        string content,
        double pageWidth,
        double pageHeight,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        IReadOnlyDictionary<string, PdfPageColorSpaceKind>? colorSpaces,
        IReadOnlyDictionary<string, PdfPageShadingResource>? shadings,
        IReadOnlyDictionary<string, PdfPageShadingPatternResource>? shadingPatterns,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null) {
        if (string.IsNullOrEmpty(content)) {
            return Array.Empty<PdfPageVisualPrimitive>();
        }

        var parser = new Parser(content, pageWidth, pageHeight, graphicsStates, colorSpaces, shadings, shadingPatterns, optionalContentVisibility, paintOrderBase, paintOrderScale, paintOrderOffset, initialClipPath);
        return parser.Parse();
    }

    private sealed class Parser {
        private readonly string _content;
        private readonly double _pageWidth;
        private readonly double _pageHeight;
        private readonly IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? _graphicsStates;
        private readonly IReadOnlyDictionary<string, PdfPageColorSpaceKind>? _colorSpaces;
        private readonly IReadOnlyDictionary<string, PdfPageShadingResource>? _shadings;
        private readonly IReadOnlyDictionary<string, PdfPageShadingPatternResource>? _shadingPatterns;
        private readonly PdfPageOptionalContentVisibility? _optionalContentVisibility;
        private readonly double _paintOrderBase;
        private readonly double _paintOrderScale;
        private readonly double _paintOrderOffset;
        private readonly List<PdfPageVisualPrimitive> _primitives = new List<PdfPageVisualPrimitive>();
        private readonly List<object> _args = new List<object>(8);
        private readonly Stack<GraphicsState> _stack = new Stack<GraphicsState>();
        private readonly Stack<bool> _hiddenContentStack = new Stack<bool>();
        private readonly List<(double X, double Y)> _path = new List<(double X, double Y)>();
        private readonly List<OfficePathCommand> _pathCommands = new List<OfficePathCommand>();
        private readonly GraphicsState _initialState;
        private GraphicsState _state;
        private int _currentSubpathStartIndex = -1;
        private bool _currentSubpathHasDraw;
        private int _index;

        public Parser(
            string content,
            double pageWidth,
            double pageHeight,
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
            IReadOnlyDictionary<string, PdfPageColorSpaceKind>? colorSpaces,
            IReadOnlyDictionary<string, PdfPageShadingResource>? shadings,
            IReadOnlyDictionary<string, PdfPageShadingPatternResource>? shadingPatterns,
            PdfPageOptionalContentVisibility? optionalContentVisibility,
            double paintOrderBase,
            double paintOrderScale,
            double paintOrderOffset,
            PdfPageClipPath? initialClipPath) {
            _content = content;
            _pageWidth = pageWidth;
            _pageHeight = pageHeight;
            _graphicsStates = graphicsStates;
            _colorSpaces = colorSpaces;
            _shadings = shadings;
            _shadingPatterns = shadingPatterns;
            _optionalContentVisibility = optionalContentVisibility;
            _paintOrderBase = paintOrderBase;
            _paintOrderScale = paintOrderScale;
            _paintOrderOffset = paintOrderOffset;
            _initialState = initialClipPath.HasValue
                ? GraphicsState.Default.WithClipPath(initialClipPath.Value)
                : GraphicsState.Default;
            _state = _initialState;
        }

        public IReadOnlyList<PdfPageVisualPrimitive> Parse() {
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
                    _args.Add(ReadNumberArray());
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

            return _primitives.Count == 0 ? Array.Empty<PdfPageVisualPrimitive>() : _primitives.AsReadOnly();
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
                case "w":
                    if (_args.Count >= 1) {
                        _state = _state.WithStrokeWidth(Math.Max(0D, NumberAt(_args.Count - 1)));
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
                    if (_args.Count >= 2 && _args[_args.Count - 2] is double[] dashArray) {
                        _state = _state.WithStrokeDashStyle(ReadDashStyle(dashArray));
                    }

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
                    if (_state.FillColorSpace == PdfPageColorSpaceKind.Pattern &&
                        _args.Count >= 1 &&
                        _args[_args.Count - 1] is string fillPatternName &&
                        TryReadShadingPattern(fillPatternName, out PdfPageShadingPatternResource fillPattern)) {
                        _state = _state.WithFillPattern(fillPattern);
                    } else if (TryReadColor(_state.FillColorSpace, out OfficeColor fillColor)) {
                        _state = _state.WithFillColor(fillColor);
                    }

                    break;
                case "SC":
                case "SCN":
                    if (_state.StrokeColorSpace == PdfPageColorSpaceKind.Pattern &&
                        _args.Count >= 1 &&
                        _args[_args.Count - 1] is string strokePatternName &&
                        TryReadShadingPattern(strokePatternName, out PdfPageShadingPatternResource strokePattern)) {
                        _state = _state.WithStrokePattern(strokePattern);
                    } else if (TryReadColor(_state.StrokeColorSpace, out OfficeColor strokeColor)) {
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
                        (double X, double Y) current = _path[_path.Count - 1];
                        CubicTo(
                            current.X,
                            current.Y,
                            NumberAt(_args.Count - 4),
                            NumberAt(_args.Count - 3),
                            NumberAt(_args.Count - 2),
                            NumberAt(_args.Count - 1),
                            firstControlAlreadyTransformed: true);
                    }

                    break;
                case "y":
                    if (_args.Count >= 4) {
                        double endX = NumberAt(_args.Count - 2);
                        double endY = NumberAt(_args.Count - 1);
                        CubicTo(
                            NumberAt(_args.Count - 4),
                            NumberAt(_args.Count - 3),
                            endX,
                            endY,
                            endX,
                            endY);
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
                case "f":
                case "F":
                    PaintPath(fill: true, stroke: false, OfficeFillRule.NonZero, paintOrder);
                    break;
                case "f*":
                    PaintPath(fill: true, stroke: false, OfficeFillRule.EvenOdd, paintOrder);
                    break;
                case "S":
                    PaintPath(fill: false, stroke: true, OfficeFillRule.NonZero, paintOrder);
                    break;
                case "s":
                    ClosePath();
                    PaintPath(fill: false, stroke: true, OfficeFillRule.NonZero, paintOrder);
                    break;
                case "B":
                    PaintPath(fill: true, stroke: true, OfficeFillRule.NonZero, paintOrder);
                    break;
                case "B*":
                    PaintPath(fill: true, stroke: true, OfficeFillRule.EvenOdd, paintOrder);
                    break;
                case "b":
                    ClosePath();
                    PaintPath(fill: true, stroke: true, OfficeFillRule.NonZero, paintOrder);
                    break;
                case "b*":
                    ClosePath();
                    PaintPath(fill: true, stroke: true, OfficeFillRule.EvenOdd, paintOrder);
                    break;
                case "sh":
                    if (_args.Count >= 1 && _args[_args.Count - 1] is string shadingName) {
                        PaintShading(shadingName, paintOrder);
                    }

                    break;
                case "n":
                    ClearPath();
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
            DiscardCurrentSubpathIfEmpty();
            var p0 = TransformPoint(x, y);
            var p1 = TransformPoint(x + width, y);
            var p2 = TransformPoint(x + width, y + height);
            var p3 = TransformPoint(x, y + height);
            _currentSubpathStartIndex = _path.Count;
            _currentSubpathHasDraw = true;
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

        private void PaintPath(bool fill, bool stroke, OfficeFillRule fillRule, double paintOrder) {
            if (_path.Count < 2) {
                ClearPath();
                return;
            }

            if (HasHiddenContent()) {
                ClearPath();
                return;
            }

            if (TryCreateAxisAlignedRectangle(out double x, out double y, out double width, out double height)) {
                OfficeLinearGradient? fillGradient = null;
                OfficeRadialGradient? fillRadialGradient = null;
                OfficeLinearGradient? strokeGradient = null;
                OfficeRadialGradient? strokeRadialGradient = null;
                if (fill && _state.FillPattern.HasValue) {
                    CreateShadingGradients(_state.FillPattern.Value.Shading, x, y, width, height, _state.FillPattern.Value.Matrix, out fillGradient, out fillRadialGradient);
                }

                if (stroke && _state.StrokePattern.HasValue) {
                    CreateShadingGradients(_state.StrokePattern.Value.Shading, x, y, width, height, _state.StrokePattern.Value.Matrix, out strokeGradient, out strokeRadialGradient);
                }

                _primitives.Add(PdfPageVisualPrimitive.Rectangle(
                    x,
                    y,
                    width,
                    height,
                    fill && fillGradient == null && fillRadialGradient == null ? _state.FillColor : null,
                    fillGradient,
                    fillRadialGradient,
                    stroke && _state.StrokeWidth > 0D && strokeGradient == null && strokeRadialGradient == null ? _state.StrokeColor : null,
                    strokeGradient,
                    strokeRadialGradient,
                    _state.StrokeWidth,
                    _state.StrokeDashStyle,
                    _state.StrokeLineCap,
                    _state.StrokeLineJoin,
                    fill ? _state.FillOpacity : null,
                    stroke && _state.StrokeWidth > 0D ? _state.StrokeOpacity : null,
                    _state.ClipPath,
                    paintOrder));
            } else if (stroke && _path.Count == 2) {
                AddLine(_path[0], _path[1], paintOrder);
            } else {
                OfficeLinearGradient? fillGradient = null;
                OfficeRadialGradient? fillRadialGradient = null;
                OfficeLinearGradient? strokeGradient = null;
                OfficeRadialGradient? strokeRadialGradient = null;
                if (fill &&
                    _state.FillPattern.HasValue &&
                    TryGetPathBounds(out double pathX, out double pathY, out double pathWidth, out double pathHeight)) {
                    CreateShadingGradients(_state.FillPattern.Value.Shading, pathX, pathY, pathWidth, pathHeight, _state.FillPattern.Value.Matrix, out fillGradient, out fillRadialGradient);
                }

                if (stroke &&
                    _state.StrokePattern.HasValue &&
                    TryGetPathBounds(out double strokePathX, out double strokePathY, out double strokePathWidth, out double strokePathHeight)) {
                    CreateShadingGradients(_state.StrokePattern.Value.Shading, strokePathX, strokePathY, strokePathWidth, strokePathHeight, _state.StrokePattern.Value.Matrix, out strokeGradient, out strokeRadialGradient);
                }

                if (PdfPageVisualPrimitive.TryCreatePath(
                    _pathCommands,
                    fill && fillGradient == null && fillRadialGradient == null ? _state.FillColor : null,
                    fillGradient,
                    fillRadialGradient,
                    stroke && _state.StrokeWidth > 0D && strokeGradient == null && strokeRadialGradient == null ? _state.StrokeColor : null,
                    strokeGradient,
                    strokeRadialGradient,
                    _state.StrokeWidth,
                    _state.StrokeDashStyle,
                    _state.StrokeLineCap,
                    _state.StrokeLineJoin,
                    fill ? _state.FillOpacity : null,
                    stroke && _state.StrokeWidth > 0D ? _state.StrokeOpacity : null,
                    fillRule,
                    _state.ClipPath,
                    paintOrder,
                    out PdfPageVisualPrimitive pathPrimitive)) {
                    _primitives.Add(pathPrimitive);
                }
            }

            ClearPath();
        }

        private bool TryReadShadingPattern(string patternName, out PdfPageShadingPatternResource pattern) {
            pattern = default;
            return _shadingPatterns != null && _shadingPatterns.TryGetValue(patternName, out pattern);
        }

        private void PaintShading(string shadingName, double paintOrder) {
            if (HasHiddenContent() ||
                _shadings == null ||
                !_shadings.TryGetValue(shadingName, out PdfPageShadingResource shading) ||
                !TryGetShadingPaintBounds(out double x, out double y, out double width, out double height)) {
                return;
            }

            CreateShadingGradients(shading, x, y, width, height, Matrix2D.Identity, out OfficeLinearGradient? linearGradient, out OfficeRadialGradient? radialGradient);
            if (radialGradient != null) {
                _primitives.Add(PdfPageVisualPrimitive.ShadedRectangle(x, y, width, height, radialGradient, _state.FillOpacity, _state.ClipPath, paintOrder));
            } else if (linearGradient != null) {
                _primitives.Add(PdfPageVisualPrimitive.ShadedRectangle(x, y, width, height, linearGradient, _state.FillOpacity, _state.ClipPath, paintOrder));
            }
        }

        private bool TryGetShadingPaintBounds(out double x, out double y, out double width, out double height) {
            if (_state.ClipPath.HasValue) {
                PdfPageClipPath clipPath = _state.ClipPath.Value;
                x = clipPath.X;
                y = clipPath.Y;
                width = clipPath.Width;
                height = clipPath.Height;
            } else {
                x = 0D;
                y = 0D;
                width = _pageWidth;
                height = _pageHeight;
            }

            return width > 0D && height > 0D;
        }

        private void CreateShadingGradients(PdfPageShadingResource shading, double x, double y, double width, double height, Matrix2D shadingTransform, out OfficeLinearGradient? linearGradient, out OfficeRadialGradient? radialGradient) {
            linearGradient = null;
            radialGradient = null;
            Matrix2D transform = Matrix2D.Multiply(_state.Transform, shadingTransform);
            (double X, double Y) start = transform.Transform(shading.X0, shading.Y0);
            (double X, double Y) end = transform.Transform(shading.X1, shading.Y1);
            double paintWidth = Math.Max(width, 0.0001D);
            double paintHeight = Math.Max(height, 0.0001D);
            double startX = Clamp01((start.X - x) / paintWidth);
            double startY = Clamp01((ToTop(start.Y) - y) / paintHeight);
            double endX = Clamp01((end.X - x) / paintWidth);
            double endY = Clamp01((ToTop(end.Y) - y) / paintHeight);
            if (shading.IsRadial) {
                double radiusScale = Math.Max(paintWidth, paintHeight);
                double startRadius = radiusScale <= 0D ? 0D : Math.Max(0D, shading.R0) / radiusScale;
                double endRadius = radiusScale <= 0D ? 0D : Math.Max(0D, shading.R1) / radiusScale;
                if (NearlyEqual(startX, endX) && NearlyEqual(startY, endY) && NearlyEqual(startRadius, endRadius)) {
                    endRadius = Math.Min(1D, startRadius + 0.5D);
                }

                radialGradient = new OfficeRadialGradient(
                    startX,
                    startY,
                    startRadius,
                    endX,
                    endY,
                    endRadius,
                    new OfficeGradientStop(0D, shading.StartColor),
                    new OfficeGradientStop(1D, shading.EndColor));
                return;
            }

            if (NearlyEqual(startX, endX) && NearlyEqual(startY, endY)) {
                linearGradient = OfficeLinearGradient.Horizontal(shading.StartColor, shading.EndColor);
                return;
            }

            linearGradient = new OfficeLinearGradient(
                startX,
                startY,
                endX,
                endY,
                new OfficeGradientStop(0D, shading.StartColor),
                new OfficeGradientStop(1D, shading.EndColor));
        }

        private bool TryGetPathBounds(out double x, out double y, out double width, out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (_path.Count == 0) {
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

            x = left;
            y = top;
            return true;
        }

        private bool TryCreateAxisAlignedRectangle(out double x, out double y, out double width, out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (_path.Count < 4) {
                return false;
            }

            if (CountMoveCommands() != 1) {
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

            x = left;
            y = top;
            return true;
        }

        private void AddLine((double X, double Y) start, (double X, double Y) end, double paintOrder) {
            double x1 = start.X;
            double y1 = ToTop(start.Y);
            double x2 = end.X;
            double y2 = ToTop(end.Y);
            if (NearlyEqual(x1, x2) && NearlyEqual(y1, y2)) {
                return;
            }

            OfficeLinearGradient? strokeGradient = null;
            OfficeRadialGradient? strokeRadialGradient = null;
            if (_state.StrokePattern.HasValue) {
                double lineX = Math.Min(x1, x2);
                double lineY = Math.Min(y1, y2);
                double lineWidth = Math.Abs(x2 - x1);
                double lineHeight = Math.Abs(y2 - y1);
                CreateShadingGradients(_state.StrokePattern.Value.Shading, lineX, lineY, lineWidth, lineHeight, _state.StrokePattern.Value.Matrix, out strokeGradient, out strokeRadialGradient);
            }

            _primitives.Add(PdfPageVisualPrimitive.Line(
                x1,
                y1,
                x2,
                y2,
                strokeGradient == null && strokeRadialGradient == null ? _state.StrokeColor : null,
                strokeGradient,
                strokeRadialGradient,
                _state.StrokeWidth,
                _state.StrokeDashStyle,
                _state.StrokeLineCap,
                _state.StrokeLineJoin,
                _state.StrokeOpacity,
                _state.ClipPath,
                paintOrder));
        }

        private void ApplyGraphicsStateResource(string name) {
            if (_graphicsStates == null ||
                !_graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource resource)) {
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
                _optionalContentVisibility?.IsHiddenAny(references.ObjectNumbers) == true));

        private void MoveTo(double x, double y) {
            DiscardCurrentSubpathIfEmpty();
            (double X, double Y) point = TransformPoint(x, y);
            _currentSubpathStartIndex = _path.Count;
            _currentSubpathHasDraw = false;
            _path.Add(point);
            _pathCommands.Add(OfficePathCommand.MoveTo(ToOfficePoint(point)));
        }

        private void CaptureClipPath(OfficeFillRule fillRule) {
            if (_path.Count < 2) {
                return;
            }

            if (TryCreateAxisAlignedRectangle(out double x, out double y, out double width, out double height)) {
                _state = _state.WithClipPath(PdfPageClipPath.ResolveActiveClip(_state.ClipPath, PdfPageClipPath.Rectangle(x, y, width, height)));
                return;
            }

            if (PdfPageClipPath.TryCreatePath(_pathCommands, fillRule, out PdfPageClipPath clipPath)) {
                _state = _state.WithClipPath(PdfPageClipPath.ResolveActiveClip(_state.ClipPath, clipPath));
            }
        }

        private void ClosePath() {
            if (_path.Count == 0 || _currentSubpathStartIndex < 0 || _currentSubpathStartIndex >= _path.Count || !_currentSubpathHasDraw) {
                return;
            }

            _path.Add(_path[_currentSubpathStartIndex]);
            _pathCommands.Add(OfficePathCommand.Close());
        }

        private void LineTo(double x, double y) {
            if (_currentSubpathStartIndex < 0) {
                MoveTo(x, y);
                return;
            }

            (double X, double Y) point = TransformPoint(x, y);
            _path.Add(point);
            _currentSubpathHasDraw = true;
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
            _currentSubpathHasDraw = true;
            _pathCommands.Add(OfficePathCommand.CubicBezierTo(ToOfficePoint(control1), ToOfficePoint(control2), ToOfficePoint(end)));
        }

        private void ClearPath() {
            _path.Clear();
            _pathCommands.Clear();
            _currentSubpathStartIndex = -1;
            _currentSubpathHasDraw = false;
        }

        private void DiscardCurrentSubpathIfEmpty() {
            if (_currentSubpathHasDraw ||
                _currentSubpathStartIndex < 0 ||
                _currentSubpathStartIndex >= _path.Count) {
                return;
            }

            _path.RemoveRange(_currentSubpathStartIndex, _path.Count - _currentSubpathStartIndex);
            if (_pathCommands.Count > 0 && _pathCommands[_pathCommands.Count - 1].Kind == OfficePathCommandKind.MoveTo) {
                _pathCommands.RemoveAt(_pathCommands.Count - 1);
            }

            _currentSubpathStartIndex = -1;
        }

        private int CountMoveCommands() {
            int count = 0;
            for (int i = 0; i < _pathCommands.Count; i++) {
                if (_pathCommands[i].Kind == OfficePathCommandKind.MoveTo) {
                    count++;
                }
            }

            return count;
        }

        private (double X, double Y) TransformPoint(double x, double y) => _state.Transform.Transform(x, y);

        private double ToTop(double pdfY) => _pageHeight - pdfY;

        private OfficePoint ToOfficePoint((double X, double Y) point) => new OfficePoint(point.X, ToTop(point.Y));

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
            if (colorSpace == PdfPageColorSpaceKind.Pattern) {
                return false;
            }

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
                case "Pattern":
                    colorSpace = PdfPageColorSpaceKind.Pattern;
                    return true;
                default:
                    if (_colorSpaces != null && _colorSpaces.TryGetValue(name, out colorSpace)) {
                        return true;
                    }

                    colorSpace = PdfPageColorSpaceKind.DeviceGray;
                    return false;
            }
        }

        private static OfficeStrokeLineCap? ReadLineCap(double value) {
            int lineCap = (int)Math.Round(value);
            switch (lineCap) {
                case 0:
                    return OfficeStrokeLineCap.Butt;
                case 1:
                    return OfficeStrokeLineCap.Round;
                case 2:
                    return OfficeStrokeLineCap.Square;
                default:
                    return null;
            }
        }

        private static OfficeStrokeLineJoin? ReadLineJoin(double value) {
            int lineJoin = (int)Math.Round(value);
            switch (lineJoin) {
                case 0:
                    return OfficeStrokeLineJoin.Miter;
                case 1:
                    return OfficeStrokeLineJoin.Round;
                case 2:
                    return OfficeStrokeLineJoin.Bevel;
                default:
                    return null;
            }
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

        private double NumberAt(int index) => _args[index] is double value ? value : 0D;

        private static byte ToByte(double value) {
            return (byte)Math.Round(Clamp01(value) * 255D);
        }

        private static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;

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

        private double[] ReadNumberArray() {
            var numbers = new List<double>();
            int depth = 1;
            _index++;
            while (_index < _content.Length && depth > 0) {
                char ch = _content[_index];
                if (ch == '(') {
                    SkipLiteralString();
                } else if (ch == '<') {
                    SkipAngleObject();
                } else if (IsNumberStart(ch)) {
                    numbers.Add(ReadNumber());
                } else {
                    if (ch == '[') {
                        depth++;
                    } else if (ch == ']') {
                        depth--;
                    }

                    _index++;
                }
            }

            return numbers.ToArray();
        }

        private static bool IsNumberStart(char ch) => ch == '-' || ch == '+' || ch == '.' || char.IsDigit(ch);

        private static bool IsDelimiter(char ch) =>
            char.IsWhiteSpace(ch) || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>' || ch == '%';

        private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
    }

    private readonly struct GraphicsState {
        private GraphicsState(Matrix2D transform, OfficeColor fillColor, PdfPageShadingPatternResource? fillPattern, OfficeColor strokeColor, PdfPageShadingPatternResource? strokePattern, PdfPageColorSpaceKind fillColorSpace, PdfPageColorSpaceKind strokeColorSpace, double strokeWidth, OfficeStrokeDashStyle strokeDashStyle, OfficeStrokeLineCap? strokeLineCap, OfficeStrokeLineJoin? strokeLineJoin, double? fillOpacity, double? strokeOpacity, PdfPageClipPath? clipPath) {
            Transform = transform;
            FillColor = fillColor;
            FillPattern = fillPattern;
            StrokeColor = strokeColor;
            StrokePattern = strokePattern;
            FillColorSpace = fillColorSpace;
            StrokeColorSpace = strokeColorSpace;
            StrokeWidth = strokeWidth;
            StrokeDashStyle = strokeDashStyle;
            StrokeLineCap = strokeLineCap;
            StrokeLineJoin = strokeLineJoin;
            FillOpacity = fillOpacity;
            StrokeOpacity = strokeOpacity;
            ClipPath = clipPath;
        }

        public Matrix2D Transform { get; }

        public OfficeColor FillColor { get; }

        public PdfPageShadingPatternResource? FillPattern { get; }

        public OfficeColor StrokeColor { get; }

        public PdfPageShadingPatternResource? StrokePattern { get; }

        public PdfPageColorSpaceKind FillColorSpace { get; }

        public PdfPageColorSpaceKind StrokeColorSpace { get; }

        public double StrokeWidth { get; }

        public OfficeStrokeDashStyle StrokeDashStyle { get; }

        public OfficeStrokeLineCap? StrokeLineCap { get; }

        public OfficeStrokeLineJoin? StrokeLineJoin { get; }

        public double? FillOpacity { get; }

        public double? StrokeOpacity { get; }

        public PdfPageClipPath? ClipPath { get; }

        public static GraphicsState Default => new GraphicsState(Matrix2D.Identity, OfficeColor.Black, null, OfficeColor.Black, null, PdfPageColorSpaceKind.DeviceGray, PdfPageColorSpaceKind.DeviceGray, 1D, OfficeStrokeDashStyle.Solid, null, null, null, null, null);

        public GraphicsState WithTransform(Matrix2D transform) => new GraphicsState(transform, FillColor, FillPattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithFillColor(OfficeColor color) => new GraphicsState(Transform, color, null, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithFillColor(OfficeColor color, PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, color, null, StrokeColor, StrokePattern, colorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithFillPattern(PdfPageShadingPatternResource pattern) => new GraphicsState(Transform, FillColor, pattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeColor(OfficeColor color) => new GraphicsState(Transform, FillColor, FillPattern, color, null, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeColor(OfficeColor color, PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, FillColor, FillPattern, color, null, FillColorSpace, colorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokePattern(PdfPageShadingPatternResource pattern) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, pattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithFillColorSpace(PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, FillColor, colorSpace == PdfPageColorSpaceKind.Pattern ? FillPattern : null, StrokeColor, StrokePattern, colorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeColorSpace(PdfPageColorSpaceKind colorSpace) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, colorSpace == PdfPageColorSpaceKind.Pattern ? StrokePattern : null, FillColorSpace, colorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeWidth(double strokeWidth) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, strokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeDashStyle(OfficeStrokeDashStyle strokeDashStyle) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, strokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeLineCap(OfficeStrokeLineCap? strokeLineCap) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, strokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeLineJoin(OfficeStrokeLineJoin? strokeLineJoin) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, strokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithOpacity(double? fillOpacity, double? strokeOpacity) =>
            new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, fillOpacity ?? FillOpacity, strokeOpacity ?? StrokeOpacity, ClipPath);

        public GraphicsState WithGraphicsStateResource(PdfPageGraphicsStateResource resource) =>
            new GraphicsState(
                Transform,
                FillColor,
                FillPattern,
                StrokeColor,
                StrokePattern,
                FillColorSpace,
                StrokeColorSpace,
                resource.StrokeWidth ?? StrokeWidth,
                resource.StrokeDashStyle ?? StrokeDashStyle,
                resource.StrokeLineCap ?? StrokeLineCap,
                resource.StrokeLineJoin ?? StrokeLineJoin,
                resource.FillOpacity ?? FillOpacity,
                resource.StrokeOpacity ?? StrokeOpacity,
                ClipPath);

        public GraphicsState WithClipPath(PdfPageClipPath clipPath) =>
            new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, clipPath);
    }
}
