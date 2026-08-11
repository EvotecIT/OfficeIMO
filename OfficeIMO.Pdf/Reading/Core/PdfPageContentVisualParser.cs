using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfPageContentVisualParser {
    private const double HairlineStrokeWidth = double.PositiveInfinity;
    private const double RenderedHairlineStrokeWidth = 0.25D;

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
        IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null) {
        return Parse(content, 0D, pageHeight, graphicsStates, colorSpaces, null, null, null, optionalContentVisibility);
    }

    public static IReadOnlyList<PdfPageVisualPrimitive> Parse(
        string content,
        double pageWidth,
        double pageHeight,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces,
        IReadOnlyDictionary<string, PdfPageShadingResource>? shadings,
        IReadOnlyDictionary<string, PdfPageShadingPatternResource>? shadingPatterns,
        IReadOnlyDictionary<string, PdfPageTilingPatternResource>? tilingPatterns,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        double? initialFillOpacity = null,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        double? initialStrokeOpacity = null,
        double? initialStrokeWidth = null,
        OfficeStrokeDashStyle? initialStrokeDashStyle = null,
        OfficeStrokeLineCap? initialStrokeLineCap = null,
        OfficeStrokeLineJoin? initialStrokeLineJoin = null,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations,
        IReadOnlyDictionary<string, PdfPageColorSpace>? patternBaseColorSpaces = null,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands,
        Action<PdfPageVisualPrimitive>? primitiveVisitor = null,
        bool retainPrimitiveData = true,
        bool scaleStrokeWidthWithTransform = false,
        Action? unsupportedShadingTransformVisitor = null,
        bool requireExactType3ShadingProjection = false) {
        if (string.IsNullOrEmpty(content)) {
            return Array.Empty<PdfPageVisualPrimitive>();
        }

        var parser = new Parser(content, pageWidth, pageHeight, graphicsStates, colorSpaces, shadings, shadingPatterns, tilingPatterns, optionalContentVisibility, paintOrderBase, paintOrderScale, paintOrderOffset, initialClipPath, initialFillColor, initialFillColorSpace, initialFillOpacity, initialStrokeColor, initialStrokeColorSpace, initialStrokeOpacity, initialStrokeWidth, initialStrokeDashStyle, initialStrokeLineCap, initialStrokeLineJoin, maxOperations, patternBaseColorSpaces, maxNestingDepth, maxOperands, primitiveVisitor, retainPrimitiveData, scaleStrokeWidthWithTransform, unsupportedShadingTransformVisitor, requireExactType3ShadingProjection);
        return parser.Parse();
    }

    internal static void CreateShadingGradients(
        PdfPageShadingPatternResource pattern,
        double x,
        double y,
        double width,
        double height,
        Matrix2D paintTransform,
        double pageHeight,
        out OfficeLinearGradient? linearGradient,
        out OfficeRadialGradient? radialGradient) {
        if (!IsSupportedShadingTransform(pattern, paintTransform)) {
            linearGradient = null;
            radialGradient = null;
            return;
        }
        Parser.CreateShadingGradients(
            pattern.Shading,
            x,
            y,
            width,
            height,
            Matrix2D.Multiply(paintTransform, pattern.Matrix),
            pageHeight,
            out linearGradient,
            out radialGradient);
    }

    internal static bool IsSupportedShadingTransform(PdfPageShadingPatternResource pattern, Matrix2D paintTransform) {
        if (!pattern.IsSupported) return false;
        if (!pattern.Shading.IsRadial) return true;
        return IsRepresentableRadialShadingTransform(Matrix2D.Multiply(paintTransform, pattern.Matrix));
    }

    internal static bool IsSupportedShadingStrokeTransform(PdfPageShadingPatternResource pattern, Matrix2D paintTransform) {
        if (!pattern.IsSupported) return false;
        if (!pattern.Shading.IsRadial) return true;
        if (!IsSupportedShadingTransform(pattern, paintTransform)) return false;
        Matrix2D transform = Matrix2D.Multiply(paintTransform, pattern.Matrix);
        double firstLengthSquared = (transform.A * transform.A) + (transform.B * transform.B);
        double secondLengthSquared = (transform.C * transform.C) + (transform.D * transform.D);
        double scale = Math.Max(firstLengthSquared, secondLengthSquared);
        double dot = (transform.A * transform.C) + (transform.B * transform.D);
        double orthogonalityTolerance = Math.Sqrt(firstLengthSquared * secondLengthSquared) * 0.000000001D;
        return firstLengthSquared > 0D &&
               secondLengthSquared > 0D &&
               Math.Abs(firstLengthSquared - secondLengthSquared) <= scale * 0.000000001D &&
               Math.Abs(dot) <= orthogonalityTolerance;
    }

    internal static bool IsSupportedExactShadingPlacement(
        PdfPageShadingResource shading,
        Matrix2D transform,
        double x,
        double y,
        double width,
        double height,
        double pageHeight) {
        if (shading.IsRadial) return true;
        if (width <= 0D || height <= 0D) return false;
        (double X, double Y) start = transform.Transform(shading.X0, shading.Y0);
        (double X, double Y) end = transform.Transform(shading.X1, shading.Y1);
        double x0 = (start.X - x) / width;
        double y0 = ((pageHeight - start.Y) - y) / height;
        double x1 = (end.X - x) / width;
        double y1 = ((pageHeight - end.Y) - y) / height;
        double dx = x1 - x0;
        double dy = y1 - y0;
        double t0 = 0D;
        double t1 = 1D;
        return ClipExactLineParameter(-dx, x0, ref t0, ref t1) &&
               ClipExactLineParameter(dx, 1D - x0, ref t0, ref t1) &&
               ClipExactLineParameter(-dy, y0, ref t0, ref t1) &&
               ClipExactLineParameter(dy, 1D - y0, ref t0, ref t1) &&
               t1 > t0;
    }

    private static bool ClipExactLineParameter(double p, double q, ref double t0, ref double t1) {
        if (Math.Abs(p) <= 0.000000001D) return q >= 0D;
        double ratio = q / p;
        if (p < 0D) {
            if (ratio > t1) return false;
            if (ratio > t0) t0 = ratio;
        } else {
            if (ratio < t0) return false;
            if (ratio < t1) t1 = ratio;
        }
        return true;
    }

    private static bool IsRepresentableRadialShadingTransform(Matrix2D transform) {
        double firstLengthSquared = (transform.A * transform.A) + (transform.B * transform.B);
        double secondLengthSquared = (transform.C * transform.C) + (transform.D * transform.D);
        if (firstLengthSquared <= 0D || secondLengthSquared <= 0D ||
            double.IsNaN(firstLengthSquared) || double.IsNaN(secondLengthSquared) ||
            double.IsInfinity(firstLengthSquared) || double.IsInfinity(secondLengthSquared)) return false;
        double scale = Math.Max(firstLengthSquared, secondLengthSquared);
        double dot = (transform.A * transform.C) + (transform.B * transform.D);
        double orthogonalityTolerance = Math.Sqrt(firstLengthSquared * secondLengthSquared) * 0.000000001D;
        if (Math.Abs(firstLengthSquared - secondLengthSquared) <= scale * 0.000000001D &&
            Math.Abs(dot) <= orthogonalityTolerance) return true;

        double componentScale = Math.Max(
            Math.Max(Math.Abs(transform.A), Math.Abs(transform.B)),
            Math.Max(Math.Abs(transform.C), Math.Abs(transform.D)));
        double componentTolerance = componentScale * 0.000000001D;
        return Math.Abs(transform.B) <= componentTolerance && Math.Abs(transform.C) <= componentTolerance;
    }

    private static double ResolveStrokeWidth(double value) {
        if (value < 0D) {
            return 0D;
        }

        return Math.Abs(value) <= 0.001D ? HairlineStrokeWidth : value;
    }

    private sealed class Parser {
        private readonly string _content;
        private readonly double _pageWidth;
        private readonly double _pageHeight;
        private readonly IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? _graphicsStates;
        private readonly IReadOnlyDictionary<string, PdfPageColorSpace>? _colorSpaces;
        private readonly IReadOnlyDictionary<string, PdfPageColorSpace>? _patternBaseColorSpaces;
        private readonly IReadOnlyDictionary<string, PdfPageShadingResource>? _shadings;
        private readonly IReadOnlyDictionary<string, PdfPageShadingPatternResource>? _shadingPatterns;
        private readonly IReadOnlyDictionary<string, PdfPageTilingPatternResource>? _tilingPatterns;
        private readonly PdfPageOptionalContentVisibility? _optionalContentVisibility;
        private readonly double _paintOrderBase;
        private readonly double _paintOrderScale;
        private readonly double _paintOrderOffset;
        private readonly List<PdfPageVisualPrimitive>? _primitives;
        private readonly Action<PdfPageVisualPrimitive>? _primitiveVisitor;
        private readonly bool _retainPrimitiveData;
        private readonly bool _scaleStrokeWidthWithTransform;
        private readonly Action? _unsupportedShadingTransformVisitor;
        private readonly bool _requireExactType3ShadingProjection;
        private readonly List<object> _args = new List<object>(8);
        private readonly Stack<GraphicsState> _stack = new Stack<GraphicsState>();
        private readonly Stack<(PdfPageTilingPatternResource? Fill, OfficeColor? FillTint, PdfPageColorSpace? FillBase, PdfPageTilingPatternResource? Stroke, OfficeColor? StrokeTint, PdfPageColorSpace? StrokeBase)> _tilingStack = new Stack<(PdfPageTilingPatternResource? Fill, OfficeColor? FillTint, PdfPageColorSpace? FillBase, PdfPageTilingPatternResource? Stroke, OfficeColor? StrokeTint, PdfPageColorSpace? StrokeBase)>();
        private readonly Stack<bool> _hiddenContentStack = new Stack<bool>();
        private readonly List<(double X, double Y)> _path = new List<(double X, double Y)>();
        private readonly List<OfficePathCommand> _pathCommands = new List<OfficePathCommand>();
        private readonly GraphicsState _initialState;
        private GraphicsState _state;
        private PdfPageTilingPatternResource? _fillTilingPattern;
        private OfficeColor? _fillTilingTint;
        private PdfPageColorSpace? _fillPatternBaseColorSpace;
        private PdfPageTilingPatternResource? _strokeTilingPattern;
        private OfficeColor? _strokeTilingTint;
        private PdfPageColorSpace? _strokePatternBaseColorSpace;
        private int _currentSubpathStartIndex = -1;
        private int _currentOperatorIndex;
        private bool _currentSubpathHasDraw;
        private readonly int _maxOperations;
        private readonly int _maxNestingDepth;
        private readonly int _maxOperands;

        public Parser(
            string content,
            double pageWidth,
            double pageHeight,
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
            IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces,
            IReadOnlyDictionary<string, PdfPageShadingResource>? shadings,
            IReadOnlyDictionary<string, PdfPageShadingPatternResource>? shadingPatterns,
            IReadOnlyDictionary<string, PdfPageTilingPatternResource>? tilingPatterns,
            PdfPageOptionalContentVisibility? optionalContentVisibility,
            double paintOrderBase,
            double paintOrderScale,
            double paintOrderOffset,
            PdfPageClipPath? initialClipPath,
            OfficeColor? initialFillColor,
            PdfPageColorSpace initialFillColorSpace,
            double? initialFillOpacity,
            OfficeColor? initialStrokeColor,
            PdfPageColorSpace initialStrokeColorSpace,
            double? initialStrokeOpacity,
            double? initialStrokeWidth,
            OfficeStrokeDashStyle? initialStrokeDashStyle,
            OfficeStrokeLineCap? initialStrokeLineCap,
            OfficeStrokeLineJoin? initialStrokeLineJoin,
            int maxOperations,
            IReadOnlyDictionary<string, PdfPageColorSpace>? patternBaseColorSpaces,
            int maxNestingDepth,
            int maxOperands,
            Action<PdfPageVisualPrimitive>? primitiveVisitor,
            bool retainPrimitiveData,
            bool scaleStrokeWidthWithTransform,
            Action? unsupportedShadingTransformVisitor,
            bool requireExactType3ShadingProjection) {
            _content = content;
            _pageWidth = pageWidth;
            _pageHeight = pageHeight;
            _graphicsStates = graphicsStates;
            _colorSpaces = colorSpaces;
            _patternBaseColorSpaces = patternBaseColorSpaces;
            _shadings = shadings;
            _shadingPatterns = shadingPatterns;
            _tilingPatterns = tilingPatterns;
            _optionalContentVisibility = optionalContentVisibility;
            _paintOrderBase = paintOrderBase;
            _paintOrderScale = paintOrderScale;
            _paintOrderOffset = paintOrderOffset;
            _maxOperations = maxOperations;
            _maxNestingDepth = maxNestingDepth;
            _maxOperands = maxOperands;
            _primitiveVisitor = primitiveVisitor;
            _retainPrimitiveData = primitiveVisitor == null || retainPrimitiveData;
            _scaleStrokeWidthWithTransform = scaleStrokeWidthWithTransform;
            _unsupportedShadingTransformVisitor = unsupportedShadingTransformVisitor;
            _requireExactType3ShadingProjection = requireExactType3ShadingProjection;
            _primitives = primitiveVisitor == null ? new List<PdfPageVisualPrimitive>() : null;
            GraphicsState initialState = initialFillColor.HasValue
                ? GraphicsState.Default.WithFillColor(initialFillColor.Value, initialFillColorSpace)
                : GraphicsState.Default;
            if (initialFillOpacity.HasValue) {
                initialState = initialState.WithOpacity(initialFillOpacity, null);
            }

            if (initialStrokeColor.HasValue) {
                initialState = initialState.WithStrokeColor(initialStrokeColor.Value, initialStrokeColorSpace);
            }

            if (initialStrokeOpacity.HasValue) {
                initialState = initialState.WithOpacity(null, initialStrokeOpacity);
            }

            if (initialStrokeWidth.HasValue) {
                initialState = initialState.WithStrokeWidth(ResolveStrokeWidth(initialStrokeWidth.Value));
            }

            if (initialStrokeDashStyle.HasValue) {
                initialState = initialState.WithStrokeDashStyle(initialStrokeDashStyle.Value);
            }

            if (initialStrokeLineCap.HasValue) {
                initialState = initialState.WithStrokeLineCap(initialStrokeLineCap);
            }

            if (initialStrokeLineJoin.HasValue) {
                initialState = initialState.WithStrokeLineJoin(initialStrokeLineJoin);
            }

            _initialState = initialClipPath.HasValue
                ? initialState.WithClipPath(initialClipPath.Value)
                : initialState;
            _state = _initialState;
        }

        public IReadOnlyList<PdfPageVisualPrimitive> Parse() {
            PdfContentStreamInterpreter.Interpret(
                _content,
                _maxOperations,
                operation => {
                    _args.Clear();
                    _args.AddRange(operation.Operands);
                    _currentOperatorIndex = operation.OperatorOffset;
                    ApplyOperator(
                        operation.Name,
                        GetPaintOrder(operation.OperatorOffset),
                        operation.HasInvalidOperands);
                },
                maxNestingDepth: _maxNestingDepth,
                maxOperands: _maxOperands);

            return _primitives == null || _primitives.Count == 0
                ? Array.Empty<PdfPageVisualPrimitive>()
                : _primitives.AsReadOnly();
        }

        private void AddPrimitive(PdfPageVisualPrimitive primitive) {
            primitive = primitive.WithSourceOperatorIndex(_currentOperatorIndex);
            if (_primitiveVisitor != null) {
                _primitiveVisitor(primitive);
            } else {
                _primitives!.Add(primitive);
            }
        }

        private double GetPaintOrder(int operatorIndex) => _paintOrderBase + ((operatorIndex + _paintOrderOffset) * _paintOrderScale);

        private void ApplyOperator(string op, double paintOrder, bool hasInvalidOperands) {
            switch (op) {
                case "q":
                    _stack.Push(_state);
                    _tilingStack.Push((_fillTilingPattern, _fillTilingTint, _fillPatternBaseColorSpace, _strokeTilingPattern, _strokeTilingTint, _strokePatternBaseColorSpace));
                    break;
                case "Q":
                    _state = _stack.Count > 0 ? _stack.Pop() : _initialState;
                    if (_tilingStack.Count > 0) {
                        (PdfPageTilingPatternResource? Fill, OfficeColor? FillTint, PdfPageColorSpace? FillBase, PdfPageTilingPatternResource? Stroke, OfficeColor? StrokeTint, PdfPageColorSpace? StrokeBase) restored = _tilingStack.Pop();
                        _fillTilingPattern = restored.Fill;
                        _fillTilingTint = restored.FillTint;
                        _fillPatternBaseColorSpace = restored.FillBase;
                        _strokeTilingPattern = restored.Stroke;
                        _strokeTilingTint = restored.StrokeTint;
                        _strokePatternBaseColorSpace = restored.StrokeBase;
                    } else {
                        _fillTilingPattern = null;
                        _fillTilingTint = null;
                        _fillPatternBaseColorSpace = null;
                        _strokeTilingPattern = null;
                        _strokeTilingTint = null;
                        _strokePatternBaseColorSpace = null;
                    }
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
                        TryReadColorSpace(fillColorSpaceName, out PdfPageColorSpace fillColorSpace)) {
                        _state = _state.WithFillColorSpace(fillColorSpace);
                        _fillPatternBaseColorSpace = ReadPatternBaseColorSpace(fillColorSpaceName, fillColorSpace);
                        _fillTilingPattern = null;
                        _fillTilingTint = null;
                    }

                    break;
                case "CS":
                    if (_args.Count >= 1 &&
                        _args[_args.Count - 1] is string strokeColorSpaceName &&
                        TryReadColorSpace(strokeColorSpaceName, out PdfPageColorSpace strokeColorSpace)) {
                        _state = _state.WithStrokeColorSpace(strokeColorSpace);
                        _strokePatternBaseColorSpace = ReadPatternBaseColorSpace(strokeColorSpaceName, strokeColorSpace);
                        _strokeTilingPattern = null;
                        _strokeTilingTint = null;
                    }

                    break;
                case "sc":
                case "scn":
                    if (_state.FillColorSpace == PdfPageColorSpaceKind.Pattern &&
                        _args.Count >= 1 &&
                        _args[_args.Count - 1] is string fillPatternName) {
                        if (TryReadTilingPattern(fillPatternName, out PdfPageTilingPatternResource fillTilingPattern)) {
                            _fillTilingPattern = fillTilingPattern;
                            _fillTilingTint = fillTilingPattern.Uncolored ? ReadPatternTint(_state.FillColor, _fillPatternBaseColorSpace) : null;
                            _state = _state.WithoutFillPattern();
                        } else if (TryReadShadingPattern(fillPatternName, out PdfPageShadingPatternResource fillPattern)) {
                            _fillTilingPattern = null;
                            _fillTilingTint = null;
                            _state = _state.WithFillPattern(fillPattern);
                        }
                    } else if (TryReadColor(_state.FillColorSpace, out OfficeColor fillColor)) {
                        _fillTilingPattern = null;
                        _fillTilingTint = null;
                        _state = _state.WithFillColor(fillColor);
                    }

                    break;
                case "SC":
                case "SCN":
                    if (_state.StrokeColorSpace == PdfPageColorSpaceKind.Pattern &&
                        _args.Count >= 1 &&
                        _args[_args.Count - 1] is string strokePatternName) {
                        if (TryReadTilingPattern(strokePatternName, out PdfPageTilingPatternResource strokeTilingPattern)) {
                            _strokeTilingPattern = strokeTilingPattern;
                            _strokeTilingTint = strokeTilingPattern.Uncolored ? ReadPatternTint(_state.StrokeColor, _strokePatternBaseColorSpace) : null;
                            _state = _state.WithoutStrokePattern();
                        } else if (TryReadShadingPattern(strokePatternName, out PdfPageShadingPatternResource strokePattern)) {
                            _strokeTilingPattern = null;
                            _strokeTilingTint = null;
                            _state = _state.WithStrokePattern(strokePattern);
                        }
                    } else if (TryReadColor(_state.StrokeColorSpace, out OfficeColor strokeColor)) {
                        _strokeTilingPattern = null;
                        _strokeTilingTint = null;
                        _state = _state.WithStrokeColor(strokeColor);
                    }

                    break;
                case "rg":
                    if (_args.Count >= 3) {
                        _fillTilingPattern = null;
                        _fillTilingTint = null;
                        _state = _state.WithFillColor(ReadRgb(_args.Count - 3), PdfPageColorSpaceKind.DeviceRgb);
                    }

                    break;
                case "RG":
                    if (_args.Count >= 3) {
                        _strokeTilingPattern = null;
                        _strokeTilingTint = null;
                        _state = _state.WithStrokeColor(ReadRgb(_args.Count - 3), PdfPageColorSpaceKind.DeviceRgb);
                    }

                    break;
                case "g":
                    if (_args.Count >= 1) {
                        _fillTilingPattern = null;
                        _fillTilingTint = null;
                        _state = _state.WithFillColor(ReadGray(_args.Count - 1), PdfPageColorSpaceKind.DeviceGray);
                    }

                    break;
                case "G":
                    if (_args.Count >= 1) {
                        _strokeTilingPattern = null;
                        _strokeTilingTint = null;
                        _state = _state.WithStrokeColor(ReadGray(_args.Count - 1), PdfPageColorSpaceKind.DeviceGray);
                    }

                    break;
                case "k":
                    if (_args.Count >= 4) {
                        _fillTilingPattern = null;
                        _fillTilingTint = null;
                        _state = _state.WithFillColor(ReadCmyk(_args.Count - 4), PdfPageColorSpaceKind.DeviceCmyk);
                    }

                    break;
                case "K":
                    if (_args.Count >= 4) {
                        _strokeTilingPattern = null;
                        _strokeTilingTint = null;
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
                case "BI":
                    break;
                case "BDC":
                    _hiddenContentStack.Push(
                        hasInvalidOperands ||
                        IsHiddenOptionalContent(
                            _args.Count > 1 ? _args[_args.Count - 2] : null,
                            _args.Count > 0 ? _args[_args.Count - 1] : null));
                    break;
                case "BMC":
                    _hiddenContentStack.Push(hasInvalidOperands);
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
                PdfPageTilingPatternPaint? fillTilingPaint = fill ? CreateTilingPatternPaint(_fillTilingPattern, _fillTilingTint, _state.FillOpacity) : null;
                PdfPageTilingPatternPaint? strokeTilingPaint = stroke && _state.StrokeWidth > 0D
                    ? CreateTilingPatternPaint(_strokeTilingPattern, _strokeTilingTint, _state.StrokeOpacity)
                    : null;
                OfficeLinearGradient? fillGradient = null;
                OfficeRadialGradient? fillRadialGradient = null;
                OfficeLinearGradient? strokeGradient = null;
                OfficeRadialGradient? strokeRadialGradient = null;
                if (fill && _state.FillPattern.HasValue) {
                    CreateShadingGradients(_state.FillPattern.Value, x, y, width, height, out fillGradient, out fillRadialGradient);
                }

                if (stroke && _state.StrokePattern.HasValue) {
                    CreateShadingGradients(_state.StrokePattern.Value, x, y, width, height, out strokeGradient, out strokeRadialGradient);
                    RejectUnsupportedShadingStroke(ref strokeGradient, ref strokeRadialGradient);
                }

                AddPrimitive(PdfPageVisualPrimitive.Rectangle(
                    x,
                    y,
                    width,
                    height,
                    fill && !_state.FillPattern.HasValue && fillGradient == null && fillRadialGradient == null && fillTilingPaint == null ? _state.FillColor : null,
                    fillGradient,
                    fillRadialGradient,
                    stroke && _state.StrokeWidth > 0D && !_state.StrokePattern.HasValue && strokeGradient == null && strokeRadialGradient == null && _strokeTilingPattern == null ? _state.StrokeColor : null,
                    strokeGradient,
                    strokeRadialGradient,
                    GetRenderedStrokeWidth(),
                    _state.StrokeDashStyle,
                    _state.StrokeLineCap,
                    _state.StrokeLineJoin,
                    fill ? _state.FillOpacity : null,
                    stroke && _state.StrokeWidth > 0D ? _state.StrokeOpacity : null,
                    _state.ClipPath,
                    paintOrder,
                    fillTilingPaint,
                    strokeTilingPaint));
            } else if (stroke && IsSingleLinePath()) {
                AddLine(_path[0], _path[1], paintOrder);
            } else {
                OfficeLinearGradient? fillGradient = null;
                OfficeRadialGradient? fillRadialGradient = null;
                OfficeLinearGradient? strokeGradient = null;
                OfficeRadialGradient? strokeRadialGradient = null;
                if (fill &&
                    _state.FillPattern.HasValue &&
                    TryGetPathBounds(out double pathX, out double pathY, out double pathWidth, out double pathHeight)) {
                    CreateShadingGradients(_state.FillPattern.Value, pathX, pathY, pathWidth, pathHeight, out fillGradient, out fillRadialGradient);
                }

                if (stroke &&
                    _state.StrokePattern.HasValue &&
                    TryGetPathBounds(out double strokePathX, out double strokePathY, out double strokePathWidth, out double strokePathHeight)) {
                    CreateShadingGradients(_state.StrokePattern.Value, strokePathX, strokePathY, strokePathWidth, strokePathHeight, out strokeGradient, out strokeRadialGradient);
                    RejectUnsupportedShadingStroke(ref strokeGradient, ref strokeRadialGradient);
                }

                IReadOnlyList<OfficePathCommand> pathCommands = fill && _retainPrimitiveData
                    ? CloseFilledSubpaths(_pathCommands)
                    : _pathCommands;
                if (PdfPageVisualPrimitive.TryCreatePath(
                    pathCommands,
                    fill && !_state.FillPattern.HasValue && fillGradient == null && fillRadialGradient == null && _fillTilingPattern == null ? _state.FillColor : null,
                    fillGradient,
                    fillRadialGradient,
                    stroke && _state.StrokeWidth > 0D && !_state.StrokePattern.HasValue && strokeGradient == null && strokeRadialGradient == null && _strokeTilingPattern == null ? _state.StrokeColor : null,
                    strokeGradient,
                    strokeRadialGradient,
                    GetRenderedStrokeWidth(),
                    _state.StrokeDashStyle,
                    _state.StrokeLineCap,
                    _state.StrokeLineJoin,
                    fill ? _state.FillOpacity : null,
                    stroke && _state.StrokeWidth > 0D ? _state.StrokeOpacity : null,
                    fillRule,
                    _state.ClipPath,
                    paintOrder,
                    fill ? CreateTilingPatternPaint(_fillTilingPattern, _fillTilingTint, _state.FillOpacity) : null,
                    stroke && _state.StrokeWidth > 0D ? CreateTilingPatternPaint(_strokeTilingPattern, _strokeTilingTint, _state.StrokeOpacity) : null,
                    _retainPrimitiveData,
                    out PdfPageVisualPrimitive pathPrimitive)) {
                    AddPrimitive(pathPrimitive);
                } else if (stroke && _state.StrokeWidth > 0D) {
                    AddStrokedPathSegments(pathCommands, paintOrder);
                }
            }

            ClearPath();
        }

        private bool TryReadShadingPattern(string patternName, out PdfPageShadingPatternResource pattern) {
            pattern = default;
            return _shadingPatterns != null && _shadingPatterns.TryGetValue(patternName, out pattern);
        }

        private bool TryReadTilingPattern(string patternName, out PdfPageTilingPatternResource pattern) {
            pattern = null!;
            return _tilingPatterns != null && _tilingPatterns.TryGetValue(patternName, out pattern!);
        }

        private PdfPageTilingPatternPaint? CreateTilingPatternPaint(PdfPageTilingPatternResource? resource, OfficeColor? tint, double? opacity) {
            if (resource == null) return null;
            var localToPattern = new Matrix2D(1D, 0D, 0D, -1D, resource.BoundingBoxX, resource.BoundingBoxTop);
            Matrix2D combined = Matrix2D.Multiply(
                new Matrix2D(1D, 0D, 0D, -1D, 0D, _pageHeight),
                Matrix2D.Multiply(_state.Transform, Matrix2D.Multiply(resource.Matrix, localToPattern)));
            return new PdfPageTilingPatternPaint(
                resource,
                new OfficeTransform(combined.A, combined.B, combined.C, combined.D, combined.E, combined.F),
                resource.Uncolored ? tint : null,
                opacity ?? 1D);
        }

        private OfficeColor ReadPatternTint(OfficeColor fallback, PdfPageColorSpace? baseColorSpace) {
            if (baseColorSpace.HasValue && TryReadColor(baseColorSpace.Value, out OfficeColor color)) return color;
            int componentCount = _args.Count > 0 && _args[_args.Count - 1] is string ? _args.Count - 1 : _args.Count;
            if (componentCount >= 3) return ReadRgb(componentCount - 3);
            if (componentCount >= 1) return ReadGray(componentCount - 1);
            return fallback;
        }

        private PdfPageColorSpace? ReadPatternBaseColorSpace(string name, PdfPageColorSpace colorSpace) {
            if (colorSpace != PdfPageColorSpaceKind.Pattern || _patternBaseColorSpaces == null) return null;
            return _patternBaseColorSpaces.TryGetValue(name, out PdfPageColorSpace baseColorSpace) ? baseColorSpace : null;
        }

        private void PaintShading(string shadingName, double paintOrder) {
            if (HasHiddenContent() ||
                _shadings == null ||
                !_shadings.TryGetValue(shadingName, out PdfPageShadingResource shading) ||
                !TryGetShadingPaintBounds(out double x, out double y, out double width, out double height)) {
                return;
            }
            if (_requireExactType3ShadingProjection && !shading.SupportsExactType3Projection) {
                _unsupportedShadingTransformVisitor?.Invoke();
                return;
            }

            CreateShadingGradients(shading, x, y, width, height, Matrix2D.Identity, out OfficeLinearGradient? linearGradient, out OfficeRadialGradient? radialGradient);
            if (radialGradient != null) {
                AddPrimitive(PdfPageVisualPrimitive.ShadedRectangle(x, y, width, height, radialGradient, _state.FillOpacity, _state.ClipPath, paintOrder));
            } else if (linearGradient != null) {
                AddPrimitive(PdfPageVisualPrimitive.ShadedRectangle(x, y, width, height, linearGradient, _state.FillOpacity, _state.ClipPath, paintOrder));
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
            Matrix2D combined = Matrix2D.Multiply(_state.Transform, shadingTransform);
            if (shading.IsRadial && !IsRepresentableRadialShadingTransform(combined)) {
                linearGradient = null;
                radialGradient = null;
                _unsupportedShadingTransformVisitor?.Invoke();
                return;
            }
            if (_requireExactType3ShadingProjection &&
                !IsSupportedExactShadingPlacement(shading, combined, x, y, width, height, _pageHeight)) {
                linearGradient = null;
                radialGradient = null;
                _unsupportedShadingTransformVisitor?.Invoke();
                return;
            }
            CreateShadingGradients(
                shading,
                x,
                y,
                width,
                height,
                combined,
                _pageHeight,
                out linearGradient,
                out radialGradient);
        }

        private void CreateShadingGradients(PdfPageShadingPatternResource pattern, double x, double y, double width, double height, out OfficeLinearGradient? linearGradient, out OfficeRadialGradient? radialGradient) {
            if (!pattern.IsSupported ||
                (_requireExactType3ShadingProjection && !pattern.SupportsExactType3Projection)) {
                linearGradient = null;
                radialGradient = null;
                _unsupportedShadingTransformVisitor?.Invoke();
                return;
            }
            CreateShadingGradients(pattern.Shading, x, y, width, height, pattern.Matrix, out linearGradient, out radialGradient);
        }

        internal static void CreateShadingGradients(
            PdfPageShadingResource shading,
            double x,
            double y,
            double width,
            double height,
            Matrix2D transform,
            double pageHeight,
            out OfficeLinearGradient? linearGradient,
            out OfficeRadialGradient? radialGradient) {
            linearGradient = null;
            radialGradient = null;
            (double X, double Y) start = transform.Transform(shading.X0, shading.Y0);
            (double X, double Y) end = transform.Transform(shading.X1, shading.Y1);
            double paintWidth = Math.Max(width, 0.0001D);
            double paintHeight = Math.Max(height, 0.0001D);
            double rawStartX = (start.X - x) / paintWidth;
            double rawStartY = ((pageHeight - start.Y) - y) / paintHeight;
            double rawEndX = (end.X - x) / paintWidth;
            double rawEndY = ((pageHeight - end.Y) - y) / paintHeight;
            double startX = Clamp01(rawStartX);
            double startY = Clamp01(rawStartY);
            double endX = Clamp01(rawEndX);
            double endY = Clamp01(rawEndY);
            if (shading.IsRadial) {
                double startRadiusX = TransformRadiusX(transform, shading.R0) / paintWidth;
                double startRadiusY = TransformRadiusY(transform, shading.R0) / paintHeight;
                double endRadiusX = TransformRadiusX(transform, shading.R1) / paintWidth;
                double endRadiusY = TransformRadiusY(transform, shading.R1) / paintHeight;
                if (NearlyEqual(rawStartX, rawEndX)
                    && NearlyEqual(rawStartY, rawEndY)
                    && NearlyEqual(startRadiusX, endRadiusX)
                    && NearlyEqual(startRadiusY, endRadiusY)) {
                    endRadiusX = startRadiusX + 0.5D;
                    endRadiusY = startRadiusY + 0.5D;
                }

                IReadOnlyList<OfficeGradientStop> stops = shading.Stops;
                radialGradient = endRadiusX > 0D && endRadiusY > 0D
                    ? new OfficeRadialGradient(rawStartX, rawStartY, startRadiusX, startRadiusY, rawEndX, rawEndY, endRadiusX, endRadiusY, stops)
                    : new OfficeRadialGradient(
                        rawStartX,
                        rawStartY,
                        Math.Max(startRadiusX, startRadiusY),
                        rawEndX,
                        rawEndY,
                        Math.Max(endRadiusX, endRadiusY),
                        stops);
                return;
            }

            if (!TryClipLinearGradientToUnitBounds(rawStartX, rawStartY, rawEndX, rawEndY, shading.Stops, out OfficeLinearGradient? clippedLinearGradient)) {
                clippedLinearGradient = null;
            }

            if (clippedLinearGradient != null) {
                linearGradient = clippedLinearGradient;
                return;
            }

            if (NearlyEqual(startX, endX) && NearlyEqual(startY, endY)) {
                linearGradient = new OfficeLinearGradient(0D, 0.5D, 1D, 0.5D, shading.Stops);
                return;
            }

            linearGradient = new OfficeLinearGradient(
                startX,
                startY,
                endX,
                endY,
                shading.Stops);
        }

        private static bool TryClipLinearGradientToUnitBounds(
            double x0,
            double y0,
            double x1,
            double y1,
            IReadOnlyList<OfficeGradientStop> stops,
            out OfficeLinearGradient? gradient) {
            gradient = null;
            double dx = x1 - x0;
            double dy = y1 - y0;
            double t0 = 0D;
            double t1 = 1D;
            if (!ClipLineParameter(-dx, x0, ref t0, ref t1) ||
                !ClipLineParameter(dx, 1D - x0, ref t0, ref t1) ||
                !ClipLineParameter(-dy, y0, ref t0, ref t1) ||
                !ClipLineParameter(dy, 1D - y0, ref t0, ref t1) ||
                t1 <= t0) {
                return false;
            }

            double clippedStartX = Clamp01(x0 + (dx * t0));
            double clippedStartY = Clamp01(y0 + (dy * t0));
            double clippedEndX = Clamp01(x0 + (dx * t1));
            double clippedEndY = Clamp01(y0 + (dy * t1));
            if (NearlyEqual(clippedStartX, clippedEndX) && NearlyEqual(clippedStartY, clippedEndY)) {
                return false;
            }

            gradient = new OfficeLinearGradient(
                clippedStartX,
                clippedStartY,
                clippedEndX,
                clippedEndY,
                ClipGradientStops(stops, t0, t1));
            return true;
        }

        private static List<OfficeGradientStop> ClipGradientStops(IReadOnlyList<OfficeGradientStop> stops, double start, double end) {
            var result = new List<OfficeGradientStop>(stops.Count + 2) {
                new OfficeGradientStop(0D, EvaluateGradientColor(stops, start))
            };
            double span = end - start;
            for (int i = 0; i < stops.Count; i++) {
                double offset = stops[i].Offset;
                if (offset > start && offset < end) {
                    result.Add(new OfficeGradientStop((offset - start) / span, stops[i].Color));
                }
            }
            result.Add(new OfficeGradientStop(1D, EvaluateGradientColor(stops, end)));
            return result;
        }

        private static OfficeColor EvaluateGradientColor(IReadOnlyList<OfficeGradientStop> stops, double offset) {
            double value = Clamp01(offset);
            for (int i = 1; i < stops.Count; i++) {
                OfficeGradientStop right = stops[i];
                if (value > right.Offset) continue;
                OfficeGradientStop left = stops[i - 1];
                double span = right.Offset - left.Offset;
                if (span <= 0D) return right.Color;
                return InterpolateColor(left.Color, right.Color, (value - left.Offset) / span);
            }
            return stops[stops.Count - 1].Color;
        }

        private static bool ClipLineParameter(double p, double q, ref double t0, ref double t1) {
            if (NearlyEqual(p, 0D)) {
                return q >= 0D;
            }

            double r = q / p;
            if (p < 0D) {
                if (r > t1) {
                    return false;
                }

                if (r > t0) {
                    t0 = r;
                }
            } else {
                if (r < t0) {
                    return false;
                }

                if (r < t1) {
                    t1 = r;
                }
            }

            return true;
        }

        private static OfficeColor InterpolateColor(OfficeColor start, OfficeColor end, double ratio) {
            double clamped = Clamp01(ratio);
            return OfficeColor.FromRgba(
                InterpolateByte(start.R, end.R, clamped),
                InterpolateByte(start.G, end.G, clamped),
                InterpolateByte(start.B, end.B, clamped),
                InterpolateByte(start.A, end.A, clamped));
        }

        private static byte InterpolateByte(byte start, byte end, double ratio) =>
            (byte)Math.Round(start + ((end - start) * ratio));

        private static double TransformRadiusX(Matrix2D transform, double radius) =>
            TransformRadius(radius, Math.Sqrt((transform.A * transform.A) + (transform.B * transform.B)));

        private static double TransformRadiusY(Matrix2D transform, double radius) =>
            TransformRadius(radius, Math.Sqrt((transform.C * transform.C) + (transform.D * transform.D)));

        private static double TransformRadius(double radius, double scale) {
            if (radius <= 0D) return 0D;
            return !double.IsNaN(scale) && !double.IsInfinity(scale) && scale > 0D ? radius * scale : radius;
        }

        private bool TryGetPathBounds(out double x, out double y, out double width, out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (_pathCommands.Count == 0) {
                return false;
            }

            bool hasPoint = false;
            double left = 0D;
            double right = 0D;
            double top = 0D;
            double bottom = 0D;

            void IncludePoint(OfficePoint point) {
                if (!hasPoint) {
                    left = point.X;
                    right = point.X;
                    top = point.Y;
                    bottom = point.Y;
                    hasPoint = true;
                    return;
                }

                left = Math.Min(left, point.X);
                right = Math.Max(right, point.X);
                top = Math.Min(top, point.Y);
                bottom = Math.Max(bottom, point.Y);
            }

            foreach (OfficePathCommand command in _pathCommands) {
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                    case OfficePathCommandKind.LineTo:
                        IncludePoint(command.Point);
                        break;
                    case OfficePathCommandKind.QuadraticBezierTo:
                        IncludePoint(command.ControlPoint1);
                        IncludePoint(command.Point);
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        IncludePoint(command.ControlPoint1);
                        IncludePoint(command.ControlPoint2);
                        IncludePoint(command.Point);
                        break;
                }
            }

            if (!hasPoint) {
                return false;
            }

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

        private bool IsSingleLinePath() =>
            _path.Count == 2 &&
            _pathCommands.Count == 2 &&
            _pathCommands[0].Kind == OfficePathCommandKind.MoveTo &&
            _pathCommands[1].Kind == OfficePathCommandKind.LineTo;

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
                CreateShadingGradients(_state.StrokePattern.Value, lineX, lineY, lineWidth, lineHeight, out strokeGradient, out strokeRadialGradient);
                RejectUnsupportedShadingStroke(ref strokeGradient, ref strokeRadialGradient);
            }

            AddPrimitive(PdfPageVisualPrimitive.Line(
                x1,
                y1,
                x2,
                y2,
                !_state.StrokePattern.HasValue && strokeGradient == null && strokeRadialGradient == null && _strokeTilingPattern == null ? _state.StrokeColor : null,
                strokeGradient,
                strokeRadialGradient,
                GetRenderedStrokeWidth(),
                _state.StrokeDashStyle,
                _state.StrokeLineCap,
                _state.StrokeLineJoin,
                _state.StrokeOpacity,
                _state.ClipPath,
                paintOrder,
                _state.StrokeWidth > 0D ? CreateTilingPatternPaint(_strokeTilingPattern, _strokeTilingTint, _state.StrokeOpacity) : null));
        }

        private void RejectUnsupportedShadingStroke(
            ref OfficeLinearGradient? strokeGradient,
            ref OfficeRadialGradient? strokeRadialGradient) {
            if (strokeGradient == null && strokeRadialGradient == null) return;
            bool unsupportedDash = _state.StrokeDashStyle != OfficeStrokeDashStyle.Solid;
            bool unsupportedRadialTransform = _state.StrokePattern.HasValue &&
                _state.StrokePattern.Value.Shading.IsRadial &&
                !IsSupportedShadingStrokeTransform(_state.StrokePattern.Value, _state.Transform);
            if (!unsupportedDash && !unsupportedRadialTransform) return;
            strokeGradient = null;
            strokeRadialGradient = null;
            _unsupportedShadingTransformVisitor?.Invoke();
        }

        private double GetRenderedStrokeWidth() {
            double strokeWidth = _state.StrokeWidth;
            if (double.IsPositiveInfinity(strokeWidth)) return RenderedHairlineStrokeWidth;
            if (!_scaleStrokeWidthWithTransform || strokeWidth <= 0D) {
                return strokeWidth;
            }

            Matrix2D transform = _state.Transform;
            double squaredScale = ((transform.A * transform.A) +
                                   (transform.B * transform.B) +
                                   (transform.C * transform.C) +
                                   (transform.D * transform.D)) / 2D;
            return squaredScale > 0D && !double.IsNaN(squaredScale) && !double.IsInfinity(squaredScale)
                ? strokeWidth * Math.Sqrt(squaredScale)
                : strokeWidth;
        }

        private void AddStrokedPathSegments(IReadOnlyList<OfficePathCommand> pathCommands, double paintOrder) {
            OfficePoint current = default;
            OfficePoint subpathStart = default;
            bool hasCurrent = false;
            bool hasSubpathStart = false;
            for (int i = 0; i < pathCommands.Count; i++) {
                OfficePathCommand command = pathCommands[i];
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        current = command.Point;
                        subpathStart = command.Point;
                        hasCurrent = true;
                        hasSubpathStart = true;
                        break;
                    case OfficePathCommandKind.LineTo:
                        if (hasCurrent) {
                            AddLine(ToPdfPoint(current), ToPdfPoint(command.Point), paintOrder);
                        }

                        current = command.Point;
                        hasCurrent = true;
                        break;
                    case OfficePathCommandKind.Close:
                        if (hasCurrent && hasSubpathStart) {
                            AddLine(ToPdfPoint(current), ToPdfPoint(subpathStart), paintOrder);
                        }

                        hasCurrent = false;
                        hasSubpathStart = false;
                        break;
                    default:
                        hasCurrent = false;
                        hasSubpathStart = false;
                        break;
                }
            }
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
                _optionalContentVisibility?.IsHidden(references) == true) ||
             (property is PdfContentDictionary dictionary &&
                dictionary.OptionalContentReferences is not null &&
                _optionalContentVisibility?.IsHidden(dictionary.OptionalContentReferences) == true));

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

        private static List<OfficePathCommand> CloseFilledSubpaths(List<OfficePathCommand> commands) {
            if (commands.Count == 0) {
                return commands;
            }

            var closed = new List<OfficePathCommand>(commands.Count + 4);
            bool hasOpenSubpath = false;
            bool subpathHasDraw = false;
            for (int i = 0; i < commands.Count; i++) {
                OfficePathCommand command = commands[i];
                if (command.Kind == OfficePathCommandKind.MoveTo) {
                    if (hasOpenSubpath && subpathHasDraw) {
                        closed.Add(OfficePathCommand.Close());
                    }

                    hasOpenSubpath = true;
                    subpathHasDraw = false;
                    closed.Add(command);
                    continue;
                }

                closed.Add(command);
                if (command.Kind == OfficePathCommandKind.Close) {
                    hasOpenSubpath = false;
                    subpathHasDraw = false;
                } else if (command.Kind == OfficePathCommandKind.LineTo ||
                    command.Kind == OfficePathCommandKind.QuadraticBezierTo ||
                    command.Kind == OfficePathCommandKind.CubicBezierTo) {
                    subpathHasDraw = true;
                }
            }

            if (hasOpenSubpath && subpathHasDraw) {
                closed.Add(OfficePathCommand.Close());
            }

            return closed;
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

        private (double X, double Y) ToPdfPoint(OfficePoint point) => (point.X, _pageHeight - point.Y);

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

        private bool TryReadColor(PdfPageColorSpace colorSpace, out OfficeColor color) {
            color = OfficeColor.Black;
            if (colorSpace == PdfPageColorSpaceKind.Pattern) {
                return false;
            }

            int componentCount = colorSpace.ComponentCount;
            int endIndex = _args.Count;
            while (endIndex > 0 && !(_args[endIndex - 1] is double)) {
                endIndex--;
            }

            if (endIndex < componentCount) {
                return false;
            }

            int startIndex = endIndex - componentCount;
            var components = new double[componentCount];
            for (int i = 0; i < componentCount; i++) components[i] = NumberAt(startIndex + i);
            return colorSpace.TryConvertColor(components, out color);
        }

        private bool TryReadColorSpace(string name, out PdfPageColorSpace colorSpace) {
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

        private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
    }

    private readonly struct GraphicsState {
        private GraphicsState(Matrix2D transform, OfficeColor fillColor, PdfPageShadingPatternResource? fillPattern, OfficeColor strokeColor, PdfPageShadingPatternResource? strokePattern, PdfPageColorSpace fillColorSpace, PdfPageColorSpace strokeColorSpace, double strokeWidth, OfficeStrokeDashStyle strokeDashStyle, OfficeStrokeLineCap? strokeLineCap, OfficeStrokeLineJoin? strokeLineJoin, double? fillOpacity, double? strokeOpacity, PdfPageClipPath? clipPath) {
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

        public PdfPageColorSpace FillColorSpace { get; }

        public PdfPageColorSpace StrokeColorSpace { get; }

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

        public GraphicsState WithFillColor(OfficeColor color, PdfPageColorSpace colorSpace) => new GraphicsState(Transform, color, null, StrokeColor, StrokePattern, colorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithFillPattern(PdfPageShadingPatternResource pattern) => new GraphicsState(Transform, FillColor, pattern, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithoutFillPattern() => new GraphicsState(Transform, FillColor, null, StrokeColor, StrokePattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeColor(OfficeColor color) => new GraphicsState(Transform, FillColor, FillPattern, color, null, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeColor(OfficeColor color, PdfPageColorSpace colorSpace) => new GraphicsState(Transform, FillColor, FillPattern, color, null, FillColorSpace, colorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokePattern(PdfPageShadingPatternResource pattern) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, pattern, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithoutStrokePattern() => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, null, FillColorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithFillColorSpace(PdfPageColorSpace colorSpace) => new GraphicsState(Transform, FillColor, colorSpace == PdfPageColorSpaceKind.Pattern ? FillPattern : null, StrokeColor, StrokePattern, colorSpace, StrokeColorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

        public GraphicsState WithStrokeColorSpace(PdfPageColorSpace colorSpace) => new GraphicsState(Transform, FillColor, FillPattern, StrokeColor, colorSpace == PdfPageColorSpaceKind.Pattern ? StrokePattern : null, FillColorSpace, colorSpace, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin, FillOpacity, StrokeOpacity, ClipPath);

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
                resource.StrokeWidth.HasValue ? ResolveStrokeWidth(resource.StrokeWidth.Value) : StrokeWidth,
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
