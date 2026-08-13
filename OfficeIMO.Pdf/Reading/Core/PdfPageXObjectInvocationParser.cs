using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfPageXObjectInvocationParser {
    private const double HairlineStrokeWidth = double.PositiveInfinity;

    private static double ResolveStrokeWidth(double value) {
        if (value < 0D) {
            return 0D;
        }

        return value == 0D ? HairlineStrokeWidth : value;
    }

    public static IReadOnlyList<PdfPageXObjectInvocation> Parse(string content, Matrix2D baseTransform, double pageHeight) {
        return Parse(content, baseTransform, pageHeight, null);
    }

    public static IReadOnlyList<PdfPageXObjectInvocation> Parse(string content, Matrix2D baseTransform, double pageHeight, IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces) {
        return Parse(content, baseTransform, pageHeight, null, colorSpaces);
    }

    public static IReadOnlyList<PdfPageXObjectInvocation> Parse(
        string content,
        Matrix2D baseTransform,
        double pageHeight,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        double? initialFillOpacity = null,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        double? initialStrokeOpacity = null,
        double? initialStrokeWidth = null,
        OfficeStrokeDashStyle? initialStrokeDashStyle = null,
        OfficeStrokeLineCap? initialStrokeLineCap = null,
        OfficeStrokeLineJoin? initialStrokeLineJoin = null,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands,
        IReadOnlyDictionary<string, PdfFontResource>? fonts = null,
        IReadOnlyDictionary<string, Func<byte[], double>>? fontWidthProviders = null,
        Func<PdfPageType3TextInvocation, bool>? type3TextVisitor = null,
        RenderedType3TextTracker? renderedType3PaintOrders = null,
        Action<int>? type3GlyphBudgetConsumer = null,
        Action? unsupportedTextVisitor = null,
        Action? unsupportedGraphicsEffectVisitor = null,
        Action? unsupportedPatternVisitor = null,
        Action? unsupportedColorVisitor = null,
        Action<string>? visibleFontVisitor = null,
        Action<string>? patternInvocationVisitor = null,
        Action<string>? authoredPatternInvocationVisitor = null,
        Action<PdfPageGraphicsStateResource, Matrix2D, OfficeColor, OfficeColor, bool, bool>? graphicsStateVisitor = null,
        bool allowSupportedGraphicsEffects = false,
        IReadOnlyDictionary<string, PdfPageColorSpace>? patternBaseColorSpaces = null,
        PdfPagePatternSelection? initialFillPattern = null,
        PdfPageColorSpace? initialFillPatternBaseColorSpace = null,
        PdfPagePatternSelection? initialStrokePattern = null,
        PdfPageColorSpace? initialStrokePatternBaseColorSpace = null,
        IReadOnlyDictionary<string, PdfPageTilingPatternResource>? tilingPatterns = null,
        IReadOnlyDictionary<string, PdfPageShadingPatternResource>? shadingPatterns = null,
        Func<PdfPageType3GlyphInvocation, PdfType3PaintChannels>? type3PaintChannelResolver = null,
        Func<string, PdfPageXObjectPaintState, PdfType3PaintChannels>? xObjectPaintChannelResolver = null,
        Func<PdfPageSoftMaskResource, Matrix2D, OfficeColor, OfficeColor, bool, bool, bool>? softMaskVisibilityResolver = null,
        Action<string>? visibleShadingVisitor = null,
        Action? invalidPatternSelectionVisitor = null,
        Action<PdfType3PaintChannels, PdfPagePatternSelection?, PdfPagePatternSelection?, bool>? ordinaryTextPaintVisitor = null,
        Action<PdfPagePatternSelection>? patternSelectionVisitor = null,
        double? pageWidth = null,
        PdfContentOrderKey? contentOrderPrefix = null) {
        if (string.IsNullOrEmpty(content)) {
            return Array.Empty<PdfPageXObjectInvocation>();
        }

        var parser = new Parser(content, baseTransform, pageHeight, pageWidth, graphicsStates, colorSpaces, optionalContentVisibility, initialFillColor, initialFillColorSpace, initialFillOpacity, paintOrderBase, paintOrderScale, paintOrderOffset, initialClipPath, initialStrokeColor, initialStrokeColorSpace, initialStrokeOpacity, initialStrokeWidth, initialStrokeDashStyle, initialStrokeLineCap, initialStrokeLineJoin, maxOperations, maxNestingDepth, maxOperands, fonts, fontWidthProviders, type3TextVisitor, renderedType3PaintOrders, type3GlyphBudgetConsumer, unsupportedTextVisitor, unsupportedGraphicsEffectVisitor, unsupportedPatternVisitor, unsupportedColorVisitor, visibleFontVisitor, patternInvocationVisitor, authoredPatternInvocationVisitor, graphicsStateVisitor, allowSupportedGraphicsEffects, patternBaseColorSpaces, initialFillPattern, initialFillPatternBaseColorSpace, initialStrokePattern, initialStrokePatternBaseColorSpace, tilingPatterns, shadingPatterns, type3PaintChannelResolver, xObjectPaintChannelResolver, softMaskVisibilityResolver, visibleShadingVisitor, invalidPatternSelectionVisitor, ordinaryTextPaintVisitor, patternSelectionVisitor, contentOrderPrefix);
        return parser.Parse();
    }

    private sealed class Parser {
        private readonly string _content;
        private readonly double _pageHeight;
        private readonly double? _pageWidth;
        private readonly Matrix2D _baseTransform;
        private readonly IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? _graphicsStates;
        private readonly IReadOnlyDictionary<string, PdfPageColorSpace>? _colorSpaces;
        private readonly IReadOnlyDictionary<string, PdfPageColorSpace>? _patternBaseColorSpaces;
        private readonly IReadOnlyDictionary<string, PdfPageTilingPatternResource>? _tilingPatterns;
        private readonly IReadOnlyDictionary<string, PdfPageShadingPatternResource>? _shadingPatterns;
        private readonly PdfPageOptionalContentVisibility? _optionalContentVisibility;
        private readonly double _paintOrderBase;
        private readonly double _paintOrderScale;
        private readonly double _paintOrderOffset;
        private readonly List<PdfPageXObjectInvocation> _invocations = new List<PdfPageXObjectInvocation>();
        private readonly List<object> _args = new List<object>(8);
        private readonly Stack<GraphicsState> _stack = new Stack<GraphicsState>();
        private readonly Stack<bool> _inexactDashStack = new Stack<bool>();
        private readonly Stack<GraphicsEffectState> _graphicsEffectStack = new Stack<GraphicsEffectState>();
        private readonly Stack<PatternState> _patternStack = new Stack<PatternState>();
        private readonly Stack<TextState> _textStack = new Stack<TextState>();
        private readonly Stack<(PdfPageSoftMaskResource? SoftMask, Matrix2D? Transform, OfficeColor FillColor, OfficeColor StrokeColor, bool HasFillPattern, bool HasStrokePattern, PdfPageGraphicsStateResource? InheritedGraphicsState)> _softMaskStack =
            new Stack<(PdfPageSoftMaskResource?, Matrix2D?, OfficeColor, OfficeColor, bool, bool, PdfPageGraphicsStateResource?)>();
        private readonly Stack<bool> _hiddenContentStack = new Stack<bool>();
        private readonly List<(double X, double Y)> _path = new List<(double X, double Y)>();
        private readonly List<OfficePathCommand> _pathCommands = new List<OfficePathCommand>();
        private readonly List<PdfPageClipPath> _pendingTextClipPaths = new List<PdfPageClipPath>();
        private readonly GraphicsState _initialState;
        private readonly PatternState _initialPatternState;
        private GraphicsState _state;
        private GraphicsEffectState _graphicsEffectState;
        private PatternState _patternState;
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
        private int _inlineImageIndex;
        private PdfContentInlineImage? _currentInlineImage;
        private readonly int _maxOperations;
        private readonly int _maxNestingDepth;
        private readonly int _maxOperands;
        private readonly IReadOnlyDictionary<string, PdfFontResource>? _fonts;
        private readonly IReadOnlyDictionary<string, Func<byte[], double>>? _fontWidthProviders;
        private readonly Func<PdfPageType3TextInvocation, bool>? _type3TextVisitor;
        private readonly RenderedType3TextTracker? _renderedType3PaintOrders;
        private readonly PdfContentOrderKey? _contentOrderPrefix;
        private readonly Action<int>? _type3GlyphBudgetConsumer;
        private readonly Action? _unsupportedTextVisitor;
        private readonly Action? _unsupportedGraphicsEffectVisitor;
        private readonly Action? _unsupportedPatternVisitor;
        private readonly Action? _unsupportedColorVisitor;
        private readonly Action<string>? _visibleFontVisitor;
        private readonly Action<string>? _patternInvocationVisitor;
        private readonly Action<string>? _authoredPatternInvocationVisitor;
        private readonly Action? _invalidPatternSelectionVisitor;
        private readonly Func<PdfPageType3GlyphInvocation, PdfType3PaintChannels>? _type3PaintChannelResolver;
        private readonly Func<string, PdfPageXObjectPaintState, PdfType3PaintChannels>? _xObjectPaintChannelResolver;
        private readonly Func<PdfPageSoftMaskResource, Matrix2D, OfficeColor, OfficeColor, bool, bool, bool>? _softMaskVisibilityResolver;
        private readonly Action<string>? _visibleShadingVisitor;
        private readonly Action<PdfPageGraphicsStateResource, Matrix2D, OfficeColor, OfficeColor, bool, bool>? _graphicsStateVisitor;
        private readonly Action<PdfType3PaintChannels, PdfPagePatternSelection?, PdfPagePatternSelection?, bool>? _ordinaryTextPaintVisitor;
        private readonly Action<PdfPagePatternSelection>? _patternSelectionVisitor;
        private readonly bool _allowSupportedGraphicsEffects;
        private string _textFont = string.Empty;
        private double _currentPaintOrder;
        private int _currentOperatorIndex;
        private PdfPageSoftMaskResource? _softMask;
        private Matrix2D? _softMaskTransform;
        private OfficeColor _softMaskFillColor = OfficeColor.Black;
        private OfficeColor _softMaskStrokeColor = OfficeColor.Black;
        private bool _softMaskHasFillPattern;
        private bool _softMaskHasStrokePattern;
        private PdfPageGraphicsStateResource? _softMaskInheritedGraphicsState;
        private bool _hasInexactDash;

        public Parser(
            string content,
            Matrix2D baseTransform,
            double pageHeight,
            double? pageWidth,
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
            IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces,
            PdfPageOptionalContentVisibility? optionalContentVisibility,
            OfficeColor? initialFillColor,
            PdfPageColorSpace initialFillColorSpace,
            double? initialFillOpacity,
            double paintOrderBase,
            double paintOrderScale,
            double paintOrderOffset,
            PdfPageClipPath? initialClipPath,
            OfficeColor? initialStrokeColor,
            PdfPageColorSpace initialStrokeColorSpace,
            double? initialStrokeOpacity,
            double? initialStrokeWidth,
            OfficeStrokeDashStyle? initialStrokeDashStyle,
            OfficeStrokeLineCap? initialStrokeLineCap,
            OfficeStrokeLineJoin? initialStrokeLineJoin,
            int maxOperations,
            int maxNestingDepth,
            int maxOperands,
            IReadOnlyDictionary<string, PdfFontResource>? fonts,
            IReadOnlyDictionary<string, Func<byte[], double>>? fontWidthProviders,
            Func<PdfPageType3TextInvocation, bool>? type3TextVisitor,
            RenderedType3TextTracker? renderedType3PaintOrders,
            Action<int>? type3GlyphBudgetConsumer,
            Action? unsupportedTextVisitor,
            Action? unsupportedGraphicsEffectVisitor,
            Action? unsupportedPatternVisitor,
            Action? unsupportedColorVisitor,
            Action<string>? visibleFontVisitor,
            Action<string>? patternInvocationVisitor,
            Action<string>? authoredPatternInvocationVisitor,
            Action<PdfPageGraphicsStateResource, Matrix2D, OfficeColor, OfficeColor, bool, bool>? graphicsStateVisitor,
            bool allowSupportedGraphicsEffects,
            IReadOnlyDictionary<string, PdfPageColorSpace>? patternBaseColorSpaces,
            PdfPagePatternSelection? initialFillPattern,
            PdfPageColorSpace? initialFillPatternBaseColorSpace,
            PdfPagePatternSelection? initialStrokePattern,
            PdfPageColorSpace? initialStrokePatternBaseColorSpace,
            IReadOnlyDictionary<string, PdfPageTilingPatternResource>? tilingPatterns,
            IReadOnlyDictionary<string, PdfPageShadingPatternResource>? shadingPatterns,
            Func<PdfPageType3GlyphInvocation, PdfType3PaintChannels>? type3PaintChannelResolver,
            Func<string, PdfPageXObjectPaintState, PdfType3PaintChannels>? xObjectPaintChannelResolver,
            Func<PdfPageSoftMaskResource, Matrix2D, OfficeColor, OfficeColor, bool, bool, bool>? softMaskVisibilityResolver,
            Action<string>? visibleShadingVisitor,
            Action? invalidPatternSelectionVisitor,
            Action<PdfType3PaintChannels, PdfPagePatternSelection?, PdfPagePatternSelection?, bool>? ordinaryTextPaintVisitor,
            Action<PdfPagePatternSelection>? patternSelectionVisitor,
            PdfContentOrderKey? contentOrderPrefix) {
            _content = content;
            _baseTransform = baseTransform;
            _graphicsStates = graphicsStates;
            _colorSpaces = colorSpaces;
            _patternBaseColorSpaces = patternBaseColorSpaces;
            _tilingPatterns = tilingPatterns;
            _shadingPatterns = shadingPatterns;
            _optionalContentVisibility = optionalContentVisibility;
            _initialState = GraphicsState.Create(baseTransform, initialFillColor, initialFillColorSpace, initialFillOpacity, initialClipPath, initialStrokeColor, initialStrokeColorSpace, initialStrokeOpacity, initialStrokeWidth, initialStrokeDashStyle, initialStrokeLineCap, initialStrokeLineJoin);
            _initialPatternState = new PatternState(
                initialFillPattern,
                initialFillPatternBaseColorSpace ?? initialFillPattern?.BaseColorSpace,
                initialStrokePattern,
                initialStrokePatternBaseColorSpace ?? initialStrokePattern?.BaseColorSpace);
            _state = _initialState;
            _patternState = _initialPatternState;
            _pageHeight = pageHeight;
            _pageWidth = pageWidth;
            _paintOrderBase = paintOrderBase;
            _paintOrderScale = paintOrderScale;
            _paintOrderOffset = paintOrderOffset;
            _maxOperations = maxOperations;
            _maxNestingDepth = maxNestingDepth;
            _maxOperands = maxOperands;
            _fonts = fonts;
            _fontWidthProviders = fontWidthProviders;
            _type3TextVisitor = type3TextVisitor;
            _renderedType3PaintOrders = renderedType3PaintOrders;
            _contentOrderPrefix = contentOrderPrefix;
            _type3GlyphBudgetConsumer = type3GlyphBudgetConsumer;
            _unsupportedTextVisitor = unsupportedTextVisitor;
            _unsupportedGraphicsEffectVisitor = unsupportedGraphicsEffectVisitor;
            _unsupportedPatternVisitor = unsupportedPatternVisitor;
            _unsupportedColorVisitor = unsupportedColorVisitor;
            _visibleFontVisitor = visibleFontVisitor;
            _patternInvocationVisitor = patternInvocationVisitor;
            _authoredPatternInvocationVisitor = authoredPatternInvocationVisitor;
            _invalidPatternSelectionVisitor = invalidPatternSelectionVisitor;
            _ordinaryTextPaintVisitor = ordinaryTextPaintVisitor;
            _patternSelectionVisitor = patternSelectionVisitor;
            _type3PaintChannelResolver = type3PaintChannelResolver;
            _xObjectPaintChannelResolver = xObjectPaintChannelResolver;
            _softMaskVisibilityResolver = softMaskVisibilityResolver;
            _visibleShadingVisitor = visibleShadingVisitor;
            _graphicsStateVisitor = graphicsStateVisitor;
            _allowSupportedGraphicsEffects = allowSupportedGraphicsEffects;
        }

        public IReadOnlyList<PdfPageXObjectInvocation> Parse() {
            PdfContentStreamInterpreter.Interpret(
                _content,
                _maxOperations,
                operation => {
                    _args.Clear();
                    _args.AddRange(operation.Operands);
                    _currentInlineImage = operation.InlineImage;
                    _currentOperatorIndex = operation.OperatorOffset;
                    ApplyOperator(
                        operation.Name,
                        GetPaintOrder(operation.OperatorOffset),
                        operation.HasInvalidOperands);
                    _currentInlineImage = null;
                },
                ResolveInlineImageComponentCount,
                _maxNestingDepth,
                _maxOperands);

            ApplyPendingTextClippingPath();

            return _invocations.Count == 0 ? Array.Empty<PdfPageXObjectInvocation>() : _invocations.AsReadOnly();
        }

        private double GetPaintOrder(int operatorIndex) => _paintOrderBase + ((operatorIndex + _paintOrderOffset) * _paintOrderScale);

        private TextState CaptureTextState() =>
            new TextState(_inText, _textFont, _textSize, _textLeading, _textCharSpacing, _textWordSpacing, _textHScale, _textRise, _textRenderingMode, _textMatrix, _lineMatrix);

        private void RestoreTextState(TextState state) {
            _inText = state.InText;
            _textFont = state.Font;
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

            bool usesType3GlyphProgram = IsActiveType3Font();
            bool isVisible = !HasHiddenContent() && !IsCurrentPaintSuppressedBySoftMask();
            PdfType3PaintChannels ordinaryTextPaintChannels = ResolveTextPaintChannels(_textRenderingMode);
            List<PdfPageType3GlyphInvocation>? glyphs = CreateType3GlyphBatch(usesType3GlyphProgram, isVisible, bytes.Length);
            (double X, double Y) advance = ProcessShownText(bytes, glyphs);
            bool ordinaryTextAffectsOutput =
                isVisible &&
                !usesType3GlyphProgram &&
                (TextAffectsOutput(_textRenderingMode) || AddsTextToClippingPath(_textRenderingMode)) &&
                IsOrdinaryTextFrameVisible(advance.X);
            if (ordinaryTextAffectsOutput) _visibleFontVisitor?.Invoke(_textFont);
            PdfType3PaintChannels type3PaintChannels = isVisible && usesType3GlyphProgram
                ? ResolveType3PaintChannels(glyphs)
                : PdfType3PaintChannels.None;
            if (_hasInexactDash && (type3PaintChannels & PdfType3PaintChannels.Stroke) != 0) {
                _unsupportedGraphicsEffectVisitor?.Invoke();
            }
            if (type3PaintChannels != PdfType3PaintChannels.None) _visibleFontVisitor?.Invoke(_textFont);
            if (isVisible && usesType3GlyphProgram) {
                AdvertiseInheritedType3Patterns(type3PaintChannels);
            }
            if (type3PaintChannels != PdfType3PaintChannels.None ||
                (ordinaryTextAffectsOutput && ordinaryTextPaintChannels != PdfType3PaintChannels.None)) {
                PublishActiveGraphicsEffectUse();
            }
            PublishType3GlyphBatch(glyphs);
            if (ordinaryTextAffectsOutput && !usesType3GlyphProgram) _unsupportedTextVisitor?.Invoke();
            bool ordinaryTextAddsClip = ordinaryTextAffectsOutput && AddsTextToClippingPath(_textRenderingMode);
            if (ordinaryTextAffectsOutput &&
                (ordinaryTextPaintChannels != PdfType3PaintChannels.None || ordinaryTextAddsClip)) {
                _ordinaryTextPaintVisitor?.Invoke(ordinaryTextPaintChannels, _patternState.Fill, _patternState.Stroke, ordinaryTextAddsClip);
            }
            if (isVisible && !usesType3GlyphProgram) ApplyTextClippingPath(advance.X);
            _textMatrix = Matrix2D.Multiply(_textMatrix, Matrix2D.Translation(advance.X, advance.Y));
        }

        private bool IsActiveType3Font() =>
            _fonts != null &&
            _fonts.TryGetValue(_textFont, out PdfFontResource? font) &&
            font.Type3 != null;

        private List<PdfPageType3GlyphInvocation>? CreateType3GlyphBatch(bool usesType3GlyphProgram, bool isVisible, int glyphCount) {
            if (!usesType3GlyphProgram || !isVisible ||
                (_type3TextVisitor == null && _type3PaintChannelResolver == null) ||
                glyphCount <= 0) return null;
            _type3GlyphBudgetConsumer?.Invoke(glyphCount);
            return new List<PdfPageType3GlyphInvocation>(glyphCount);
        }

        private void PublishType3GlyphBatch(List<PdfPageType3GlyphInvocation>? glyphs) {
            if (_type3TextVisitor != null && glyphs != null && glyphs.Count > 0 &&
                _type3TextVisitor!(new PdfPageType3TextInvocation(glyphs, _currentPaintOrder, _currentOperatorIndex))) {
                _renderedType3PaintOrders?.Add(_currentPaintOrder, _contentOrderPrefix?.Append(_currentOperatorIndex));
            }
        }

        private (double X, double Y) ProcessShownText(byte[] bytes, List<PdfPageType3GlyphInvocation>? glyphs) {
            PdfFontResource? font = null;
            bool isType3 = _fonts != null &&
                _fonts.TryGetValue(_textFont, out font) &&
                font.Type3 != null;
            if (!isType3) {
                double width1000;
                if (_fontWidthProviders != null &&
                    _fontWidthProviders.TryGetValue(_textFont, out Func<byte[], double>? provider)) {
                    width1000 = provider(bytes);
                } else {
                    width1000 = string.Equals(font?.FontSubtype, "Type0", StringComparison.Ordinal)
                        ? (bytes.Length / 2D) * 1000D
                        : bytes.Length * 500D;
                }

                bool isComposite = string.Equals(font?.FontSubtype, "Type0", StringComparison.Ordinal);
                int glyphCount = isComposite ? bytes.Length / 2 : bytes.Length;
                int spaceCount = 0;
                if (!isComposite && Math.Abs(_textWordSpacing) > 0D) {
                    for (int i = 0; i < bytes.Length; i++) {
                        if (bytes[i] == 32) spaceCount++;
                    }
                }

                double spacing = (glyphCount * _textCharSpacing) + (spaceCount * _textWordSpacing);
                return ((((width1000 / 1000D) * _textSize) + spacing) * _textHScale, 0D);
            }

            double advanceX = 0D;
            double advanceY = 0D;
            for (int i = 0; i < bytes.Length; i++) {
                byte code = bytes[i];
                if (glyphs != null) {
                    Matrix2D glyphTextMatrix = Matrix2D.Multiply(_textMatrix, Matrix2D.Translation(advanceX, advanceY));
                    Matrix2D textState = Matrix2D.Multiply(
                        _state.Transform,
                        Matrix2D.Multiply(
                            glyphTextMatrix,
                            new Matrix2D(_textSize * _textHScale, 0D, 0D, _textSize, 0D, _textRise)));
                    glyphs.Add(new PdfPageType3GlyphInvocation(
                        font!, code, textState, _state.ClipPath,
                        _state.FillColor, _state.FillColorSpace, _patternState.Fill, _patternState.FillBaseColorSpace, _state.FillOpacity,
                        _state.StrokeColor, _state.StrokeColorSpace, _patternState.Stroke, _patternState.StrokeBaseColorSpace, _state.StrokeOpacity,
                        _patternState.Fill?.Name, _patternState.Stroke?.Name,
                        _state.StrokeWidth, _state.StrokeDashStyle, _state.StrokeLineCap, _state.StrokeLineJoin));
                }

                double spacing = _textCharSpacing + (code == 32 ? _textWordSpacing : 0D);
                PdfType3FontResource type3 = font!.Type3!;
                (double X, double Y) displacement = type3.GetGlyphDisplacement(code);
                advanceX += (displacement.X * _textSize + spacing) * _textHScale;
                advanceY += displacement.Y * _textSize;
            }

            return (advanceX, advanceY);
        }

        private void ShowTextArray(object arrayObject) {
            if (arrayObject is not List<object> items) {
                ShowText(arrayObject);
                return;
            }

            bool usesType3GlyphProgram = IsActiveType3Font();
            bool isVisible = !HasHiddenContent() && !IsCurrentPaintSuppressedBySoftMask();
            long glyphCount = 0;
            for (int i = 0; i < items.Count; i++) {
                if (items[i] is byte[] bytes) glyphCount += bytes.Length;
            }
            if (glyphCount > int.MaxValue) {
                _type3GlyphBudgetConsumer?.Invoke(int.MaxValue);
                throw PdfReadLimitException.Create(PdfReadLimitKind.Type3GlyphInvocations, int.MaxValue, glyphCount);
            }
            List<PdfPageType3GlyphInvocation>? glyphs = CreateType3GlyphBatch(usesType3GlyphProgram, isVisible, (int)glyphCount);
            PdfType3PaintChannels ordinaryTextPaintChannels = ResolveTextPaintChannels(_textRenderingMode);
            bool ordinaryTextAffectsOutput = false;
            for (int i = 0; i < items.Count; i++) {
                if (items[i] is byte[] bytes) {
                    (double X, double Y) advance = ProcessShownText(bytes, glyphs);
                    ordinaryTextAffectsOutput |=
                        isVisible &&
                        !usesType3GlyphProgram &&
                        (TextAffectsOutput(_textRenderingMode) || AddsTextToClippingPath(_textRenderingMode)) &&
                        IsOrdinaryTextFrameVisible(advance.X);
                    if (isVisible && !usesType3GlyphProgram) ApplyTextClippingPath(advance.X);
                    _textMatrix = Matrix2D.Multiply(_textMatrix, Matrix2D.Translation(advance.X, advance.Y));
                } else if (items[i] is double kerning) {
                    double delta = -kerning / 1000D * _textSize * _textHScale;
                    _textMatrix = Matrix2D.Multiply(_textMatrix, Matrix2D.Translation(delta, 0D));
                }
            }
            if (ordinaryTextAffectsOutput && glyphCount > 0) _visibleFontVisitor?.Invoke(_textFont);
            PdfType3PaintChannels type3PaintChannels =
                isVisible && usesType3GlyphProgram && glyphCount > 0
                    ? ResolveType3PaintChannels(glyphs)
                    : PdfType3PaintChannels.None;
            if (_hasInexactDash && (type3PaintChannels & PdfType3PaintChannels.Stroke) != 0) {
                _unsupportedGraphicsEffectVisitor?.Invoke();
            }
            if (type3PaintChannels != PdfType3PaintChannels.None) _visibleFontVisitor?.Invoke(_textFont);
            if (isVisible && usesType3GlyphProgram && glyphCount > 0) {
                AdvertiseInheritedType3Patterns(type3PaintChannels);
            }
            if (type3PaintChannels != PdfType3PaintChannels.None ||
                (ordinaryTextAffectsOutput && ordinaryTextPaintChannels != PdfType3PaintChannels.None)) {
                PublishActiveGraphicsEffectUse();
            }
            PublishType3GlyphBatch(glyphs);
            if (ordinaryTextAffectsOutput && glyphCount > 0 && !usesType3GlyphProgram) _unsupportedTextVisitor?.Invoke();
            bool ordinaryTextAddsClip = ordinaryTextAffectsOutput && AddsTextToClippingPath(_textRenderingMode);
            if (ordinaryTextAffectsOutput &&
                (ordinaryTextPaintChannels != PdfType3PaintChannels.None || ordinaryTextAddsClip) &&
                glyphCount > 0) {
                _ordinaryTextPaintVisitor?.Invoke(ordinaryTextPaintChannels, _patternState.Fill, _patternState.Stroke, ordinaryTextAddsClip);
            }
        }

        private void AdvertiseInheritedType3Patterns(PdfType3PaintChannels channels) {
            PublishDeferredPatternUse(
                (channels & PdfType3PaintChannels.Fill) != 0,
                (channels & PdfType3PaintChannels.Stroke) != 0);
            if ((channels & PdfType3PaintChannels.Fill) != 0 && _patternState.Fill.HasValue) {
                _patternInvocationVisitor?.Invoke(_patternState.Fill.Value.Name);
            }
            if ((channels & PdfType3PaintChannels.Stroke) != 0 && _patternState.Stroke.HasValue) {
                _patternInvocationVisitor?.Invoke(_patternState.Stroke.Value.Name);
            }
        }

        private PdfType3PaintChannels ResolveType3PaintChannels(List<PdfPageType3GlyphInvocation>? glyphs) {
            if (_type3PaintChannelResolver == null || glyphs == null || glyphs.Count == 0) {
                return PdfType3PaintChannels.Both;
            }
            PdfType3PaintChannels channels = PdfType3PaintChannels.None;
            for (int i = 0; i < glyphs.Count; i++) {
                channels |= _type3PaintChannelResolver(glyphs[i]);
                if (channels == PdfType3PaintChannels.Both) break;
            }
            return channels;
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
                _pendingTextClipPaths.Add(textClipPath);
            }
        }

        private bool IsOrdinaryTextFrameVisible(double advance) {
            if (_textSize <= 0D) return false;
            // A zero advance is not proof that a glyph paints no pixels. Fonts may
            // legitimately expose non-empty outlines with a zero text advance.
            if (Math.Abs(advance) <= 0.000001D) {
                return _pageWidth.HasValue
                    ? PdfReadPage.HasPositiveVisibleClipArea(_state.ClipPath, _pageWidth.Value, _pageHeight)
                    : !_state.ClipPath.HasValue ||
                      (_state.ClipPath.Value.Width > 0.000001D && _state.ClipPath.Value.Height > 0.000001D);
            }
            double left = advance < 0D ? advance : 0D;
            double width = Math.Abs(advance);
            double descent = Math.Max(0.001D, _textSize * 0.25D);
            double height = Math.Max(0.001D, _textSize + descent);
            Matrix2D textToPage = Matrix2D.Multiply(_state.Transform, _textMatrix);
            var textFrameBuilder = new PdfPageClipPathBuilder(_pageHeight);
            textFrameBuilder.AddRectanglePath(textToPage, left, _textRise - descent, width, height);
            if (!textFrameBuilder.TryCreateClipPath(OfficeFillRule.NonZero, out PdfPageClipPath textFrame)) {
                return true;
            }
            PdfPageClipPath visibleFrame = PdfPageClipPath.ResolveActiveClip(_state.ClipPath, textFrame);
            if (_pageWidth.HasValue) {
                visibleFrame = PdfPageClipPath.ResolveActiveClip(
                    visibleFrame,
                    PdfPageClipPath.Rectangle(0D, 0D, _pageWidth.Value, _pageHeight));
            }
            return visibleFrame.Width > 0.000001D && visibleFrame.Height > 0.000001D;
        }

        private void ApplyPendingTextClippingPath() {
            if (PdfPageClipPath.TryCombineTextClippingPaths(_pendingTextClipPaths, out PdfPageClipPath textClipPath)) {
                _state = _state.WithClipPath(PdfPageClipPath.ResolveActiveClip(_state.ClipPath, textClipPath));
            }
            _pendingTextClipPaths.Clear();
        }

        private void ApplyOperator(string op, double paintOrder, bool hasInvalidOperands) {
            _currentPaintOrder = paintOrder;
            switch (op) {
                case "q":
                    _stack.Push(_state);
                    _inexactDashStack.Push(_hasInexactDash);
                    _graphicsEffectStack.Push(_graphicsEffectState);
                    _patternStack.Push(_patternState);
                    _textStack.Push(CaptureTextState());
                    _softMaskStack.Push((
                        _softMask,
                        _softMaskTransform,
                        _softMaskFillColor,
                        _softMaskStrokeColor,
                        _softMaskHasFillPattern,
                        _softMaskHasStrokePattern,
                        _softMaskInheritedGraphicsState));
                    break;
                case "Q":
                    _state = _stack.Count > 0 ? _stack.Pop() : _initialState;
                    _hasInexactDash = _inexactDashStack.Count > 0 && _inexactDashStack.Pop();
                    _graphicsEffectState = _graphicsEffectStack.Count > 0 ? _graphicsEffectStack.Pop() : default;
                    _patternState = _patternStack.Count > 0 ? _patternStack.Pop() : _initialPatternState;
                    RestoreTextState(_textStack.Count > 0 ? _textStack.Pop() : TextState.Default);
                    (PdfPageSoftMaskResource? SoftMask, Matrix2D? Transform, OfficeColor FillColor, OfficeColor StrokeColor, bool HasFillPattern, bool HasStrokePattern, PdfPageGraphicsStateResource? InheritedGraphicsState) restoredSoftMask = _softMaskStack.Count > 0
                        ? _softMaskStack.Pop()
                        : (null, null, OfficeColor.Black, OfficeColor.Black, false, false, null);
                    _softMask = restoredSoftMask.SoftMask;
                    _softMaskTransform = restoredSoftMask.Transform;
                    _softMaskFillColor = restoredSoftMask.FillColor;
                    _softMaskStrokeColor = restoredSoftMask.StrokeColor;
                    _softMaskHasFillPattern = restoredSoftMask.HasFillPattern;
                    _softMaskHasStrokePattern = restoredSoftMask.HasStrokePattern;
                    _softMaskInheritedGraphicsState = restoredSoftMask.InheritedGraphicsState;
                    break;
                case "cm":
                    if (HasExactFiniteNumbers(6)) {
                        Matrix2D matrix = new Matrix2D(
                            NumberAt(0),
                            NumberAt(1),
                            NumberAt(2),
                            NumberAt(3),
                            NumberAt(4),
                            NumberAt(5));
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
                    if (_args.Count >= 2 &&
                        TryGetNumberArray(_args[_args.Count - 2], out double[] dashArray) &&
                        _args[_args.Count - 1] is double dashPhase) {
                        if (PdfPageContentVisualParser.TryReadExactDashStyle(dashArray, dashPhase, out OfficeStrokeDashStyle exactStyle)) {
                            _state = _state.WithStrokeDashStyle(exactStyle);
                            _hasInexactDash = false;
                        } else {
                            _state = _state.WithStrokeDashStyle(ReadDashStyle(dashArray));
                            _hasInexactDash = true;
                        }
                    } else {
                        _hasInexactDash = true;
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
                    if (op == "s" || op == "b" || op == "b*") {
                        ClosePath();
                    }
                    if (!HasHiddenContent() && !IsCurrentPaintSuppressedBySoftMask() && _pathCommands.Count > 0) {
                        PdfType3PaintChannels channels = ResolveVisiblePathPaintChannels(
                            OperatorFillsPath(op),
                            OperatorStrokesPath(op),
                            op == "f*" || op == "B*" || op == "b*" ? OfficeFillRule.EvenOdd : OfficeFillRule.NonZero);
                        PublishDeferredPatternUse(
                            (channels & PdfType3PaintChannels.Fill) != 0,
                            (channels & PdfType3PaintChannels.Stroke) != 0);
                        if (channels != PdfType3PaintChannels.None) PublishActiveGraphicsEffectUse();
                        if (_hasInexactDash && (channels & PdfType3PaintChannels.Stroke) != 0) {
                            _unsupportedGraphicsEffectVisitor?.Invoke();
                        }
                    }
                    if (!HasHiddenContent() &&
                        !IsCurrentPaintSuppressedBySoftMask() &&
                        OperatorStrokesPath(op) &&
                        _state.StrokeWidth > 0D &&
                        !double.IsPositiveInfinity(_state.StrokeWidth) &&
                        !IsConformalStrokeTransform(_state.Transform)) {
                        _unsupportedGraphicsEffectVisitor?.Invoke();
                    }
                    ClearPath();
                    break;
                case "gs":
                    if (_args.Count >= 1 && _args[_args.Count - 1] is string graphicsStateName) {
                        ApplyGraphicsStateResource(graphicsStateName);
                    }

                    break;
                case "sh":
                    if (_args.Count != 1 || _args[0] is not string) {
                        if (!HasHiddenContent()) _unsupportedGraphicsEffectVisitor?.Invoke();
                    } else if (!HasHiddenContent() &&
                        !IsCurrentPaintSuppressedBySoftMask() &&
                        (_state.FillOpacity ?? 1D) > 0D &&
                        IsCurrentClipPotentiallyVisible() &&
                        _args[0] is string shadingName &&
                        !string.IsNullOrEmpty(shadingName)) {
                        _visibleShadingVisitor?.Invoke(shadingName);
                        PublishActiveGraphicsEffectUse();
                    }

                    break;
                case "cs":
                    if (_args.Count == 1 &&
                        _args[0] is string fillColorSpaceName &&
                        TryReadColorSpace(fillColorSpaceName, out PdfPageColorSpace fillColorSpace)) {
                        _state = _state.WithFillColorSpace(fillColorSpace);
                        _patternState = _patternState.WithFill(
                            null,
                            ReadPatternBaseColorSpace(fillColorSpaceName, fillColorSpace));
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "CS":
                    if (_args.Count == 1 &&
                        _args[0] is string strokeColorSpaceName &&
                        TryReadColorSpace(strokeColorSpaceName, out PdfPageColorSpace strokeColorSpace)) {
                        _state = _state.WithStrokeColorSpace(strokeColorSpace);
                        _patternState = _patternState.WithStroke(
                            null,
                            ReadPatternBaseColorSpace(strokeColorSpaceName, strokeColorSpace));
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "sc":
                case "scn":
                    if (_state.FillColorSpace == PdfPageColorSpaceKind.Pattern &&
                        (op == "sc" || _args.Count == 0 || _args[_args.Count - 1] is not string)) {
                        string invalidFillPatternName = _args.Count > 0 && _args[_args.Count - 1] is string candidateName
                            ? candidateName
                            : string.Empty;
                        _patternState = _patternState.WithFill(
                            new PdfPagePatternSelection(
                                invalidFillPatternName,
                                null,
                                _patternState.FillBaseColorSpace,
                                null,
                                null,
                                _state.Transform,
                                componentCount: -1),
                            _patternState.FillBaseColorSpace,
                            deferredVisibleUse: true);
                    }
                    if (op == "scn" && _args.Count > 0 && _args[_args.Count - 1] is string fillPatternName) {
                        if (_state.FillColorSpace == PdfPageColorSpaceKind.Pattern) {
                            OfficeColor? tint = _patternState.FillBaseColorSpace.HasValue &&
                                TryReadColor(_patternState.FillBaseColorSpace.Value, out OfficeColor fillPatternTint)
                                    ? fillPatternTint
                                    : (OfficeColor?)null;
                            PdfPageTilingPatternResource? tilingPattern = ResolveTilingPattern(fillPatternName);
                            PdfPageShadingPatternResource? shadingPattern = ResolveShadingPattern(fillPatternName);
                            _patternState = _patternState.WithFill(
                                new PdfPagePatternSelection(
                                    fillPatternName,
                                    tint,
                                    _patternState.FillBaseColorSpace,
                                    tilingPattern,
                                    shadingPattern,
                                    _state.Transform,
                                    CountPatternComponents()),
                                _patternState.FillBaseColorSpace,
                                deferredVisibleUse: true);
                        } else if (!HasHiddenContent()) {
                            _unsupportedColorVisitor?.Invoke();
                        }
                    }
                    if (_state.FillColorSpace != PdfPageColorSpaceKind.Pattern &&
                        TryReadColor(_state.FillColorSpace, out OfficeColor fillColor)) {
                        _state = _state.WithFillColor(fillColor);
                        _patternState = _patternState.WithFill(null, null);
                    } else if (!HasHiddenContent() && !(_args.Count > 0 && _args[_args.Count - 1] is string)) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "SC":
                case "SCN":
                    if (_state.StrokeColorSpace == PdfPageColorSpaceKind.Pattern &&
                        (op == "SC" || _args.Count == 0 || _args[_args.Count - 1] is not string)) {
                        string invalidStrokePatternName = _args.Count > 0 && _args[_args.Count - 1] is string candidateName
                            ? candidateName
                            : string.Empty;
                        _patternState = _patternState.WithStroke(
                            new PdfPagePatternSelection(
                                invalidStrokePatternName,
                                null,
                                _patternState.StrokeBaseColorSpace,
                                null,
                                null,
                                _state.Transform,
                                componentCount: -1),
                            _patternState.StrokeBaseColorSpace,
                            deferredVisibleUse: true);
                    }
                    if (op == "SCN" && _args.Count > 0 && _args[_args.Count - 1] is string strokePatternName) {
                        if (_state.StrokeColorSpace == PdfPageColorSpaceKind.Pattern) {
                            OfficeColor? tint = _patternState.StrokeBaseColorSpace.HasValue &&
                                TryReadColor(_patternState.StrokeBaseColorSpace.Value, out OfficeColor strokePatternTint)
                                    ? strokePatternTint
                                    : (OfficeColor?)null;
                            PdfPageTilingPatternResource? tilingPattern = ResolveTilingPattern(strokePatternName);
                            PdfPageShadingPatternResource? shadingPattern = ResolveShadingPattern(strokePatternName);
                            _patternState = _patternState.WithStroke(
                                new PdfPagePatternSelection(
                                    strokePatternName,
                                    tint,
                                    _patternState.StrokeBaseColorSpace,
                                    tilingPattern,
                                    shadingPattern,
                                    _state.Transform,
                                    CountPatternComponents()),
                                _patternState.StrokeBaseColorSpace,
                                deferredVisibleUse: true);
                        } else if (!HasHiddenContent()) {
                            _unsupportedColorVisitor?.Invoke();
                        }
                    }
                    if (_state.StrokeColorSpace != PdfPageColorSpaceKind.Pattern &&
                        TryReadColor(_state.StrokeColorSpace, out OfficeColor strokeColor)) {
                        _state = _state.WithStrokeColor(strokeColor);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent() && !(_args.Count > 0 && _args[_args.Count - 1] is string)) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "rg":
                    if (HasTrailingFiniteNumbers(3)) {
                        _state = _state.WithFillColor(ReadRgb(_args.Count - 3), PdfPageColorSpaceKind.DeviceRgb);
                        _patternState = _patternState.WithFill(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "RG":
                    if (HasTrailingFiniteNumbers(3)) {
                        _state = _state.WithStrokeColor(ReadRgb(_args.Count - 3), PdfPageColorSpaceKind.DeviceRgb);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "g":
                    if (HasTrailingFiniteNumbers(1)) {
                        _state = _state.WithFillColor(ReadGray(_args.Count - 1), PdfPageColorSpaceKind.DeviceGray);
                        _patternState = _patternState.WithFill(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "G":
                    if (HasTrailingFiniteNumbers(1)) {
                        _state = _state.WithStrokeColor(ReadGray(_args.Count - 1), PdfPageColorSpaceKind.DeviceGray);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "k":
                    if (HasTrailingFiniteNumbers(4)) {
                        _state = _state.WithFillColor(ReadCmyk(_args.Count - 4), PdfPageColorSpaceKind.DeviceCmyk);
                        _patternState = _patternState.WithFill(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "K":
                    if (HasTrailingFiniteNumbers(4)) {
                        _state = _state.WithStrokeColor(ReadCmyk(_args.Count - 4), PdfPageColorSpaceKind.DeviceCmyk);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "BT":
                    ApplyPendingTextClippingPath();
                    _inText = true;
                    _textMatrix = Matrix2D.Identity;
                    _lineMatrix = Matrix2D.Identity;
                    break;
                case "ET":
                    ApplyPendingTextClippingPath();
                    _inText = false;
                    break;
                case "Tf":
                    if (_args.Count >= 2) {
                        _textFont = _args[_args.Count - 2] as string ?? string.Empty;
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
                        bool needsPaintAnalysis =
                            _patternState.FillDeferredVisibleUse ||
                            _patternState.StrokeDeferredVisibleUse ||
                            _graphicsEffectState.HasEffect ||
                            _hasInexactDash;
                        PdfType3PaintChannels channels = needsPaintAnalysis
                            ? IsCurrentPaintSuppressedBySoftMask()
                                ? PdfType3PaintChannels.None
                                : _xObjectPaintChannelResolver?.Invoke(
                                    name,
                                    new PdfPageXObjectPaintState(
                                        _state.Transform,
                                        _state.ClipPath,
                                        _state.FillOpacity,
                                        _state.StrokeOpacity,
                                        _state.StrokeWidth,
                                        _state.StrokeDashStyle,
                                        _state.StrokeLineCap,
                                        _state.StrokeLineJoin)) ?? PdfType3PaintChannels.Both
                            : PdfType3PaintChannels.None;
                        if (_patternState.FillDeferredVisibleUse || _patternState.StrokeDeferredVisibleUse) {
                            PublishDeferredPatternUse(
                                (channels & PdfType3PaintChannels.Fill) != 0,
                                (channels & PdfType3PaintChannels.Stroke) != 0);
                        }
                        if (_graphicsEffectState.HasEffect && channels != PdfType3PaintChannels.None) {
                            PublishActiveGraphicsEffectUse();
                        }
                        if (_hasInexactDash && (channels & PdfType3PaintChannels.Stroke) != 0) {
                            _unsupportedGraphicsEffectVisitor?.Invoke();
                        }
                        _invocations.Add(new PdfPageXObjectInvocation(name, _state.Transform, _state.ClipPath, _state.FillColor, _state.FillColorSpace, _patternState.Fill, _patternState.FillBaseColorSpace, _state.FillOpacity, _state.StrokeColor, _state.StrokeColorSpace, _patternState.Stroke, _patternState.StrokeBaseColorSpace, _state.StrokeOpacity, _state.StrokeWidth, _state.StrokeDashStyle, _state.StrokeLineCap, _state.StrokeLineJoin, paintOrder, _currentOperatorIndex));
                    }

                    break;
                case "BI":
                    if (_currentInlineImage is not null && !HasHiddenContent()) {
                        PdfType3PaintChannels channels = IsCurrentPaintSuppressedBySoftMask()
                            ? PdfType3PaintChannels.None
                            : ResolveVisibleInlineImagePaintChannels();
                        bool consumesFillColor =
                            channels != PdfType3PaintChannels.None &&
                            _currentInlineImage.Dictionary.Items.TryGetValue("ImageMask", out PdfObject? imageMaskObject) &&
                            imageMaskObject is PdfBoolean { Value: true };
                        PublishDeferredPatternUse(
                            fill: consumesFillColor,
                            stroke: false);
                        if (channels != PdfType3PaintChannels.None) PublishActiveGraphicsEffectUse();
                        var stream = new PdfStream(_currentInlineImage.Dictionary, _currentInlineImage.Data);
                        var inlineImage = new PdfPageInlineImage(
                            "__inline" + (++_inlineImageIndex).ToString(CultureInfo.InvariantCulture),
                            stream);
                        _invocations.Add(new PdfPageXObjectInvocation(inlineImage, _state.Transform, _state.ClipPath, _state.FillColor, _state.FillColorSpace, _patternState.Fill, _patternState.FillBaseColorSpace, _state.FillOpacity, _state.StrokeColor, _state.StrokeColorSpace, _patternState.Stroke, _patternState.StrokeBaseColorSpace, _state.StrokeOpacity, _state.StrokeWidth, _state.StrokeDashStyle, _state.StrokeLineCap, _state.StrokeLineJoin, paintOrder, _currentOperatorIndex));
                    }

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

        private bool HasTrailingFiniteNumbers(int count) {
            if (_args.Count < count) return false;
            for (int index = _args.Count - count; index < _args.Count; index++) {
                if (_args[index] is not double value || double.IsNaN(value) || double.IsInfinity(value)) return false;
            }
            return true;
        }

        private bool HasExactFiniteNumbers(int count) =>
            _args.Count == count && HasTrailingFiniteNumbers(count);

        private void ApplyGraphicsStateResource(string name) {
            if (_graphicsStates == null || !_graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource resource)) {
                if (!HasHiddenContent()) _unsupportedGraphicsEffectVisitor?.Invoke();
                return;
            }

            _state = _state.WithGraphicsStateResource(resource);
            if (resource.HasSoftMask) {
                _softMask = resource.SoftMask;
                _softMaskTransform = resource.SoftMask == null ? null : _state.Transform;
                _softMaskFillColor = _state.FillColor;
                _softMaskStrokeColor = _state.StrokeColor;
                _softMaskHasFillPattern = _patternState.Fill.HasValue;
                _softMaskHasStrokePattern = _patternState.Stroke.HasValue;
                _softMaskInheritedGraphicsState = CaptureCurrentGraphicsState();
            }
            _graphicsEffectState = _graphicsEffectState.Apply(resource, _state.Transform);
        }

        private PdfPageGraphicsStateResource CaptureCurrentGraphicsState() => new PdfPageGraphicsStateResource(
            _state.FillOpacity,
            _state.StrokeOpacity,
            _state.StrokeWidth,
            _state.StrokeDashStyle,
            _state.StrokeLineCap,
            _state.StrokeLineJoin);

        private void PublishActiveGraphicsEffectUse() {
            if (!_graphicsEffectState.HasEffect) return;
            PdfPageGraphicsStateResource effect = _graphicsEffectState.ToResource();
            PdfPageGraphicsStateResource? softMaskState = effect.HasSoftMask ? _softMaskInheritedGraphicsState : null;
            var resource = new PdfPageGraphicsStateResource(
                softMaskState.HasValue ? softMaskState.Value.FillOpacity : _state.FillOpacity,
                softMaskState.HasValue ? softMaskState.Value.StrokeOpacity : _state.StrokeOpacity,
                softMaskState.HasValue ? softMaskState.Value.StrokeWidth : _state.StrokeWidth,
                softMaskState.HasValue ? softMaskState.Value.StrokeDashStyle : _state.StrokeDashStyle,
                softMaskState.HasValue ? softMaskState.Value.StrokeLineCap : _state.StrokeLineCap,
                softMaskState.HasValue ? softMaskState.Value.StrokeLineJoin : _state.StrokeLineJoin,
                effect.BlendMode,
                effect.HasSoftMask,
                effect.SoftMask,
                effect.HasUnsupportedSoftMask,
                effect.HasUnsupportedBlendMode,
                effect.HasUnsupportedEntries);
            _graphicsStateVisitor?.Invoke(
                resource,
                _graphicsEffectState.SoftMaskTransform,
                _softMaskFillColor,
                _softMaskStrokeColor,
                _softMaskHasFillPattern,
                _softMaskHasStrokePattern);
            if ((!_allowSupportedGraphicsEffects &&
                 ((resource.BlendMode.HasValue && resource.BlendMode.Value != OfficeBlendMode.Normal) ||
                  (resource.HasSoftMask && resource.SoftMask != null))) ||
                resource.HasUnsupportedSoftMask ||
                resource.HasUnsupportedBlendMode ||
                resource.HasUnsupportedEntries) {
                _unsupportedGraphicsEffectVisitor?.Invoke();
            }
        }

        private bool HasHiddenContent() {
            foreach (bool hidden in _hiddenContentStack) {
                if (hidden) {
                    return true;
                }
            }

            return false;
        }

        private bool IsCurrentClipPotentiallyVisible() {
            if (_pageWidth.HasValue) {
                return PdfReadPage.HasPositiveVisibleClipArea(
                    _state.ClipPath,
                    _pageWidth.Value,
                    _pageHeight);
            }

            return !_state.ClipPath.HasValue ||
                   (_state.ClipPath.Value.Width > 0D && _state.ClipPath.Value.Height > 0D);
        }

        private bool IsCurrentPaintSuppressedBySoftMask() =>
            _softMask != null &&
            _softMaskVisibilityResolver != null &&
            !_softMaskVisibilityResolver(
                _softMask,
                _softMaskTransform ?? _state.Transform,
                _softMaskFillColor,
                _softMaskStrokeColor,
                _softMaskHasFillPattern,
                _softMaskHasStrokePattern);

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

        private int ResolveInlineImageComponentCount(string colorSpaceName) {
            if (_colorSpaces != null &&
                _colorSpaces.TryGetValue(colorSpaceName, out PdfPageColorSpace colorSpace)) {
                return GetComponentCount(colorSpace);
            }

            return 1;
        }

        private static int GetComponentCount(PdfPageColorSpace colorSpace) {
            return colorSpace.ComponentCount;
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

        private bool TryReadColor(PdfPageColorSpace colorSpace, out OfficeColor color) {
            color = OfficeColor.Black;
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

        private PdfPageColorSpace? ReadPatternBaseColorSpace(string name, PdfPageColorSpace colorSpace) {
            if (colorSpace != PdfPageColorSpaceKind.Pattern || _patternBaseColorSpaces == null) {
                return null;
            }

            return _patternBaseColorSpaces.TryGetValue(name, out PdfPageColorSpace baseColorSpace)
                ? baseColorSpace
                : (PdfPageColorSpace?)null;
        }

        private PdfPageTilingPatternResource? ResolveTilingPattern(string name) {
            return _tilingPatterns != null && _tilingPatterns.TryGetValue(name, out PdfPageTilingPatternResource? pattern)
                ? pattern
                : null;
        }

        private static bool IsValidPatternSelection(
            PdfPageTilingPatternResource? tilingPattern,
            PdfPageShadingPatternResource? shadingPattern,
            PdfPageColorSpace? baseColorSpace,
            OfficeColor? tint,
            int componentCount) {
            if (tilingPattern != null) {
                return tilingPattern.Uncolored
                    ? baseColorSpace.HasValue && tint.HasValue && componentCount == baseColorSpace.Value.ComponentCount
                    : !baseColorSpace.HasValue && !tint.HasValue && componentCount == 0;
            }
            return shadingPattern?.IsSupported == true && !baseColorSpace.HasValue && !tint.HasValue && componentCount == 0;
        }

        private int CountPatternComponents() {
            int count = 0;
            for (int index = 0; index < _args.Count - 1; index++) {
                if (_args[index] is not double) return -1;
                count++;
            }
            return count;
        }

        private PdfPageShadingPatternResource? ResolveShadingPattern(
            string name,
            Matrix2D? paintTransform = null) {
            if (_shadingPatterns == null ||
                !_shadingPatterns.TryGetValue(name, out PdfPageShadingPatternResource pattern)) return null;
            return PdfPageContentVisualParser.IsSupportedShadingTransform(
                    pattern,
                    paintTransform ?? _state.Transform)
                ? pattern
                : PdfPageShadingPatternResource.Unsupported;
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

        private static int ReadTextRenderingMode(double value) {
            int mode = (int)Math.Round(value);
            return mode < 0 || mode > 7 ? 0 : mode;
        }

        private static bool AddsTextToClippingPath(int renderingMode) =>
            renderingMode >= 4 && renderingMode <= 7;

        private bool TextAffectsOutput(int renderingMode) {
            if (AddsTextToClippingPath(renderingMode)) return true;
            bool fills = renderingMode == 0 || renderingMode == 2;
            bool strokes = renderingMode == 1 || renderingMode == 2;
            return (fills && (_state.FillOpacity ?? 1D) > 0D) ||
                   (strokes && (_state.StrokeOpacity ?? 1D) > 0D);
        }

        private static PdfType3PaintChannels ResolveTextPaintChannels(int renderingMode) => renderingMode switch {
            0 or 4 => PdfType3PaintChannels.Fill,
            1 or 5 => PdfType3PaintChannels.Stroke,
            2 or 6 => PdfType3PaintChannels.Both,
            _ => PdfType3PaintChannels.None
        };

        private static bool OperatorStrokesPath(string op) =>
            op == "S" || op == "s" || op == "B" || op == "B*" || op == "b" || op == "b*";

        private static bool OperatorFillsPath(string op) =>
            op == "f" || op == "F" || op == "f*" || op == "B" || op == "B*" || op == "b" || op == "b*";

        private PdfType3PaintChannels ResolveVisiblePathPaintChannels(
            bool fill,
            bool stroke,
            OfficeFillRule fillRule) {
            if (!fill && !stroke) return PdfType3PaintChannels.None;
            double visibilityStrokeWidth = GetProjectedStrokeWidth();
            if (!PdfPageVisualPrimitive.TryCreatePath(
                    _pathCommands,
                    fill ? OfficeColor.Black : (OfficeColor?)null,
                    null,
                    null,
                    stroke ? OfficeColor.Black : (OfficeColor?)null,
                    null,
                    null,
                    visibilityStrokeWidth,
                    _state.StrokeDashStyle ?? OfficeStrokeDashStyle.Solid,
                    _state.StrokeLineCap,
                    _state.StrokeLineJoin,
                    _state.FillOpacity,
                    _state.StrokeOpacity,
                    fillRule,
                    _state.ClipPath,
                    0D,
                    null,
                    null,
                    retainPathCommands: true,
                    out PdfPageVisualPrimitive primitive)) {
                return stroke
                    ? ResolveVisibleDegenerateStrokePaintChannels(visibilityStrokeWidth)
                    : PdfType3PaintChannels.None;
            }

            return PdfReadPage.ResolveVisibleType3PrimitivePaintChannels(
                primitive,
                _pageWidth,
                _pageHeight);
        }

        private PdfType3PaintChannels ResolveVisibleDegenerateStrokePaintChannels(double strokeWidth) {
            PdfType3PaintChannels channels = PdfType3PaintChannels.None;
            OfficePoint current = default;
            OfficePoint subpathStart = default;
            bool hasCurrent = false;
            bool hasSubpathStart = false;
            for (int i = 0; i < _pathCommands.Count; i++) {
                OfficePathCommand command = _pathCommands[i];
                if (command.Kind == OfficePathCommandKind.MoveTo) {
                    current = command.Point;
                    subpathStart = command.Point;
                    hasCurrent = true;
                    hasSubpathStart = true;
                    continue;
                }
                OfficePoint next = command.Kind == OfficePathCommandKind.Close && hasSubpathStart
                    ? subpathStart
                    : command.Point;
                if (hasCurrent && (!NearlyEqual(current.X, next.X) || !NearlyEqual(current.Y, next.Y))) {
                    PdfPageVisualPrimitive segment = PdfPageVisualPrimitive.Line(
                        current.X,
                        current.Y,
                        next.X,
                        next.Y,
                        OfficeColor.Black,
                        null,
                        null,
                        strokeWidth,
                        _state.StrokeDashStyle ?? OfficeStrokeDashStyle.Solid,
                        _state.StrokeLineCap,
                        _state.StrokeLineJoin,
                        _state.StrokeOpacity,
                        _state.ClipPath);
                    channels |= PdfReadPage.ResolveVisibleType3PrimitivePaintChannels(segment, _pageWidth, _pageHeight);
                    if (channels == PdfType3PaintChannels.Stroke) return channels;
                }
                current = next;
                hasCurrent = true;
            }
            return channels;
        }

        private double GetProjectedStrokeWidth() {
            if (double.IsPositiveInfinity(_state.StrokeWidth)) return 0.25D;
            double squaredScale = ((_state.Transform.A * _state.Transform.A) +
                                   (_state.Transform.B * _state.Transform.B) +
                                   (_state.Transform.C * _state.Transform.C) +
                                   (_state.Transform.D * _state.Transform.D)) / 2D;
            return squaredScale > 0D && !double.IsNaN(squaredScale) && !double.IsInfinity(squaredScale)
                ? _state.StrokeWidth * Math.Sqrt(squaredScale)
                : _state.StrokeWidth;
        }

        private PdfType3PaintChannels ResolveVisibleInlineImagePaintChannels() {
            (double X, double Y) p0 = _state.Transform.Transform(0D, 0D);
            (double X, double Y) p1 = _state.Transform.Transform(1D, 0D);
            (double X, double Y) p2 = _state.Transform.Transform(1D, 1D);
            (double X, double Y) p3 = _state.Transform.Transform(0D, 1D);
            var commands = new[] {
                OfficePathCommand.MoveTo(p0.X, ToTop(p0.Y)),
                OfficePathCommand.LineTo(p1.X, ToTop(p1.Y)),
                OfficePathCommand.LineTo(p2.X, ToTop(p2.Y)),
                OfficePathCommand.LineTo(p3.X, ToTop(p3.Y)),
                OfficePathCommand.Close()
            };
            if (!PdfPageVisualPrimitive.TryCreatePath(
                    commands,
                    OfficeColor.Black,
                    null,
                    null,
                    null,
                    null,
                    null,
                    0D,
                    OfficeStrokeDashStyle.Solid,
                    null,
                    null,
                    _state.FillOpacity,
                    null,
                    OfficeFillRule.NonZero,
                    _state.ClipPath,
                    0D,
                    null,
                    null,
                    retainPathCommands: true,
                    out PdfPageVisualPrimitive primitive)) {
                return PdfType3PaintChannels.None;
            }

            return PdfReadPage.ResolveVisibleType3PrimitivePaintChannels(
                primitive,
                _pageWidth,
                _pageHeight);
        }

        private void PublishDeferredPatternUse(bool fill, bool stroke) {
            if (fill && _patternState.FillDeferredVisibleUse && _patternState.Fill.HasValue) {
                ValidateDeferredPatternSelection(_patternState.Fill.Value);
                PublishPatternUse(_patternState.Fill.Value.Name);
                _patternState = _patternState.WithFillDeferredVisibleUse(false);
            }
            if (stroke && _patternState.StrokeDeferredVisibleUse && _patternState.Stroke.HasValue) {
                ValidateDeferredPatternSelection(_patternState.Stroke.Value);
                PublishPatternUse(_patternState.Stroke.Value.Name);
                _patternState = _patternState.WithStrokeDeferredVisibleUse(false);
            }
        }

        private void ValidateDeferredPatternSelection(PdfPagePatternSelection selection) {
            _patternSelectionVisitor?.Invoke(selection);
            if (!IsValidPatternSelection(
                    ResolveTilingPattern(selection.Name),
                    ResolveShadingPattern(selection.Name, selection.PaintTransform),
                    selection.BaseColorSpace,
                    selection.Tint,
                    selection.ComponentCount)) {
                _invalidPatternSelectionVisitor?.Invoke();
            }
        }

        private void PublishPatternUse(string name) {
            _authoredPatternInvocationVisitor?.Invoke(name);
            _unsupportedPatternVisitor?.Invoke();
            _patternInvocationVisitor?.Invoke(name);
        }

        private static bool IsConformalStrokeTransform(Matrix2D transform) {
            double firstLengthSquared = (transform.A * transform.A) + (transform.B * transform.B);
            double secondLengthSquared = (transform.C * transform.C) + (transform.D * transform.D);
            if (firstLengthSquared <= 0D || secondLengthSquared <= 0D ||
                double.IsNaN(firstLengthSquared) || double.IsNaN(secondLengthSquared) ||
                double.IsInfinity(firstLengthSquared) || double.IsInfinity(secondLengthSquared)) return false;

            double scale = Math.Max(firstLengthSquared, secondLengthSquared);
            double dot = (transform.A * transform.C) + (transform.B * transform.D);
            return Math.Abs(firstLengthSquared - secondLengthSquared) <= scale * 0.000000001D &&
                   Math.Abs(dot) <= Math.Sqrt(firstLengthSquared * secondLengthSquared) * 0.000000001D;
        }

        private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
    }

    private readonly struct TextState {
        public TextState(bool inText, string font, double size, double leading, double charSpacing, double wordSpacing, double hScale, double textRise, int textRenderingMode, Matrix2D textMatrix, Matrix2D lineMatrix) {
            InText = inText;
            Font = font;
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

        public static TextState Default { get; } = new TextState(false, string.Empty, 12D, 14.4D, 0D, 0D, 1D, 0D, 0, Matrix2D.Identity, Matrix2D.Identity);

        public bool InText { get; }

        public string Font { get; }

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

    private readonly struct PatternState {
        internal PatternState(
            PdfPagePatternSelection? fill,
            PdfPageColorSpace? fillBaseColorSpace,
            PdfPagePatternSelection? stroke,
            PdfPageColorSpace? strokeBaseColorSpace,
            bool fillDeferredVisibleUse = false,
            bool strokeDeferredVisibleUse = false) {
            Fill = fill;
            FillBaseColorSpace = fillBaseColorSpace;
            Stroke = stroke;
            StrokeBaseColorSpace = strokeBaseColorSpace;
            FillDeferredVisibleUse = fillDeferredVisibleUse;
            StrokeDeferredVisibleUse = strokeDeferredVisibleUse;
        }

        internal PdfPagePatternSelection? Fill { get; }
        internal PdfPageColorSpace? FillBaseColorSpace { get; }
        internal PdfPagePatternSelection? Stroke { get; }
        internal PdfPageColorSpace? StrokeBaseColorSpace { get; }
        internal bool FillDeferredVisibleUse { get; }
        internal bool StrokeDeferredVisibleUse { get; }

        internal PatternState WithFill(PdfPagePatternSelection? fill, PdfPageColorSpace? baseColorSpace, bool deferredVisibleUse = false) =>
            new PatternState(fill, baseColorSpace, Stroke, StrokeBaseColorSpace, deferredVisibleUse, StrokeDeferredVisibleUse);

        internal PatternState WithStroke(PdfPagePatternSelection? stroke, PdfPageColorSpace? baseColorSpace, bool deferredVisibleUse = false) =>
            new PatternState(Fill, FillBaseColorSpace, stroke, baseColorSpace, FillDeferredVisibleUse, deferredVisibleUse);

        internal PatternState WithFillDeferredVisibleUse(bool value) =>
            new PatternState(Fill, FillBaseColorSpace, Stroke, StrokeBaseColorSpace, value, StrokeDeferredVisibleUse);

        internal PatternState WithStrokeDeferredVisibleUse(bool value) =>
            new PatternState(Fill, FillBaseColorSpace, Stroke, StrokeBaseColorSpace, FillDeferredVisibleUse, value);
    }

    private readonly struct GraphicsState {
        private GraphicsState(
            Matrix2D transform,
            PdfPageClipPath? clipPath,
            OfficeColor fillColor,
            PdfPageColorSpace fillColorSpace,
            double? fillOpacity,
            OfficeColor strokeColor,
            PdfPageColorSpace strokeColorSpace,
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

        public PdfPageColorSpace FillColorSpace { get; }

        public double? FillOpacity { get; }

        public OfficeColor StrokeColor { get; }

        public PdfPageColorSpace StrokeColorSpace { get; }

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
            PdfPageColorSpace fillColorSpace,
            double? fillOpacity,
            PdfPageClipPath? clipPath,
            OfficeColor? strokeColor,
            PdfPageColorSpace strokeColorSpace,
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

        public GraphicsState WithFillColor(OfficeColor color, PdfPageColorSpace colorSpace) => new GraphicsState(Transform, ClipPath, color, colorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithFillColorSpace(PdfPageColorSpace colorSpace) => new GraphicsState(Transform, ClipPath, FillColor, colorSpace, FillOpacity, StrokeColor, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeColor(OfficeColor color) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, color, StrokeColorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeColor(OfficeColor color, PdfPageColorSpace colorSpace) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, color, colorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

        public GraphicsState WithStrokeColorSpace(PdfPageColorSpace colorSpace) => new GraphicsState(Transform, ClipPath, FillColor, FillColorSpace, FillOpacity, StrokeColor, colorSpace, StrokeOpacity, StrokeWidth, StrokeDashStyle, StrokeLineCap, StrokeLineJoin);

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

    private readonly struct GraphicsEffectState {
        private GraphicsEffectState(
            OfficeBlendMode? blendMode,
            bool hasUnsupportedBlendMode,
            bool hasSoftMask,
            PdfPageSoftMaskResource? softMask,
            bool hasUnsupportedSoftMask,
            bool hasUnsupportedEntries,
            Matrix2D softMaskTransform) {
            BlendMode = blendMode;
            HasUnsupportedBlendMode = hasUnsupportedBlendMode;
            HasSoftMask = hasSoftMask;
            SoftMask = softMask;
            HasUnsupportedSoftMask = hasUnsupportedSoftMask;
            HasUnsupportedEntries = hasUnsupportedEntries;
            SoftMaskTransform = softMaskTransform;
        }

        private OfficeBlendMode? BlendMode { get; }
        private bool HasUnsupportedBlendMode { get; }
        private bool HasSoftMask { get; }
        private PdfPageSoftMaskResource? SoftMask { get; }
        private bool HasUnsupportedSoftMask { get; }
        private bool HasUnsupportedEntries { get; }
        internal Matrix2D SoftMaskTransform { get; }
        internal bool HasEffect =>
            (BlendMode.HasValue && BlendMode.Value != OfficeBlendMode.Normal) ||
            HasUnsupportedBlendMode ||
            HasSoftMask ||
            HasUnsupportedSoftMask ||
            HasUnsupportedEntries;

        internal GraphicsEffectState Apply(PdfPageGraphicsStateResource resource, Matrix2D transform) {
            OfficeBlendMode? blendMode = BlendMode;
            bool hasUnsupportedBlendMode = HasUnsupportedBlendMode;
            if (resource.BlendMode.HasValue || resource.HasUnsupportedBlendMode) {
                blendMode = resource.BlendMode;
                hasUnsupportedBlendMode = resource.HasUnsupportedBlendMode;
            }

            bool hasSoftMask = HasSoftMask;
            PdfPageSoftMaskResource? softMask = SoftMask;
            bool hasUnsupportedSoftMask = HasUnsupportedSoftMask;
            bool hasUnsupportedEntries = HasUnsupportedEntries || resource.HasUnsupportedEntries;
            Matrix2D softMaskTransform = SoftMaskTransform;
            if (resource.HasSoftMask || resource.HasUnsupportedSoftMask) {
                hasSoftMask = resource.HasSoftMask && resource.SoftMask != null;
                softMask = resource.SoftMask;
                hasUnsupportedSoftMask = resource.HasUnsupportedSoftMask;
                softMaskTransform = transform;
            }

            return new GraphicsEffectState(
                blendMode,
                hasUnsupportedBlendMode,
                hasSoftMask,
                softMask,
                hasUnsupportedSoftMask,
                hasUnsupportedEntries,
                softMaskTransform);
        }

        internal PdfPageGraphicsStateResource ToResource() => new PdfPageGraphicsStateResource(
            null,
            null,
            null,
            null,
            null,
            null,
            BlendMode,
            HasSoftMask,
            SoftMask,
            HasUnsupportedSoftMask,
            HasUnsupportedBlendMode,
            HasUnsupportedEntries);
    }
}

internal readonly struct PdfPageXObjectPaintState {
    internal PdfPageXObjectPaintState(
        Matrix2D transform,
        PdfPageClipPath? clipPath,
        double? fillOpacity,
        double? strokeOpacity,
        double strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin) {
        Transform = transform;
        ClipPath = clipPath;
        FillOpacity = fillOpacity;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
    }

    internal Matrix2D Transform { get; }
    internal PdfPageClipPath? ClipPath { get; }
    internal double? FillOpacity { get; }
    internal double? StrokeOpacity { get; }
    internal double StrokeWidth { get; }
    internal OfficeStrokeDashStyle? StrokeDashStyle { get; }
    internal OfficeStrokeLineCap? StrokeLineCap { get; }
    internal OfficeStrokeLineJoin? StrokeLineJoin { get; }

    internal PdfPageXObjectPaintState WithTransformAndClip(Matrix2D transform, PdfPageClipPath? clipPath) =>
        new PdfPageXObjectPaintState(
            transform,
            clipPath,
            FillOpacity,
            StrokeOpacity,
            StrokeWidth,
            StrokeDashStyle,
            StrokeLineCap,
            StrokeLineJoin);
}

internal readonly struct PdfPageXObjectInvocation {
    public PdfPageXObjectInvocation(
        string name,
        Matrix2D transform,
        PdfPageClipPath? clipPath,
        OfficeColor fillColor,
        PdfPageColorSpace fillColorSpace,
        PdfPagePatternSelection? fillPattern,
        PdfPageColorSpace? fillPatternBaseColorSpace,
        double? fillOpacity,
        OfficeColor strokeColor,
        PdfPageColorSpace strokeColorSpace,
        PdfPagePatternSelection? strokePattern,
        PdfPageColorSpace? strokePatternBaseColorSpace,
        double? strokeOpacity,
        double strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin,
        double paintOrder = 0D,
        int sourceOperatorIndex = 0) {
        Name = name;
        InlineImage = null;
        Transform = transform;
        ClipPath = clipPath;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillPattern = fillPattern;
        FillPatternBaseColorSpace = fillPatternBaseColorSpace;
        FillOpacity = fillOpacity;
        StrokeColor = strokeColor;
        StrokeColorSpace = strokeColorSpace;
        StrokePattern = strokePattern;
        StrokePatternBaseColorSpace = strokePatternBaseColorSpace;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        PaintOrder = paintOrder;
        SourceOperatorIndex = sourceOperatorIndex;
    }

    public PdfPageXObjectInvocation(
        PdfPageInlineImage inlineImage,
        Matrix2D transform,
        PdfPageClipPath? clipPath,
        OfficeColor fillColor,
        PdfPageColorSpace fillColorSpace,
        PdfPagePatternSelection? fillPattern,
        PdfPageColorSpace? fillPatternBaseColorSpace,
        double? fillOpacity,
        OfficeColor strokeColor,
        PdfPageColorSpace strokeColorSpace,
        PdfPagePatternSelection? strokePattern,
        PdfPageColorSpace? strokePatternBaseColorSpace,
        double? strokeOpacity,
        double strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin,
        double paintOrder = 0D,
        int sourceOperatorIndex = 0) {
        Name = inlineImage.ResourceName;
        InlineImage = inlineImage;
        Transform = transform;
        ClipPath = clipPath;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillPattern = fillPattern;
        FillPatternBaseColorSpace = fillPatternBaseColorSpace;
        FillOpacity = fillOpacity;
        StrokeColor = strokeColor;
        StrokeColorSpace = strokeColorSpace;
        StrokePattern = strokePattern;
        StrokePatternBaseColorSpace = strokePatternBaseColorSpace;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        PaintOrder = paintOrder;
        SourceOperatorIndex = sourceOperatorIndex;
    }

    public string Name { get; }

    public PdfPageInlineImage? InlineImage { get; }

    public Matrix2D Transform { get; }

    public PdfPageClipPath? ClipPath { get; }

    public OfficeColor FillColor { get; }

    public PdfPageColorSpace FillColorSpace { get; }

    internal PdfPagePatternSelection? FillPattern { get; }

    internal PdfPageColorSpace? FillPatternBaseColorSpace { get; }

    public double? FillOpacity { get; }

    public OfficeColor StrokeColor { get; }

    public PdfPageColorSpace StrokeColorSpace { get; }

    internal PdfPagePatternSelection? StrokePattern { get; }

    internal PdfPageColorSpace? StrokePatternBaseColorSpace { get; }

    public double? StrokeOpacity { get; }

    public double StrokeWidth { get; }

    public OfficeStrokeDashStyle? StrokeDashStyle { get; }

    public OfficeStrokeLineCap? StrokeLineCap { get; }

    public OfficeStrokeLineJoin? StrokeLineJoin { get; }

    internal PdfPageXObjectPaintState PaintState => new PdfPageXObjectPaintState(
        Transform,
        ClipPath,
        FillOpacity,
        StrokeOpacity,
        StrokeWidth,
        StrokeDashStyle,
        StrokeLineCap,
        StrokeLineJoin);

    public double PaintOrder { get; }

    internal int SourceOperatorIndex { get; }
}
