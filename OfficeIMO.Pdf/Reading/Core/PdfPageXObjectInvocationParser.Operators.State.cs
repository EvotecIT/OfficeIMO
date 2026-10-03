using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageXObjectInvocationParser {
    private sealed partial class Parser {
        private void ApplyStateOperator(string op) {
            switch (op) {
                case "q":
                    _stack.Push(_state);
                    _inexactDashStack.Push(_hasInexactDash);
                    _graphicsEffectStack.Push(_graphicsEffectState);
                    _patternStack.Push(_patternState);
                    _renderingIntentStack.Push((_hasAuthoredRenderingIntent, _renderingIntent, _fillColorSelection, _strokeColorSelection));
                    _textStack.Push(CaptureTextState());
                    _colorSpaceNameStack.Push((_fillColorSpaceResourceName, _strokeColorSpaceResourceName));
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
                    if (_renderingIntentStack.Count > 0) {
                        (bool Authored, OfficeIccRenderingIntent Intent, PdfPaintColorSelection? Fill, PdfPaintColorSelection? Stroke) restoredColor = _renderingIntentStack.Pop();
                        _hasAuthoredRenderingIntent = restoredColor.Authored;
                        _renderingIntent = restoredColor.Intent;
                        _fillColorSelection = restoredColor.Fill;
                        _strokeColorSelection = restoredColor.Stroke;
                    } else {
                        _hasAuthoredRenderingIntent = _initialHasAuthoredRenderingIntent;
                        _renderingIntent = _initialRenderingIntent;
                        _fillColorSelection = _initialFillColorSelection;
                        _strokeColorSelection = _initialStrokeColorSelection;
                    }
                    RestoreTextState(_textStack.Count > 0 ? _textStack.Pop() : TextState.Default);
                    if (_colorSpaceNameStack.Count > 0) {
                        (string? Fill, string? Stroke) restoredNames = _colorSpaceNameStack.Pop();
                        _fillColorSpaceResourceName = restoredNames.Fill;
                        _strokeColorSpaceResourceName = restoredNames.Stroke;
                    } else {
                        _fillColorSpaceResourceName = null;
                        _strokeColorSpaceResourceName = null;
                    }
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
                        var dashPattern = new PdfStrokeDashPattern(dashArray, dashPhase);
                        if (PdfPageContentVisualParser.TryReadExactDashStyle(dashArray, dashPhase, out OfficeStrokeDashStyle exactStyle)) {
                            _state = _state.WithStrokeDash(exactStyle, dashPattern);
                            _hasInexactDash = false;
                        } else {
                            _state = _state.WithStrokeDash(ReadDashStyle(dashArray), dashPattern);
                            _hasInexactDash = true;
                        }
                    } else {
                        _hasInexactDash = true;
                    }

                    break;
                case "gs":
                    if (_args.Count >= 1 && _args[_args.Count - 1] is string graphicsStateName) {
                        ApplyGraphicsStateResource(graphicsStateName);
                    }

                    break;
                case "ri":
                    if (_args.Count == 1 && _args[0] is string renderingIntentName) {
                        _hasAuthoredRenderingIntent = true;
                        ApplyRenderingIntent(PdfRenderingIntentResolver.FromName(renderingIntentName));
                    }
                    break;
            }
        }
    }
}
