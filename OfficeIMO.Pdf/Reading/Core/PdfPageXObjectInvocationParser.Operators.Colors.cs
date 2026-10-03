using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageXObjectInvocationParser {
    private sealed partial class Parser {
        private void ApplyColorsOperator(string op) {
            switch (op) {
                case "cs":
                    if (_args.Count == 1 &&
                        _args[0] is string fillColorSpaceName &&
                        TryReadColorSpace(fillColorSpaceName, out PdfPageColorSpace fillColorSpace)) {
                        _fillColorSelection = null;
                        _fillColorSpaceResourceName = IsNamedColorSpaceResource(fillColorSpaceName)
                            ? fillColorSpaceName
                            : null;
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
                        _strokeColorSelection = null;
                        _strokeColorSpaceResourceName = IsNamedColorSpaceResource(strokeColorSpaceName)
                            ? strokeColorSpaceName
                            : null;
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
                                componentCount: -1,
                                renderingIntent: _renderingIntent),
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
                                    CountPatternComponents(),
                                    _renderingIntent),
                                _patternState.FillBaseColorSpace,
                                deferredVisibleUse: true);
                        } else if (!HasHiddenContent()) {
                            _unsupportedColorVisitor?.Invoke();
                        }
                    }
                    if (_state.FillColorSpace != PdfPageColorSpaceKind.Pattern &&
                        TryApplyFillColor(_state.FillColorSpace, out OfficeColor fillColor)) {
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
                                componentCount: -1,
                                renderingIntent: _renderingIntent),
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
                                    CountPatternComponents(),
                                    _renderingIntent),
                                _patternState.StrokeBaseColorSpace,
                                deferredVisibleUse: true);
                        } else if (!HasHiddenContent()) {
                            _unsupportedColorVisitor?.Invoke();
                        }
                    }
                    if (_state.StrokeColorSpace != PdfPageColorSpaceKind.Pattern &&
                        TryApplyStrokeColor(_state.StrokeColorSpace, out OfficeColor strokeColor)) {
                        _state = _state.WithStrokeColor(strokeColor);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent() && !(_args.Count > 0 && _args[_args.Count - 1] is string)) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "rg":
                    if (HasTrailingFiniteNumbers(3) && TryApplyFillColor(PdfPageColorSpaceKind.DeviceRgb, out OfficeColor rgbFill)) {
                        _fillColorSpaceResourceName = null;
                        _state = _state.WithFillColor(rgbFill, PdfPageColorSpaceKind.DeviceRgb);
                        _patternState = _patternState.WithFill(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "RG":
                    if (HasTrailingFiniteNumbers(3) && TryApplyStrokeColor(PdfPageColorSpaceKind.DeviceRgb, out OfficeColor rgbStroke)) {
                        _strokeColorSpaceResourceName = null;
                        _state = _state.WithStrokeColor(rgbStroke, PdfPageColorSpaceKind.DeviceRgb);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "g":
                    if (HasTrailingFiniteNumbers(1) && TryApplyFillColor(PdfPageColorSpaceKind.DeviceGray, out OfficeColor grayFill)) {
                        _fillColorSpaceResourceName = null;
                        _state = _state.WithFillColor(grayFill, PdfPageColorSpaceKind.DeviceGray);
                        _patternState = _patternState.WithFill(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "G":
                    if (HasTrailingFiniteNumbers(1) && TryApplyStrokeColor(PdfPageColorSpaceKind.DeviceGray, out OfficeColor grayStroke)) {
                        _strokeColorSpaceResourceName = null;
                        _state = _state.WithStrokeColor(grayStroke, PdfPageColorSpaceKind.DeviceGray);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "k":
                    if (HasTrailingFiniteNumbers(4) && TryApplyFillColor(PdfPageColorSpaceKind.DeviceCmyk, out OfficeColor cmykFill)) {
                        _fillColorSpaceResourceName = null;
                        _state = _state.WithFillColor(cmykFill, PdfPageColorSpaceKind.DeviceCmyk);
                        _patternState = _patternState.WithFill(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
                case "K":
                    if (HasTrailingFiniteNumbers(4) && TryApplyStrokeColor(PdfPageColorSpaceKind.DeviceCmyk, out OfficeColor cmykStroke)) {
                        _strokeColorSpaceResourceName = null;
                        _state = _state.WithStrokeColor(cmykStroke, PdfPageColorSpaceKind.DeviceCmyk);
                        _patternState = _patternState.WithStroke(null, null);
                    } else if (!HasHiddenContent()) {
                        _unsupportedColorVisitor?.Invoke();
                    }

                    break;
            }
        }
    }
}
