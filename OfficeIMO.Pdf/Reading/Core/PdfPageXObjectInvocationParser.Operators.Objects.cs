using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageXObjectInvocationParser {
    private sealed partial class Parser {
        private void ApplyObjectsOperator(string op, double paintOrder) {
            switch (op) {
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
                        _visibleShadingWithIntentVisitor?.Invoke(shadingName, _renderingIntent);
                        PublishActiveGraphicsEffectUse(PdfType3PaintChannels.Fill);
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
                            _hasInexactDash ||
                            _visibleColorSpaceVisitor != null;
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
                                        _state.StrokeLineJoin,
                                        _state.StrokeDashPattern)) ?? PdfType3PaintChannels.Both
                            : PdfType3PaintChannels.None;
                        PublishVisibleColorSpaceUse(channels);
                        if (_patternState.FillDeferredVisibleUse || _patternState.StrokeDeferredVisibleUse) {
                            PublishDeferredPatternUse(
                                (channels & PdfType3PaintChannels.Fill) != 0,
                                (channels & PdfType3PaintChannels.Stroke) != 0);
                        }
                        if (_graphicsEffectState.HasEffect && channels != PdfType3PaintChannels.None) {
                            PublishActiveGraphicsEffectUse(channels);
                        }
                        if (_hasInexactDash && (channels & PdfType3PaintChannels.Stroke) != 0) {
                            _unsupportedGraphicsEffectVisitor?.Invoke();
                        }
                        _invocations.Add(new PdfPageXObjectInvocation(
                            name,
                            _state.Transform,
                            _state.ClipPath,
                            _state.FillColor,
                            _state.FillColorSpace,
                            _patternState.Fill,
                            _patternState.FillBaseColorSpace,
                            _state.FillOpacity,
                            _state.StrokeColor,
                            _state.StrokeColorSpace,
                            _patternState.Stroke,
                            _patternState.StrokeBaseColorSpace,
                            _state.StrokeOpacity,
                            _state.StrokeWidth,
                            _state.StrokeDashStyle,
                            _state.StrokeLineCap,
                            _state.StrokeLineJoin,
                            paintOrder,
                            _currentOperatorIndex,
                            _state.BlendMode,
                            _state.AuthoredBlendMode,
                            _state.HasUnsupportedBlendMode,
                            _state.HasUnsupportedPaintState,
                            _state.HasUnsupportedImagePaintEffect,
                            _state.HasSoftMask,
                            _hasAuthoredRenderingIntent,
                            _renderingIntent,
                            _fillColorSelection,
                            _strokeColorSelection,
                            _state.StrokeDashPattern,
                            GetActiveMcid(),
                            HasArtifactContent(),
                            _state.ImagePaintEffectState,
                            new PdfTextStateSnapshot(_textFont, _textSize, _textLeading, _textCharSpacing,
                                _textWordSpacing, _textHScale, _textRise, _textRenderingMode)));
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
                        if (consumesFillColor) PublishVisibleColorSpaceUse(PdfType3PaintChannels.Fill);
                        if (channels != PdfType3PaintChannels.None && !consumesFillColor &&
                            TryGetInlineImageColorSpaceName(_currentInlineImage.Dictionary, out string? imageColorSpaceName) &&
                            !IsBuiltInColorSpaceName(imageColorSpaceName!)) {
                            _visibleColorSpaceVisitor?.Invoke(imageColorSpaceName!);
                        }
                        PublishDeferredPatternUse(
                            fill: consumesFillColor,
                            stroke: false);
                        if (channels != PdfType3PaintChannels.None) PublishActiveGraphicsEffectUse(channels);
                        var stream = new PdfStream(_currentInlineImage.Dictionary, _currentInlineImage.Data);
                        var inlineImage = new PdfPageInlineImage(
                            "__inline" + (++_inlineImageIndex).ToString(CultureInfo.InvariantCulture),
                            stream);
                        _invocations.Add(new PdfPageXObjectInvocation(
                            inlineImage,
                            _state.Transform,
                            _state.ClipPath,
                            _state.FillColor,
                            _state.FillColorSpace,
                            _patternState.Fill,
                            _patternState.FillBaseColorSpace,
                            _state.FillOpacity,
                            _state.StrokeColor,
                            _state.StrokeColorSpace,
                            _patternState.Stroke,
                            _patternState.StrokeBaseColorSpace,
                            _state.StrokeOpacity,
                            _state.StrokeWidth,
                            _state.StrokeDashStyle,
                            _state.StrokeLineCap,
                            _state.StrokeLineJoin,
                            paintOrder,
                            _currentOperatorIndex,
                            _state.BlendMode,
                            _state.AuthoredBlendMode,
                            _state.HasUnsupportedBlendMode,
                            _state.HasUnsupportedPaintState,
                            _state.HasUnsupportedImagePaintEffect,
                            _state.HasSoftMask,
                            _hasAuthoredRenderingIntent,
                            _renderingIntent,
                            _fillColorSelection,
                            _strokeColorSelection,
                            _state.StrokeDashPattern,
                            GetActiveMcid(),
                            HasArtifactContent(),
                            _state.ImagePaintEffectState));
                    }

                    break;
            }
        }
    }
}
