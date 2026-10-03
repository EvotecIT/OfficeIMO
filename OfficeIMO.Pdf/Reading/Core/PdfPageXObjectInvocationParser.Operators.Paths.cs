using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageXObjectInvocationParser {
    private sealed partial class Parser {
        private void ApplyPathsOperator(string op) {
            switch (op) {
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
                        PublishVisibleColorSpaceUse(channels);
                        PublishDeferredPatternUse(
                            (channels & PdfType3PaintChannels.Fill) != 0,
                            (channels & PdfType3PaintChannels.Stroke) != 0);
                        if (channels != PdfType3PaintChannels.None) PublishActiveGraphicsEffectUse(channels);
                        if (_hasInexactDash && (channels & PdfType3PaintChannels.Stroke) != 0) {
                            _unsupportedGraphicsEffectVisitor?.Invoke();
                        }
                    }
                    if (!HasHiddenContent() &&
                        !IsCurrentPaintSuppressedBySoftMask() &&
                        OperatorStrokesPath(op) &&
                        _state.StrokeWidth > 0D &&
                        !double.IsPositiveInfinity(_state.StrokeWidth) &&
                        !_state.Transform.IsConformalStrokeTransform()) {
                        _unsupportedGraphicsEffectVisitor?.Invoke();
                    }
                    ClearPath();
                    break;
            }
        }
    }
}
