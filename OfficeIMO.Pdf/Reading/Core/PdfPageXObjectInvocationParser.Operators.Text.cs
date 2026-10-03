using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageXObjectInvocationParser {
    private sealed partial class Parser {
        private void ApplyTextOperator(string op) {
            switch (op) {
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
            }
        }
    }
}
