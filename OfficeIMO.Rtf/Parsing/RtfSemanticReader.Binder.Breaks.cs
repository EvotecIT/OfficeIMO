using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private bool TryApplyBreakControl(RtfControlWord control, CharacterState state) {
            switch (control.Name) {
                case "line":
                    AppendBreak(RtfBreakKind.Line, state);
                    return true;
                case "softline":
                    AppendBreak(RtfBreakKind.SoftLine, state);
                    return true;
                case "page":
                    AppendBreak(RtfBreakKind.Page, state);
                    return true;
                case "softpage":
                    AppendBreak(RtfBreakKind.SoftPage, state);
                    return true;
                case "column":
                    AppendBreak(RtfBreakKind.Column, state);
                    return true;
                default:
                    return false;
            }
        }

        private void AppendBreak(RtfBreakKind kind, CharacterState state) {
            ApplyParagraphState(_currentParagraph, state);
            _currentParagraph.AddBreak(kind);
        }
    }
}
